using System.Globalization;
using System.Text.RegularExpressions;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 解析器</summary>
/// <remarks>
/// 支持 CommonMark 核心语法 + GFM 扩展（表格/任务列表/删除线）。
/// 采用两遍扫描：第一遍识别块结构，第二遍解析每个叶子块内的行内内容。
/// </remarks>
internal sealed class MarkdownParser
{
    #region 字段
    private String[] _lines = [];
    private Int32 _pos;
    private Int32 _listDepth;
    private MarkdownDocument? _doc;

    /// <summary>引用链接定义（[id]: url "title"），键为引用标识</summary>
    private Dictionary<String, (String url, String title)> _refs = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>需跳过的行（引用链接定义行）</summary>
    private HashSet<Int32> _skipLines = [];

    /// <summary>处理管线（MD06），为 null 时使用默认全功能管线</summary>
    public MarkdownPipeline? Pipeline { get; set; }

    /// <summary>是否启用往返模式（捕获块原始源码与快照，供 MD06-03 往返渲染；非往返省去捕获开销）</summary>
    internal Boolean Roundtrip;

    /// <summary>往返快照序列化器（MD26：复用实例，省每块 new MarkdownWriter；状态在 finally 中复位，单线程安全）</summary>
    private MarkdownWriter? _snapshotWriter;
    #endregion

    #region 常量（静态缓存，避免热点路径分配，MD19）
    /// <summary>裸 URL 前缀（TryParseBareUrl 高频调用，静态化避免每次调用 new 数组分配）</summary>
    private static readonly String[] _urlPrefixes = ["https://", "http://", "ftp://", "www."];

    /// <summary>表格行快速路径判定字符（反引号/反斜杠，MD21b：IndexOfAny 一次扫描替代两次 IndexOf）</summary>
    private static readonly Char[] _tableEscapeChars = ['`', '\\'];
    #endregion

    #region 正则（静态缓存，避免热点路径重复编译）
    private static readonly Regex _atxHeading = new(@"^(#{1,6})(\s+|$)(.*)$", RegexOptions.Compiled);
    private static readonly Regex _trailingHashes = new(@"\s+#+\s*$", RegexOptions.Compiled);
    private static readonly Regex _orderedList = new(@"^(\d{1,9})([.)]) ", RegexOptions.Compiled);
    private static readonly Regex _taskMarker = new(@"^\[([ xX])\] (.*)$", RegexOptions.Compiled);
    private static readonly Regex _setext1 = new(@"^[=]+\s*$", RegexOptions.Compiled);
    private static readonly Regex _setext2 = new(@"^[-]+\s*$", RegexOptions.Compiled);
    private static readonly Regex _tableSep = new(@"^[|\-: ]+$", RegexOptions.Compiled);
    private static readonly Regex _footnoteDef = new(@"^\[\^([^\]]+)\]:\s*(.*)$", RegexOptions.Compiled);
    private static readonly Regex _abbrDef = new(@"^\*\[([^\]]+)\]:\s*(.*)$", RegexOptions.Compiled);
    private static readonly Regex _referenceDef = new(@"^\[([^\]]+)\]:\s*(?:<([^>]+)>|(\S+))(?:\s+(.*))?$", RegexOptions.Compiled);
    #endregion

    #region 入口
    /// <summary>解析 Markdown 文本，返回文档对象</summary>
    /// <param name="text">Markdown 文本</param>
    /// <returns>解析后的文档</returns>
    public MarkdownDocument Parse(String text)
    {
        // 统一换行符，展开 Tab（CommonMark 规范：Tab = 4 空格）
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        return ParseCore(text.Split('\n'));
    }

    /// <summary>从行列表解析（内部复用，子解析器避免 Join+Split 重复分配）</summary>
    /// <param name="lines">已按 \n 规范化的行列表</param>
    /// <param name="skipFrontMatter">是否跳过 YAML Front Matter（子解析器内容不可能有文档级 Front Matter）</param>
    /// <returns>解析后的文档</returns>
    internal MarkdownDocument ParseLines(List<String> lines, Boolean skipFrontMatter = false)
        => ParseCore(lines.ToArray(), skipFrontMatter);

    /// <summary>解析核心：从行数组构建文档</summary>
    /// <param name="lines">已按 \n 规范化的行数组</param>
    /// <param name="skipFrontMatter">是否跳过 YAML Front Matter（子解析器内容不可能有文档级 Front Matter）</param>
    /// <returns>解析后的文档</returns>
    private MarkdownDocument ParseCore(String[] lines, Boolean skipFrontMatter = false)
    {
        _lines = lines;
        _pos = 0;

        var doc = new MarkdownDocument();
        _doc = doc;

        // YAML Front Matter (MD05-02)；子解析器（列表/引用内容）跳过——块内容不可能有文档级 Front Matter，且避免 `- ---` 被误判
        if (!skipFrontMatter && (Pipeline == null || Pipeline.EnableFrontMatter))
            ParseFrontMatter(doc);

        // 预扫描引用链接定义（支持定义在后使用在前；列表项内引用定义需保留）
        PreScanReferences(doc);

        while (_pos < _lines.Length)
        {
            // 跳过引用链接定义行（已在预扫描中提取）；无定义时短路，避免每行哈希查找（MD16）
            if (_skipLines.Count > 0 && _skipLines.Contains(_pos)) { _pos++; continue; }

            var blockStart = _pos;
            var blockCount = doc.Blocks.Count;
            ParseBlock(doc.Blocks);

            // 设置新生成块的源码位置（MD06-04）并捕获原始文本（MD06-03，仅往返模式需要）
            for (var bi = blockCount; bi < doc.Blocks.Count; bi++)
            {
                var block = doc.Blocks[bi];
                block.SourceLine = blockStart;
                block.SourceColumn = 0;

                if (!Roundtrip) continue;

                // 捕获原始源码文本用于往返渲染（跳过前导空白行，含尾部空行以保留块间分隔）
                var srcStart = blockStart;
                while (srcStart < _pos && String.IsNullOrWhiteSpace(_lines[srcStart]))
                    srcStart++;
                var srcEnd = _pos;
                while (srcEnd < _lines.Length && String.IsNullOrWhiteSpace(_lines[srcEnd]))
                    srcEnd++;
                if (srcEnd > srcStart)
                {
                    // 单行快路径（MD26）：避免 new String[1] + String.Join 的数组与拼接分配，直接拼接单行
                    if (srcEnd - srcStart == 1)
                        block.SourceText = _lines[srcStart] + "\n";
                    else
                    {
                        var lineCount = srcEnd - srcStart;
                        var rawLines = new String[lineCount];
                        Array.Copy(_lines, srcStart, rawLines, 0, lineCount);
                        // 拼接时补上末尾换行，保留块间空行（含被捕获的尾部空行）
                        block.SourceText = String.Join("\n", rawLines) + "\n";
                    }
                    // 记录规范序列化快照，用于往返渲染时检测块是否被修改
                    block.SourceSnapshot = (_snapshotWriter ??= new MarkdownWriter()).SerializeSnapshot(block);
                }
            }
        }

        _doc = null;
        return doc;
    }
    #endregion

    #region 块级解析
    /// <summary>添加块到容器，并检查后续行是否有自定义属性 {.class #id} (MD05-06)</summary>
    private void AddBlock(List<MarkdownBlock> container, MarkdownBlock block)
    {
        container.Add(block);

        // 检查下一行是否是自定义属性
        if ((Pipeline == null || Pipeline.EnableCustomAttributes) && _pos < _lines.Length)
        {
            var nextLine = _lines[_pos].TrimStart();
            if (nextLine.StartsWith("{") && nextLine.EndsWith("}") && nextLine.Length > 2)
            {
                var inner = nextLine[1..^1].Trim();
                if (inner.StartsWith(".") || inner.StartsWith("#") || inner.Contains("="))
                {
                    block.Attributes = inner;
                    _pos++;
                }
            }
        }
    }

    private void ParseBlock(List<MarkdownBlock> container)
    {
        if (_pos >= _lines.Length) return;

        var line = _lines[_pos];
        var trimmed = line.TrimStart();
        var indent = GetIndent(line);

        // 空行
        if (String.IsNullOrWhiteSpace(line)) { _pos++; return; }

        // 缩进代码块（4+ 空格或制表；CommonMark 规定不能中断段落）
        if (indent >= 4)
        {
            AddBlock(container, ParseIndentedCode());
            return;
        }

        // ATX 标题  # text
        var atx = TryParseAtxHeading(trimmed);
        if (atx != null) { AddBlock(container, atx); _pos++; return; }

        // 围栏代码块 ``` or ~~~
        if (trimmed.StartsWith("```") || trimmed.StartsWith("~~~"))
        {
            AddBlock(container, ParseFencedCode(trimmed));
            return;
        }

        // HTML 块 (以 < 开头的完整块)
        if (trimmed.StartsWith("<"))
        {
            var html = TryParseHtmlBlock(trimmed);
            if (html != null) { AddBlock(container, html); return; }
        }

        // 分隔线  --- / *** / ___
        if (IsThematicBreak(trimmed)) { AddBlock(container, MarkdownBlock.CreateThematicBreak()); _pos++; return; }

        // 引用块 >
        if (trimmed.StartsWith(">"))
        {
            AddBlock(container, ParseBlockQuote());
            return;
        }

        // 无序列表
        if (IsBulletListMarker(trimmed, out var _))
        {
            AddBlock(container, ParseList(false));
            return;
        }

        // 有序列表
        if (IsOrderedListMarker(trimmed, out var startNum, out var _))
        {
            AddBlock(container, ParseList(true, startNum));
            return;
        }

        // 脚注定义 [^id]: text (MD05-03)
        if ((Pipeline == null || Pipeline.EnableFootnotes) && TryParseFootnoteDefinition(trimmed, out var fnDef))
        {
            AddBlock(container, fnDef!);
            _pos++;
            return;
        }

        // 数学公式块 $$...$$ (MD05-04)
        if ((Pipeline == null || Pipeline.EnableMath) && (trimmed == "$$" || trimmed.StartsWith("$$")))
        {
            AddBlock(container, ParseMathBlock(trimmed));
            return;
        }

        // 缩写定义 *[ABBR]: (MD05-08)
        if ((Pipeline == null || Pipeline.EnableAbbreviations) && trimmed.StartsWith("*[") && trimmed.Contains("]:"))
        {
            ParseAbbreviation(trimmed);
            _pos++;
            return;
        }

        // 定义列表 : description (MD05-07)
        if ((Pipeline == null || Pipeline.EnableDefinitionLists) && (trimmed.StartsWith(": ") || trimmed == ":"))
        {
            var dl = TryParseDefinitionList();
            if (dl != null) { AddBlock(container, dl); return; }
        }

        // Setext 标题（只在段落内检查）
        // 段落（可能包含 Setext 标题）
        AddBlock(container, ParseParagraphOrSetext(indent));
    }

    private MarkdownBlock? TryParseAtxHeading(String trimmed)
    {
        // 快速失败（MD20）：ATX 标题必以 # 开头，避免普通行每行跑正则
        if (trimmed.Length == 0 || trimmed[0] != '#') return null;
        var match = _atxHeading.Match(trimmed);
        if (!match.Success) return null;
        var hashes = match.Groups[1].Value;
        // trailing #? remove it
        var text = match.Groups[3].Value.TrimEnd();
        text = _trailingHashes.Replace(text, "").TrimEnd();
        var inlines = ParseInlineWithRefs(text);
        return MarkdownBlock.CreateHeading(hashes.Length, inlines);
    }

    private MarkdownBlock ParseFencedCode(String firstLine)
    {
        var fence = firstLine.StartsWith("~~~") ? "~~~" : "```";
        var lang = firstLine.TrimStart('~', '`').Trim();
        // extract just the language (first word)
        var spaceIdx = lang.IndexOf(' ');
        if (spaceIdx >= 0) lang = lang[..spaceIdx];

        _pos++;
        var sb = new System.Text.StringBuilder();
        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith(fence))
            {
                _pos++;
                break;
            }
            sb.Append(line).Append('\n');
            _pos++;
        }
        var code = sb.ToString();
        if (code.EndsWith("\n")) code = code[..^1];
        if (code.EndsWith("\r")) code = code[..^1];
        return MarkdownBlock.CreateCodeBlock(code, lang);
    }

    /// <summary>解析 HTML 块（CommonMark 类型 ①②③④⑤⑥⑦）</summary>
    /// <param name="trimmed">当前行（去前导空白）</param>
    /// <returns>HTML 块，非 HTML 块返回 null</returns>
    private MarkdownBlock? TryParseHtmlBlock(String trimmed)
    {
        // ① 特殊标签 <pre|script|style|textarea：直到匹配闭合标签（可跨空行）
        if (IsSpecialTagStart(trimmed))
            return ParseSpecialHtmlBlock();

        // ②③④⑤：带结束标记的类型
        String? endMarker = null;
        if (trimmed.StartsWith("<!--")) endMarker = "-->";
        else if (trimmed.StartsWith("<?")) endMarker = "?>";
        else if (trimmed.StartsWith("<![CDATA[")) endMarker = "]]>";
        else if (trimmed.Length > 2 && trimmed[1] == '!' && Char.IsUpper(trimmed[2])) endMarker = ">";

        if (endMarker != null)
        {
            // 直到结束标记（可跨空行），未找到则到文档末尾
            var sb = new System.Text.StringBuilder();
            while (_pos < _lines.Length)
            {
                var line = _lines[_pos];
                sb.Append(line).Append('\n');
                _pos++;
                if (line.Contains(endMarker)) break;
            }
            var html = sb.ToString();
            if (html.EndsWith("\n")) html = html[..^1];
            return MarkdownBlock.CreateHtmlBlock(html);
        }

        // ⑥ 块级标签起始：直到空行
        if (IsBlockTagStart(trimmed))
        {
            var sb = new System.Text.StringBuilder();
            while (_pos < _lines.Length && !String.IsNullOrWhiteSpace(_lines[_pos]))
            {
                sb.Append(_lines[_pos]).Append('\n');
                _pos++;
            }
            var html = sb.ToString();
            if (html.EndsWith("\n")) html = html[..^1];
            return MarkdownBlock.CreateHtmlBlock(html);
        }

        // ⑦ 单行完整标签 + 后跟空行
        if (IsCompleteSingleTag(trimmed) && _pos + 1 < _lines.Length && String.IsNullOrWhiteSpace(_lines[_pos + 1]))
        {
            var line = _lines[_pos];
            _pos++;
            return MarkdownBlock.CreateHtmlBlock(line);
        }

        return null;
    }

    /// <summary>解析特殊 HTML 块（&lt;pre/script/style/textarea 直到匹配闭合标签）</summary>
    /// <returns>HTML 块</returns>
    private MarkdownBlock ParseSpecialHtmlBlock()
    {
        var trimmed = _lines[_pos].TrimStart();
        var tagName = "";
        var i = 1;
        while (i < trimmed.Length && Char.IsLetter(trimmed[i])) { tagName += trimmed[i]; i++; }
        var closeTag = "</" + tagName;

        var sb = new System.Text.StringBuilder();
        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            sb.Append(line).Append('\n');
            _pos++;
            if (line.Contains(closeTag)) break;
        }
        var html = sb.ToString();
        if (html.EndsWith("\n")) html = html[..^1];
        return MarkdownBlock.CreateHtmlBlock(html);
    }

    /// <summary>判断是否为特殊标签起始行（&lt;pre|script|style|textarea）</summary>
    private static Boolean IsSpecialTagStart(String trimmed)
    {
        foreach (var name in new[] { "pre", "script", "style", "textarea" })
        {
            if (trimmed.Length == name.Length + 1 &&
                trimmed.StartsWith("<" + name, StringComparison.OrdinalIgnoreCase))
                return true;
            if (trimmed.Length > name.Length + 1 &&
                trimmed.StartsWith("<" + name, StringComparison.OrdinalIgnoreCase))
            {
                var after = trimmed[name.Length + 1];
                if (after is ' ' or '\t' or '>') return true;
            }
        }
        return false;
    }

    /// <summary>块级标签名集合（CommonMark 类型 ⑥）</summary>
    private static readonly HashSet<String> BlockTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "address", "article", "aside", "blockquote", "body", "caption", "center", "col",
        "colgroup", "dd", "details", "dialog", "div", "dl", "dt", "fieldset",
        "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6",
        "head", "header", "hr", "html", "iframe", "legend", "li", "main", "menu",
        "nav", "ol", "optgroup", "option", "p", "section", "summary", "table", "tbody",
        "td", "tfoot", "th", "thead", "title", "tr", "track", "ul",
    };

    /// <summary>判断是否为块级标签起始行（CommonMark 类型 ⑥）</summary>
    private static Boolean IsBlockTagStart(String trimmed)
    {
        if (trimmed.Length < 2 || trimmed[0] != '<') return false;
        var nameStart = 1;
        var nameEnd = nameStart;
        while (nameEnd < trimmed.Length && (Char.IsLetter(trimmed[nameEnd]) || Char.IsDigit(trimmed[nameEnd])))
            nameEnd++;
        if (nameEnd == nameStart) return false;
        var name = trimmed[nameStart..nameEnd];
        if (!BlockTags.Contains(name)) return false;
        if (nameEnd >= trimmed.Length) return true;
        var after = trimmed[nameEnd];
        return after is ' ' or '\t' or '>' or '/';
    }

    /// <summary>判断行是否为 HTML 块起始（CommonMark type 1-6，可中断段落）</summary>
    /// <param name="trimmed">去前导空白后的行文本</param>
    /// <returns>是否为可中断段落的 HTML 块起始</returns>
    private static Boolean IsHtmlBlockStart(String trimmed)
    {
        if (trimmed.Length < 2 || trimmed[0] != '<') return false;

        // 类型① 特殊标签（pre/script/style/textarea）
        if (IsSpecialTagStart(trimmed)) return true;

        // 类型②③④⑤ 注释/处理指令/CDATA/声明
        if (trimmed.StartsWith("<!--") || trimmed.StartsWith("<?") || trimmed.StartsWith("<![CDATA["))
            return true;
        if (trimmed.Length > 2 && trimmed[1] == '!' && Char.IsUpper(trimmed[2])) return true;

        // 类型⑥ 块级标签
        return IsBlockTagStart(trimmed);
    }

    /// <summary>判断是否为单行完整标签（CommonMark 类型 ⑦）</summary>
    private static Boolean IsCompleteSingleTag(String trimmed)
    {
        if (trimmed.Length < 3 || trimmed[0] != '<') return false;
        // 完整开始/结束标签：<tag ...> 或 </tag>，无跨行
        var close = trimmed.IndexOf('>');
        if (close < 0 || close != trimmed.Length - 1) return false;
        var inner = trimmed[1..close];
        if (inner.Length == 0) return false;

        var isClose = inner[0] == '/';
        var nameStart = isClose ? 1 : 0;
        if (nameStart >= inner.Length || !Char.IsLetter(inner[nameStart])) return false;

        var nameEnd = nameStart;
        while (nameEnd < inner.Length && (Char.IsLetterOrDigit(inner[nameEnd]) || inner[nameEnd] == '-'))
            nameEnd++;
        var rest = inner[nameEnd..];

        // 结束标签：剩余必须为空
        if (isClose) return rest.Trim().Length == 0;

        // 开始标签：剩余为空、自闭合 / 或合法属性列表
        if (rest.Length == 0 || rest.Trim() == "/") return true;
        return IsValidAttributes(rest);
    }

    /// <summary>校验属性列表：name 或 name="value" 以空白分隔（避免 URL 误判为标签）</summary>
    /// <param name="rest">标签名后的剩余文本</param>
    /// <returns>是否为合法属性序列</returns>
    private static Boolean IsValidAttributes(String rest)
    {
        var i = 0;
        while (i < rest.Length)
        {
            // 跳过空白
            while (i < rest.Length && (rest[i] == ' ' || rest[i] == '\t')) i++;
            if (i >= rest.Length) return true;
            // 自闭合
            if (rest[i] == '/' && i == rest.Length - 1) return true;
            // 属性名
            var nameStart = i;
            while (i < rest.Length && (Char.IsLetterOrDigit(rest[i]) || rest[i] == '-' || rest[i] == '_'))
                i++;
            if (i == nameStart) return false;
            // 值部分
            if (i < rest.Length && rest[i] == '=')
            {
                i++;
                if (i < rest.Length && (rest[i] == '"' || rest[i] == '\''))
                {
                    var quote = rest[i];
                    i++;
                    var closeQuote = rest.IndexOf(quote, i);
                    if (closeQuote < 0) return false;
                    i = closeQuote + 1;
                }
                else
                {
                    // 无引号值（简单非空白序列）
                    var vStart = i;
                    while (i < rest.Length && !Char.IsWhiteSpace(rest[i])) i++;
                    if (i == vStart) return false;
                }
            }
        }
        return true;
    }

    private static Boolean IsThematicBreak(String trimmed)
    {
        if (trimmed.Length < 3) return false;
        var ch = trimmed[0];
        if (ch != '-' && ch != '*' && ch != '_') return false;
        var count = 0;
        foreach (var c in trimmed)
        {
            if (c == ch) count++;
            else if (c != ' ') return false;
        }
        return count >= 3;
    }

    private MarkdownBlock ParseBlockQuote()
    {
        var quotedLines = new List<String>();
        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            var trimmed = line.TrimStart();
            if (String.IsNullOrWhiteSpace(line)) { _pos++; break; }
            if (!trimmed.StartsWith(">") && quotedLines.Count > 0 && !String.IsNullOrWhiteSpace(line))
            {
                // lazy continuation: include as-is
                quotedLines.Add(line);
                _pos++;
                continue;
            }
            if (trimmed.StartsWith("> "))
                quotedLines.Add(trimmed[2..]);
            else if (trimmed == ">")
                quotedLines.Add("");
            else if (trimmed.StartsWith(">"))       // ">>" nested, strip one level
                quotedLines.Add(trimmed[1..]);
            else
                break;
            _pos++;
        }
        // 快速路径（MD15）：单行普通内容直接行内解析，避免创建子解析器（引用块分配主成本，与 MD14 列表项同理）
        // 语义等价：子解析器 BuildParagraphText 对每行 Trim()，单行内容即 Trim() 后行内解析；
        // IsBlockStart 含 '[' 排除引用定义/脚注定义行，空行走原路径保持原行为
        if (quotedLines.Count == 1 && !String.IsNullOrWhiteSpace(quotedLines[0]) && !IsBlockStart(quotedLines[0]))
        {
            var inlines = ParseInlineWithRefs(quotedLines[0].Trim());
            return MarkdownBlock.CreateBlockQuote([MarkdownBlock.CreateParagraph(inlines)]);
        }
        // Re-parse quoted content
        var inner = new MarkdownParser { _refs = _refs, Pipeline = Pipeline };
        var innerDoc = inner.ParseLines(quotedLines);
        return MarkdownBlock.CreateBlockQuote(innerDoc.Blocks);
    }

    private static Boolean IsBulletListMarker(String trimmed, out Char marker)
    {
        marker = ' ';
        if (trimmed.Length < 2) return false;
        var ch = trimmed[0];
        if (ch != '-' && ch != '*' && ch != '+') return false;
        if (trimmed[1] != ' ' && trimmed[1] != '\t') return false;
        marker = ch;
        return true;
    }

    private static Boolean IsOrderedListMarker(String trimmed, out Int32 num, out Char delimiter)
    {
        num = 1; delimiter = '.';
        // 快速失败（MD20）：有序列表必以数字开头，避免普通行每行跑正则
        if (trimmed.Length == 0 || !Char.IsDigit(trimmed[0])) return false;
        var match = _orderedList.Match(trimmed);
        if (!match.Success) return false;
        num = Int32.Parse(match.Groups[1].Value);
        delimiter = match.Groups[2].Value[0];
        return true;
    }

    /// <summary>解析缩进代码块（4+ 空格或制表缩进）</summary>
    /// <returns>代码块</returns>
    private MarkdownBlock ParseIndentedCode()
    {
        var sb = new System.Text.StringBuilder();
        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            if (String.IsNullOrWhiteSpace(line))
            {
                // 空行：若后续仍为缩进行则作为代码内空行，否则结束代码块
                var next = _pos + 1;
                while (next < _lines.Length && String.IsNullOrWhiteSpace(_lines[next])) next++;
                if (next < _lines.Length && IsIndented(_lines[next]))
                {
                    sb.Append('\n');
                    _pos++;
                    continue;
                }
                break;
            }
            if (IsIndented(line))
            {
                sb.Append(StripIndent(line)).Append('\n');
                _pos++;
            }
            else
            {
                break;
            }
        }
        var code = sb.ToString();
        if (code.EndsWith("\n")) code = code[..^1];
        return MarkdownBlock.CreateCodeBlock(code);
    }

    /// <summary>计算行前导空白列数（Tab = 4 列，CommonMark 规范）</summary>
    /// <param name="line">行文本</param>
    /// <returns>缩进列数</returns>
    private static Int32 GetIndent(String line)
    {
        var col = 0;
        foreach (var c in line)
        {
            if (c == ' ') col++;
            else if (c == '\t') col = (col / 4 + 1) * 4;
            else break;
        }
        return col;
    }

    /// <summary>剥离行前导空白到指定列（保留相对缩进）</summary>
    /// <param name="line">行文本</param>
    /// <param name="columns">目标列数</param>
    /// <returns>剥离后文本</returns>
    private static String StripColumns(String line, Int32 columns)
    {
        var col = 0;
        var idx = 0;
        while (idx < line.Length && col < columns)
        {
            var c = line[idx];
            if (c == ' ') { col++; idx++; }
            else if (c == '\t') { col = (col / 4 + 1) * 4; idx++; }
            else break;
        }
        return line[idx..];
    }

    /// <summary>判断行是否为缩进代码行（4 空格或制表）</summary>
    private static Boolean IsIndented(String line)
    {
        if (line.StartsWith("\t")) return true;
        return line.Length - line.TrimStart().Length >= 4;
    }

    /// <summary>剥离缩进代码行的 4 空格或 1 制表</summary>
    private static String StripIndent(String line)
    {
        if (line.StartsWith("\t")) return line[1..];
        return line[4..];
    }

    private MarkdownBlock ParseList(Boolean ordered, Int32 startNum = 1)
    {
        var itemContents = new List<List<String>>();
        var taskFlags = new List<(Boolean isTask, Boolean isChecked)>();
        var bulletChar = ' ';
        var orderedDelim = ' ';
        var markerIndent = -1;
        var contentColumn = 2;
        var firstNum = startNum;

        var currentLines = new List<String>();
        var isTask = false;
        var isChecked = false;
        var first = true;
        var loose = false; // CommonMark 松散列表：项间/项内出现空行

        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            var trimmed = line.TrimStart();
            var indent = GetIndent(line);

            // 空行：预读判断列表是否延续（后续为同缩进标记或缩进续行）
            if (String.IsNullOrWhiteSpace(line))
            {
                var next = _pos + 1;
                while (next < _lines.Length && String.IsNullOrWhiteSpace(_lines[next])) next++;
                if (next < _lines.Length)
                {
                    var nextTrimmed = _lines[next].TrimStart();
                    var nextIndent = _lines[next].Length - nextTrimmed.Length;
                    var nextMarker = !ordered
                        ? IsBulletListMarker(nextTrimmed, out var _) && (markerIndent < 0 || nextIndent == markerIndent)
                        : IsOrderedListMarker(nextTrimmed, out var _, out var _) && (markerIndent < 0 || nextIndent == markerIndent);
                    if (nextMarker || nextIndent >= contentColumn)
                    {
                        // 列表延续：跳过空行，空行作为项内段落分隔 → 松散列表
                        loose = true;
                        while (_pos < next) _pos++;
                        currentLines.Add("");
                        continue;
                    }
                }
                break;
            }

            // 列表标记
            var isMarker = false;
            Char m = ' ';
            Int32 num = 0;
            Char delim = ' ';
            if (!ordered)
                isMarker = IsBulletListMarker(trimmed, out m);
            else
                isMarker = IsOrderedListMarker(trimmed, out num, out delim);

            if (isMarker && (markerIndent < 0 || indent == markerIndent))
            {
                var spaceIdx = trimmed.IndexOf(' ');
                var markerLen = spaceIdx >= 0 ? spaceIdx + 1 : trimmed.Length;

                if (markerIndent < 0)
                {
                    markerIndent = indent;
                    if (!ordered) bulletChar = m;
                    else { orderedDelim = delim; firstNum = num; }
                    contentColumn = markerIndent + markerLen;
                }
                else if (!ordered && m != bulletChar)
                {
                    break;
                }
                else if (ordered && delim != orderedDelim)
                {
                    break;
                }

                // 收尾当前列表项
                if (!first)
                {
                    itemContents.Add(currentLines);
                    taskFlags.Add((isTask, isChecked));
                }
                first = false;

                // 新列表项首行：剥离标记与任务标记
                _pos++;
                var content = spaceIdx >= 0 ? trimmed[(spaceIdx + 1)..] : "";
                currentLines = [];
                isTask = false;
                isChecked = false;
                if (Pipeline == null || Pipeline.EnableTaskLists)
                    ParseTaskMarker(content, out content, out isTask, out isChecked);
                if (content.Length > 0) currentLines.Add(content);
                continue;
            }

            // 缩进续行：并入当前列表项（剥离内容起始列）
            if (!first && indent >= contentColumn)
            {
                currentLines.Add(StripColumns(line, contentColumn));
                _pos++;
                continue;
            }

            break;
        }

        // 收尾最后一项
        if (!first)
        {
            itemContents.Add(currentLines);
            taskFlags.Add((isTask, isChecked));
        }

        // 构建列表项
        var items = new List<MarkdownBlock>();
        for (var i = 0; i < itemContents.Count; i++)
        {
            items.Add(BuildItemFromContent(itemContents[i], taskFlags[i].isTask, taskFlags[i].isChecked));
        }

        if (!ordered) return new BulletListBlock(items) { IsLoose = loose };
        return new OrderedListBlock(items, firstNum) { IsLoose = loose };
    }

    /// <summary>判断行是否为块级元素起始（列表项快速路径使用，MD14；避免单行普通内容误走子解析器）</summary>
    /// <param name="line">内容行（已剥离列表标记）</param>
    /// <returns>可能是块级起始时为 true</returns>
    private static Boolean IsBlockStart(String line)
    {
        var t = line.TrimStart();
        if (t.Length == 0) return false;
        var c = t[0];
        // ATX 标题 / 引用 / 列表 / 围栏 / HTML 块 / 表格 / 定义列表 / 引用定义 / 脚注
        if (c is '#' or '>' or '-' or '*' or '+' or '`' or '~' or '<' or '|' or ':' or '[' or '$') return true;
        // 有序列表（数字 + . 或 )）
        if (Char.IsDigit(c)) return true;
        return false;
    }

    /// <summary>从列表项内容行构建块（单个段落用行内表示，多块/嵌套用子块）</summary>
    /// <param name="contentLines">剥离标记后的内容行</param>
    /// <param name="isTask">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    /// <returns>列表项块</returns>
    private MarkdownBlock BuildItemFromContent(List<String> contentLines, Boolean isTask, Boolean isChecked)
    {
        // 去掉首尾空行
        while (contentLines.Count > 0 && String.IsNullOrWhiteSpace(contentLines[0])) contentLines.RemoveAt(0);
        while (contentLines.Count > 0 && String.IsNullOrWhiteSpace(contentLines[^1])) contentLines.RemoveAt(contentLines.Count - 1);

        if (contentLines.Count == 0)
            return MarkdownBlock.CreateListItem([], isTask, isChecked);

        // 深度保护：防止恶意深层嵌套导致栈溢出
        if (_listDepth >= 32)
        {
            var text = String.Join(" ", contentLines).Trim();
            return MarkdownBlock.CreateListItem(ParseInlineWithRefs(text), isTask, isChecked);
        }

        // 快速路径（MD14）：单行普通内容（非块级起始）直接行内解析，避免创建子解析器（列表分配主成本）
        if (contentLines.Count == 1 && !IsBlockStart(contentLines[0]))
        {
            var inlines = ParseInlineWithRefs(contentLines[0].Trim());
            return MarkdownBlock.CreateListItem(inlines, isTask, isChecked);
        }

        // 重新解析为块，支持嵌套列表/引用/代码块/多段落等
        var inner = new MarkdownParser
        {
            _listDepth = _listDepth + 1,
            _refs = _refs,
            Pipeline = Pipeline,
        };
        var innerDoc = inner.ParseLines(contentLines, true);

        if (innerDoc.Blocks.Count == 1 && innerDoc.Blocks[0].Type == MarkdownBlockType.Paragraph)
        {
            // 简单列表项：单个段落 → 用行内表示
            return MarkdownBlock.CreateListItem(innerDoc.Blocks[0].Inlines, isTask, isChecked);
        }
        return MarkdownBlock.CreateListItemWithBlocks(innerDoc.Blocks, isTask, isChecked);
    }

    /// <summary>解析任务列表标记 [ ] / [x] / [X]</summary>
    /// <param name="content">原始内容</param>
    /// <param name="rest">剥离标记后的内容</param>
    /// <param name="isTask">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    private static void ParseTaskMarker(String content, out String rest, out Boolean isTask, out Boolean isChecked)
    {
        rest = content;
        isTask = false;
        isChecked = false;
        var m = _taskMarker.Match(content);
        if (m.Success)
        {
            isTask = true;
            isChecked = m.Groups[1].Value != " ";
            rest = m.Groups[2].Value;
        }
    }

    private MarkdownBlock ParseParagraphOrSetext(Int32 indentHint)
    {
        var lines = new List<String>();
        var startPos = _pos;

        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            if (String.IsNullOrWhiteSpace(line)) { _pos++; break; }
            var trimmedNext = line.TrimStart();

            // Setext heading underline
            if (lines.Count > 0 && _setext1.IsMatch(trimmedNext))
            {
                _pos++;
                var headingText = String.Join(" ", lines).Trim();
                return MarkdownBlock.CreateHeading(1, ParseInlineWithRefs(headingText));
            }
            // Setext heading underline（单个 `-` 是空列表项标记，不构成 underline，CommonMark 例 77）
            if (lines.Count > 0 && trimmedNext != "-" && _setext2.IsMatch(trimmedNext))
            {
                _pos++;
                var headingText = String.Join(" ", lines).Trim();
                return MarkdownBlock.CreateHeading(2, ParseInlineWithRefs(headingText));
            }

            // ATX heading interrupts paragraph
            if (TryParseAtxHeading(trimmedNext) != null) break;
            // Thematic break interrupts paragraph
            if (IsThematicBreak(trimmedNext)) break;
            // Fence code interrupts
            if (trimmedNext.StartsWith("```") || trimmedNext.StartsWith("~~~")) break;
            // Block quote interrupts
            if (trimmedNext.StartsWith(">")) break;
            // HTML 块（type 1-6）可中断段落；type 7 单行完整标签不可中断
            if (trimmedNext.StartsWith("<") && IsHtmlBlockStart(trimmedNext)) break;
            // List markers interrupt (after first line)；表格 delimiter 行仅在首行也含 |（构成表格的前提）时豁免中断，避免 `Foo\n- | -` 误合并为段落
            if (lines.Count > 0 && (IsBulletListMarker(trimmedNext, out var _)
                || IsOrderedListMarker(trimmedNext, out var _, out var _)) &&
                !(lines[0].Contains("|") && IsTableDelimiterLine(trimmedNext)))
                break;

            // Hard break: trailing 2+ spaces before newline
            if (line.EndsWith("  ") && lines.Count > 0)
                lines.Add(line.TrimEnd() + "  ");
            else
                lines.Add(line);
            _pos++;
        }

        if (lines.Count == 0) return MarkdownBlock.CreateParagraph([]);

        // Check if it looks like a GFM table（header 与 delimiter 行均须含 `|`，GFM 规范；避免 `Foo\n-` / `Foo\n- | -` 误判）
        if (lines.Count >= 2 && (Pipeline == null || Pipeline.EnableTables) &&
            lines[0].Contains("|") && lines[1].Contains("|"))
        {
            var table = TryParseTable(lines);
            if (table != null) return table;
        }

        // Join with spaces (CommonMark paragraph continuation)
        var paragraphText = BuildParagraphText(lines);
        return MarkdownBlock.CreateParagraph(ParseInlineWithRefs(paragraphText));
    }

    private static String BuildParagraphText(List<String> lines)
    {
        // 单行快路径（MD24）：绝大多数段落为单行，直接 Trim 返回，避免 StringBuilder + ToString 分配
        if (lines.Count == 1) return lines[0].Trim();

        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            if (i > 0)
            {
                // 上一行行尾两空格 → 硬换行；否则软换行（交由行内解析生成 SoftBreak/HardBreak）
                sb.Append(lines[i - 1].EndsWith("  ") ? "  \n" : "\n");
            }
            sb.Append(line.Trim());
        }
        return sb.ToString();
    }

    private MarkdownBlock? TryParseTable(List<String> lines)
    {
        if (lines.Count < 2) return null;
        var sep = lines[1].Trim();

        // 分隔行单元格解析与校验（GFM：每格必须为 :?-+:?，至少 1 个连字符）
        var alignCells = SplitTableRow(sep);
        if (alignCells.Count == 0) return null;
        var alignments = new List<String>();
        foreach (var cell in alignCells)
        {
            var c = cell.Trim();
            if (!IsTableDelimiterCell(c)) return null;
            if (c.StartsWith(":") && c.EndsWith(":")) alignments.Add("center");
            else if (c.EndsWith(":")) alignments.Add("right");
            else if (c.StartsWith(":")) alignments.Add("left");
            else alignments.Add("");
        }

        var table = new TableBlock();

        // Header row（列数以分隔行为准，多出丢弃、缺少补空，GFM 规范）
        var headerRow = new TableRowBlock();
        var headerCells = SplitTableRow(lines[0]);
        for (var i = 0; i < alignments.Count; i++)
        {
            var cellText = i < headerCells.Count ? headerCells[i].Trim() : "";
            var cell = new TableCellBlock(
                ParseInlineWithRefs(cellText),
                isHeader: true,
                alignment: alignments[i]);
            headerRow.Children.Add(cell);
        }
        table.Children.Add(headerRow);

        // Data rows
        for (var r = 2; r < lines.Count; r++)
        {
            var dataRow = new TableRowBlock();
            var dataCells = SplitTableRow(lines[r]);
            for (var i = 0; i < alignments.Count; i++)
            {
                var cellText = i < dataCells.Count ? dataCells[i].Trim() : "";
                var cell = new TableCellBlock(
                    ParseInlineWithRefs(cellText),
                    isHeader: false,
                    alignment: alignments[i]);
                dataRow.Children.Add(cell);
            }
            table.Children.Add(dataRow);
        }

        return table;
    }

    /// <summary>校验表格分隔行单元格（GFM：1+ 个连字符，可选左右冒号）</summary>
    /// <param name="cell">单元格文本（去首尾空白）</param>
    /// <returns>是否合法分隔单元格</returns>
    private static Boolean IsTableDelimiterCell(String cell)
    {
        if (cell.Length == 0) return false;
        var i = 0;
        if (cell[i] == ':') i++;
        if (i >= cell.Length) return false;
        var hasHyphen = false;
        while (i < cell.Length && cell[i] == '-') { hasHyphen = true; i++; }
        if (!hasHyphen) return false;
        if (i < cell.Length && cell[i] == ':') i++;
        return i == cell.Length;
    }

    /// <summary>判断整行是否为合法表格分隔行（GFM：含 | 且每个单元格均为 :?-+:?）</summary>
    /// <param name="line">行文本</param>
    /// <returns>是否为分隔行</returns>
    private static Boolean IsTableDelimiterLine(String line)
    {
        if (!line.Contains("|")) return false;
        var cells = SplitTableRow(line);
        if (cells.Count == 0) return false;
        foreach (var cell in cells)
        {
            if (!IsTableDelimiterCell(cell.Trim())) return false;
        }
        return true;
    }

    /// <summary>拆分表格行，支持 \| 转义管道符与代码段内管道符不拆分（GFM）</summary>
    /// <param name="line">表格行文本</param>
    /// <returns>单元格列表</returns>
    private static List<String> SplitTableRow(String line)
    {
        // 快速路径（MD18）：行内无反引号/反斜杠时，| 全为分隔符，直接切分（绝大多数简单表格）。
        // MD21b：IndexOfAny 一次扫描替代两次 IndexOf（表格为最高耗时密度块，省 1 次全行扫描）
        if (line.IndexOfAny(_tableEscapeChars) < 0)
        {
            var start = 0;
            var end = line.Length;
            while (start < end && Char.IsWhiteSpace(line[start])) start++;
            while (end > start && Char.IsWhiteSpace(line[end - 1])) end--;
            if (start < end && line[start] == '|') start++;
            if (end > start && line[end - 1] == '|') end--;
            var cells = new List<String>();
            var cs = start;
            for (var i = start; i <= end; i++)
            {
                if (i == end || line[i] == '|')
                {
                    cells.Add(line.Substring(cs, i - cs).Trim());
                    cs = i + 1;
                }
            }
            return cells;
        }
        else
        {
            // 索引边界剥离首尾空白与管道符，避免 Trim/[1..]/[..^1] 子串分配（MD16）
            var start = 0;
            var end = line.Length;
            while (start < end && Char.IsWhiteSpace(line[start])) start++;
            while (end > start && Char.IsWhiteSpace(line[end - 1])) end--;
            if (start < end && line[start] == '|') start++;
            if (end > start && line[end - 1] == '|') end--;

            var cells = new List<String>();
            var cellStart = start;
            var slashCount = 0;
            var codeTicks = 0; // >0 表示处于代码段内，值为开头反引号个数
            var i = start;
        while (i < end)
        {
            var c = line[i];
            if (codeTicks > 0)
            {
                // 代码段内：管道符/转义不生效，等长反引号串闭合
                if (c == '`')
                {
                    var run = 0;
                    while (i + run < end && line[i + run] == '`') run++;
                    if (run == codeTicks) codeTicks = 0;
                    i += run;
                    continue;
                }
                i++;
                continue;
            }

            if (c == '`')
            {
                // 进入代码段：记录反引号个数
                var run = 0;
                while (i + run < end && line[i + run] == '`') run++;
                codeTicks = run;
                i += run;
                continue;
            }
            if (c == '\\') { slashCount++; i++; continue; }
            if (c == '|')
            {
                // 奇数个连续反斜杠 → 转义管道符，不拆分（\| 保留在单元格内，由 ProcessCell 去反斜杠）
                if (slashCount % 2 == 0)
                {
                    cells.Add(ProcessCell(line.Substring(cellStart, i - cellStart)));
                    cellStart = i + 1;
                }
                slashCount = 0;
                i++;
                continue;
            }
            slashCount = 0;
            i++;
        }
            cells.Add(ProcessCell(line.Substring(cellStart, end - cellStart)));
            return cells;
        }
    }

    /// <summary>处理单元格：去首尾空白 + 去转义管道符反斜杠（\| → |）</summary>
    /// <param name="raw">单元格原始子串</param>
    /// <returns>处理后的单元格文本</returns>
    private static String ProcessCell(String raw)
    {
        var cell = raw.Trim();
        // 仅含 \| 转义时处理（普通单元格零额外分配）
        var idx = cell.IndexOf("\\|", StringComparison.Ordinal);
        if (idx < 0) return cell;
        var sb = new System.Text.StringBuilder(cell.Length);
        var i = 0;
        while (i < cell.Length)
        {
            if (i + 1 < cell.Length && cell[i] == '\\' && cell[i + 1] == '|')
            {
                sb.Append('|');
                i += 2;
            }
            else
            {
                sb.Append(cell[i]);
                i++;
            }
        }
        return sb.ToString();
    }
    #endregion

    #region MD05 高级解析（FrontMatter/脚注/数学公式/自动链接）
    /// <summary>Emoji 短码映射表 (MD05-05)</summary>
    private static readonly Dictionary<String, String> EmojiMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["smile"] = "😄", ["laughing"] = "😆", ["joy"] = "😂", ["grin"] = "😁",
        ["wink"] = "😉", ["blush"] = "😊", ["heart_eyes"] = "😍", ["kissing_heart"] = "😘",
        ["thinking"] = "🤔", ["neutral_face"] = "😐", ["expressionless"] = "😑",
        ["angry"] = "😠", ["rage"] = "😡", ["tired_face"] = "😫", ["cry"] = "😢",
        ["sob"] = "😭", ["scream"] = "😱", ["sleeping"] = "😴", ["sunglasses"] = "😎",
        ["+1"] = "👍", ["thumbsup"] = "👍", ["-1"] = "👎", ["thumbsdown"] = "👎",
        ["clap"] = "👏", ["wave"] = "👋", ["ok_hand"] = "👌", ["pray"] = "🙏",
        ["muscle"] = "💪", ["fire"] = "🔥", ["star"] = "⭐", ["heart"] = "❤️",
        ["broken_heart"] = "💔", ["zap"] = "⚡", ["boom"] = "💥", ["rocket"] = "🚀",
        ["check"] = "✅", ["x"] = "❌", ["warning"] = "⚠️", ["question"] = "❓",
        ["bulb"] = "💡", ["book"] = "📖", ["memo"] = "📝", ["package"] = "📦",
        ["lock"] = "🔒", ["unlock"] = "🔓", ["key"] = "🔑", ["hammer"] = "🔨",
        ["link"] = "🔗", ["email"] = "📧", ["phone"] = "📱", ["computer"] = "💻",
        ["tada"] = "🎉", ["gift"] = "🎁", ["coffee"] = "☕", ["beer"] = "🍺",
        ["pizza"] = "🍕", ["apple"] = "🍎", ["car"] = "🚗", ["airplane"] = "✈️",
        ["sunny"] = "☀️", ["cloud"] = "☁️", ["rain"] = "🌧️", ["snow"] = "❄️",
        ["100"] = "💯", ["heavy_check_mark"] = "✔️", ["heavy_multiplication_x"] = "✖️",
    };

    private void ParseFrontMatter(MarkdownDocument doc)
    {
        if (_pos >= _lines.Length) return;
        if (_lines[_pos].TrimEnd() != "---") return;

        var start = _pos;
        _pos++;

        // 数组值支持：key:
        //   - item1
        //   - item2
        String? listKey = null;
        List<String>? listValue = null;

        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            if (line.TrimEnd() == "---")
            {
                // 收尾数组值
                if (listKey != null && listValue is { Count: > 0 })
                    doc.FrontMatter[listKey] = String.Join(", ", listValue);
                _pos++;
                return;
            }

            var trimmed = line.Trim();
            // 注释行
            if (trimmed.StartsWith("#")) { _pos++; continue; }

            // 数组项（延续上一键）
            if (trimmed.StartsWith("- "))
            {
                if (listKey != null)
                {
                    listValue ??= [];
                    listValue.Add(trimmed[2..].Trim().Trim('"', '\''));
                }
                _pos++;
                continue;
            }

            // 新键值对（引号值内冒号不截断）
            var colonIdx = FindColonOutsideQuotes(trimmed);
            if (colonIdx > 0)
            {
                // 收尾上一数组
                if (listKey != null && listValue is { Count: > 0 })
                    doc.FrontMatter[listKey] = String.Join(", ", listValue);
                listKey = null;
                listValue = null;

                var key = trimmed[..colonIdx].Trim();
                var value = trimmed[(colonIdx + 1)..].Trim();

                // 数组开头 key:（值为空）
                if (value.Length == 0 || value is "[" or "]")
                {
                    listKey = key;
                    listValue = [];
                    _pos++;
                    continue;
                }

                // 去掉引号（双引号值还原反斜杠转义，与 EscapeYamlValue 对称）
                if (value.Length >= 2 &&
                    ((value[0] == '"' && value[^1] == '"') ||
                     (value[0] == '\'' && value[^1] == '\'')))
                {
                    var quote = value[0];
                    value = value[1..^1];
                    if (quote == '"') value = UnescapeYaml(value);
                }
                if (key.Length > 0)
                    doc.FrontMatter[key] = value;
            }
            _pos++;
        }

        _pos = start;
    }

    /// <summary>查找引号外的冒号位置（引号值内冒号不截断）</summary>
    /// <param name="line">行文本</param>
    /// <returns>冒号位置，未找到返回 -1</returns>
    private static Int32 FindColonOutsideQuotes(String line)
    {
        var inQuote = '\0';
        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (inQuote != '\0')
            {
                if (c == inQuote) inQuote = '\0';
                continue;
            }
            if (c is '"' or '\'') { inQuote = c; continue; }
            if (c == ':') return i;
        }
        return -1;
    }

    /// <summary>还原 YAML 双引号值的反斜杠转义（\\ → \，\" → "）</summary>
    /// <param name="value">带转义的值</param>
    /// <returns>还原后值</returns>
    private static String UnescapeYaml(String value)
    {
        if (!value.Contains('\\')) return value;
        var sb = new System.Text.StringBuilder(value.Length);
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (c == '\\' && i + 1 < value.Length)
            {
                var n = value[i + 1];
                if (n == '\\' || n == '"') { sb.Append(n); i++; continue; }
            }
            sb.Append(c);
        }
        return sb.ToString();
    }

    private Boolean TryParseFootnoteDefinition(String trimmed, out MarkdownBlock? block)
    {
        block = null;

        // 快速失败（MD20）：脚注定义必以 [^ 开头，避免普通行每行跑正则
        if (trimmed.Length < 3 || trimmed[0] != '[' || trimmed[1] != '^') return false;

        // [^id]: text
        var match = _footnoteDef.Match(trimmed);
        if (!match.Success) return false;

        var id = match.Groups[1].Value;
        var defText = match.Groups[2].Value;
        var inlines = ParseInlineWithRefs(defText);
        block = new FootnoteDefinitionBlock(id, inlines);
        return true;
    }

    /// <summary>解析缩写定义 *[ABBR]: 全称 (MD05-08)</summary>
    private void ParseAbbreviation(String trimmed)
    {
        // 格式: *[ABBR]: Full Definition
        var match = _abbrDef.Match(trimmed);
        if (!match.Success) return;

        var abbr = match.Groups[1].Value.Trim();
        var full = match.Groups[2].Value.Trim();
        if (abbr.Length > 0 && full.Length > 0 && _doc != null)
        {
            _doc.Abbreviations[abbr] = full;
        }
    }

    /// <summary>解析定义列表 (MD05-07)</summary>
    /// <remarks>回溯上一行作为术语，当前行 `: description` 作为描述</remarks>
    private MarkdownBlock? TryParseDefinitionList()
    {
        // 当前行是 `: description`，上一行应该是术语
        if (_pos < 1) return null;

        var prevLine = _lines[_pos - 1].Trim();
        if (String.IsNullOrWhiteSpace(prevLine)) return null;

        // 术语不能是特殊行
        if (prevLine.StartsWith("#") || prevLine.StartsWith(">") || prevLine.StartsWith("-")
            || prevLine.StartsWith("*") || prevLine.StartsWith("```") || prevLine.StartsWith("|"))
            return null;

        var dl = new DefinitionListBlock();

        // 解析术语（从上一行）
        var termInlines = ParseInlineWithRefs(prevLine);
        var term = new DefinitionTermBlock(termInlines);
        dl.Children.Add(term);

        // 解析当前描述
        var descLine = _lines[_pos].TrimStart();
        var descText = descLine.StartsWith(": ") ? descLine[2..] : descLine[1..];
        var descInlines = ParseInlineWithRefs(descText.Trim());
        var desc = new DefinitionDescriptionBlock(descInlines);
        dl.Children.Add(desc);
        _pos++;

        // 继续解析后续 `: ` 行（同一术语的多段描述）
        while (_pos < _lines.Length)
        {
            var nextLine = _lines[_pos].TrimStart();
            if (!nextLine.StartsWith(": ")) break;
            var moreText = nextLine[2..];
            var moreInlines = ParseInlineWithRefs(moreText.Trim());
            dl.Children.Add(new DefinitionDescriptionBlock(moreInlines));
            _pos++;
        }

        return dl;
    }

    private MarkdownBlock ParseMathBlock(String firstLine)
    {
        // $$ 单独一行开头
        var sb = new System.Text.StringBuilder();
        if (firstLine.Length > 2 && firstLine != "$$")
            sb.AppendLine(firstLine[2..].TrimStart());

        _pos++;
        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            _pos++;
            if (line.TrimEnd() == "$$") break;
            sb.AppendLine(line);
        }

        return new MathBlock(sb.ToString().Trim());
    }
    #endregion

    #region 行内解析
    /// <summary>解析行内 Markdown 内容，返回行内元素列表</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <returns>行内元素列表</returns>
    public static List<MarkdownInline> ParseInline(String text)
    {
        var result = new List<MarkdownInline>();
        if (String.IsNullOrEmpty(text)) return result;
        // 快速路径（MD18）：无行内标记字符时直接单文本令牌，避免 TokenizeInline 逐字符全分支检查
        if (IsPlainText(text))
        {
            result.Add(MarkdownInline.CreateText(text));
            return result;
        }
        ParseInlineCore(text, 0, text.Length, result, null, null, false);
        return result;
    }

    /// <summary>解析行内 Markdown 内容（携带引用链接定义上下文）</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <returns>行内元素列表</returns>
    private List<MarkdownInline> ParseInlineWithRefs(String text)
    {
        var result = new List<MarkdownInline>();
        if (String.IsNullOrEmpty(text)) return result;
        // 快速路径（MD18）：无行内标记字符时直接单文本令牌，避免 TokenizeInline 逐字符全分支检查
        if (IsPlainText(text))
        {
            result.Add(MarkdownInline.CreateText(text));
            return result;
        }
        ParseInlineCore(text, 0, text.Length, result, _refs, Pipeline, false);
        return result;
    }

    /// <summary>判断文本是否为纯文本（无行内标记触发字符），行内快速路径使用（MD18）</summary>
    /// <param name="text">行内文本</param>
    /// <returns>纯文本时为 true</returns>
    private static Boolean IsPlainText(String text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            // 排除行内标记触发字符：转义/实体/代码/图片/链接/强调/删除线/自动链接(< >)/数学/emoji/email/裸URL
            switch (c)
            {
                case '\\': case '&': case '`': case '!': case '[': case ']':
                case '*': case '_': case '~': case '<': case '>': case '$':
                case ':': case '@': case '\n': case '\r':
                    return false;
                case 'h': case 'f': case 'w':
                    // 裸 URL 前缀（MD19c）：仅当构成 URL 前缀（https:// http:// ftp:// www.）才排除，
                    // 避免含 h/f/w 的纯英文段落（如 "The quick brown fox..."）被误排除快速路径
                    if (IsUrlPrefixAt(text, i)) return false;
                    break;
            }
        }
        return true;
    }

    /// <summary>判断位置是否为裸 URL 前缀（与 TryParseBareUrl 的前缀集一致）</summary>
    /// <param name="text">文本</param>
    /// <param name="i">当前位置</param>
    /// <returns>构成 URL 前缀时为 true</returns>
    private static Boolean IsUrlPrefixAt(String text, Int32 i)
    {
        var c = text[i];
        if (c == 'w') return i + 4 <= text.Length && text.AsSpan(i, 4).Equals("www.".AsSpan(), StringComparison.Ordinal);
        if (c == 'f') return i + 6 <= text.Length && text.AsSpan(i, 6).Equals("ftp://".AsSpan(), StringComparison.OrdinalIgnoreCase);
        if (i + 7 <= text.Length && text.AsSpan(i, 7).Equals("https://".AsSpan(), StringComparison.OrdinalIgnoreCase)) return true;
        return i + 7 <= text.Length && text.AsSpan(i, 7).Equals("http://".AsSpan(), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>解析行内 Markdown 内容（核心递归实现）</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <param name="start">开始位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="result">输出列表</param>
    /// <param name="refs">引用链接定义（可为 null）</param>
    /// <param name="pipeline">处理管线（可为 null，视为全开）</param>
    /// <param name="inLink">是否处于链接内部（GFM：链接内不识别裸 URL 自动链接）</param>
    private static void ParseInlineCore(String text, Int32 start, Int32 end, List<MarkdownInline> result,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline, Boolean inLink)
    {
        // 两阶段解析（CommonMark §6.2）：
        // 1) 词法令牌化：文本拆分为令牌列表，* / _ 分隔符 run 收集为强调候选
        // 2) 分隔符栈解析：按 flanking + rule of 3 配对强调/粗体，嵌套结构
        // result 恒以空列表传入（顶层与递归子列表），tokens 别名复用 result，
        // 消除临时列表分配与末尾 AddRange 复制（MD16）
        var tokens = result;
        // delims 延迟创建（MD19）：TokenizeInline 仅遇 * / _ 分隔符时才分配，无强调文本省一个空 List
        var delims = TokenizeInline(text, start, end, tokens, refs, pipeline, inLink);
        ResolveEmphasis(tokens, delims);
    }

    /// <summary>行内令牌化：文本解析为令牌列表，* / _ 分隔符 run 收集为强调候选</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <param name="start">开始位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="tokens">输出令牌列表</param>
    /// <param name="refs">引用链接定义（可为 null）</param>
    /// <param name="pipeline">处理管线（可为 null，视为全开）</param>
    /// <param name="inLink">是否处于链接内部（GFM：链接内不识别裸 URL 自动链接）</param>
    /// <returns>强调分隔符信息列表（无强调时返回 null，MD19 延迟创建）</returns>
    private static List<DelimInfo>? TokenizeInline(String text, Int32 start, Int32 end, List<MarkdownInline> tokens,
        Dictionary<String, (String url, String title)>? refs,
        MarkdownPipeline? pipeline, Boolean inLink)
    {
        var i = start;
        var textStart = start;
        List<DelimInfo>? delims = null;

        // 管线开关：null 视为全开
        Boolean IsOn(Func<MarkdownPipeline, Boolean> get) => pipeline == null || get(pipeline);

        while (i < end)
        {
            var ch = text[i];

            // Hard break: "  \n" or "\\\n"
            if (ch == '\n')
            {
                var isHard = i >= 2 && text[i - 1] == ' ' && text[i - 2] == ' ';
                var isBackslashHard = i >= 1 && text[i - 1] == '\\';
                // 硬换行：剥离标记用的两个空格/反斜杠，避免进入文本令牌（CommonMark 硬换行标记不属于内容）
                if (isHard)
                    FlushText(text, textStart, i - 2, tokens);
                else
                    FlushText(text, textStart, i, tokens);
                if (isHard || isBackslashHard)
                    tokens.Add(MarkdownInline.CreateHardBreak());
                else
                    tokens.Add(MarkdownInline.CreateSoftBreak());
                i++;
                textStart = i;
                continue;
            }

            // Escape
            if (ch == '\\' && i + 1 < end && IsEscapable(text[i + 1]))
            {
                FlushText(text, textStart, i, tokens);
                tokens.Add(MarkdownInline.CreateText(text[i + 1].ToString()));
                i += 2;
                textStart = i;
                continue;
            }

            // HTML 实体解码：&amp; / &lt; / &#65; 等（CommonMark 最长实体名 ~31 字符）
            if (ch == '&')
            {
                // MD21b：扫描范围限 ≤33（semi - i <= 32 才有效，否则无需扫描到段尾）
                var semi = text.IndexOf(';', i + 1, Math.Min(33, end - i - 1));
                if (semi > i + 1 && semi - i <= 32)
                {
                    var entity = text.Substring(i + 1, semi - i - 1);
                    var decoded = DecodeEntity(entity);
                    if (decoded != null)
                    {
                        FlushText(text, textStart, i, tokens);
                        tokens.Add(MarkdownInline.CreateText(decoded));
                        i = semi + 1;
                        textStart = i;
                        continue;
                    }
                }
            }

            // Inline code: `...` （支持多反引号包裹，CommonMark 兼容）
            if (ch == '`')
            {
                // 统计开头反引号串长度
                var ticks = 0;
                while (i + ticks < end && text[i + ticks] == '`') ticks++;
                // 查找等长反引号串闭合
                var close = FindBacktickClose(text, i + ticks, end, ticks);
                if (close >= 0)
                {
                    FlushText(text, textStart, i, tokens);
                    var code = text.Substring(i + ticks, close - i - ticks);
                    // CommonMark：内容首尾各一个空格且非纯空格时去掉一个（如 ` foo ` / `` ` `` 场景）
                    if (code.Length >= 2 && code[0] == ' ' && code[^1] == ' ' && code.Trim(' ').Length > 0)
                        code = code[1..^1];
                    tokens.Add(MarkdownInline.CreateCode(code));
                    i = close + ticks;
                    textStart = i;
                    continue;
                }
            }

            // Image: ![alt](url "title") 或引用图片 ![alt][id] / ![alt][] / ![alt]
            if (ch == '!' && i + 1 < end && text[i + 1] == '[')
            {
                if (TryParseImage(text, i, end, refs, pipeline, out var imgInline, out var imgEnd))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(imgInline!);
                    i = imgEnd;
                    textStart = i;
                    continue;
                }
            }

            // Link: [text](url "title")
            if (ch == '[')
            {
                if (TryParseLink(text, i, end, out var linkInline, out var linkEnd, refs, pipeline))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(linkInline!);
                    i = linkEnd;
                    textStart = i;
                    continue;
                }
                // 引用链接 [text][id] / [text][] / [text]（shortcut）
                if (TryParseReference(text, i, end, refs, pipeline, out var refInline, out var refEnd))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(refInline!);
                    i = refEnd;
                    textStart = i;
                    continue;
                }
            }

            // 强调分隔符 run：* 或 _（收集为令牌 + 分隔符信息，交由 ResolveEmphasis 配对）
            if (ch == '*' || ch == '_')
            {
                var run = 0;
                while (i + run < end && text[i + run] == ch) run++;
                FlushText(text, textStart, i, tokens);
                var token = MarkdownInline.CreateText(new String(ch, run));
                tokens.Add(token);
                var (canOpen, canClose) = ClassifyDelimiter(text, i, run, end);
                delims ??= [];
                delims.Add(new DelimInfo
                {
                    Ch = ch,
                    Len = run,
                    Original = run,
                    CanOpen = canOpen,
                    CanClose = canClose,
                    Token = token,
                    Index = tokens.Count - 1,
                });
                i += run;
                textStart = i;
                continue;
            }

            // Strikethrough: ~~text~~（GFM：内容非空且不以空白首尾）
            if (ch == '~' && i + 1 < end && text[i + 1] == '~' && IsOn(p => p.EnableStrikethrough))
            {
                if (i + 2 < end && !Char.IsWhiteSpace(text[i + 2]))
                {
                    var closeIdx = FindStrikethroughClose(text, i + 2, end);
                    if (closeIdx > i + 2)
                    {
                        FlushText(text, textStart, i, tokens);
                        var inner = new List<MarkdownInline>();
                        ParseInlineCore(text, i + 2, closeIdx, inner, refs, pipeline, inLink);
                        tokens.Add(MarkdownInline.CreateStrikethrough(inner));
                        i = closeIdx + 2;
                        textStart = i;
                        continue;
                    }
                }
            }

            // Autolink: <url> or <email>
            if (ch == '<' && IsOn(p => p.EnableAutoLinks))
            {
                if (TryParseAutoLink(text, i, end, out var autoInline, out var autoEnd))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(autoInline!);
                    i = autoEnd;
                    textStart = i;
                    continue;
                }
                // 行内 HTML 标签：<tag> / </tag> / <!-- --> 等
                if (TryParseInlineHtml(text, i, end, out var htmlInline, out var htmlEnd))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(htmlInline!);
                    i = htmlEnd;
                    textStart = i;
                    continue;
                }
            }

            // 裸 URL 自动链接 (MD05-01): http:// / https:// / ftp:// / www.
            if (IsOn(p => p.EnableAutoLinks) && (ch == 'h' || ch == 'f' || ch == 'w'))
            {
                if (TryParseBareUrl(text, i, end, inLink, out var bareInline, out var bareEnd))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(bareInline!);
                    i = bareEnd;
                    textStart = i;
                    continue;
                }
            }

            // Email 自动链接 (MD05-01): user@domain
            // MD19：按连续 email 字符 run 处理——run 内只 IndexOf('@') 一次，
            // 避免对每个字母数字字符重复 IndexOf('@') 造成 O(n²)（英文无 @ 文本 ~6% 成本）。
            // 注意：IsEmailChar 含 _（强调标记/email 用户名下划线），run 内出现 _ 时回退原逻辑
            // （逐字符尝试），保证强调解析与含下划线 email（john_doe@x.com）兼容。
            // MD22：触发判定加 ASCII 快速路径（中文等非 ASCII 才走 Unicode 分类）
            if (IsOn(p => p.EnableAutoLinks) && (IsAsciiLetterOrDigit(ch) || ch >= 128 && Char.IsLetterOrDigit(ch)))
            {
                var runEnd2 = i + 1;
                while (runEnd2 < end && (IsEmailChar(text[runEnd2]) || text[runEnd2] == '@')) runEnd2++;
                var hasUnderscore = text.AsSpan(i, runEnd2 - i).IndexOf('_') >= 0;
                if (hasUnderscore || runEnd2 - i <= 1)
                {
                    // 含 _ 或单字符：回退原逻辑（逐字符尝试）
                    if (TryParseBareEmail(text, i, end, out var emailInline, out var emailEnd))
                    {
                        FlushText(text, textStart, i, tokens);
                        tokens.Add(emailInline!);
                        i = emailEnd;
                        textStart = i;
                        continue;
                    }
                }
                else if (text.IndexOf('@', i, runEnd2 - i) >= 0 &&
                    TryParseBareEmail(text, i, end, out var emailInline2, out var emailEnd2))
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(emailInline2!);
                    i = emailEnd2;
                    textStart = i;
                    continue;
                }
                else
                {
                    // run 内无 @ 且无下划线：整段作为普通文本，i 跳到 run 末尾（避免逐字符重复尝试）
                    i = runEnd2;
                    continue;
                }
            }

            // 行内数学公式 $...$ (MD05-04)
            if (ch == '$' && i + 1 < end && text[i + 1] != '$' && IsOn(p => p.EnableMath))
            {
                var mathEnd = text.IndexOf('$', i + 1);
                if (mathEnd > i + 1)
                {
                    FlushText(text, textStart, i, tokens);
                    tokens.Add(new MarkdownInline { Type = MarkdownInlineType.MathInline, Text = text[(i + 1)..mathEnd] });
                    i = mathEnd + 1;
                    textStart = i;
                    continue;
                }
            }

            // 脚注引用 [^id] (MD05-03)
            if (ch == '[' && i + 1 < end && text[i + 1] == '^' && IsOn(p => p.EnableFootnotes))
            {
                var closeBracket = text.IndexOf(']', i + 2);
                if (closeBracket > i + 2)
                {
                    var fnId = text[(i + 2)..closeBracket];
                    // 确保是有效脚注ID（非空且无换行）
                    if (fnId.Length > 0 && !fnId.Contains('\n') && !fnId.Contains('\r'))
                    {
                        FlushText(text, textStart, i, tokens);
                        tokens.Add(new MarkdownInline { Type = MarkdownInlineType.FootnoteRef, Text = fnId });
                        i = closeBracket + 1;
                        textStart = i;
                        continue;
                    }
                }
            }

            // Emoji 短码 :smile: (MD05-05)
            if (ch == ':' && i > textStart && i + 2 < end && IsOn(p => p.EnableEmoji))
            {
                // 快速失败（MD21）：emoji 短码 :word: 必以字母/数字/下划线开头；
                // URL 协议分隔符 ://、时间 12:30 等冒号直接跳过，避免 IndexOf/子串/字典查找
                var c1 = text[i + 1];
                if (Char.IsLetterOrDigit(c1) || c1 == '_')
                {
                    // 扫描范围限制（MD21）：code 长度 ≤30 才有效，最坏只需扫 31 字符，避免扫到段尾
                    var colonEnd = text.IndexOf(':', i + 1, Math.Min(31, end - i - 1));
                    if (colonEnd > i + 1 && colonEnd - i - 1 <= 30)
                    {
                        var code = text[(i + 1)..colonEnd];
                        if (EmojiMap.TryGetValue(code, out var emoji))
                        {
                            FlushText(text, textStart, i, tokens);
                            tokens.Add(MarkdownInline.CreateText(emoji));
                            i = colonEnd + 1;
                            textStart = i;
                            continue;
                        }
                    }
                }
            }

            i++;
        }

        FlushText(text, textStart, end, tokens);
        return delims;
    }

    private static void FlushText(String text, Int32 start, Int32 end, List<MarkdownInline> result)
    {
        if (start >= end) return;
        var s = text[start..end];
        // 硬换行前的尾部反斜杠标记不进入文本令牌（MD11）
        if (s.EndsWith("\\")) s = s[..^1];
        if (s.Length > 0)
            result.Add(MarkdownInline.CreateText(s));
    }

    private static Boolean IsEscapable(Char ch) =>
        "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~".IndexOf(ch) >= 0;

    private static Boolean TryParseImage(String text, Int32 start, Int32 end,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;
        // start points to '!'
        if (start + 1 >= end || text[start + 1] != '[') return false;

        var closeAlt = FindClosingBracket(text, start + 1, end);
        if (closeAlt < 0) return false;
        var alt = text.Substring(start + 2, closeAlt - start - 2);

        // 行内图片 ![alt](url "title")
        if (closeAlt + 1 < end && text[closeAlt + 1] == '(')
        {
            var (url, title, endParen) = ParseLinkDestination(text, closeAlt + 1, end);
            if (endParen < 0) return false;

            result = MarkdownInline.CreateImage(url, alt, title);
            newPos = endParen + 1;
            return true;
        }

        // 引用图片 ![alt][id] / ![alt][] / ![alt]（shortcut，与 TryParseReference 对称）
        if (refs == null || refs.Count == 0 || alt.Length == 0) return false;

        if (closeAlt + 1 < end && text[closeAlt + 1] == '[')
        {
            // ![alt][id] 或 ![alt][]（折叠引用）
            var close2 = FindClosingBracket(text, closeAlt + 1, end);
            if (close2 < 0) return false;
            var id2 = text.Substring(closeAlt + 2, close2 - closeAlt - 2);
            if (id2.Length > 0 && id2.Contains('^')) return false;
            var id = id2.Length > 0 ? id2 : alt;
            if (!refs.TryGetValue(id, out var def)) return false;

            result = MarkdownInline.CreateImage(def.url, alt, def.title);
            newPos = close2 + 1;
            return true;
        }

        // ![alt] shortcut 引用（仅当 label 已定义，且后不跟 (）；允许位于文本末尾
        if (closeAlt + 1 < end && text[closeAlt + 1] == '(') return false;
        if (!refs.TryGetValue(alt, out var def2)) return false;

        result = MarkdownInline.CreateImage(def2.url, alt, def2.title);
        newPos = closeAlt + 1;
        return true;
    }

    private static Boolean TryParseLink(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline)
    {
        result = null; newPos = start;
        var closeText = FindClosingBracket(text, start, end);
        if (closeText < 0 || closeText + 1 >= end || text[closeText + 1] != '(') return false;

        var linkText = text.Substring(start + 1, closeText - start - 1);
        // CommonMark §6.3：链接文本不能包含嵌套链接（图片允许）。如 [foo [bar](/uri)](/uri) 内层链接优先，外层失败
        if (HasNestedLink(linkText)) return false;

        var (url, title, endParen) = ParseLinkDestination(text, closeText + 1, end);
        if (endParen < 0) return false;

        var innerInlines = new List<MarkdownInline>();
        ParseInlineCore(linkText, 0, linkText.Length, innerInlines, refs, pipeline, true);
        result = MarkdownInline.CreateLink(url, title, innerInlines);
        newPos = endParen + 1;
        return true;
    }

    /// <summary>检测链接文本中是否包含嵌套链接（CommonMark §6.3：链接不能包含链接，图片除外）</summary>
    /// <param name="text">链接文本（不含外层 [ ]）</param>
    /// <returns>包含内联/引用链接时为 true</returns>
    private static Boolean HasNestedLink(String text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '[') continue;

            // 图片 ![alt](url) / ![alt][id]：允许嵌套，跳过整个图片
            var isImage = i > 0 && text[i - 1] == '!';
            if (isImage)
            {
                var close = FindClosingBracket(text, i, text.Length);
                if (close >= 0 && close + 1 < text.Length && text[close + 1] == '(')
                {
                    var (_, _, endParen) = ParseLinkDestination(text, close + 1, text.Length);
                    if (endParen >= 0) { i = endParen; continue; }
                }
                continue;
            }

            // [text](url) 内联链接 或 [text][id] 引用链接 → 嵌套链接，外层失败
            var close2 = FindClosingBracket(text, i, text.Length);
            if (close2 < 0) return false;
            if (close2 + 1 < text.Length && (text[close2 + 1] == '(' || text[close2 + 1] == '['))
                return true;
            // shortcut [text]：保守允许（是否引用需 refs 判定，无法精确判断）
            i = close2;
        }
        return false;
    }

    private static (String url, String title, Int32 endParen) ParseLinkDestination(String text, Int32 openParen, Int32 end)
    {
        if (openParen >= end || text[openParen] != '(') return ("", "", -1);
        var i = openParen + 1;

        // skip whitespace
        while (i < end && text[i] == ' ') i++;

        // URL: either <url> or unbracketed until space/close
        Int32 urlStart;
        Int32 urlEnd;
        if (i < end && text[i] == '<')
        {
            urlStart = i + 1;
            urlEnd = text.IndexOf('>', urlStart);
            if (urlEnd < 0) return ("", "", -1);
            i = urlEnd + 1;
        }
        else
        {
            urlStart = i;
            while (i < end && text[i] != ' ' && text[i] != ')' && text[i] != '"' && text[i] != '\'') i++;
            urlEnd = i;
        }
        var url = text[urlStart..urlEnd];

        // optional title
        while (i < end && text[i] == ' ') i++;
        var title = "";
        if (i < end && (text[i] == '"' || text[i] == '\'' || text[i] == '('))
        {
            var close = text[i] == '(' ? ')' : text[i];
            i++;
            var titleStart = i;
            // 支持反斜杠转义的引号（如 \"），扫描到未转义的闭合字符为止
            while (i < end)
            {
                if (text[i] == '\\' && i + 1 < end) { i += 2; continue; }
                if (text[i] == close) break;
                i++;
            }
            if (i < end)
            {
                title = text[titleStart..i].Replace("\\\"", "\"").Replace("\\\\", "\\");
                i++;
            }
        }

        while (i < end && text[i] == ' ') i++;
        if (i >= end || text[i] != ')') return ("", "", -1);
        return (url, title, i);
    }

    private static Int32 FindClosingBracket(String text, Int32 open, Int32 end)
    {
        var depth = 0;
        for (var i = open; i < end; i++)
        {
            if (text[i] == '[' || text[i] == '(') depth++;
            else if (text[i] == ']' || text[i] == ')')
            {
                depth--;
                if (depth == 0) return i;
            }
        }
        return -1;
    }

    /// <summary>查找指定长度的反引号串闭合位置（行内代码，CommonMark 兼容）</summary>
    /// <remarks>闭合反引号串必须恰好等长：run 前后均不能是反引号（CommonMark §6.4 例 344，如 `foo `` bar` 的闭合是末尾单反引号而非双反引号）</remarks>
    /// <param name="text">文本</param>
    /// <param name="start">开始搜索位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="ticks">反引号个数</param>
    /// <returns>闭合起始位置，未找到返回 -1</returns>
    private static Int32 FindBacktickClose(String text, Int32 start, Int32 end, Int32 ticks)
    {
        for (var i = start; i <= end - ticks; i++)
        {
            // 闭合 run 必须恰好等长：run 前后均不能是反引号，否则属于更长 run
            if (i > start && text[i - 1] == '`') continue;
            if (i + ticks < end && text[i + ticks] == '`') continue;
            var ok = true;
            for (var k = 0; k < ticks; k++)
            {
                if (text[i + k] != '`') { ok = false; break; }
            }
            if (ok) return i;
        }
        return -1;
    }

    /// <summary>强调分隔符信息（CommonMark 分隔符栈元素）</summary>
    /// <summary>分隔符信息（MD23：class→struct，省强调分隔符对象分配）</summary>
    private struct DelimInfo
    {
        /// <summary>分隔符字符（* 或 _）</summary>
        public Char Ch;

        /// <summary>当前剩余长度</summary>
        public Int32 Len;

        /// <summary>原始长度（用于 rule of 3）</summary>
        public Int32 Original;

        /// <summary>是否可开启强调</summary>
        public Boolean CanOpen;

        /// <summary>是否可关闭强调</summary>
        public Boolean CanClose;

        /// <summary>对应令牌（Text 节点）</summary>
        public MarkdownInline Token;

        /// <summary>令牌在 tokens 列表中的当前索引（避免 IndexOf 重复查找）</summary>
        public Int32 Index;
    }

    /// <summary>计算分隔符 run 的开闭能力（CommonMark 左/右 flanking 规则）</summary>
    /// <param name="text">文本</param>
    /// <param name="start">run 起始位置</param>
    /// <param name="len">run 长度</param>
    /// <param name="end">结束位置（不含）</param>
    /// <returns>是否可开启 / 是否可关闭</returns>
    private static (Boolean canOpen, Boolean canClose) ClassifyDelimiter(String text, Int32 start, Int32 len, Int32 end)
    {
        var ch = text[start];
        var before = start > 0 ? text[start - 1] : '\0';
        var after = start + len < end ? text[start + len] : '\0';

        // 行首/行尾视为空白（CommonMark 规范）
        var beforeIsWs = before == '\0' || Char.IsWhiteSpace(before);
        var afterIsWs = after == '\0' || Char.IsWhiteSpace(after);
        var beforeIsPunct = before != '\0' && IsPunctuation(before);
        var afterIsPunct = after != '\0' && IsPunctuation(after);

        // 左 flanking：后非空白，且（后非标点 或 前为空白/标点）
        var leftFlanking = !afterIsWs && (!afterIsPunct || beforeIsWs || beforeIsPunct);
        // 右 flanking：前非空白，且（前非标点 或 后为空白/标点）
        var rightFlanking = !beforeIsWs && (!beforeIsPunct || afterIsWs || afterIsPunct);

        if (ch == '*')
            return (leftFlanking, rightFlanking);

        // '_'：可开要求前非 ASCII 字母数字（防 snake_case）；可闭要求后非 ASCII 字母数字
        return (leftFlanking && !IsAsciiAlnum(before), rightFlanking && !IsAsciiAlnum(after));
    }

    /// <summary>按 CommonMark 分隔符栈算法解析强调（flanking + rule of 3 + 嵌套）</summary>
    /// <param name="tokens">令牌列表（原地修改）</param>
    /// <param name="delims">分隔符信息列表</param>
    private static void ResolveEmphasis(List<MarkdownInline> tokens, List<DelimInfo>? delims)
    {
        // 无强调分隔符（MD19 延迟创建，delims 为 null）时直接返回
        if (delims == null || delims.Count == 0) return;
        var ci = 0;
        while (ci < delims.Count)
        {
            var closer = delims[ci];
            if (closer.Len <= 0 || (closer.Ch != '*' && closer.Ch != '_') || !closer.CanClose)
            {
                ci++;
                continue;
            }

            // 向前查找最近的同字符可开分隔符（CommonMark：跳过已耗尽与被 rule-of-3 排除者）
            Int32 oi = ci - 1;
            DelimInfo? opener = null;
            for (; oi >= 0; oi--)
            {
                var d = delims[oi];
                if (d.Len <= 0 || d.Ch != closer.Ch || !d.CanOpen) continue;

                // Rule of 3：若开或闭分隔符可同时开与闭，则两 run 长度和不能为 3 的倍数（除非两者都是 3 的倍数）
                if ((d.CanOpen && d.CanClose) || (closer.CanOpen && closer.CanClose))
                {
                    var sum = d.Original + closer.Original;
                    if (sum % 3 == 0 && !(d.Original % 3 == 0 && closer.Original % 3 == 0))
                        continue;
                }
                opener = d;
                break;
            }
            if (!opener.HasValue)
            {
                ci++;
                continue;
            }
            // MD23：struct 副本，所有修改最后统一写回 delims[oi]/delims[ci]
            var o = opener.Value;

            // 使用字符数：两者皆为 3 的倍数用 2，长度和为 3 的倍数用 1，其余用 2
            var use = (o.Original % 3 == 0 && closer.Original % 3 == 0)
                ? 2
                : (o.Original + closer.Original) % 3 == 0 ? 1 : 2;
            use = Math.Min(use, Math.Min(closer.Len, o.Len));
            if (use <= 0)
            {
                ci++;
                continue;
            }

            var oi2 = o.Index;
            var ci2 = closer.Index;
            if (oi2 < 0 || ci2 < 0 || oi2 >= ci2)
            {
                ci++;
                continue;
            }

            // 收集中间内容（含嵌套令牌）
            var children = new List<MarkdownInline>();
            for (var k = oi2 + 1; k < ci2; k++) children.Add(tokens[k]);

            // 创建强调/粗体节点
            MarkdownInline emph = use == 2
                ? MarkdownInline.CreateStrong(children)
                : MarkdownInline.CreateEmphasis(children);

            // 重建 [oi2..ci2] 区间：剩余 opener 字符 + 强调节点 + 剩余 closer 字符
            var openerLeftover = o.Len - use;
            var closerLeftover = closer.Len - use;
            var newItems = new List<MarkdownInline>();
            if (openerLeftover > 0)
            {
                var t = MarkdownInline.CreateText(new String(o.Ch, openerLeftover));
                o.Token = t;
                o.Index = oi2;
                newItems.Add(t);
            }
            newItems.Add(emph);
            if (closerLeftover > 0)
            {
                var t2 = MarkdownInline.CreateText(new String(closer.Ch, closerLeftover));
                closer.Token = t2;
                closer.Index = oi2 + newItems.Count;
                newItems.Add(t2);
            }

            // 失活 opener 与 closer 之间的分隔符（进入强调内部，CommonMark 规范不可再参与配对）
            for (var k = oi + 1; k < ci; k++)
            {
                var dk = delims[k];
                dk.Len = 0;
                delims[k] = dk;
            }

            // 更新 tokens 并维护索引（oi2 之后的分隔符整体平移 delta）
            var removedCount = ci2 - oi2 + 1;
            var delta = newItems.Count - removedCount;
            tokens.RemoveRange(oi2, removedCount);
            tokens.InsertRange(oi2, newItems);
            for (var k = ci + 1; k < delims.Count; k++)
            {
                var dk = delims[k];
                dk.Index += delta;
                delims[k] = dk;
            }

            o.Len = openerLeftover;
            closer.Len = closerLeftover;
            delims[oi] = o;
            delims[ci] = closer;

            // closer 仍有剩余 → 回退一格继续匹配同一 closer（如 ***foo*** 先粗后斜）
            if (closer.Len > 0) ci--;
            ci++;
        }
    }

    /// <summary>判断字符是否为标点（CommonMark：ASCII 标点 + Unicode P* 类别）</summary>
    /// <param name="c">字符</param>
    /// <returns>是否标点</returns>
    private static Boolean IsPunctuation(Char c)
    {
        if (c is '!' or '"' or '#' or '$' or '%' or '&' or '\'' or '(' or ')' or '*' or '+' or ',' or '-' or '.' or '/' or
            ':' or ';' or '<' or '=' or '>' or '?' or '@' or '[' or '\\' or ']' or '^' or '_' or '`' or '{' or '|' or '}' or '~')
            return true;
        var cat = Char.GetUnicodeCategory(c);
        return cat is UnicodeCategory.ConnectorPunctuation or UnicodeCategory.DashPunctuation
            or UnicodeCategory.OpenPunctuation or UnicodeCategory.ClosePunctuation
            or UnicodeCategory.InitialQuotePunctuation or UnicodeCategory.FinalQuotePunctuation
            or UnicodeCategory.OtherPunctuation;
    }

    /// <summary>判断字符是否为 ASCII 字母数字</summary>
    /// <param name="c">字符</param>
    /// <returns>是否 ASCII 字母数字</returns>
    private static Boolean IsAsciiAlnum(Char c) =>
        c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9';

    /// <summary>查找删除线闭合标记（GFM：恰好两个波浪线、且内容不以空白结尾）</summary>
    /// <param name="text">文本</param>
    /// <param name="start">开始搜索位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <returns>闭合起始位置，未找到返回 -1</returns>
    private static Int32 FindStrikethroughClose(String text, Int32 start, Int32 end)
    {
        var i = start;
        while (i < end - 1)
        {
            if (text[i] == '~')
            {
                var run = 0;
                while (i + run < end && text[i + run] == '~') run++;
                // 恰好 2 个波浪线且前一字符非空白（内容不以空白结尾）
                if (run == 2 && i > start && !Char.IsWhiteSpace(text[i - 1]))
                    return i;
                i += run;
                continue;
            }
            i++;
        }
        return -1;
    }

    private static Boolean TryParseAutoLink(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;
        var closeAngle = text.IndexOf('>', start + 1);
        if (closeAngle < 0 || closeAngle >= end) return false;
        var inner = text.Substring(start + 1, closeAngle - start - 1);
        if (inner.Length == 0) return false;

        // URI autolink：<scheme:...>，scheme 为 2-32 位字母数字（CommonMark）
        var colonIdx = inner.IndexOf(':');
        if (colonIdx > 0 && colonIdx <= 32 && IsValidScheme(inner[..colonIdx]))
        {
            var children = new List<MarkdownInline> { MarkdownInline.CreateText(inner) };
            result = MarkdownInline.CreateLink(inner, "", children);
            newPos = closeAngle + 1;
            return true;
        }

        // Email autolink：<user@domain>
        var atIdx = inner.IndexOf('@');
        if (atIdx > 0 && atIdx < inner.Length - 1 &&
            IsValidEmail(inner[..atIdx], inner[(atIdx + 1)..]))
        {
            var children = new List<MarkdownInline> { MarkdownInline.CreateText(inner) };
            result = MarkdownInline.CreateLink("mailto:" + inner, "", children);
            newPos = closeAngle + 1;
            return true;
        }
        return false;
    }

    /// <summary>校验 URI scheme（CommonMark：2-32 位，字母开头，字母数字+.-）</summary>
    /// <param name="scheme">scheme 部分</param>
    /// <returns>是否合法</returns>
    private static Boolean IsValidScheme(String scheme)
    {
        if (scheme.Length < 2 || scheme.Length > 32) return false;
        if (!Char.IsLetter(scheme[0])) return false;
        for (var i = 1; i < scheme.Length; i++)
        {
            var c = scheme[i];
            if (!(Char.IsLetterOrDigit(c) || c is '+' or '.' or '-')) return false;
        }
        return true;
    }

    /// <summary>校验邮箱 user@domain（CommonMark 简化规则）</summary>
    /// <param name="user">@ 前部分</param>
    /// <param name="domain">@ 后部分</param>
    /// <returns>是否合法</returns>
    private static Boolean IsValidEmail(String user, String domain)
    {
        if (user.Length == 0 || domain.Length == 0) return false;
        for (var i = 0; i < user.Length; i++)
        {
            var c = user[i];
            if (!(Char.IsLetterOrDigit(c) || "!#$%&'*+/=?^_`{|}~.-".IndexOf(c) >= 0)) return false;
        }
        if (!Char.IsLetterOrDigit(domain[0]) || !Char.IsLetterOrDigit(domain[^1])) return false;
        for (var i = 0; i < domain.Length; i++)
        {
            var c = domain[i];
            if (!(Char.IsLetterOrDigit(c) || c is '-' or '.')) return false;
        }
        return true;
    }

    /// <summary>尝试解析裸 URL（http/https/ftp/www）自动链接 (MD05-01, GFM)</summary>
    private static Boolean TryParseBareUrl(String text, Int32 start, Int32 end, Boolean inLink,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;

        // GFM：链接内部不识别裸 URL
        if (inLink) return false;

        // 检查是否是 URL 前缀（Span 前缀比较，避免 text[start..end] 子串分配，MD17；前缀静态化避免每次 new 数组，MD19）
        var span = text.AsSpan(start, end - start);
        String? matchedPrefix = null;
        foreach (var prefix in _urlPrefixes)
        {
            if (span.StartsWith(prefix.AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                matchedPrefix = prefix;
                break;
            }
        }
        if (matchedPrefix == null) return false;

        // GFM：www. 链接前不能是字母数字字符
        if (matchedPrefix == "www." && start > 0 &&
            (Char.IsLetterOrDigit(text[start - 1]) || text[start - 1] == '_'))
            return false;

        // 收集 URL 字符直到空白或边界；支持括号平衡（GFM）
        var urlEnd = start;
        var parenDepth = 0;
        while (urlEnd < end)
        {
            var c = text[urlEnd];
            if (Char.IsWhiteSpace(c) || c == '<' || c == '>' || c == '"' || c == '\'') break;
            if (c == '(') { parenDepth++; urlEnd++; continue; }
            if (c == ')' || c == ']' || c == '}')
            {
                // 未闭合的右括号/方括号/花括号 → 结束 URL（不含该字符）
                if (parenDepth == 0) break;
                if (c == ')') parenDepth--;
                urlEnd++;
                continue;
            }
            // 末尾标点不包含在 URL 中（后跟空白/结束）
            if (c is '.' or ',' or ';' or ':' or '!' or '?' or '*' or '_' or '~' &&
                urlEnd + 1 < end && (Char.IsWhiteSpace(text[urlEnd + 1]) || text[urlEnd + 1] == '<'))
                break;
            urlEnd++;
        }

        // 去除末尾未闭合开括号与标点
        while (urlEnd > start)
        {
            var c = text[urlEnd - 1];
            if (c == '(' || c is '.' or ',' or ';' or ':' or '!' or '?' or '*' or '_' or '~')
                urlEnd--;
            else
                break;
        }

        if (urlEnd <= start + matchedPrefix.Length) return false;

        var url = text[start..urlEnd];

        // GFM：www. 需包含点号的域名（点不能是末尾）
        if (matchedPrefix == "www.")
        {
            var dotIdx = url.IndexOf('.', 4);
            if (dotIdx < 0 || dotIdx == url.Length - 1) return false;
        }

        var href = matchedPrefix == "www." ? "http://" + url : url;
        var children = new List<MarkdownInline> { MarkdownInline.CreateText(url) };
        result = MarkdownInline.CreateLink(href, "", children);
        newPos = urlEnd;
        return true;
    }

    /// <summary>尝试解析裸 Email 自动链接 (MD05-01)</summary>
    private static Boolean TryParseBareEmail(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;

        // 检查周围是否有 @ 符号
        var atIdx = text.IndexOf('@', start);
        if (atIdx < 0 || atIdx >= end) return false;

        // 找到 email 的起始位置（回溯到单词边界）
        var emailStart = atIdx;
        while (emailStart > start && IsEmailChar(text[emailStart - 1]))
            emailStart--;

        // 检查 @ 前后是否都有有效字符
        if (emailStart >= atIdx || atIdx + 1 >= end) return false;

        // 找到 email 的结束位置
        var emailEnd = atIdx + 1;
        while (emailEnd < end && IsEmailChar(text[emailEnd]))
            emailEnd++;

        if (emailEnd <= atIdx + 1) return false;

        var email = text[emailStart..emailEnd];
        // 基本验证
        if (!email.Contains('.') || email.StartsWith("@") || email.EndsWith("@")) return false;

        result = MarkdownInline.CreateLink("mailto:" + email, "", [MarkdownInline.CreateText(email)]);
        newPos = emailEnd;
        return true;
    }

    private static Boolean IsEmailChar(Char c)
    {
        // ASCII 快速路径（MD22）：字母数字与常用符号直接位运算，避免 Unicode 分类调用；
        // 中文等非 ASCII 字符才走 Char.IsLetterOrDigit（中文是 Unicode Letter）
        if (c < 128)
        {
            return c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9'
                or '.' or '-' or '_' or '+' or '%';
        }
        return Char.IsLetterOrDigit(c);
    }

    /// <summary>判断 ASCII 字母数字（email 触发快速路径，MD22）</summary>
    /// <param name="c">字符</param>
    /// <returns>是否为 ASCII 字母或数字</returns>
    private static Boolean IsAsciiLetterOrDigit(Char c) =>
        c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9';

    /// <summary>尝试解析引用链接 [text][id] / [text][] / [text]（CommonMark）</summary>
    /// <param name="text">文本</param>
    /// <param name="start">开始位置（指向 [）</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="refs">引用定义字典（可为 null）</param>
    /// <param name="pipeline">处理管线（可为 null）</param>
    /// <param name="result">解析结果</param>
    /// <param name="newPos">新位置</param>
    /// <returns>是否解析成功</returns>
    private static Boolean TryParseReference(String text, Int32 start, Int32 end,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;
        if (refs == null || refs.Count == 0) return false;

        // [label]
        var close1 = FindClosingBracket(text, start, end);
        if (close1 < 0) return false;
        var label = text.Substring(start + 1, close1 - start - 1);
        if (label.Length == 0 || label.Contains('^')) return false;

        var after = close1 + 1;
        if (after < end && text[after] == '[')
        {
            // [label][id] 或 [label][]（折叠引用）
            var close2 = FindClosingBracket(text, after, end);
            if (close2 < 0) return false;
            var id2 = text.Substring(after + 1, close2 - after - 1);
            if (id2.Length > 0 && id2.Contains('^')) return false;
            var id = id2.Length > 0 ? id2 : label;
            if (!refs.TryGetValue(id, out var def)) return false;

            var innerInlines = new List<MarkdownInline>();
            ParseInlineCore(text, start + 1, close1, innerInlines, refs, pipeline, true);
            result = MarkdownInline.CreateLink(def.url, def.title, innerInlines);
            newPos = close2 + 1;
            return true;
        }

        // [label] shortcut 引用（仅当 label 已定义，且后不跟 (）
        if (after < end && text[after] == '(') return false;
        if (!refs.TryGetValue(label, out var def2)) return false;

        var inner2 = new List<MarkdownInline>();
        ParseInlineCore(text, start + 1, close1, inner2, refs, pipeline, true);
        result = MarkdownInline.CreateLink(def2.url, def2.title, inner2);
        newPos = close1 + 1;
        return true;
    }

    /// <summary>尝试解析行内 HTML 标签：&lt;tag&gt; / &lt;/tag&gt; / &lt;!-- --&gt; 等</summary>
    /// <param name="text">文本</param>
    /// <param name="start">开始位置（指向 &lt;）</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="result">解析结果</param>
    /// <param name="newPos">新位置</param>
    /// <returns>是否解析成功</returns>
    private static Boolean TryParseInlineHtml(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;

        // 注释 <!-- ... -->
        if (start + 4 <= end && text.Substring(start, 4) == "<!--")
        {
            var closeComment = text.IndexOf("-->", start + 4);
            if (closeComment >= 0 && closeComment < end)
            {
                result = MarkdownInline.CreateRawHtml(text.Substring(start, closeComment + 3 - start));
                newPos = closeComment + 3;
                return true;
            }
            return false;
        }

        // 起始标签 <tag 或结束标签 </tag
        var isClose = start + 1 < end && text[start + 1] == '/';
        var nameStart = start + (isClose ? 2 : 1);
        if (nameStart >= end || !Char.IsLetter(text[nameStart])) return false;

        var nameEnd = nameStart;
        while (nameEnd < end && (Char.IsLetterOrDigit(text[nameEnd]) || text[nameEnd] == '-' || text[nameEnd] == ':'))
            nameEnd++;

        var close = text.IndexOf('>', nameEnd);
        if (close < 0 || close >= end) return false;

        // 属性区不允许包含嵌套 < >
        for (var k = nameEnd; k < close; k++)
        {
            if (text[k] == '<' || text[k] == '>') return false;
        }

        result = MarkdownInline.CreateRawHtml(text.Substring(start, close - start + 1));
        newPos = close + 1;
        return true;
    }

    /// <summary>解码单个 HTML 实体（含数字字符引用），未识别返回 null</summary>
    /// <param name="entity">实体名（不含 &amp; 与 ;）</param>
    /// <returns>解码后字符串，未识别返回 null</returns>
    /// <remarks>委托 <see cref="EntityDecoder"/> 共享实现：具名实体全表（HTML5）+ 数字引用含补充平面代理对</remarks>
    private static String? DecodeEntity(String entity) => EntityDecoder.Decode(entity);

    /// <summary>预扫描引用链接定义 [id]: url "title"（支持定义在后使用在前）</summary>
    /// <param name="doc">文档对象（填充 References）</param>
    private void PreScanReferences(MarkdownDocument doc)
    {
        for (var i = 0; i < _lines.Length; i++)
        {
            var line = _lines[i];
            if (line.Length == 0) continue;
            // 引用定义必须以 [ 开头（CommonMark 允许最多 3 空格缩进；4+ 缩进是代码块）。
            // 快速排除其余行，避免逐行 Trim + 正则匹配（普通文档绝大多数行非引用定义）
            var first = line[0];
            if (first != '[' && first != ' ' && first != '\t') continue;
            var indent = GetIndent(line);
            if (indent >= 4 || indent >= line.Length || line[indent] != '[') continue;
            var trimmed = line.Trim();
            // 排除脚注定义 [^id]:
            if (trimmed.StartsWith("[^") || trimmed.Length > 1 && trimmed[1] == '^') continue;
            var match = _referenceDef.Match(trimmed);
            if (!match.Success) continue;

            var id = match.Groups[1].Value;
            if (id.Length == 0) continue;
            var url = match.Groups[2].Success ? match.Groups[2].Value : match.Groups[3].Value;
            var title = "";
            var rest = match.Groups[4].Value.Trim();
            if (rest.Length >= 2 &&
                ((rest[0] == '"' && rest[^1] == '"') ||
                 (rest[0] == '\'' && rest[^1] == '\'') ||
                 (rest[0] == '(' && rest[^1] == ')')))
            {
                title = rest[1..^1];
            }
            else if (rest.Length > 0)
            {
                // 非引号结尾的多余内容视为非法定义
                continue;
            }

            _refs[id] = (url, title);
            doc.References[id] = url;
            _skipLines.Add(i);
            // 捕获原始行（含标题与格式）及其后连续空行，供往返渲染原位输出（MD06-03）
            var src = new System.Text.StringBuilder();
            src.Append(_lines[i]).Append('\n');
            var j = i + 1;
            while (j < _lines.Length && String.IsNullOrWhiteSpace(_lines[j]))
            {
                src.Append('\n');
                j++;
            }
            doc.ReferenceSourceLines.Add(src.ToString());
            doc.ReferenceLineIndexes.Add(i);
        }
    }
    #endregion
}
