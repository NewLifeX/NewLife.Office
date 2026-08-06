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
    #endregion

    #region 入口
    /// <summary>解析 Markdown 文本，返回文档对象</summary>
    /// <param name="text">Markdown 文本</param>
    /// <returns>解析后的文档</returns>
    public MarkdownDocument Parse(String text)
    {
        // 统一换行符，展开 Tab（CommonMark 规范：Tab = 4 空格）
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        _lines = text.Split('\n');
        _pos = 0;

        var doc = new MarkdownDocument();
        _doc = doc;

        // YAML Front Matter (MD05-02)
        if (Pipeline == null || Pipeline.EnableFrontMatter)
            ParseFrontMatter(doc);

        // 预扫描引用链接定义（支持定义在后使用在前）
        PreScanReferences(doc);

        while (_pos < _lines.Length)
        {
            // 跳过引用链接定义行（已在预扫描中提取）
            if (_skipLines.Contains(_pos)) { _pos++; continue; }

            var blockStart = _pos;
            var blockCount = doc.Blocks.Count;
            ParseBlock(doc.Blocks);

            // 设置新生成块的源码位置（MD06-04）并捕获原始文本（MD06-03）
            for (var bi = blockCount; bi < doc.Blocks.Count; bi++)
            {
                var block = doc.Blocks[bi];
                block.SourceLine = blockStart;
                block.SourceColumn = 0;

                // 捕获原始源码文本用于往返渲染（跳过前导空白行，含尾部空行以保留块间分隔）
                var srcStart = blockStart;
                while (srcStart < _pos && String.IsNullOrWhiteSpace(_lines[srcStart]))
                    srcStart++;
                var srcEnd = _pos;
                while (srcEnd < _lines.Length && String.IsNullOrWhiteSpace(_lines[srcEnd]))
                    srcEnd++;
                if (srcEnd > srcStart)
                {
                    var lineCount = srcEnd - srcStart;
                    var rawLines = new String[lineCount];
                    Array.Copy(_lines, srcStart, rawLines, 0, lineCount);
                    // 拼接时补上末尾换行，保留块间空行（含被捕获的尾部空行）
                    block.SourceText = String.Join("\n", rawLines) + "\n";
                    // 记录规范序列化快照，用于往返渲染时检测块是否被修改
                    block.SourceSnapshot = MarkdownWriter.SerializeSnapshot(block);
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
        var match = Regex.Match(trimmed, @"^(#{1,6})(\s+|$)(.*)$");
        if (!match.Success) return null;
        var hashes = match.Groups[1].Value;
        // trailing #? remove it
        var text = match.Groups[3].Value.TrimEnd();
        text = Regex.Replace(text, @"\s+#+\s*$", "").TrimEnd();
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
        // Re-parse quoted content
        var inner = new MarkdownParser { _refs = _refs, Pipeline = Pipeline };
        var innerDoc = inner.Parse(String.Join("\n", quotedLines));
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
        var match = Regex.Match(trimmed, @"^(\d{1,9})([.)]) ");
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
                        // 列表延续：跳过空行，空行作为项内段落分隔
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

        if (!ordered) return MarkdownBlock.CreateBulletList(items);
        return MarkdownBlock.CreateOrderedList(items, firstNum);
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

        // 重新解析为块，支持嵌套列表/引用/代码块/多段落等
        var inner = new MarkdownParser
        {
            _listDepth = _listDepth + 1,
            _refs = _refs,
            Pipeline = Pipeline,
        };
        var innerDoc = inner.Parse(String.Join("\n", contentLines));

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
        var m = Regex.Match(content, @"^\[([ xX])\] (.*)$");
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
            if (lines.Count > 0 && Regex.IsMatch(trimmedNext, @"^[=]+\s*$"))
            {
                _pos++;
                var headingText = String.Join(" ", lines).Trim();
                return MarkdownBlock.CreateHeading(1, ParseInlineWithRefs(headingText));
            }
            if (lines.Count > 0 && Regex.IsMatch(trimmedNext, @"^[-]+\s*$"))
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
            // List markers interrupt (after first line)
            if (lines.Count > 0 && (IsBulletListMarker(trimmedNext, out var _)
                || IsOrderedListMarker(trimmedNext, out var _, out var _))) break;

            // Hard break: trailing 2+ spaces before newline
            if (line.EndsWith("  ") && lines.Count > 0)
                lines.Add(line.TrimEnd() + "  ");
            else
                lines.Add(line);
            _pos++;
        }

        if (lines.Count == 0) return MarkdownBlock.CreateParagraph([]);

        // Check if it looks like a GFM table
        if (lines.Count >= 2 && (Pipeline == null || Pipeline.EnableTables) &&
            (lines[1].TrimStart().StartsWith("|") || Regex.IsMatch(lines[1].Trim(), @"^[|\-: ]+$")))
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
        if (!sep.Contains("---") && !sep.Contains(":--") && !sep.Contains("--:") && !sep.Contains(":-:"))
            return null;

        // Parse alignments
        var alignCells = SplitTableRow(sep);
        var alignments = new List<String>();
        foreach (var cell in alignCells)
        {
            var c = cell.Trim();
            if (c.StartsWith(":") && c.EndsWith(":")) alignments.Add("center");
            else if (c.EndsWith(":")) alignments.Add("right");
            else if (c.StartsWith(":")) alignments.Add("left");
            else alignments.Add("");
        }

        var table = new TableBlock();

        // Header row
        var headerRow = new TableRowBlock();
        var headerCells = SplitTableRow(lines[0]);
        for (var i = 0; i < headerCells.Count; i++)
        {
            var cell = new TableCellBlock(
                ParseInlineWithRefs(headerCells[i].Trim()),
                isHeader: true,
                alignment: i < alignments.Count ? alignments[i] : "");
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
                    alignment: i < alignments.Count ? alignments[i] : "");
                dataRow.Children.Add(cell);
            }
            table.Children.Add(dataRow);
        }

        return table;
    }

    /// <summary>拆分表格行，支持 \| 转义管道符（GFM）</summary>
    /// <param name="line">表格行文本</param>
    /// <returns>单元格列表</returns>
    private static List<String> SplitTableRow(String line)
    {
        line = line.Trim();
        if (line.StartsWith("|")) line = line[1..];
        if (line.EndsWith("|")) line = line[..^1];

        var cells = new List<String>();
        var sb = new System.Text.StringBuilder();
        var slashCount = 0;
        foreach (var c in line)
        {
            if (c == '\\') { sb.Append(c); slashCount++; continue; }
            if (c == '|')
            {
                // 奇数个连续反斜杠 → 转义管道符，去掉转义用反斜杠
                if (slashCount % 2 == 1)
                {
                    sb.Remove(sb.Length - 1, 1);
                    sb.Append('|');
                }
                else
                {
                    cells.Add(sb.ToString());
                    sb.Clear();
                }
                slashCount = 0;
                continue;
            }
            slashCount = 0;
            sb.Append(c);
        }
        cells.Add(sb.ToString());
        return cells;
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

                // 去掉引号
                if (value.Length >= 2 &&
                    ((value[0] == '"' && value[^1] == '"') ||
                     (value[0] == '\'' && value[^1] == '\'')))
                    value = value[1..^1];
                if (key.Length > 0 && value.Length > 0)
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

    private Boolean TryParseFootnoteDefinition(String trimmed, out MarkdownBlock? block)
    {
        block = null;

        // [^id]: text
        var match = Regex.Match(trimmed, @"^\[\^([^\]]+)\]:\s*(.*)$");
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
        var match = Regex.Match(trimmed, @"^\*\[([^\]]+)\]:\s*(.*)$");
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
        ParseInlineCore(text, 0, text.Length, result, null, null);
        return result;
    }

    /// <summary>解析行内 Markdown 内容（携带引用链接定义上下文）</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <returns>行内元素列表</returns>
    private List<MarkdownInline> ParseInlineWithRefs(String text)
    {
        var result = new List<MarkdownInline>();
        if (String.IsNullOrEmpty(text)) return result;
        ParseInlineCore(text, 0, text.Length, result, _refs, Pipeline);
        return result;
    }

    /// <summary>解析行内 Markdown 内容（核心递归实现）</summary>
    /// <param name="text">行内文本（不含换行）</param>
    /// <param name="start">开始位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="result">输出列表</param>
    /// <param name="refs">引用链接定义（可为 null）</param>
    /// <param name="pipeline">处理管线（可为 null，视为全开）</param>
    private static void ParseInlineCore(String text, Int32 start, Int32 end, List<MarkdownInline> result,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline)
    {
        var i = start;
        var textStart = start;

        // 管线开关：null 视为全开
        Boolean IsOn(Func<MarkdownPipeline, Boolean> get) => pipeline == null || get(pipeline);

        while (i < end)
        {
            var ch = text[i];

            // Hard break: "  \n" or "\\\n"
            if (ch == '\n')
            {
                FlushText(text, textStart, i, result);
                var isHard = i >= 2 && text[i - 1] == ' ' && text[i - 2] == ' ';
                var isBackslashHard = i >= 1 && text[i - 1] == '\\';
                if (isHard || isBackslashHard)
                    result.Add(MarkdownInline.CreateHardBreak());
                else
                    result.Add(MarkdownInline.CreateSoftBreak());
                i++;
                textStart = i;
                continue;
            }

            // Escape
            if (ch == '\\' && i + 1 < end && IsEscapable(text[i + 1]))
            {
                FlushText(text, textStart, i, result);
                result.Add(MarkdownInline.CreateText(text[i + 1].ToString()));
                i += 2;
                textStart = i;
                continue;
            }

            // HTML 实体解码：&amp; / &lt; / &#65; 等
            if (ch == '&')
            {
                var semi = text.IndexOf(';', i + 1);
                if (semi > i + 1 && semi - i <= 12)
                {
                    var entity = text.Substring(i + 1, semi - i - 1);
                    var decoded = DecodeEntity(entity);
                    if (decoded != null)
                    {
                        FlushText(text, textStart, i, result);
                        result.Add(MarkdownInline.CreateText(decoded));
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
                    FlushText(text, textStart, i, result);
                    var code = text.Substring(i + ticks, close - i - ticks);
                    // CommonMark：内容首尾各一个空格且非纯空格时去掉一个（如 ` foo ` / `` ` `` 场景）
                    if (code.Length >= 2 && code[0] == ' ' && code[^1] == ' ' && code.Trim(' ').Length > 0)
                        code = code[1..^1];
                    result.Add(MarkdownInline.CreateCode(code));
                    i = close + ticks;
                    textStart = i;
                    continue;
                }
            }

            // Image: ![alt](url "title")
            if (ch == '!' && i + 1 < end && text[i + 1] == '[')
            {
                if (TryParseImage(text, i, end, out var imgInline, out var imgEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(imgInline!);
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
                    FlushText(text, textStart, i, result);
                    result.Add(linkInline!);
                    i = linkEnd;
                    textStart = i;
                    continue;
                }
                // 引用链接 [text][id] / [text][] / [text]（shortcut）
                if (TryParseReference(text, i, end, refs, pipeline, out var refInline, out var refEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(refInline!);
                    i = refEnd;
                    textStart = i;
                    continue;
                }
            }

            // Strong+Emphasis ***
            if (ch == '*' || ch == '_')
            {
                if (TryParseEmphasis(text, i, end, out var emphInline, out var emphEnd, refs, pipeline))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(emphInline!);
                    i = emphEnd;
                    textStart = i;
                    continue;
                }
            }

            // Strikethrough: ~~text~~
            if (ch == '~' && i + 1 < end && text[i + 1] == '~' && IsOn(p => p.EnableStrikethrough))
            {
                var closeIdx = text.IndexOf("~~", i + 2, StringComparison.Ordinal);
                if (closeIdx > i + 1)
                {
                    FlushText(text, textStart, i, result);
                    var inner = new List<MarkdownInline>();
                    ParseInlineCore(text, i + 2, closeIdx, inner, refs, pipeline);
                    result.Add(MarkdownInline.CreateStrikethrough(inner));
                    i = closeIdx + 2;
                    textStart = i;
                    continue;
                }
            }

            // Autolink: <url> or <email>
            if (ch == '<' && IsOn(p => p.EnableAutoLinks))
            {
                if (TryParseAutoLink(text, i, end, out var autoInline, out var autoEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(autoInline!);
                    i = autoEnd;
                    textStart = i;
                    continue;
                }
                // 行内 HTML 标签：<tag> / </tag> / <!-- --> 等
                if (TryParseInlineHtml(text, i, end, out var htmlInline, out var htmlEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(htmlInline!);
                    i = htmlEnd;
                    textStart = i;
                    continue;
                }
            }

            // 裸 URL 自动链接 (MD05-01): http:// / https:// / ftp://
            if (IsOn(p => p.EnableAutoLinks) && (ch == 'h' || ch == 'f' || ch == 'w'))
            {
                if (TryParseBareUrl(text, i, end, out var bareInline, out var bareEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(bareInline!);
                    i = bareEnd;
                    textStart = i;
                    continue;
                }
            }

            // Email 自动链接 (MD05-01): user@domain
            if (IsOn(p => p.EnableAutoLinks) && Char.IsLetterOrDigit(ch))
            {
                if (TryParseBareEmail(text, i, end, out var emailInline, out var emailEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(emailInline!);
                    i = emailEnd;
                    textStart = i;
                    continue;
                }
            }

            // 行内数学公式 $...$ (MD05-04)
            if (ch == '$' && i + 1 < end && text[i + 1] != '$' && IsOn(p => p.EnableMath))
            {
                var mathEnd = text.IndexOf('$', i + 1);
                if (mathEnd > i + 1)
                {
                    FlushText(text, textStart, i, result);
                    result.Add(new MarkdownInline { Type = MarkdownInlineType.MathInline, Text = text[(i + 1)..mathEnd] });
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
                        FlushText(text, textStart, i, result);
                        result.Add(new MarkdownInline { Type = MarkdownInlineType.FootnoteRef, Text = fnId });
                        i = closeBracket + 1;
                        textStart = i;
                        continue;
                    }
                }
            }

            // Emoji 短码 :smile: (MD05-05)
            if (ch == ':' && i > textStart && i + 2 < end && IsOn(p => p.EnableEmoji))
            {
                var colonEnd = text.IndexOf(':', i + 1);
                if (colonEnd > i + 1 && colonEnd - i - 1 <= 30)
                {
                    var code = text[(i + 1)..colonEnd];
                    if (EmojiMap.TryGetValue(code, out var emoji))
                    {
                        FlushText(text, textStart, i, result);
                        result.Add(MarkdownInline.CreateText(emoji));
                        i = colonEnd + 1;
                        textStart = i;
                        continue;
                    }
                }
            }

            i++;
        }

        FlushText(text, textStart, end, result);
    }

    private static void FlushText(String text, Int32 start, Int32 end, List<MarkdownInline> result)
    {
        if (start >= end) return;
        var s = text[start..end];
        // Remove trailing backslash before hard break
        if (s.EndsWith("\\")) s = s[..^1];
        if (s.Length > 0)
            result.Add(MarkdownInline.CreateText(s));
    }

    private static Boolean IsEscapable(Char ch) =>
        "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~".IndexOf(ch) >= 0;

    private static Boolean TryParseImage(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;
        // start points to '!'
        if (start + 1 >= end || text[start + 1] != '[') return false;

        var closeAlt = FindClosingBracket(text, start + 1, end);
        if (closeAlt < 0 || closeAlt + 1 >= end || text[closeAlt + 1] != '(') return false;

        var alt = text.Substring(start + 2, closeAlt - start - 2);
        var (url, title, endParen) = ParseLinkDestination(text, closeAlt + 1, end);
        if (endParen < 0) return false;

        result = MarkdownInline.CreateImage(url, alt, title);
        newPos = endParen + 1;
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
        var (url, title, endParen) = ParseLinkDestination(text, closeText + 1, end);
        if (endParen < 0) return false;

        var innerInlines = new List<MarkdownInline>();
        ParseInlineCore(linkText, 0, linkText.Length, innerInlines, refs, pipeline);
        result = MarkdownInline.CreateLink(url, title, innerInlines);
        newPos = endParen + 1;
        return true;
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
    /// <param name="text">文本</param>
    /// <param name="start">开始搜索位置</param>
    /// <param name="end">结束位置（不含）</param>
    /// <param name="ticks">反引号个数</param>
    /// <returns>闭合起始位置，未找到返回 -1</returns>
    private static Int32 FindBacktickClose(String text, Int32 start, Int32 end, Int32 ticks)
    {
        for (var i = start; i <= end - ticks; i++)
        {
            var ok = true;
            for (var k = 0; k < ticks; k++)
            {
                if (text[i + k] != '`') { ok = false; break; }
            }
            if (ok) return i;
        }
        return -1;
    }

    private static Boolean TryParseEmphasis(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos,
        Dictionary<String, (String url, String title)>? refs, MarkdownPipeline? pipeline)
    {
        result = null; newPos = start;
        var ch = text[start];

        // Count opening markers
        var count = 0;
        var i = start;
        while (i < end && text[i] == ch) { count++; i++; }

        if (count >= 3)
        {
            // ***...*** strong+emphasis
            var closeMarker = new String(ch, 3);
            var closeIdx = FindCloseNotInCode(text, i, end, closeMarker);
            if (closeIdx >= 0)
            {
                var inner = new List<MarkdownInline>();
                ParseInlineCore(text, i, closeIdx, inner, refs, pipeline);
                result = MarkdownInline.CreateStrongEmphasis(inner);
                newPos = closeIdx + 3;
                return true;
            }
        }
        if (count >= 2)
        {
            // **...** or __...__
            var closeMarker = new String(ch, 2);
            var closeIdx = FindCloseNotInCode(text, i, end, closeMarker);
            if (closeIdx >= 0)
            {
                var inner = new List<MarkdownInline>();
                ParseInlineCore(text, i, closeIdx, inner, refs, pipeline);
                result = MarkdownInline.CreateStrong(inner);
                newPos = closeIdx + 2;
                return true;
            }
        }
        if (count >= 1)
        {
            // *...* or _..._
            if (ch == '_')
            {
                // _ only starts emphasis if not surrounded by word chars
                if (start > 0 && (Char.IsLetterOrDigit(text[start - 1]) || text[start - 1] == '_'))
                    return false;
            }
            var closeIdx = FindCloseNotInCode(text, i, end, ch.ToString());
            if (closeIdx >= 0)
            {
                var inner = new List<MarkdownInline>();
                ParseInlineCore(text, i, closeIdx, inner, refs, pipeline);
                result = MarkdownInline.CreateEmphasis(inner);
                newPos = closeIdx + 1;
                return true;
            }
        }
        return false;
    }

    private static Int32 FindCloseNotInCode(String text, Int32 start, Int32 end, String marker)
    {
        var inCode = false;
        var i = start;
        var markerLen = marker.Length;
        while (i < end - markerLen + 1)
        {
            if (text[i] == '`') { inCode = !inCode; i++; continue; }
            if (inCode) { i++; continue; }

            if (text[i] == marker[0])
            {
                // 统计连续标记串长度：等长才闭合，否则作为嵌套强调/粗体内容跳过
                var runLen = 1;
                while (i + runLen < end && text[i + runLen] == marker[0]) runLen++;
                if (runLen == markerLen) return i;
                i += runLen;
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
        // URL autolink
        if (inner.Contains("://") || inner.Contains('@'))
        {
            var href = inner.Contains('@') ? "mailto:" + inner : inner;
            var children = new List<MarkdownInline> { MarkdownInline.CreateText(inner) };
            result = MarkdownInline.CreateLink(href, "", children);
            newPos = closeAngle + 1;
            return true;
        }
        return false;
    }

    /// <summary>尝试解析裸 URL（http/https/ftp）自动链接 (MD05-01)</summary>
    private static Boolean TryParseBareUrl(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;

        // 检查是否是 URL 前缀
        var remaining = text[start..end];
        var prefixes = new[] { "https://", "http://", "ftp://", "www." };
        String? matchedPrefix = null;
        foreach (var prefix in prefixes)
        {
            if (remaining.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                matchedPrefix = prefix;
                break;
            }
        }
        if (matchedPrefix == null) return false;

        // 收集 URL 字符直到空白或标点
        var urlEnd = start;
        while (urlEnd < end)
        {
            var c = text[urlEnd];
            if (Char.IsWhiteSpace(c) || c == '<' || c == '>' || c == '"' || c == '\'') break;
            // 末尾标点不包含在 URL 中
            if (c is '.' or ',' or ';' or ':' or '!' or '?' && urlEnd + 1 < end && (Char.IsWhiteSpace(text[urlEnd + 1]) || text[urlEnd + 1] == '<'))
                break;
            urlEnd++;
        }

        // 去除末尾标点
        while (urlEnd > start && text[urlEnd - 1] is '.' or ',' or ';' or ')' or '}' or ']')
            urlEnd--;

        if (urlEnd <= start + matchedPrefix.Length) return false;

        var url = text[start..urlEnd];
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
        return Char.IsLetterOrDigit(c) || c == '.' || c == '-' || c == '_' || c == '+' || c == '%';
    }

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
            ParseInlineCore(text, start + 1, close1, innerInlines, refs, pipeline);
            result = MarkdownInline.CreateLink(def.url, def.title, innerInlines);
            newPos = close2 + 1;
            return true;
        }

        // [label] shortcut 引用（仅当 label 已定义，且后不跟 (）
        if (after < end && text[after] == '(') return false;
        if (!refs.TryGetValue(label, out var def2)) return false;

        var inner2 = new List<MarkdownInline>();
        ParseInlineCore(text, start + 1, close1, inner2, refs, pipeline);
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
    /// <returns>解码后字符，未识别返回 null</returns>
    private static String? DecodeEntity(String entity)
    {
        // 数字字符引用 &#123; / &#x1F;
        if (entity.StartsWith("#"))
        {
            try
            {
                Int32 code;
                if (entity.Length > 2 && (entity[1] == 'x' || entity[1] == 'X'))
                    code = Convert.ToInt32(entity[2..], 16);
                else
                    code = Int32.Parse(entity[1..]);
                // 超出 BMP 的代理对不处理（保留原文）
                if (code <= 0 || code > 0xFFFF || code >= 0xD800 && code <= 0xDFFF) return null;
                return ((Char)code).ToString();
            }
            catch
            {
                return null;
            }
        }

        // 具名实体（常用子集）
        return entity switch
        {
            "amp" => "&",
            "lt" => "<",
            "gt" => ">",
            "quot" => "\"",
            "apos" => "'",
            "nbsp" => " ",
            "copy" => "©",
            "reg" => "®",
            "trade" => "™",
            "mdash" => "—",
            "ndash" => "–",
            "hellip" => "…",
            "lsquo" => "'",
            "rsquo" => "'",
            "ldquo" => "\"",
            "rdquo" => "\"",
            "laquo" => "«",
            "raquo" => "»",
            "bull" => "•",
            "middot" => "·",
            "sect" => "§",
            "para" => "¶",
            "deg" => "°",
            "plusmn" => "±",
            "times" => "×",
            "divide" => "÷",
            "micro" => "µ",
            _ => null,
        };
    }

    /// <summary>预扫描引用链接定义 [id]: url "title"（支持定义在后使用在前）</summary>
    /// <param name="doc">文档对象（填充 References）</param>
    private void PreScanReferences(MarkdownDocument doc)
    {
        for (var i = 0; i < _lines.Length; i++)
        {
            var trimmed = _lines[i].Trim();
            // 排除脚注定义 [^id]:
            if (trimmed.StartsWith("[^") || trimmed.StartsWith("[") && trimmed.Length > 1 && trimmed[1] == '^') continue;
            var match = Regex.Match(trimmed, @"^\[([^\]]+)\]:\s*(?:<([^>]+)>|(\S+))(?:\s+(.*))?$");
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
        }
    }
    #endregion
}
