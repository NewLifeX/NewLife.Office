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
    private MarkdownDocument? _doc;

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
        ParseFrontMatter(doc);

        while (_pos < _lines.Length)
        {
            ParseBlock(doc.Blocks);
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
        if (_pos < _lines.Length)
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
        var indent = line.Length - trimmed.Length;

        // 空行
        if (String.IsNullOrWhiteSpace(line)) { _pos++; return; }

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
        if (TryParseFootnoteDefinition(trimmed, out var fnDef))
        {
            AddBlock(container, fnDef!);
            _pos++;
            return;
        }

        // 数学公式块 $$...$$ (MD05-04)
        if (trimmed == "$$" || trimmed.StartsWith("$$"))
        {
            AddBlock(container, ParseMathBlock(trimmed));
            return;
        }

        // 缩写定义 *[ABBR]: (MD05-08)
        if (trimmed.StartsWith("*[") && trimmed.Contains("]:"))
        {
            ParseAbbreviation(trimmed);
            _pos++;
            return;
        }

        // 定义列表 : description (MD05-07)
        if (trimmed.StartsWith(": ") || trimmed == ":")
        {
            var dl = TryParseDefinitionList();
            if (dl != null) { AddBlock(container, dl); return; }
        }

        // Setext 标题（只在段落内检查）
        // 段落（可能包含 Setext 标题）
        AddBlock(container, ParseParagraphOrSetext(indent));
    }

    private static MarkdownBlock? TryParseAtxHeading(String trimmed)
    {
        var match = Regex.Match(trimmed, @"^(#{1,6})(\s+|$)(.*)$");
        if (!match.Success) return null;
        var hashes = match.Groups[1].Value;
        // trailing #? remove it
        var text = match.Groups[3].Value.TrimEnd();
        text = Regex.Replace(text, @"\s+#+\s*$", "").TrimEnd();
        var inlines = ParseInline(text);
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
            sb.AppendLine(line);
            _pos++;
        }
        var code = sb.ToString();
        if (code.EndsWith("\n")) code = code[..^1];
        if (code.EndsWith("\r")) code = code[..^1];
        return MarkdownBlock.CreateCodeBlock(code, lang);
    }

    private MarkdownBlock? TryParseHtmlBlock(String trimmed)
    {
        // Simple: consume lines until empty line or end
        var sb = new System.Text.StringBuilder();
        var start = _pos;
        while (_pos < _lines.Length && !String.IsNullOrWhiteSpace(_lines[_pos]))
        {
            sb.AppendLine(_lines[_pos]);
            _pos++;
        }
        var html = sb.ToString().TrimEnd();
        // Only treat as HTML block if it looks like a tag
        if (!html.TrimStart().StartsWith("<")) { _pos = start; return null; }
        return MarkdownBlock.CreateHtmlBlock(html);
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
        var inner = new MarkdownParser();
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

    private MarkdownBlock ParseList(Boolean ordered, Int32 startNum = 1)
    {
        var items = new List<MarkdownBlock>();
        var bulletChar = ' ';
        var orderedDelim = ' ';

        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            if (String.IsNullOrWhiteSpace(line)) { _pos++; continue; }
            var trimmed = line.TrimStart();

            if (!ordered)
            {
                if (!IsBulletListMarker(trimmed, out var m)) break;
                if (bulletChar == ' ') bulletChar = m;
                else if (m != bulletChar) break;
                _pos++;
                items.Add(ParseListItem(trimmed[2..].TrimStart(), false));
            }
            else
            {
                if (!IsOrderedListMarker(trimmed, out var _, out var delim)) break;
                if (orderedDelim == ' ') orderedDelim = delim;
                else if (delim != orderedDelim) break;
                var markerMatch = Regex.Match(trimmed, @"^(\d{1,9}[.)]) (.*)$");
                var content = markerMatch.Success ? markerMatch.Groups[2].Value : trimmed[3..];
                _pos++;
                items.Add(ParseListItem(content, false));
            }

            // Continuation lines (indented 2+ spaces)
            while (_pos < _lines.Length)
            {
                var cont = _lines[_pos];
                if (String.IsNullOrWhiteSpace(cont)) break;
                if (!cont.StartsWith("  ") && !cont.StartsWith("\t")) break;
                // We'll handle this as part of the item content in a full implementation
                // For now, just skip
                _pos++;
            }
        }

        if (!ordered) return MarkdownBlock.CreateBulletList(items);
        return MarkdownBlock.CreateOrderedList(items, startNum);
    }

    private static MarkdownBlock ParseListItem(String content, Boolean tight)
    {
        // Check for task list item: - [ ] or - [x]
        var taskMatch = Regex.Match(content, @"^\[([ xX])\] (.*)$");
        var isTask = taskMatch.Success;
        var isChecked = isTask && taskMatch.Groups[1].Value != " ";
        var itemText = isTask ? taskMatch.Groups[2].Value : content;
        var inlines = ParseInline(itemText);
        return MarkdownBlock.CreateListItem(inlines, isTask, isChecked);
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
                return MarkdownBlock.CreateHeading(1, ParseInline(headingText));
            }
            if (lines.Count > 0 && Regex.IsMatch(trimmedNext, @"^[-]+\s*$"))
            {
                _pos++;
                var headingText = String.Join(" ", lines).Trim();
                return MarkdownBlock.CreateHeading(2, ParseInline(headingText));
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
        if (lines.Count >= 2 && lines[1].TrimStart().StartsWith("|") ||
            (lines.Count >= 2 && Regex.IsMatch(lines[1].Trim(), @"^[|\-: ]+$")))
        {
            var table = TryParseTable(lines);
            if (table != null) return table;
        }

        // Join with spaces (CommonMark paragraph continuation)
        var paragraphText = BuildParagraphText(lines);
        return MarkdownBlock.CreateParagraph(ParseInline(paragraphText));
    }

    private static String BuildParagraphText(List<String> lines)
    {
        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            if (i > 0)
            {
                // Hard break if trailing spaces
                if (line.EndsWith("  "))
                    sb.Append("  \n");
                else
                    sb.Append(' ');
            }
            sb.Append(line.Trim());
        }
        return sb.ToString();
    }

    private static MarkdownBlock? TryParseTable(List<String> lines)
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
                ParseInline(headerCells[i].Trim()),
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
                    ParseInline(cellText),
                    isHeader: false,
                    alignment: i < alignments.Count ? alignments[i] : "");
                dataRow.Children.Add(cell);
            }
            table.Children.Add(dataRow);
        }

        return table;
    }

    private static List<String> SplitTableRow(String line)
    {
        line = line.Trim();
        if (line.StartsWith("|")) line = line[1..];
        if (line.EndsWith("|")) line = line[..^1];
        return [.. line.Split('|')];
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

        while (_pos < _lines.Length)
        {
            var line = _lines[_pos];
            if (line.TrimEnd() == "---")
            {
                for (var i = start + 1; i < _pos; i++)
                {
                    var kvLine = _lines[i];
                    var colonIdx = kvLine.IndexOf(':');
                    if (colonIdx > 0 && colonIdx < kvLine.Length - 1)
                    {
                        var key = kvLine[..colonIdx].Trim();
                        var value = kvLine[(colonIdx + 1)..].Trim();
                        if (value.Length >= 2 &&
                            ((value[0] == '"' && value[^1] == '"') ||
                             (value[0] == '\'' && value[^1] == '\'')))
                            value = value[1..^1];
                        if (key.Length > 0 && value.Length > 0)
                            doc.FrontMatter[key] = value;
                    }
                }
                _pos++;
                return;
            }
            _pos++;
        }

        _pos = start;
    }

    private static Boolean TryParseFootnoteDefinition(String trimmed, out MarkdownBlock? block)
    {
        block = null;

        // [^id]: text
        var match = Regex.Match(trimmed, @"^\[\^([^\]]+)\]:\s*(.*)$");
        if (!match.Success) return false;

        var id = match.Groups[1].Value;
        var defText = match.Groups[2].Value;
        var inlines = ParseInline(defText);
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
        var termInlines = ParseInline(prevLine);
        var term = new DefinitionTermBlock(termInlines);
        dl.Children.Add(term);

        // 解析当前描述
        var descLine = _lines[_pos].TrimStart();
        var descText = descLine.StartsWith(": ") ? descLine[2..] : descLine[1..];
        var descInlines = ParseInline(descText.Trim());
        var desc = new DefinitionDescriptionBlock(descInlines);
        dl.Children.Add(desc);
        _pos++;

        // 继续解析后续 `: ` 行（同一术语的多段描述）
        while (_pos < _lines.Length)
        {
            var nextLine = _lines[_pos].TrimStart();
            if (!nextLine.StartsWith(": ")) break;
            var moreText = nextLine[2..];
            var moreInlines = ParseInline(moreText.Trim());
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
        ParseInlineCore(text, 0, text.Length, result);
        return result;
    }

    private static void ParseInlineCore(String text, Int32 start, Int32 end, List<MarkdownInline> result)
    {
        var i = start;
        var textStart = start;

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

            // Inline code: `...`
            if (ch == '`')
            {
                var tickEnd = text.IndexOf('`', i + 1);
                if (tickEnd > i)
                {
                    FlushText(text, textStart, i, result);
                    result.Add(MarkdownInline.CreateCode(text.Substring(i + 1, tickEnd - i - 1)));
                    i = tickEnd + 1;
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
                if (TryParseLink(text, i, end, out var linkInline, out var linkEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(linkInline!);
                    i = linkEnd;
                    textStart = i;
                    continue;
                }
            }

            // Strong+Emphasis ***
            if (ch == '*' || ch == '_')
            {
                if (TryParseEmphasis(text, i, end, out var emphInline, out var emphEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(emphInline!);
                    i = emphEnd;
                    textStart = i;
                    continue;
                }
            }

            // Strikethrough: ~~text~~
            if (ch == '~' && i + 1 < end && text[i + 1] == '~')
            {
                var closeIdx = text.IndexOf("~~", i + 2, StringComparison.Ordinal);
                if (closeIdx > i + 1)
                {
                    FlushText(text, textStart, i, result);
                    var inner = new List<MarkdownInline>();
                    ParseInlineCore(text, i + 2, closeIdx, inner);
                    result.Add(MarkdownInline.CreateStrikethrough(inner));
                    i = closeIdx + 2;
                    textStart = i;
                    continue;
                }
            }

            // Autolink: <url> or <email>
            if (ch == '<')
            {
                if (TryParseAutoLink(text, i, end, out var autoInline, out var autoEnd))
                {
                    FlushText(text, textStart, i, result);
                    result.Add(autoInline!);
                    i = autoEnd;
                    textStart = i;
                    continue;
                }
            }

            // 裸 URL 自动链接 (MD05-01): http:// / https:// / ftp://
            if (ch == 'h' || ch == 'f' || ch == 'w')
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
            if (Char.IsLetterOrDigit(ch) && i > textStart)
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
            if (ch == '$' && i + 1 < end && text[i + 1] != '$')
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
            if (ch == '[' && i + 1 < end && text[i + 1] == '^')
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
            if (ch == ':' && i > textStart && i + 2 < end)
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
        out MarkdownInline? result, out Int32 newPos)
    {
        result = null; newPos = start;
        var closeText = FindClosingBracket(text, start, end);
        if (closeText < 0 || closeText + 1 >= end || text[closeText + 1] != '(') return false;

        var linkText = text.Substring(start + 1, closeText - start - 1);
        var (url, title, endParen) = ParseLinkDestination(text, closeText + 1, end);
        if (endParen < 0) return false;

        var innerInlines = new List<MarkdownInline>();
        ParseInlineCore(linkText, 0, linkText.Length, innerInlines);
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
            var titleStart = i + 1;
            var titleEnd = text.IndexOf(close, titleStart);
            if (titleEnd > 0)
            {
                title = text[titleStart..titleEnd];
                i = titleEnd + 1;
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

    private static Boolean TryParseEmphasis(String text, Int32 start, Int32 end,
        out MarkdownInline? result, out Int32 newPos)
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
                ParseInlineCore(text, i, closeIdx, inner);
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
                ParseInlineCore(text, i, closeIdx, inner);
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
                ParseInlineCore(text, i, closeIdx, inner);
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
        while (i < end - marker.Length + 1)
        {
            if (text[i] == '`') inCode = !inCode;
            if (!inCode && text.Substring(i, Math.Min(marker.Length, end - i)) == marker)
            {
                // Ensure not opening delimiter
                if (i + marker.Length <= end)
                    return i;
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
    #endregion
}
