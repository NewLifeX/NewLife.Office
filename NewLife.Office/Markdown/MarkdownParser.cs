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
        while (_pos < _lines.Length)
            ParseBlock(doc.Blocks);
        return doc;
    }
    #endregion

    #region 块级解析
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
        if (atx != null) { container.Add(atx); _pos++; return; }

        // 围栏代码块 ``` or ~~~
        if (trimmed.StartsWith("```") || trimmed.StartsWith("~~~"))
        {
            container.Add(ParseFencedCode(trimmed));
            return;
        }

        // HTML 块 (以 < 开头的完整块)
        if (trimmed.StartsWith("<"))
        {
            var html = TryParseHtmlBlock(trimmed);
            if (html != null) { container.Add(html); return; }
        }

        // 分隔线  --- / *** / ___
        if (IsThematicBreak(trimmed)) { container.Add(MarkdownBlock.CreateThematicBreak()); _pos++; return; }

        // 引用块 >
        if (trimmed.StartsWith(">"))
        {
            container.Add(ParseBlockQuote());
            return;
        }

        // 无序列表
        if (IsBulletListMarker(trimmed, out var _))
        {
            container.Add(ParseList(false));
            return;
        }

        // 有序列表
        if (IsOrderedListMarker(trimmed, out var startNum, out var _))
        {
            container.Add(ParseList(true, startNum));
            return;
        }

        // Setext 标题（只在段落内检查）
        // 段落（可能包含 Setext 标题）
        container.Add(ParseParagraphOrSetext(indent));
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
        if (spaceIdx >= 0) lang = lang.Substring(0, spaceIdx);

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
        if (code.EndsWith("\n")) code = code.Substring(0, code.Length - 1);
        if (code.EndsWith("\r")) code = code.Substring(0, code.Length - 1);
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
                quotedLines.Add(trimmed.Substring(2));
            else if (trimmed == ">")
                quotedLines.Add("");
            else if (trimmed.StartsWith(">"))       // ">>" nested, strip one level
                quotedLines.Add(trimmed.Substring(1));
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
                items.Add(ParseListItem(trimmed.Substring(2).TrimStart(), false));
            }
            else
            {
                if (!IsOrderedListMarker(trimmed, out var _, out var delim)) break;
                if (orderedDelim == ' ') orderedDelim = delim;
                else if (delim != orderedDelim) break;
                var markerMatch = Regex.Match(trimmed, @"^(\d{1,9}[.)]) (.*)$");
                var content = markerMatch.Success ? markerMatch.Groups[2].Value : trimmed.Substring(3);
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

        var table = new MarkdownBlock { Type = MarkdownBlockType.Table };

        // Header row
        var headerRow = new MarkdownBlock { Type = MarkdownBlockType.TableRow };
        var headerCells = SplitTableRow(lines[0]);
        for (var i = 0; i < headerCells.Count; i++)
        {
            var cell = new MarkdownBlock
            {
                Type = MarkdownBlockType.TableCell,
                IsHeader = true,
                Alignment = i < alignments.Count ? alignments[i] : "",
            };
            cell.Inlines.AddRange(ParseInline(headerCells[i].Trim()));
            headerRow.Children.Add(cell);
        }
        table.Children.Add(headerRow);

        // Data rows
        for (var r = 2; r < lines.Count; r++)
        {
            var dataRow = new MarkdownBlock { Type = MarkdownBlockType.TableRow };
            var dataCells = SplitTableRow(lines[r]);
            for (var i = 0; i < alignments.Count; i++)
            {
                var cell = new MarkdownBlock
                {
                    Type = MarkdownBlockType.TableCell,
                    IsHeader = false,
                    Alignment = i < alignments.Count ? alignments[i] : "",
                };
                var cellText = i < dataCells.Count ? dataCells[i].Trim() : "";
                cell.Inlines.AddRange(ParseInline(cellText));
                dataRow.Children.Add(cell);
            }
            table.Children.Add(dataRow);
        }

        return table;
    }

    private static List<String> SplitTableRow(String line)
    {
        line = line.Trim();
        if (line.StartsWith("|")) line = line.Substring(1);
        if (line.EndsWith("|")) line = line.Substring(0, line.Length - 1);
        return [.. line.Split('|')];
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

            i++;
        }

        FlushText(text, textStart, end, result);
    }

    private static void FlushText(String text, Int32 start, Int32 end, List<MarkdownInline> result)
    {
        if (start >= end) return;
        var s = text.Substring(start, end - start);
        // Remove trailing backslash before hard break
        if (s.EndsWith("\\")) s = s.Substring(0, s.Length - 1);
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
        var url = text.Substring(urlStart, urlEnd - urlStart);

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
                title = text.Substring(titleStart, titleEnd - titleStart);
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
    #endregion
}
