using System.Text;
using NewLife.Collections;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 序列化器</summary>
/// <remarks>将 <see cref="MarkdownDocument"/> 序列化为 CommonMark + GFM 格式的 Markdown 文本。</remarks>
internal sealed class MarkdownWriter
{
    #region 属性
    private readonly String _bulletChar;

    /// <summary>是否正在写入表格单元格（用于转义管道符）</summary>
    private Boolean _inTable;

    /// <summary>下一个行内元素是否处于行首（用于保护列表/标题等起始标记）</summary>
    private Boolean _atLineStart;
    #endregion

    #region 构造
    /// <summary>实例化</summary>
    /// <param name="bulletChar">无序列表项目符号</param>
    public MarkdownWriter(String bulletChar = "-")
    {
        _bulletChar = bulletChar ?? "-";
    }
    #endregion

    #region 入口
    /// <summary>将文档序列化为 Markdown 字符串</summary>
    /// <param name="doc">文档</param>
    /// <returns>Markdown 文本</returns>
    public String ToMarkdown(MarkdownDocument doc)
    {
        var sb = new StringBuilder();
        var roundtrip = doc.Roundtrip;

        var refLines = doc.ReferenceSourceLines;
        var refIndexes = doc.ReferenceLineIndexes;
        var refIdx = 0;

        // 往返模式：输出指定行号之前的引用定义原始行（原位保留，不统一移到底部）
        void EmitReferencesBefore(Int32 lineIndex)
        {
            if (!roundtrip) return;
            while (refIdx < refIndexes.Count && refIndexes[refIdx] < lineIndex)
            {
                if (sb.Length > 0 && sb[^1] != '\n') sb.Append('\n');
                sb.Append(refLines[refIdx]);
                refIdx++;
            }
        }

        for (var i = 0; i < doc.Blocks.Count; i++)
        {
            var block = doc.Blocks[i];
            EmitReferencesBefore(block.SourceLine);

            // 往返模式：若块有原始源码且内容未被修改，原样输出以保留原始格式；已修改则重新序列化
            if (roundtrip && block.SourceText != null && !IsModified(block))
            {
                sb.Append(block.SourceText);
                if (i < doc.Blocks.Count - 1 && !block.SourceText.EndsWith("\n")) sb.AppendLine();
                continue;
            }
            WriteBlock(sb, block, 0);
            if (i < doc.Blocks.Count - 1) sb.AppendLine();
        }

        // 引用链接定义（[id]: url）：往返模式输出剩余（原本在文档末尾的）；非往返模式追加到末尾
        if (refLines.Count > 0)
        {
            if (!roundtrip) refIdx = 0;
            while (refIdx < refLines.Count)
            {
                if (sb.Length > 0 && sb[^1] != '\n') sb.Append('\n');
                sb.Append(refLines[refIdx]);
                refIdx++;
            }
        }
        else if (doc.References.Count > 0)
        {
            // 程序化构造的文档（无原始行）：从字典生成
            if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
            foreach (var kv in doc.References)
            {
                sb.Append('[').Append(kv.Key).Append("]: ").AppendLine(kv.Value);
            }
        }

        return sb.ToString().Replace("\r\n", "\n").TrimEnd() + "\n";
    }

    /// <summary>判断块是否被修改：当前规范序列化与解析时快照比较（MD06-03 往返渲染）</summary>
    /// <param name="block">块节点</param>
    /// <returns>内容与解析时是否不一致</returns>
    private Boolean IsModified(MarkdownBlock block)
        => block.SourceSnapshot == null || SerializeBlock(block) != block.SourceSnapshot;

    /// <summary>将块序列化为字符串（供修改检测使用；MD26：StringBuilder 池化，往返模式每块 1 次调用）</summary>
    /// <param name="block">块节点</param>
    /// <returns>序列化文本</returns>
    private String SerializeBlock(MarkdownBlock block)
    {
        var sb = Pool.StringBuilder.Get();
        try
        {
            WriteBlock(sb, block, 0);
            return sb.ToString();
        }
        finally
        {
            Pool.Put(sb, false);
        }
    }

    /// <summary>生成块的规范序列化快照（供解析器在往返渲染时记录，MD06-03；MD26：实例方法复用 writer，省每块 new）</summary>
    /// <param name="block">块节点</param>
    /// <returns>规范序列化文本</returns>
    internal String SerializeSnapshot(MarkdownBlock block) => SerializeBlock(block);
    #endregion

    #region 块
    private void WriteBlock(StringBuilder sb, MarkdownBlock block, Int32 indent)
    {
        var pad = new String(' ', indent);
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
                var h = (HeadingBlock)block;
                sb.Append(new String('#', h.Level)).Append(' ');
                WriteInlines(sb, block.Inlines);
                sb.AppendLine();
                sb.AppendLine();
                break;

            case MarkdownBlockType.Paragraph:
                sb.Append(pad);
                _atLineStart = true;
                WriteInlines(sb, block.Inlines);
                _atLineStart = false;
                sb.AppendLine();
                sb.AppendLine();
                break;

            case MarkdownBlockType.CodeBlock:
                var cb = (CodeBlock)block;
                var fence = cb.Language == null || !cb.Language.Contains("~") ? "```" : "~~~";
                sb.Append(pad).Append(fence).AppendLine(cb.Language ?? "");
                foreach (var codeLine in cb.RawText.Split('\n'))
                {
                    sb.Append(pad).AppendLine(codeLine);
                }
                sb.Append(pad).AppendLine(fence);
                sb.AppendLine();
                break;

            case MarkdownBlockType.BlockQuote:
                foreach (var child in block.Children)
                {
                    var inner = new StringBuilder();
                    WriteBlock(inner, child, 0);
                    // 去掉末尾空行，避免产生多余的空 ">" 行（保证往返快照一致）
                    var text = inner.ToString().TrimEnd('\r', '\n');
                    foreach (var line in text.Split('\n'))
                    {
                        if (line.Length == 0) sb.AppendLine(">");
                        else sb.Append("> ").AppendLine(line);
                    }
                }
                sb.AppendLine();
                break;

            case MarkdownBlockType.BulletList:
                var bl = (BulletListBlock)block;
                for (var i = 0; i < block.Children.Count; i++)
                {
                    WriteListItem(sb, block.Children[i], false, 0, indent);
                    // 松散列表：项间空行（CommonMark 保真）
                    if (bl.IsLoose && i < block.Children.Count - 1) sb.AppendLine();
                }
                sb.AppendLine();
                break;

            case MarkdownBlockType.OrderedList:
                var ol = (OrderedListBlock)block;
                var num = ol.OrderedStart;
                for (var i = 0; i < block.Children.Count; i++)
                {
                    WriteListItem(sb, block.Children[i], true, num++, indent);
                    // 松散列表：项间空行（CommonMark 保真）
                    if (ol.IsLoose && i < block.Children.Count - 1) sb.AppendLine();
                }
                sb.AppendLine();
                break;

            case MarkdownBlockType.ThematicBreak:
                sb.AppendLine("---");
                sb.AppendLine();
                break;

            case MarkdownBlockType.HtmlBlock:
                var hb = (HtmlBlock)block;
                sb.AppendLine(hb.RawText);
                sb.AppendLine();
                break;

            case MarkdownBlockType.Table:
                WriteTable(sb, block);
                sb.AppendLine();
                break;

            case MarkdownBlockType.MathBlock:
                var mb = (MathBlock)block;
                sb.AppendLine("$$");
                sb.AppendLine(mb.Content);
                sb.AppendLine("$$");
                sb.AppendLine();
                break;

            case MarkdownBlockType.FootnoteDefinition:
                var fd = (FootnoteDefinitionBlock)block;
                sb.Append("[^").Append(fd.Id).Append("]: ");
                WriteInlines(sb, fd.Definition);
                sb.AppendLine();
                sb.AppendLine();
                break;

            case MarkdownBlockType.DefinitionList:
                foreach (var child in block.Children)
                {
                    if (child is DefinitionTermBlock dt)
                    {
                        sb.Append(pad);
                        WriteInlines(sb, dt.Inlines);
                        sb.AppendLine();
                    }
                    else if (child is DefinitionDescriptionBlock dd)
                    {
                        sb.Append(pad).Append(": ");
                        WriteInlines(sb, dd.Inlines);
                        sb.AppendLine();
                    }
                }
                sb.AppendLine();
                break;
        }
    }

    private void WriteListItem(StringBuilder sb, MarkdownBlock item, Boolean ordered,
        Int32 number, Int32 indent)
    {
        var pad = new String(' ', indent);
        String marker;
        if (ordered)
            marker = $"{number}. ";
        else
            marker = _bulletChar + " ";

        sb.Append(pad).Append(marker);

        var li = item as ListItemBlock;
        if (li != null && li.IsTaskItem)
            sb.Append(li.IsChecked ? "[x] " : "[ ] ");

        if (item.Inlines.Count > 0)
        {
            _atLineStart = true;
            WriteInlines(sb, item.Inlines);
            _atLineStart = false;
            sb.AppendLine();

            // 如果有子块（嵌套列表等），继续输出子块
            if (item.Children.Count > 0)
            {
                var childIndent = indent + marker.Length;
                foreach (var child in item.Children)
                {
                    WriteBlock(sb, child, childIndent);
                }
            }
        }
        else if (item.Children.Count > 0)
        {
            var firstBlock = item.Children[0];
            if (firstBlock.Type == MarkdownBlockType.Paragraph && firstBlock.Inlines.Count > 0)
            {
                _atLineStart = true;
                WriteInlines(sb, firstBlock.Inlines);
                _atLineStart = false;
                sb.AppendLine();
            }
            else
            {
                sb.AppendLine();
            }
            var childIndent = indent + marker.Length;
            for (var i = 1; i < item.Children.Count; i++)
            {
                WriteBlock(sb, item.Children[i], childIndent);
            }
        }
        else
        {
            sb.AppendLine();
        }
    }

    private void WriteTable(StringBuilder sb, MarkdownBlock table)
    {
        if (table.Children.Count == 0) return;
        var headerRow = table.Children[0];
        // 列数以表头行为准：数据行补空/截断，保证输出列数一致（GFM）
        var headerCols = headerRow.Children.Count;

        // 表格单元格内转义管道符，避免破坏列结构
        _inTable = true;
        try
        {
            // Header
            sb.Append('|');
            foreach (var cell in headerRow.Children)
            {
                sb.Append(' ');
                WriteInlines(sb, cell.Inlines);
                sb.Append(" |");
            }
            sb.AppendLine();

            // Separator
            sb.Append('|');
            foreach (var cell in headerRow.Children)
            {
                var tc = cell as TableCellBlock;
                var align = tc?.Alignment ?? "";
                sb.Append(align == "center" ? " :---: " : align == "right" ? " ---: " : align == "left" ? " :--- " : " --- ");
                sb.Append('|');
            }
            sb.AppendLine();

            // Data rows
            for (var r = 1; r < table.Children.Count; r++)
            {
                var row = table.Children[r];
                sb.Append('|');
                for (var c = 0; c < headerCols; c++)
                {
                    sb.Append(' ');
                    if (c < row.Children.Count)
                        WriteInlines(sb, row.Children[c].Inlines);
                    sb.Append(" |");
                }
                sb.AppendLine();
            }
        }
        finally
        {
            _inTable = false;
        }
    }
    #endregion

    #region 行内
    private void WriteInlines(StringBuilder sb, List<MarkdownInline> inlines)
    {
        var saved = _atLineStart;
        try
        {
            for (var i = 0; i < inlines.Count; i++)
            {
                // 仅第一个行内元素处于行首，避免嵌套内容误转义
                if (i > 0) _atLineStart = false;
                WriteInline(sb, inlines[i]);
            }
        }
        finally
        {
            _atLineStart = saved;
        }
    }

    private void WriteInline(StringBuilder sb, MarkdownInline inline)
    {
        switch (inline.Type)
        {
            case MarkdownInlineType.Text:
                WriteText(sb, inline.Text ?? "", _inTable);
                break;
            case MarkdownInlineType.Code:
                WriteCode(sb, inline.Text ?? "");
                break;
            case MarkdownInlineType.Strong:
                sb.Append("**");
                WriteInlines(sb, inline.Children);
                sb.Append("**");
                break;
            case MarkdownInlineType.Emphasis:
                sb.Append('*');
                WriteInlines(sb, inline.Children);
                sb.Append('*');
                break;
            case MarkdownInlineType.StrongEmphasis:
                sb.Append("***");
                WriteInlines(sb, inline.Children);
                sb.Append("***");
                break;
            case MarkdownInlineType.Strikethrough:
                sb.Append("~~");
                WriteInlines(sb, inline.Children);
                sb.Append("~~");
                break;
            case MarkdownInlineType.Link:
                sb.Append('[');
                WriteInlines(sb, inline.Children);
                sb.Append("](").Append(WriteDestination(inline.Href));
                if (!String.IsNullOrEmpty(inline.Title))
                    sb.Append(" \"").Append(EscapeTitle(inline.Title)).Append('"');
                sb.Append(')');
                break;
            case MarkdownInlineType.Image:
                sb.Append("![").Append(inline.Alt).Append("](").Append(WriteDestination(inline.Href));
                if (!String.IsNullOrEmpty(inline.Title))
                    sb.Append(" \"").Append(EscapeTitle(inline.Title)).Append('"');
                sb.Append(')');
                break;
            case MarkdownInlineType.HardBreak:
                sb.Append("  \n");
                break;
            case MarkdownInlineType.SoftBreak:
                sb.Append('\n');
                break;
            case MarkdownInlineType.RawHtml:
                sb.Append(inline.Text);
                break;
            case MarkdownInlineType.AutoLink:
                sb.Append('<').Append(inline.Text).Append('>');
                break;
            case MarkdownInlineType.MathInline:
                sb.Append('$').Append(inline.Text).Append('$');
                break;
            case MarkdownInlineType.FootnoteRef:
                sb.Append("[^").Append(inline.Text).Append(']');
                break;
        }
    }

    /// <summary>写入文本内容并做上下文相关转义，避免特殊字符破坏 Markdown 结构</summary>
    /// <param name="sb">输出</param>
    /// <param name="text">文本</param>
    /// <param name="inTable">是否在表格单元格内（转义管道符）</param>
    private void WriteText(StringBuilder sb, String text, Boolean inTable)
    {
        if (String.IsNullOrEmpty(text)) return;

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            var next = i + 1 < text.Length ? text[i + 1] : '\0';
            switch (ch)
            {
                case '\\': sb.Append("\\\\"); break;
                case '`': sb.Append("\\`"); break;
                case '*': sb.Append("\\*"); break;
                case '_': sb.Append("\\_"); break;
                case '[': sb.Append("\\["); break;
                case ']': sb.Append("\\]"); break;
                case '|':
                    sb.Append(inTable ? "\\|" : "|");
                    break;
                case '<':
                    // 仅当可能构成 HTML 标签/自动链接/注释时转义
                    if (IsAsciiLetter(next) || next is '/' or '!' or '?' or '>')
                        sb.Append("\\<");
                    else
                        sb.Append(ch);
                    break;
                case '~':
                    // 双波浪线可能构成删除线
                    if (next == '~') sb.Append("\\~");
                    else sb.Append('~');
                    break;
                case '-':
                case '+':
                    // 行首的列表/分隔线标记需转义
                    if (_atLineStart && i == 0 && (next == ' ' || next == '\t' || next == ch))
                        sb.Append('\\').Append(ch);
                    else
                        sb.Append(ch);
                    break;
                case '#':
                    // 行首的标题标记需转义
                    if (_atLineStart && i == 0 && (next == ' ' || next == '\t' || next == '#'))
                        sb.Append("\\#");
                    else
                        sb.Append(ch);
                    break;
                case '>':
                    // 行首的引用标记需转义
                    if (_atLineStart && i == 0 && (next == ' ' || next == '\t'))
                        sb.Append("\\>");
                    else
                        sb.Append(ch);
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }
    }

    /// <summary>写入行内代码，含反引号或首尾空格时用双反引号包裹（CommonMark 兼容）</summary>
    /// <param name="sb">输出</param>
    /// <param name="code">代码文本</param>
    private static void WriteCode(StringBuilder sb, String code)
    {
        if (code.Contains('`') || code.StartsWith(" ") || code.EndsWith(" "))
        {
            var ticks = "``";
            while (code.Contains(ticks)) ticks += "`";
            sb.Append(ticks).Append(' ').Append(code).Append(' ').Append(ticks);
        }
        else
        {
            sb.Append('`').Append(code).Append('`');
        }
    }

    /// <summary>写入链接目标：含空格或括号时用尖括号包裹，避免解析失败</summary>
    /// <param name="href">目标 URL</param>
    /// <returns>可安全输出的目标文本</returns>
    private static String WriteDestination(String href)
    {
        if (String.IsNullOrEmpty(href)) return "";
        if (href.Contains(' ') || href.Contains('(') || href.Contains(')'))
            return "<" + href.Replace("<", "\\<").Replace(">", "\\>") + ">";
        return href;
    }

    /// <summary>转义链接标题中的引号与反斜杠</summary>
    /// <param name="title">标题</param>
    /// <returns>转义后文本</returns>
    private static String EscapeTitle(String title) =>
        title.Replace("\\", "\\\\").Replace("\"", "\\\"");

    /// <summary>ASCII 字母判断（兼容低版本框架，不使用 Char.IsAsciiLetter）</summary>
    private static Boolean IsAsciiLetter(Char ch) =>
        ch >= 'a' && ch <= 'z' || ch >= 'A' && ch <= 'Z';
    #endregion
}
