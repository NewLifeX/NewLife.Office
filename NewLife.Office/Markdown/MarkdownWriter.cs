using System.Text;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 序列化器</summary>
/// <remarks>将 <see cref="MarkdownDocument"/> 序列化为 CommonMark + GFM 格式的 Markdown 文本。</remarks>
internal sealed class MarkdownWriter
{
    #region 入口
    /// <summary>将文档序列化为 Markdown 字符串</summary>
    /// <param name="doc">文档</param>
    /// <returns>Markdown 文本</returns>
    public String ToMarkdown(MarkdownDocument doc)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < doc.Blocks.Count; i++)
        {
            WriteBlock(sb, doc.Blocks[i], 0);
            if (i < doc.Blocks.Count - 1) sb.AppendLine();
        }
        return sb.ToString().TrimEnd() + "\n";
    }
    #endregion

    #region 块
    private static void WriteBlock(StringBuilder sb, MarkdownBlock block, Int32 indent)
    {
        var pad = new String(' ', indent);
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
                sb.Append(new String('#', block.Level)).Append(' ');
                WriteInlines(sb, block.Inlines);
                sb.AppendLine();
                sb.AppendLine();
                break;

            case MarkdownBlockType.Paragraph:
                sb.Append(pad);
                WriteInlines(sb, block.Inlines);
                sb.AppendLine();
                sb.AppendLine();
                break;

            case MarkdownBlockType.CodeBlock:
                var fence = block.Language == null || !block.Language.Contains("~") ? "```" : "~~~";
                sb.Append(pad).Append(fence).AppendLine(block.Language ?? "");
                foreach (var codeLine in block.RawText.Split('\n'))
                    sb.Append(pad).AppendLine(codeLine);
                sb.Append(pad).AppendLine(fence);
                sb.AppendLine();
                break;

            case MarkdownBlockType.BlockQuote:
                foreach (var child in block.Children)
                {
                    var inner = new StringBuilder();
                    WriteBlock(inner, child, 0);
                    foreach (var line in inner.ToString().Split('\n'))
                    {
                        if (line.Length == 0) sb.AppendLine(">");
                        else sb.Append("> ").AppendLine(line);
                    }
                }
                sb.AppendLine();
                break;

            case MarkdownBlockType.BulletList:
                foreach (var item in block.Children)
                    WriteListItem(sb, item, false, 0, indent);
                sb.AppendLine();
                break;

            case MarkdownBlockType.OrderedList:
                var num = block.OrderedStart;
                foreach (var item in block.Children)
                    WriteListItem(sb, item, true, num++, indent);
                sb.AppendLine();
                break;

            case MarkdownBlockType.ThematicBreak:
                sb.AppendLine("---");
                sb.AppendLine();
                break;

            case MarkdownBlockType.HtmlBlock:
                sb.AppendLine(block.RawText);
                sb.AppendLine();
                break;

            case MarkdownBlockType.Table:
                WriteTable(sb, block);
                sb.AppendLine();
                break;
        }
    }

    private static void WriteListItem(StringBuilder sb, MarkdownBlock item, Boolean ordered,
        Int32 number, Int32 indent)
    {
        var pad = new String(' ', indent);
        String marker;
        if (ordered)
            marker = $"{number}. ";
        else
            marker = "- ";

        sb.Append(pad).Append(marker);

        if (item.IsTaskItem)
            sb.Append(item.IsChecked ? "[x] " : "[ ] ");

        if (item.Inlines.Count > 0)
        {
            WriteInlines(sb, item.Inlines);
            sb.AppendLine();
        }
        else if (item.Children.Count > 0)
        {
            var firstBlock = item.Children[0];
            if (firstBlock.Type == MarkdownBlockType.Paragraph && firstBlock.Inlines.Count > 0)
            {
                WriteInlines(sb, firstBlock.Inlines);
                sb.AppendLine();
            }
            else
            {
                sb.AppendLine();
            }
            var childIndent = indent + marker.Length;
            for (var i = 1; i < item.Children.Count; i++)
                WriteBlock(sb, item.Children[i], childIndent);
        }
        else
        {
            sb.AppendLine();
        }
    }

    private static void WriteTable(StringBuilder sb, MarkdownBlock table)
    {
        if (table.Children.Count == 0) return;
        var headerRow = table.Children[0];

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
            var align = cell.Alignment;
            sb.Append(align == "center" ? " :---: " : align == "right" ? " ---: " : " --- ");
            sb.Append('|');
        }
        sb.AppendLine();

        // Data rows
        for (var r = 1; r < table.Children.Count; r++)
        {
            var row = table.Children[r];
            sb.Append('|');
            foreach (var cell in row.Children)
            {
                sb.Append(' ');
                WriteInlines(sb, cell.Inlines);
                sb.Append(" |");
            }
            sb.AppendLine();
        }
    }
    #endregion

    #region 行内
    private static void WriteInlines(StringBuilder sb, List<MarkdownInline> inlines)
    {
        foreach (var inline in inlines)
            WriteInline(sb, inline);
    }

    private static void WriteInline(StringBuilder sb, MarkdownInline inline)
    {
        switch (inline.Type)
        {
            case MarkdownInlineType.Text:
                sb.Append(inline.Text);
                break;
            case MarkdownInlineType.Code:
                sb.Append('`').Append(inline.Text).Append('`');
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
                sb.Append("](").Append(inline.Href);
                if (!String.IsNullOrEmpty(inline.Title))
                    sb.Append(" \"").Append(inline.Title).Append('"');
                sb.Append(')');
                break;
            case MarkdownInlineType.Image:
                sb.Append("![").Append(inline.Alt).Append("](").Append(inline.Href);
                if (!String.IsNullOrEmpty(inline.Title))
                    sb.Append(" \"").Append(inline.Title).Append('"');
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
        }
    }
    #endregion
}
