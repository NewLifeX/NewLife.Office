using System.Collections.Generic;
using System.Linq;
using NewLife.Office;

namespace NewLife.Office.Markdown;

/// <summary>Markdown → Word（.docx）转换器（MD03-02）</summary>
/// <remarks>
/// 将 <see cref="MarkdownDocument"/> 转换为 Word docx 文件，
/// 标题/段落/列表/表格/代码块/引用块映射到对应的 Word 元素。
/// </remarks>
public sealed class MarkdownWordConverter
{
    #region 公开方法
    /// <summary>将 Markdown 文档写入 Word 写入器</summary>
    /// <param name="doc">Markdown 文档</param>
    /// <param name="writer">目标 Word 写入器</param>
    public void Convert(MarkdownDocument doc, WordWriter writer)
    {
        foreach (var block in doc.Blocks)
            WriteBlock(block, writer, 0);
    }

    /// <summary>将 Markdown 文档转换为 .docx 字节数组</summary>
    /// <param name="doc">Markdown 文档</param>
    /// <returns>docx 字节数组</returns>
    public Byte[] ToBytes(MarkdownDocument doc)
    {
        using var writer = new WordWriter();
        Convert(doc, writer);
        using var ms = new MemoryStream();
        writer.Save(ms);
        return ms.ToArray();
    }
    #endregion

    #region 块处理
    private static void WriteBlock(MarkdownBlock block, WordWriter writer, Int32 depth)
    {
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
                var headingText = InlinesToPlainText(block.Inlines);
                writer.AppendHeading(headingText, block.Level);
                break;

            case MarkdownBlockType.Paragraph:
                var runs = InlinesToWordRuns(block.Inlines);
                if (runs.Count > 0)
                    writer.AppendFormattedParagraph(runs);
                break;

            case MarkdownBlockType.CodeBlock:
                var codePara = writer.AppendParagraph(block.RawText, WordParagraphStyle.Normal,
                    new WordRunProperties { FontName = "Courier New", FontSize = 10f });
                break;

            case MarkdownBlockType.BlockQuote:
                foreach (var child in block.Children)
                    WriteBlock(child, writer, depth + 1);
                break;

            case MarkdownBlockType.BulletList:
                var bulletItems = CollectListItemTexts(block);
                writer.AppendBulletList(bulletItems);
                break;

            case MarkdownBlockType.OrderedList:
                var orderedItems = CollectListItemTexts(block);
                writer.AppendOrderedList(orderedItems);
                break;

            case MarkdownBlockType.Table:
                WriteTable(block, writer);
                break;

            case MarkdownBlockType.ThematicBreak:
                writer.AppendParagraph("─────────────────────────────────");
                break;

            case MarkdownBlockType.HtmlBlock:
                // HTML 块以纯文本形式保留
                if (!String.IsNullOrWhiteSpace(block.RawText))
                    writer.AppendParagraph(block.RawText.Trim());
                break;
        }
    }

    private static void WriteTable(MarkdownBlock tableBlock, WordWriter writer)
    {
        var rows = new List<IEnumerable<String>>();
        foreach (var row in tableBlock.Children)
        {
            if (row.Type != MarkdownBlockType.TableRow) continue;
            var cells = row.Children
                .Where(c => c.Type == MarkdownBlockType.TableCell)
                .Select(c => InlinesToPlainText(c.Inlines));
            rows.Add(cells);
        }
        if (rows.Count > 0)
            writer.AppendTable(rows, firstRowHeader: true);
    }
    #endregion

    #region 行内处理
    private static List<WordRun> InlinesToWordRuns(List<MarkdownInline> inlines,
        Boolean bold = false, Boolean italic = false, Boolean code = false)
    {
        var result = new List<WordRun>();
        foreach (var inline in inlines)
        {
            switch (inline.Type)
            {
                case MarkdownInlineType.Text:
                    if (!String.IsNullOrEmpty(inline.Text))
                        result.Add(MakeRun(inline.Text, bold, italic, code));
                    break;

                case MarkdownInlineType.Strong:
                    result.AddRange(InlinesToWordRuns(inline.Children, bold: true, italic: italic, code: code));
                    break;

                case MarkdownInlineType.Emphasis:
                    result.AddRange(InlinesToWordRuns(inline.Children, bold: bold, italic: true, code: code));
                    break;

                case MarkdownInlineType.StrongEmphasis:
                    result.AddRange(InlinesToWordRuns(inline.Children, bold: true, italic: true, code: code));
                    break;

                case MarkdownInlineType.Code:
                    result.Add(MakeRun(inline.Text, bold, italic, code: true));
                    break;

                case MarkdownInlineType.Strikethrough:
                    // Word 不原生支持删除线映射，以灰色文字代替
                    result.Add(new WordRun
                    {
                        Text = InlinesToPlainText(inline.Children),
                        Properties = new WordRunProperties { ForeColor = "808080" },
                    });
                    break;

                case MarkdownInlineType.Link:
                    // 超链接文字 + 链接地址作为括号文本
                    var linkText = inline.Children.Count > 0
                        ? InlinesToPlainText(inline.Children)
                        : inline.Href;
                    result.Add(MakeRun(linkText, bold, italic, code));
                    break;

                case MarkdownInlineType.Image:
                    result.Add(MakeRun($"[{inline.Alt}]", bold, italic, code));
                    break;

                case MarkdownInlineType.HardBreak:
                case MarkdownInlineType.SoftBreak:
                    result.Add(new WordRun { Text = " " });
                    break;

                case MarkdownInlineType.RawHtml:
                    // 忽略 HTML 标签，只保留文本
                    break;
            }
        }
        return result;
    }

    private static WordRun MakeRun(String text, Boolean bold, Boolean italic, Boolean code)
    {
        if (!bold && !italic && !code)
            return new WordRun { Text = text };
        return new WordRun
        {
            Text = text,
            Properties = new WordRunProperties
            {
                Bold = bold,
                Italic = italic,
                FontName = code ? "Courier New" : null,
                FontSize = code ? 10f : null,
            },
        };
    }

    private static String InlinesToPlainText(List<MarkdownInline> inlines)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var inline in inlines)
        {
            switch (inline.Type)
            {
                case MarkdownInlineType.Text:
                case MarkdownInlineType.Code:
                case MarkdownInlineType.RawHtml:
                    sb.Append(inline.Text);
                    break;
                case MarkdownInlineType.Strong:
                case MarkdownInlineType.Emphasis:
                case MarkdownInlineType.StrongEmphasis:
                case MarkdownInlineType.Strikethrough:
                case MarkdownInlineType.Link:
                    sb.Append(InlinesToPlainText(inline.Children));
                    break;
                case MarkdownInlineType.Image:
                    sb.Append(inline.Alt);
                    break;
                case MarkdownInlineType.HardBreak:
                case MarkdownInlineType.SoftBreak:
                    sb.Append(' ');
                    break;
            }
        }
        return sb.ToString();
    }

    private static List<String> CollectListItemTexts(MarkdownBlock listBlock)
    {
        var items = new List<String>();
        foreach (var item in listBlock.Children)
        {
            if (item.Type != MarkdownBlockType.ListItem) continue;
            // 取第一段文本
            if (item.Inlines.Count > 0)
                items.Add(InlinesToPlainText(item.Inlines));
            else if (item.Children.Count > 0)
                items.Add(InlinesToPlainText(item.Children[0].Inlines));
        }
        return items;
    }
    #endregion
}
