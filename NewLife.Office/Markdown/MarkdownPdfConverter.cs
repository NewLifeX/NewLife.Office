using System.Collections.Generic;
using System.Linq;

namespace NewLife.Office.Markdown;

/// <summary>Markdown → PDF 转换器（MD03-03）</summary>
/// <remarks>
/// 将 <see cref="MarkdownDocument"/> 转换为 PDF 字节数组，
/// 标题/段落/列表/表格/代码块/引用块映射到对应的 PDF 元素。
/// </remarks>
public sealed class MarkdownPdfConverter
{
    #region 属性
    /// <summary>H1 字号，默认 20pt</summary>
    public Single H1FontSize { get; set; } = 20f;
    /// <summary>H2 字号，默认 16pt</summary>
    public Single H2FontSize { get; set; } = 16f;
    /// <summary>H3 字号，默认 14pt</summary>
    public Single H3FontSize { get; set; } = 14f;
    /// <summary>H4–H6 字号，默认 12pt</summary>
    public Single H46FontSize { get; set; } = 12f;
    /// <summary>正文字号，默认 11pt</summary>
    public Single BodyFontSize { get; set; } = 11f;
    /// <summary>代码块字号，默认 9pt</summary>
    public Single CodeFontSize { get; set; } = 9f;
    #endregion

    #region 公开方法
    /// <summary>将 Markdown 文档转换为 PDF 字节数组</summary>
    /// <param name="doc">Markdown 文档</param>
    /// <returns>PDF 字节数组</returns>
    public Byte[] ToBytes(MarkdownDocument doc)
    {
        using var pdf = new PdfFluentDocument();
        Convert(doc, pdf);
        using var ms = new MemoryStream();
        pdf.Save(ms);
        return ms.ToArray();
    }

    /// <summary>将 Markdown 文档写入 PDF 文档对象</summary>
    /// <param name="doc">Markdown 文档</param>
    /// <param name="pdf">目标 PDF 文档</param>
    public void Convert(MarkdownDocument doc, PdfFluentDocument pdf)
    {
        foreach (var block in doc.Blocks)
        {
            WriteBlock(block, pdf, 0);
        }
    }
    #endregion

    #region 块处理
    private void WriteBlock(MarkdownBlock block, PdfFluentDocument pdf, Int32 depth)
    {
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
                var headingText = InlinesToPlainText(block.Inlines);
                var headingSize = block.Level switch
                {
                    1 => H1FontSize,
                    2 => H2FontSize,
                    3 => H3FontSize,
                    _ => H46FontSize,
                };
                pdf.AddEmptyLine(4f);
                pdf.AddText(headingText, headingSize);
                pdf.AddEmptyLine(4f);
                break;

            case MarkdownBlockType.Paragraph:
                var paraText = InlinesToPlainText(block.Inlines);
                if (!String.IsNullOrWhiteSpace(paraText))
                {
                    pdf.AddText(paraText, BodyFontSize);
                    pdf.AddEmptyLine(4f);
                }
                break;

            case MarkdownBlockType.CodeBlock:
                pdf.AddEmptyLine(2f);
                foreach (var line in block.RawText.Split('\n'))
                {
                    pdf.AddText(line.TrimEnd(), CodeFontSize, indentX: 20f);
                }
                pdf.AddEmptyLine(2f);
                break;

            case MarkdownBlockType.BlockQuote:
                foreach (var child in block.Children)
                {
                    // 引用块内容以缩进形式呈现
                    if (child.Type == MarkdownBlockType.Paragraph)
                    {
                        var quoteText = InlinesToPlainText(child.Inlines);
                        pdf.AddText("  " + quoteText, BodyFontSize, indentX: 20f);
                        pdf.AddEmptyLine(2f);
                    }
                    else
                        WriteBlock(child, pdf, depth + 1);
                }
                break;

            case MarkdownBlockType.BulletList:
                foreach (var item in block.Children)
                {
                    if (item.Type != MarkdownBlockType.ListItem) continue;
                    var itemText = item.Inlines.Count > 0
                        ? InlinesToPlainText(item.Inlines)
                        : (item.Children.Count > 0 ? InlinesToPlainText(item.Children[0].Inlines) : "");
                    pdf.AddText("• " + itemText, BodyFontSize, indentX: 20f);
                }
                pdf.AddEmptyLine(4f);
                break;

            case MarkdownBlockType.OrderedList:
                var idx = block.OrderedStart;
                foreach (var item in block.Children)
                {
                    if (item.Type != MarkdownBlockType.ListItem) continue;
                    var itemText = item.Inlines.Count > 0
                        ? InlinesToPlainText(item.Inlines)
                        : (item.Children.Count > 0 ? InlinesToPlainText(item.Children[0].Inlines) : "");
                    pdf.AddText($"{idx++}. " + itemText, BodyFontSize, indentX: 20f);
                }
                pdf.AddEmptyLine(4f);
                break;

            case MarkdownBlockType.Table:
                WriteTable(block, pdf);
                break;

            case MarkdownBlockType.ThematicBreak:
                pdf.AddEmptyLine(4f);
                pdf.AddText("─────────────────────────────────", BodyFontSize);
                pdf.AddEmptyLine(4f);
                break;

            case MarkdownBlockType.HtmlBlock:
                if (!String.IsNullOrWhiteSpace(block.RawText))
                    pdf.AddText(block.RawText.Trim(), BodyFontSize);
                break;
        }
    }

    private void WriteTable(MarkdownBlock tableBlock, PdfFluentDocument pdf)
    {
        var rows = new List<String[]>();
        foreach (var row in tableBlock.Children)
        {
            if (row.Type != MarkdownBlockType.TableRow) continue;
            var cells = row.Children
                .Where(c => c.Type == MarkdownBlockType.TableCell)
                .Select(c => InlinesToPlainText(c.Inlines))
                .ToArray();
            rows.Add(cells);
        }
        if (rows.Count > 0)
        {
            pdf.AddTable(rows, firstRowHeader: true);
            pdf.AddEmptyLine(4f);
        }
    }
    #endregion

    #region 行内处理
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
    #endregion
}
