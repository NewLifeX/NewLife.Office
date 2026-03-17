using System;
using System.Collections.Generic;
using System.Text;

namespace NewLife.Office.Markdown;

/// <summary>Markdown转HTML转换器</summary>
internal sealed class MarkdownHtmlConverter
{
    #region 属性
    private readonly MarkdownHtmlOptions _options;
    #endregion

    #region 构造
    /// <summary>实例化转换器</summary>
    /// <param name="options">HTML渲染选项</param>
    public MarkdownHtmlConverter(MarkdownHtmlOptions options) => _options = options;
    #endregion

    #region 方法
    /// <summary>将Markdown文档转换为HTML片段</summary>
    /// <param name="doc">Markdown文档</param>
    /// <returns>HTML字符串</returns>
    public String Convert(MarkdownDocument doc)
    {
        var sb = new StringBuilder();
        foreach (var block in doc.Blocks)
        {
            RenderBlock(sb, block);
        }

        return sb.ToString();
    }

    /// <summary>渲染单个块节点</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">块节点</param>
    private void RenderBlock(StringBuilder sb, MarkdownBlock block)
    {
        switch (block.Type)
        {
            case MarkdownBlockType.Heading:
                RenderHeading(sb, block);
                break;
            case MarkdownBlockType.Paragraph:
                sb.Append("<p>");
                RenderInlines(sb, block.Inlines);
                sb.AppendLine("</p>");
                break;
            case MarkdownBlockType.CodeBlock:
                RenderCodeBlock(sb, block);
                break;
            case MarkdownBlockType.BlockQuote:
                sb.AppendLine("<blockquote>");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</blockquote>");
                break;
            case MarkdownBlockType.BulletList:
                sb.AppendLine("<ul>");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</ul>");
                break;
            case MarkdownBlockType.OrderedList:
                var start = block.OrderedStart > 1 ? " start=\"" + block.OrderedStart + "\"" : "";
                sb.AppendLine("<ol" + start + ">");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</ol>");
                break;
            case MarkdownBlockType.ListItem:
                RenderListItem(sb, block);
                break;
            case MarkdownBlockType.Table:
                RenderTable(sb, block);
                break;
            case MarkdownBlockType.ThematicBreak:
                sb.AppendLine("<hr />");
                break;
            case MarkdownBlockType.HtmlBlock:
                sb.AppendLine(block.RawText);
                break;
            default:
                break;
        }
    }

    /// <summary>渲染标题块</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">标题块</param>
    private void RenderHeading(StringBuilder sb, MarkdownBlock block)
    {
        var level = block.Level < 1 ? 1 : block.Level > 6 ? 6 : block.Level;
        var tag = "h" + level;
        // 生成可链接的锚点 id
        var id = block.GetPlainText().ToLower()
            .Replace(' ', '-')
            .Replace(".", "")
            .Replace(",", "")
            .Replace("(", "")
            .Replace(")", "")
            .Replace("/", "");
        sb.Append("<" + tag + " id=\"" + HtmlEncode(id) + "\">");
        RenderInlines(sb, block.Inlines);
        sb.AppendLine("</" + tag + ">");
    }

    /// <summary>渲染代码块</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">代码块</param>
    private void RenderCodeBlock(StringBuilder sb, MarkdownBlock block)
    {
        var codeAttr = "";
        if (_options.AddLanguageClass && !String.IsNullOrEmpty(block.Language))
            codeAttr = " class=\"language-" + HtmlEncode(block.Language) + "\"";

        sb.Append("<pre><code" + codeAttr + ">");
        sb.Append(HtmlEncode(block.RawText ?? ""));
        sb.AppendLine("</code></pre>");
    }

    /// <summary>渲染列表项</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">列表项块</param>
    private void RenderListItem(StringBuilder sb, MarkdownBlock block)
    {
        sb.Append("<li>");
        if (block.IsTaskItem)
        {
            var checked_ = block.IsChecked ? " checked=\"\"" : "";
            sb.Append("<input type=\"checkbox\" disabled=\"\"" + checked_ + " /> ");
        }

        if (block.Children.Count == 0)
        {
            RenderInlines(sb, block.Inlines);
        }
        else
        {
            RenderInlines(sb, block.Inlines);
            foreach (var child in block.Children)
            {
                RenderBlock(sb, child);
            }
        }

        sb.AppendLine("</li>");
    }

    /// <summary>渲染表格</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">表格块</param>
    private void RenderTable(StringBuilder sb, MarkdownBlock block)
    {
        sb.AppendLine("<table>");
        var rows = block.Children;
        if (rows.Count == 0)
        {
            sb.AppendLine("</table>");
            return;
        }

        // 第一行为表头
        sb.AppendLine("<thead>");
        RenderTableRow(sb, rows[0], isHeader: true);
        sb.AppendLine("</thead>");

        if (rows.Count > 1)
        {
            sb.AppendLine("<tbody>");
            for (var i = 1; i < rows.Count; i++)
            {
                RenderTableRow(sb, rows[i], isHeader: false);
            }
            sb.AppendLine("</tbody>");
        }

        sb.AppendLine("</table>");
    }

    /// <summary>渲染表格行</summary>
    /// <param name="sb">输出</param>
    /// <param name="row">行块</param>
    /// <param name="isHeader">是否为表头行</param>
    private void RenderTableRow(StringBuilder sb, MarkdownBlock row, Boolean isHeader)
    {
        sb.AppendLine("<tr>");
        var tag = isHeader ? "th" : "td";
        foreach (var cell in row.Children)
        {
            var align = cell.Alignment == null ? "" : " style=\"text-align:" + cell.Alignment + "\"";
            sb.Append("<" + tag + align + ">");
            RenderInlines(sb, cell.Inlines);
            sb.AppendLine("</" + tag + ">");
        }
        sb.AppendLine("</tr>");
    }

    /// <summary>渲染内联节点列表</summary>
    /// <param name="sb">输出</param>
    /// <param name="inlines">内联列表</param>
    private void RenderInlines(StringBuilder sb, List<MarkdownInline> inlines)
    {
        foreach (var inline in inlines)
        {
            RenderInline(sb, inline);
        }
    }

    /// <summary>渲染单个内联节点</summary>
    /// <param name="sb">输出</param>
    /// <param name="inline">内联节点</param>
    private void RenderInline(StringBuilder sb, MarkdownInline inline)
    {
        switch (inline.Type)
        {
            case MarkdownInlineType.Text:
                sb.Append(HtmlEncode(inline.Text ?? ""));
                break;
            case MarkdownInlineType.Strong:
                sb.Append("<strong>");
                RenderInlines(sb, inline.Children);
                sb.Append("</strong>");
                break;
            case MarkdownInlineType.Emphasis:
                sb.Append("<em>");
                RenderInlines(sb, inline.Children);
                sb.Append("</em>");
                break;
            case MarkdownInlineType.StrongEmphasis:
                sb.Append("<strong><em>");
                RenderInlines(sb, inline.Children);
                sb.Append("</em></strong>");
                break;
            case MarkdownInlineType.Code:
                sb.Append("<code>");
                sb.Append(HtmlEncode(inline.Text ?? ""));
                sb.Append("</code>");
                break;
            case MarkdownInlineType.Strikethrough:
                sb.Append("<del>");
                RenderInlines(sb, inline.Children);
                sb.Append("</del>");
                break;
            case MarkdownInlineType.Link:
                RenderLink(sb, inline);
                break;
            case MarkdownInlineType.Image:
                RenderImage(sb, inline);
                break;
            case MarkdownInlineType.HardBreak:
                sb.AppendLine("<br />");
                break;
            case MarkdownInlineType.SoftBreak:
                sb.Append(" ");
                break;
            case MarkdownInlineType.RawHtml:
                sb.Append(inline.Text ?? "");
                break;
            default:
                sb.Append(HtmlEncode(inline.Text ?? ""));
                break;
        }
    }

    /// <summary>渲染链接</summary>
    /// <param name="sb">输出</param>
    /// <param name="inline">链接内联节点</param>
    private void RenderLink(StringBuilder sb, MarkdownInline inline)
    {
        var href = inline.Href ?? "";
        if (_options.SafeLinks && IsDangerousUrl(href))
        {
            // 危险链接仅输出文本
            RenderInlines(sb, inline.Children);
            return;
        }

        var target = "";
        var rel = "";
        if (_options.ExternalLinkTarget && IsExternalLink(href))
        {
            target = " target=\"_blank\"";
            rel = " rel=\"noopener noreferrer\"";
        }

        var title = String.IsNullOrEmpty(inline.Title) ? "" : " title=\"" + HtmlEncode(inline.Title) + "\"";
        sb.Append("<a href=\"" + HtmlEncode(href) + "\"" + title + target + rel + ">");
        if (inline.Children.Count > 0)
            RenderInlines(sb, inline.Children);
        else
            sb.Append(HtmlEncode(href));
        sb.Append("</a>");
    }

    /// <summary>渲染图片</summary>
    /// <param name="sb">输出</param>
    /// <param name="inline">图片内联节点</param>
    private void RenderImage(StringBuilder sb, MarkdownInline inline)
    {
        var src = inline.Href ?? "";
        if (_options.SafeLinks && IsDangerousUrl(src))
        {
            sb.Append("![" + HtmlEncode(inline.Alt ?? "") + "]");
            return;
        }

        var alt = HtmlEncode(inline.Alt ?? "");
        var title = String.IsNullOrEmpty(inline.Title) ? "" : " title=\"" + HtmlEncode(inline.Title) + "\"";
        sb.Append("<img src=\"" + HtmlEncode(src) + "\" alt=\"" + alt + "\"" + title + " />");
    }
    #endregion

    #region 辅助
    /// <summary>HTML编码文本内容</summary>
    /// <param name="text">原始文本</param>
    /// <returns>编码后文本</returns>
    private static String HtmlEncode(String text)
    {
        if (text == null || text.Length == 0) return "";
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    /// <summary>判断是否为危险URL（javascript:/data:）</summary>
    /// <param name="url">URL字符串</param>
    /// <returns>是否危险</returns>
    private static Boolean IsDangerousUrl(String url)
    {
        if (String.IsNullOrEmpty(url)) return false;
        var lower = url.TrimStart().ToLower();
        return lower.StartsWith("javascript:") || lower.StartsWith("vbscript:") || lower.StartsWith("data:");
    }

    /// <summary>判断是否为外部链接</summary>
    /// <param name="url">URL字符串</param>
    /// <returns>是否外部链接</returns>
    private static Boolean IsExternalLink(String url)
    {
        if (String.IsNullOrEmpty(url)) return false;
        return url.StartsWith("http://") || url.StartsWith("https://") || url.StartsWith("//");
    }
    #endregion
}
