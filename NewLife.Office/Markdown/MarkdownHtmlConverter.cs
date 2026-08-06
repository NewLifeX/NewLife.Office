using System;
using System.Collections.Generic;
using System.Text;

namespace NewLife.Office.Markdown;

/// <summary>Markdown转HTML转换器</summary>
internal sealed class MarkdownHtmlConverter
{
    #region 属性
    private readonly MarkdownHtmlOptions _options;

    /// <summary>当前列表是否松散（CommonMark：松散列表项内容渲染为 &lt;p&gt;）</summary>
    private Boolean _looseList;
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
                sb.Append("<p").Append(BlockAttributes(block)).Append('>');
                RenderInlines(sb, block.Inlines);
                sb.AppendLine("</p>");
                break;
            case MarkdownBlockType.CodeBlock:
                RenderCodeBlock(sb, block);
                break;
            case MarkdownBlockType.BlockQuote:
                sb.Append("<blockquote").Append(BlockAttributes(block)).AppendLine(">");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</blockquote>");
                break;
            case MarkdownBlockType.BulletList:
                var bl = (BulletListBlock)block;
                var savedLoose = _looseList;
                _looseList = bl.IsLoose;
                sb.AppendLine("<ul>");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</ul>");
                _looseList = savedLoose;
                break;
            case MarkdownBlockType.OrderedList:
                var ol = (OrderedListBlock)block;
                var start = ol.OrderedStart > 1 ? " start=\"" + ol.OrderedStart + "\"" : "";
                savedLoose = _looseList;
                _looseList = ol.IsLoose;
                sb.AppendLine("<ol" + start + ">");
                foreach (var child in block.Children)
                {
                    RenderBlock(sb, child);
                }
                sb.AppendLine("</ol>");
                _looseList = savedLoose;
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
                var hb = (HtmlBlock)block;
                sb.AppendLine(hb.RawText);
                break;
            case MarkdownBlockType.MathBlock:
                var mb = (MathBlock)block;
                sb.AppendLine("<div class=\"math\">$$");
                sb.AppendLine(HtmlEncode(mb.Content));
                sb.AppendLine("$$</div>");
                break;
            case MarkdownBlockType.FootnoteDefinition:
                var fd = (FootnoteDefinitionBlock)block;
                sb.Append("<div class=\"footnote\" id=\"fn-").Append(HtmlEncode(fd.Id)).Append("\"><sup>")
                  .Append(HtmlEncode(fd.Id)).Append("</sup>: ");
                RenderInlines(sb, fd.Definition);
                sb.AppendLine("</div>");
                break;
            case MarkdownBlockType.DefinitionList:
                sb.AppendLine("<dl>");
                foreach (var child in block.Children)
                {
                    if (child is DefinitionTermBlock dt)
                    {
                        sb.Append("<dt>");
                        RenderInlines(sb, dt.Inlines);
                        sb.AppendLine("</dt>");
                    }
                    else if (child is DefinitionDescriptionBlock dd)
                    {
                        sb.Append("<dd>");
                        RenderInlines(sb, dd.Inlines);
                        sb.AppendLine("</dd>");
                    }
                }
                sb.AppendLine("</dl>");
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
        var h = (HeadingBlock)block;
        var level = h.Level < 1 ? 1 : h.Level > 6 ? 6 : h.Level;
        var tag = "h" + level;
        // 生成可链接的锚点 id
        var id = block.GetPlainText().ToLower()
            .Replace(' ', '-')
            .Replace(".", "")
            .Replace(",", "")
            .Replace("(", "")
            .Replace(")", "")
            .Replace("/", "");
        sb.Append("<" + tag + " id=\"" + HtmlEncode(id) + "\"" + BlockAttributes(block) + ">");
        RenderInlines(sb, block.Inlines);
        sb.AppendLine("</" + tag + ">");
    }

    /// <summary>渲染代码块</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">代码块</param>
    private void RenderCodeBlock(StringBuilder sb, MarkdownBlock block)
    {
        var cb = (CodeBlock)block;
        var codeAttr = "";
        if (_options.AddLanguageClass && !String.IsNullOrEmpty(cb.Language))
            codeAttr = " class=\"language-" + HtmlEncode(cb.Language) + "\"";

        sb.Append("<pre><code" + codeAttr + ">");
        sb.Append(HtmlEncode(cb.RawText ?? ""));
        sb.AppendLine("</code></pre>");
    }

    /// <summary>渲染列表项</summary>
    /// <param name="sb">输出</param>
    /// <param name="block">列表项块</param>
    private void RenderListItem(StringBuilder sb, MarkdownBlock block)
    {
        var li = (ListItemBlock)block;
        sb.Append("<li>");
        if (li.IsTaskItem)
        {
            var checked_ = li.IsChecked ? " checked=\"\"" : "";
            sb.Append("<input type=\"checkbox\" disabled=\"\"" + checked_ + " /> ");
        }

        if (block.Children.Count == 0)
        {
            // 松散列表：简单项内容包裹 <p>（CommonMark）
            if (_looseList && block.Inlines.Count > 0)
            {
                sb.Append("<p>");
                RenderInlines(sb, block.Inlines);
                sb.Append("</p>");
            }
            else
            {
                RenderInlines(sb, block.Inlines);
            }
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
            var tc = (TableCellBlock)cell;
            var align = tc.Alignment == null ? "" : " style=\"text-align:" + tc.Alignment + "\"";
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
                RenderTextWithAbbr(sb, inline.Text ?? "");
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
            case MarkdownInlineType.AutoLink:
                var href2 = inline.Href ?? inline.Text ?? "";
                sb.Append("<a href=\"").Append(HtmlEncode(href2)).Append("\">")
                  .Append(HtmlEncode(inline.Text ?? "")).Append("</a>");
                break;
            case MarkdownInlineType.MathInline:
                sb.Append("<span class=\"math\">$").Append(HtmlEncode(inline.Text ?? "")).Append("$</span>");
                break;
            case MarkdownInlineType.FootnoteRef:
                sb.Append("<sup><a href=\"#fn-").Append(HtmlEncode(inline.Text ?? "")).Append("\" id=\"fnref-")
                  .Append(HtmlEncode(inline.Text ?? "")).Append("\">[")
                  .Append(HtmlEncode(inline.Text ?? "")).Append("]</a></sup>");
                break;
            default:
                RenderTextWithAbbr(sb, inline.Text ?? "");
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
    internal static String HtmlEncode(String text)
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
    /// <summary>生成块级元素的属性字符串 (MD05-06)</summary>
    private static String BlockAttributes(MarkdownBlock block)
    {
        if (String.IsNullOrEmpty(block.Attributes)) return String.Empty;

        var attr = block.Attributes;
        var sb = new StringBuilder();
        var parts = attr.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var classes = new List<String>();
        var id = String.Empty;

        foreach (var part in parts)
        {
            if (part.StartsWith("."))
            {
                var cls = part[1..];
                if (cls.Length > 0) classes.Add(cls);
            }
            else if (part.StartsWith("#"))
            {
                id = part[1..];
            }
            else if (part.Contains('='))
            {
                var eqIdx = part.IndexOf('=');
                var key = part[..eqIdx];
                var value = part[(eqIdx + 1)..].Trim('"', '\'');
                sb.Append(' ').Append(key).Append("=\"").Append(HtmlEncode(value)).Append('"');
            }
        }

        if (classes.Count > 0)
            sb.Append(" class=\"").Append(HtmlEncode(String.Join(" ", classes))).Append('"');
        if (id.Length > 0)
            sb.Append(" id=\"").Append(HtmlEncode(id)).Append('"');

        return sb.ToString();
    }

    /// <summary>渲染文本，对已知缩写自动添加 &lt;abbr&gt; 标签 (MD05-08)</summary>
    private void RenderTextWithAbbr(StringBuilder sb, String text)
    {
        if (String.IsNullOrEmpty(text))
        {
            sb.Append(HtmlEncode(text));
            return;
        }

        var abbrs = _options.Abbreviations;
        if (abbrs == null || abbrs.Count == 0)
        {
            sb.Append(HtmlEncode(text));
            return;
        }

        // 简单替换：按缩写文本长度降序排列，避免短缩写误匹配
        var sorted = abbrs.OrderByDescending(kv => kv.Key.Length).ToList();
        var i = 0;
        while (i < text.Length)
        {
            var matched = false;
            foreach (var kv in sorted)
            {
                if (i + kv.Key.Length <= text.Length &&
                    String.Compare(text, i, kv.Key, 0, kv.Key.Length, StringComparison.Ordinal) == 0)
                {
                    // 确保是单词边界
                    var beforeOk = i == 0 || !Char.IsLetterOrDigit(text[i - 1]);
                    var afterOk = i + kv.Key.Length >= text.Length || !Char.IsLetterOrDigit(text[i + kv.Key.Length]);
                    if (beforeOk && afterOk)
                    {
                        sb.Append("<abbr title=\"").Append(HtmlEncode(kv.Value)).Append("\">")
                          .Append(HtmlEncode(kv.Key)).Append("</abbr>");
                        i += kv.Key.Length;
                        matched = true;
                        break;
                    }
                }
            }
            if (!matched)
            {
                sb.Append(HtmlEncode(text[i].ToString()));
                i++;
            }
        }
    }
    #endregion
}
