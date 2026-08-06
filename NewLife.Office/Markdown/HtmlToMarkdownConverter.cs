using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office.Markdown;

/// <summary>HTML → Markdown 转换器</summary>
/// <remarks>
/// 将 HTML 片段解析并转换为 Markdown 文档对象模型（<see cref="MarkdownDocument"/>），
/// 支持 CommonMark + GFM 扩展语法。
/// <para>支持的 HTML 标签：h1-h6/p/em/strong/b/i/a/img/ul/ol/li/table/tr/td/th/blockquote/code/pre/hr/br/del/s/sub/sup</para>
/// </remarks>
public sealed class HtmlToMarkdownConverter
{
    #region 字段
    private readonly HtmlToMarkdownOptions _options;
    private String _html = String.Empty;
    private Int32 _pos;
    private readonly HashSet<String> _allowedSchemes = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>待输出块队列（details 容器解析出多个内部块时使用）</summary>
    private List<MarkdownBlock>? _pendingBlocks;
    #endregion

    #region 构造
    /// <summary>实例化转换器</summary>
    /// <param name="options">转换选项（null 使用默认）</param>
    public HtmlToMarkdownConverter(HtmlToMarkdownOptions? options = null)
    {
        _options = options ?? new HtmlToMarkdownOptions();
        foreach (var scheme in _options.WhitelistUriSchemes.Split(';'))
        {
            var s = scheme.Trim();
            if (s.Length > 0) _allowedSchemes.Add(s);
        }
    }
    #endregion

    #region 公共方法
    /// <summary>将 HTML 字符串转换为 MarkdownDocument</summary>
    /// <param name="html">HTML 文本</param>
    /// <returns>Markdown 文档对象</returns>
    public MarkdownDocument Convert(String html)
    {
        if (String.IsNullOrEmpty(html)) return new MarkdownDocument();

        _html = html;
        _pos = 0;
        _pendingBlocks = null;

        var doc = new MarkdownDocument();
        SkipWhitespaceAndNewlines();

        while (_pos < _html.Length || _pendingBlocks is { Count: > 0 })
        {
            var block = _pendingBlocks is { Count: > 0 } ? DrainPendingBlock() : ParseBlockElement();
            if (block != null)
                doc.Blocks.Add(block);
            if (_pos < _html.Length) SkipWhitespaceAndNewlines();
        }

        return doc;
    }

    /// <summary>将 HTML 字符串直接转换为 Markdown 文本</summary>
    /// <param name="html">HTML 文本</param>
    /// <returns>Markdown 字符串</returns>
    public String ConvertToString(String html)
    {
        var doc = Convert(html);
        return doc.ToMarkdown();
    }
    #endregion

    #region 块级解析
    private MarkdownBlock? ParseBlockElement()
    {
        if (_pos >= _html.Length) return null;

        // 跳过空白
        SkipWhitespaceAndNewlines();

        if (_pos >= _html.Length) return null;

        // 不是标签，作为段落处理
        if (_html[_pos] != '<')
            return ParseParagraph();

        // 注释
        if (IsAt("<!--"))
        {
            var end = _html.IndexOf("-->", _pos, StringComparison.Ordinal);
            if (end >= 0)
            {
                _pos = end + 3;
                SkipWhitespaceAndNewlines();
                return ParseBlockElement();
            }
            // 未闭合的注释，跳过剩余
            _pos = _html.Length;
            return null;
        }

        // 解析标签
        var tag = ParseTag();
        if (tag == null)
        {
            // 无效的 <，作为文本处理
            return ParseParagraph();
        }

        return ConvertTagToBlock(tag);
    }

    private MarkdownBlock? ConvertTagToBlock(HtmlTag tag)
    {
        if (tag.IsClosing) return null;

        var name = tag.Name;

        // 自闭合块级标签（hr/br）
        if (tag.IsSelfClosing)
        {
            return name switch
            {
                "hr" => ParseHr(tag),
                "br" => ParseBr(),
                _ => null,
            };
        }

        switch (name)
        {
            case "h1": return ParseHeading(1);
            case "h2": return ParseHeading(2);
            case "h3": return ParseHeading(3);
            case "h4": return ParseHeading(4);
            case "h5": return ParseHeading(5);
            case "h6": return ParseHeading(6);
            case "p": return ParseParagraph();
            case "pre": return ParsePreBlock();
            case "blockquote": return ParseBlockQuote();
            case "ul": return ParseList(false);
            case "ol": return ParseList(true);
            case "table": return ParseTable();
            case "dl": return ParseDefinitionList();
            case "details": return ParseDetailsContainer();
            case "summary": return ParseSummary();
            case "hr": return ParseHr(tag);
            case "br": return ParseBr();
            // 行内标签：退回到段落解析（由 ParseParagraph 内部处理行内元素）
            case "strong" or "b" or "em" or "i" or "a" or "img" or "code"
                or "del" or "s" or "strike" or "sub" or "sup"
                or "span" or "label" or "small" or "mark" or "u"
                or "abbr" or "cite" or "dfn" or "kbd" or "samp" or "var"
                or "time" or "q" or "ins" or "ruby" or "rt" or "rp":
                return ParseParagraph();
            default:
                return HandleUnknownTag(tag);
        }
    }

    private MarkdownBlock? HandleUnknownTag(HtmlTag tag)
    {
        switch (_options.UnknownTags)
        {
            case UnknownTagStrategy.PassThrough:
                return MarkdownBlock.CreateHtmlBlock(tag.RawText);
            case UnknownTagStrategy.Drop:
                // 丢弃外层标签，返回 null 让上层循环继续解析内部内容
                // 闭合标签由外层 ParseBlockElement 的 while 循环通过 IsBlockTag 检查跳过
                return null;
            case UnknownTagStrategy.Escape:
                return new ParagraphBlock([MarkdownInline.CreateText(EscapeHtml(tag.RawText))]);
            default:
                return null;
        }
    }
    #endregion

    #region 标题
    private MarkdownBlock ParseHeading(Int32 level)
    {
        var inlines = ParseInlineUntilClose("h" + level);
        return new HeadingBlock(level, inlines);
    }
    #endregion

    #region 段落
    private MarkdownBlock ParseParagraph()
    {
        var inlines = new List<MarkdownInline>();

        while (_pos < _html.Length)
        {
            if (_html[_pos] == '<')
            {
                // 检查是否是块级标签结束
                if (IsAt("</"))
                {
                    var peekTag = PeekClosingTag();
                    if (peekTag != null && IsBlockTag(peekTag))
                        break;
                }
                if (IsAt("<br") || IsAt("<BR"))
                {
                    var brTag = ParseTag();
                    if (brTag != null && !brTag.IsClosing)
                    {
                        inlines.Add(MarkdownInline.CreateHardBreak());
                        continue;
                    }
                }
                if (IsAt("<hr") || IsAt("<HR"))
                {
                    break;
                }
                // 检查是否是块级开始标签
                if (PeekBlockStartTag() != null)
                    break;

                var tag = ParseTag();
                if (tag == null)
                {
                    inlines.Add(MarkdownInline.CreateText("<"));
                    _pos++;
                    continue;
                }
                // 自闭合标签：作为行内元素处理（img/br 等）
                if (tag.IsSelfClosing)
                {
                    var inlinesFromTag = ConvertTagToInlines(tag);
                    if (inlinesFromTag != null)
                        inlines.AddRange(inlinesFromTag);
                    continue;
                }
                if (tag.IsClosing)
                {
                    // 结束标签属于外层，停止段落
                    _pos -= tag.RawText.Length; // 回退
                    break;
                }
                var inlinesFromTag2 = ConvertTagToInlines(tag);
                if (inlinesFromTag2 != null)
                    inlines.AddRange(inlinesFromTag2);
                continue;
            }

            // 普通文本
            var text = ReadUntil('<');
            if (text.Length > 0)
                inlines.Add(MarkdownInline.CreateText(DecodeHtmlEntities(text)));
        }

        return new ParagraphBlock(NormalizeInlines(inlines));
    }
    #endregion

    #region 代码块
    private MarkdownBlock ParsePreBlock()
    {
        var sb = new StringBuilder();
        var lang = String.Empty;

        // 检查是否有 <code> 子元素
        var savedPos = _pos;
        SkipWhitespaceAndNewlines();
        if (_pos < _html.Length && _html[_pos] == '<')
        {
            var codeTag = ParseTag();
            if (codeTag != null && !codeTag.IsClosing &&
                (codeTag.Name == "code" || codeTag.Name == "CODE"))
            {
                // 提取 class 中的语言
                lang = ExtractLanguageFromClass(codeTag.GetAttribute("class"));
            }
            else
            {
                _pos = savedPos;
            }
        }

        // 读取内容直到 </pre>
        while (_pos < _html.Length)
        {
            if (IsAt("</pre>") || IsAt("</PRE>"))
            {
                _pos += 6;
                break;
            }
            if (IsAt("</code>") || IsAt("</CODE>"))
            {
                _pos += 7;
                continue;
            }
            if (_html[_pos] == '<')
            {
                var saved = _pos;
                var tag = ParseTag();
                if (tag != null && !tag.IsClosing && tag.Name == "br")
                {
                    sb.AppendLine();
                    continue;
                }
                if (tag != null && !tag.IsClosing && tag.Name == "code")
                {
                    continue;
                }
                // 其他标签跳过
                _pos = saved;
                sb.Append(_html[_pos]);
                _pos++;
                continue;
            }
            sb.Append(_html[_pos]);
            _pos++;
        }

        var code = DecodeHtmlEntities(sb.ToString().Trim());
        return new CodeBlock(code, lang);
    }
    #endregion

    #region 引用块
    private MarkdownBlock ParseBlockQuote()
    {
        var children = new List<MarkdownBlock>();

        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;

            if (IsAt("</blockquote>") || IsAt("</BLOCKQUOTE>"))
            {
                _pos += 13;
                break;
            }

            var block = ParseBlockElement();
            if (block != null)
                children.Add(block);
        }

        return new BlockQuoteBlock(children);
    }
    #endregion

    #region 列表
    private MarkdownBlock ParseList(Boolean ordered)
    {
        var items = new List<MarkdownBlock>();

        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;

            var closeTag = ordered ? "</ol>" : "</ul>";
            if (IsAt(closeTag, StringComparison.OrdinalIgnoreCase))
            {
                _pos += closeTag.Length;
                break;
            }

            // 查找 <li>
            if (_html[_pos] == '<')
            {
                var tag = ParseTag();
                if (tag == null)
                {
                    _pos++;
                    continue;
                }

                if (!tag.IsClosing && (tag.Name == "li" || tag.Name == "LI"))
                {
                    var item = ParseListItem();
                    if (item != null)
                        items.Add(item);
                    continue;
                }
                if (tag.IsClosing && (tag.Name == "li" || tag.Name == "LI"))
                {
                    continue;
                }
                // 其他标签如 <ul>/<ol>（嵌套列表已由 ParseListItem 处理），跳过
                continue;
            }

            // 忽略非 < 字符
            _pos++;
        }

        if (items.Count == 0) return new ParagraphBlock([]);

        return ordered
            ? MarkdownBlock.CreateOrderedList(items)
            : MarkdownBlock.CreateBulletList(items);
    }

    private MarkdownBlock? ParseListItem()
    {
        var inlines = new List<MarkdownInline>();
        var isTask = false;
        var isChecked = false;

        // GFM 任务列表：<li> [空白] <input type="checkbox" [checked]> 文本
        var savedInputPos = _pos;
        SkipWhitespaceAndNewlines();
        if (IsAt("<input", StringComparison.OrdinalIgnoreCase))
        {
            var inputTag = ParseTag();
            if (inputTag != null && !inputTag.IsClosing &&
                inputTag.GetAttribute("type").Equals("checkbox", StringComparison.OrdinalIgnoreCase))
            {
                isTask = true;
                isChecked = inputTag.Attributes.ContainsKey("checked");
                SkipWhitespaceAndNewlines();
            }
            else
            {
                _pos = savedInputPos;
            }
        }
        else
        {
            _pos = savedInputPos;
        }

        while (_pos < _html.Length)
        {
            if (_html[_pos] == '<')
            {
                if (IsAt("</li>") || IsAt("</LI>"))
                {
                    _pos += 5;
                    break;
                }
                // 检查嵌套列表
                if (IsAt("<ul") || IsAt("<UL") || IsAt("<ol") || IsAt("<OL"))
                {
                    break;
                }
                // 检查 <li> 开始（嵌套列表）
                if (IsAt("<li") || IsAt("<LI"))
                {
                    break;
                }

                var tag = ParseTag();
                if (tag == null)
                {
                    inlines.Add(MarkdownInline.CreateText("<"));
                    _pos++;
                    continue;
                }
                if (tag.IsClosing) continue; // 跳过未知闭合标签

                var tagInlines = ConvertTagToInlines(tag);
                if (tagInlines != null)
                    inlines.AddRange(tagInlines);
                continue;
            }

            var text = ReadUntil('<');
            if (text.Length > 0)
                inlines.Add(MarkdownInline.CreateText(DecodeHtmlEntities(text)));
        }

        // 处理嵌套列表
        var children = new List<MarkdownBlock>();
        var hasChildren = false;
        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;

            if (_html[_pos] == '<')
            {
                if (IsAt("</li>") || IsAt("</LI>"))
                {
                    _pos += 5;
                    break;
                }
                if (IsAt("<ul") || IsAt("<UL"))
                {
                    var nestedList = ParseList(false);
                    children.Add(nestedList);
                    hasChildren = true;
                    continue;
                }
                if (IsAt("<ol") || IsAt("<OL"))
                {
                    var nestedList = ParseList(true);
                    children.Add(nestedList);
                    hasChildren = true;
                    continue;
                }
                break;
            }
            break;
        }

        if (hasChildren)
        {
            var item = ListItemBlock.CreateWithBlocks(children, isTask, isChecked);
            item.Inlines.AddRange(NormalizeInlines(inlines));
            return item;
        }

        return new ListItemBlock(NormalizeInlines(inlines), isTask, isChecked);
    }

    /// <summary>解析定义列表 dl/dt/dd → DefinitionListBlock（MD05-07）</summary>
    /// <returns>定义列表块</returns>
    private MarkdownBlock ParseDefinitionList()
    {
        var dl = new DefinitionListBlock();
        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;
            if (IsAt("</dl>") || IsAt("</DL>"))
            {
                _pos += 5;
                break;
            }
            if (_html[_pos] == '<')
            {
                var tag = ParseTag();
                if (tag == null) { _pos++; continue; }
                if (tag.IsClosing) continue;
                if (tag.Name == "dt" || tag.Name == "DT")
                {
                    dl.Children.Add(new DefinitionTermBlock(ParseInlineUntilClose("dt")));
                    continue;
                }
                if (tag.Name == "dd" || tag.Name == "DD")
                {
                    dl.Children.Add(new DefinitionDescriptionBlock(ParseInlineUntilClose("dd")));
                    continue;
                }
                // 未知标签跳过
                continue;
            }
            _pos++;
        }
        return dl;
    }

    /// <summary>解析 details 折叠容器：内部内容作为普通块解析（Drop 折叠语义，保留内容）</summary>
    /// <returns>首个内部块，其余经待输出队列返回</returns>
    private MarkdownBlock? ParseDetailsContainer()
    {
        var inner = ReadUntilCloseTag("details");
        _pendingBlocks = new HtmlToMarkdownConverter(_options).Convert(inner).Blocks;
        return DrainPendingBlock();
    }

    /// <summary>解析 summary 折叠标题 → 粗体段落</summary>
    /// <returns>段落块</returns>
    private MarkdownBlock ParseSummary()
    {
        var inlines = ParseInlineUntilClose("summary");
        if (inlines.Count == 0) return new ParagraphBlock([]);
        return new ParagraphBlock([MarkdownInline.CreateStrong(inlines)]);
    }

    /// <summary>取出待输出队列中的下一个块</summary>
    /// <returns>块，队列空返回 null</returns>
    private MarkdownBlock? DrainPendingBlock()
    {
        if (_pendingBlocks == null || _pendingBlocks.Count == 0) return null;
        var block = _pendingBlocks[0];
        _pendingBlocks.RemoveAt(0);
        return block;
    }
    #endregion

    #region 表格
    private MarkdownBlock ParseTable()
    {
        var rows = new List<MarkdownBlock>();
        var isHeader = true;

        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;

            if (IsAt("</table>") || IsAt("</TABLE>"))
            {
                _pos += 8;
                break;
            }

            if (_html[_pos] == '<')
            {
                var savedPos = _pos;
                var tag = ParseTag();
                if (tag == null) { _pos++; continue; }

                if (!tag.IsClosing && (tag.Name == "thead" || tag.Name == "THEAD"))
                {
                    isHeader = true;
                    continue;
                }
                if (tag.IsClosing && (tag.Name == "thead" || tag.Name == "THEAD"))
                {
                    continue;
                }
                if (!tag.IsClosing && (tag.Name == "tbody" || tag.Name == "TBODY"))
                {
                    isHeader = false;
                    continue;
                }
                if (tag.IsClosing && (tag.Name == "tbody" || tag.Name == "TBODY"))
                {
                    continue;
                }
                if (!tag.IsClosing && (tag.Name == "tr" || tag.Name == "TR"))
                {
                    var row = ParseTableRow(isHeader);
                    if (row != null)
                        rows.Add(row);
                    isHeader = false;
                    continue;
                }
                if (tag.IsClosing && (tag.Name == "tr" || tag.Name == "TR"))
                {
                    continue;
                }
                // 其他标签忽略
                _pos = savedPos + 1;
            }
            else
            {
                _pos++;
            }
        }

        if (rows.Count == 0) return new ParagraphBlock([]);

        var table = new TableBlock();
        table.Children.AddRange(rows);
        return table;
    }

    private MarkdownBlock? ParseTableRow(Boolean isHeader)
    {
        var cells = new List<MarkdownBlock>();
        var tagName = isHeader ? "th" : "td";

        while (_pos < _html.Length)
        {
            SkipWhitespaceAndNewlines();
            if (_pos >= _html.Length) break;

            if (IsAt("</tr>") || IsAt("</TR>"))
            {
                _pos += 5;
                break;
            }

            if (_html[_pos] == '<')
            {
                var savedPos = _pos;
                var tag = ParseTag();
                if (tag == null) { _pos++; continue; }

                if (!tag.IsClosing && (tag.Name == "th" || tag.Name == "TH" || tag.Name == "td" || tag.Name == "TD"))
                {
                    var cellInlines = ParseInlineUntilClose(tag.Name);
                    cells.Add(new TableCellBlock(NormalizeInlines(cellInlines), isHeader: isHeader || tag.Name == "th" || tag.Name == "TH"));
                    continue;
                }
                if (tag.IsClosing && (tag.Name == "th" || tag.Name == "TH" || tag.Name == "td" || tag.Name == "TD"))
                {
                    continue;
                }
                _pos = savedPos + 1;
            }
            else
            {
                _pos++;
            }
        }

        if (cells.Count == 0) return null;
        var row = new TableRowBlock();
        row.Children.AddRange(cells);
        return row;
    }
    #endregion

    #region 分隔线与换行
    private MarkdownBlock ParseHr(HtmlTag tag)
    {
        return new ThematicBreakBlock();
    }

    private MarkdownBlock ParseBr()
    {
        return new ParagraphBlock([MarkdownInline.CreateHardBreak()]);
    }
    #endregion

    #region 行内元素解析
    private List<MarkdownInline>? ConvertTagToInlines(HtmlTag tag)
    {
        if (tag.IsClosing && !tag.IsSelfClosing) return null;

        var name = tag.Name;

        // 自闭合行内标签（img/br）
        if (tag.IsSelfClosing)
        {
            return name switch
            {
                "img" => ParseImage(tag),
                "br" => [MarkdownInline.CreateHardBreak()],
                "hr" => null,
                _ => null,
            };
        }

        switch (name)
        {
            case "strong":
            case "b":
            case "B":
                var strongChildren = ParseInlineUntilClose(name);
                if (strongChildren.Count == 0) return null;
                return [MarkdownInline.CreateStrong(strongChildren)];

            case "em":
            case "i":
            case "I":
                var emChildren = ParseInlineUntilClose(name);
                if (emChildren.Count == 0) return null;
                return [MarkdownInline.CreateEmphasis(emChildren)];

            case "del":
            case "s":
            case "S":
            case "strike":
                var delChildren = ParseInlineUntilClose(name);
                if (delChildren.Count == 0) return null;
                return [MarkdownInline.CreateStrikethrough(delChildren)];

            case "code":
            case "tt":
                var codeText = ReadUntilCloseTag(name);
                return [MarkdownInline.CreateCode(DecodeHtmlEntities(codeText.Trim()))];

            case "a":
            case "A":
                return ParseLink(tag);

            case "img":
            case "IMG":
                return ParseImage(tag);

            case "br":
            case "BR":
                return [MarkdownInline.CreateHardBreak()];

            case "sub":
            case "SUB":
                var subChildren = ParseInlineUntilClose(name);
                return subChildren.Count > 0
                    ? [MarkdownInline.CreateRawHtml("<sub>" + InlinesToPlainText(subChildren) + "</sub>")]
                    : null;

            case "sup":
            case "SUP":
                var supChildren = ParseInlineUntilClose(name);
                return supChildren.Count > 0
                    ? [MarkdownInline.CreateRawHtml("<sup>" + InlinesToPlainText(supChildren) + "</sup>")]
                    : null;

            default:
                return HandleUnknownInlineTag(tag);
        }
    }

    private List<MarkdownInline>? HandleUnknownInlineTag(HtmlTag tag)
    {
        switch (_options.UnknownTags)
        {
            case UnknownTagStrategy.PassThrough:
                var inner = ParseInlineUntilClose(tag.Name);
                if (inner.Count == 0)
                    return [MarkdownInline.CreateRawHtml(tag.RawText)];
                // 包裹在 RawHtml 中
                var result = new List<MarkdownInline>
                {
                    MarkdownInline.CreateRawHtml("<" + tag.Name + GetAttributesString(tag) + ">")
                };
                result.AddRange(inner);
                result.Add(MarkdownInline.CreateRawHtml("</" + tag.Name + ">"));
                return result;

            case UnknownTagStrategy.Drop:
                return ParseInlineUntilClose(tag.Name);

            case UnknownTagStrategy.Escape:
                return [MarkdownInline.CreateText(EscapeHtml(tag.RawText))];

            default:
                return null;
        }
    }

    private List<MarkdownInline>? ParseLink(HtmlTag tag)
    {
        var href = tag.GetAttribute("href");
        var title = tag.GetAttribute("title");

        if (String.IsNullOrEmpty(href))
        {
            // 无 href 的 a 标签当作普通文本
            return ParseInlineUntilClose("a");
        }

        href = DecodeHtmlEntities(href).Trim();

        // 检查协议白名单
        if (!IsAllowedScheme(href))
        {
            // 不允许的协议，仅保留文本
            return ParseInlineUntilClose("a");
        }

        var children = ParseInlineUntilClose("a");
        if (children.Count == 0)
        {
            // 空链接，使用 href 作为文本
            children = [MarkdownInline.CreateText(href)];
        }

        return [MarkdownInline.CreateLink(href, title ?? String.Empty, children)];
    }

    private List<MarkdownInline>? ParseImage(HtmlTag tag)
    {
        var src = tag.GetAttribute("src");
        var alt = tag.GetAttribute("alt");
        var title = tag.GetAttribute("title");

        if (String.IsNullOrEmpty(src))
        {
            if (!String.IsNullOrEmpty(alt))
                return [MarkdownInline.CreateText(alt)];
            return null;
        }

        src = DecodeHtmlEntities(src).Trim();

        if (!IsAllowedScheme(src))
        {
            if (!String.IsNullOrEmpty(alt))
                return [MarkdownInline.CreateText("[Image: " + alt + "]")];
            return null;
        }

        return [MarkdownInline.CreateImage(src, alt ?? String.Empty, title ?? String.Empty)];
    }

    private List<MarkdownInline> ParseInlineUntilClose(String tagName)
    {
        var inlines = new List<MarkdownInline>();
        var closeTag = "</" + tagName + ">";
        var closeTagUpper = "</" + tagName.ToUpperInvariant() + ">";

        while (_pos < _html.Length)
        {
            if (IsAt(closeTag, StringComparison.OrdinalIgnoreCase))
            {
                _pos += closeTag.Length;
                break;
            }

            if (_html[_pos] == '<')
            {
                var tag = ParseTag();
                if (tag == null)
                {
                    inlines.Add(MarkdownInline.CreateText("<"));
                    _pos++;
                    continue;
                }
                if (tag.IsClosing)
                {
                    // 意外的结束标签（非目标标签），作为文本
                    _pos -= tag.RawText.Length;
                    inlines.Add(MarkdownInline.CreateText("<"));
                    _pos++;
                    continue;
                }
                var tagInlines = ConvertTagToInlines(tag);
                if (tagInlines != null)
                    inlines.AddRange(tagInlines);
                continue;
            }

            var text = ReadUntil('<');
            if (text.Length > 0)
                inlines.Add(MarkdownInline.CreateText(DecodeHtmlEntities(text)));
        }

        return NormalizeInlines(inlines);
    }

    private String ReadUntilCloseTag(String tagName)
    {
        var sb = new StringBuilder();
        var closeTag = "</" + tagName + ">";

        while (_pos < _html.Length)
        {
            if (IsAt(closeTag, StringComparison.OrdinalIgnoreCase))
            {
                _pos += closeTag.Length;
                break;
            }
            sb.Append(_html[_pos]);
            _pos++;
        }

        return sb.ToString();
    }

    private void SkipTagContent(String tagName)
    {
        var closeTag = "</" + tagName + ">";
        var idx = _html.IndexOf(closeTag, _pos, StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
            _pos = idx + closeTag.Length;
    }
    #endregion

    #region HTML 词法分析
    private HtmlTag? ParseTag()
    {
        if (_pos >= _html.Length || _html[_pos] != '<') return null;

        var start = _pos;
        _pos++; // 跳过 <

        var isClosing = false;
        if (_pos < _html.Length && _html[_pos] == '/')
        {
            isClosing = true;
            _pos++;
        }

        // 读取标签名
        var nameStart = _pos;
        while (_pos < _html.Length && IsTagNameChar(_html[_pos]))
            _pos++;

        if (_pos == nameStart)
        {
            _pos = start + 1;
            return null; // 无效标签
        }

        var name = _html[nameStart.._pos].ToLowerInvariant();

        // 跳过空白
        SkipWhitespace();

        // 读取属性
        var attributes = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        while (_pos < _html.Length && _html[_pos] != '>' && _html[_pos] != '/')
        {
            SkipWhitespace();

            if (_pos >= _html.Length || _html[_pos] == '>' || _html[_pos] == '/')
                break;

            // 读取属性名
            var attrStart = _pos;
            while (_pos < _html.Length && IsAttrNameChar(_html[_pos]))
                _pos++;

            if (_pos == attrStart)
            {
                _pos++;
                continue;
            }

            var attrName = _html[attrStart.._pos].ToLowerInvariant();

            SkipWhitespace();

            // 读取属性值
            var attrValue = attrName; // 布尔属性默认值
            if (_pos < _html.Length && _html[_pos] == '=')
            {
                _pos++;
                SkipWhitespace();

                if (_pos < _html.Length && (_html[_pos] == '"' || _html[_pos] == '\''))
                {
                    var quote = _html[_pos];
                    _pos++;
                    var valStart = _pos;
                    while (_pos < _html.Length && _html[_pos] != quote)
                        _pos++;
                    attrValue = _html[valStart.._pos];
                    if (_pos < _html.Length) _pos++; // 跳过闭合引号
                }
                else
                {
                    var valStart = _pos;
                    while (_pos < _html.Length && !Char.IsWhiteSpace(_html[_pos]) && _html[_pos] != '>')
                        _pos++;
                    attrValue = _html[valStart.._pos];
                }
            }

            attributes[attrName] = attrValue;
        }

        // 跳过 > 或 />
        var isSelfClosing = false;
        if (_pos < _html.Length && _html[_pos] == '/')
        {
            isSelfClosing = true;
            _pos++;
        }
        if (_pos < _html.Length && _html[_pos] == '>')
            _pos++;

        // 自闭合标签
        if (isSelfClosing) isClosing = true;

        var rawText = _html[start.._pos];
        return new HtmlTag(name, isClosing, isSelfClosing, attributes, rawText);
    }

    private String? PeekClosingTag()
    {
        if (_pos + 2 >= _html.Length) return null;
        if (_html[_pos] != '<' || _html[_pos + 1] != '/') return null;

        var saved = _pos;
        _pos += 2;
        var nameStart = _pos;
        while (_pos < _html.Length && IsTagNameChar(_html[_pos]))
            _pos++;
        var name = _html[nameStart.._pos].ToLowerInvariant();
        _pos = saved;
        return name;
    }

    private String? PeekBlockStartTag()
    {
        if (_pos >= _html.Length || _html[_pos] != '<') return null;

        var saved = _pos;
        _pos++;
        var nameStart = _pos;
        while (_pos < _html.Length && IsTagNameChar(_html[_pos]))
            _pos++;

        if (_pos == nameStart) { _pos = saved; return null; }
        var name = _html[nameStart.._pos].ToLowerInvariant();
        _pos = saved;

        return IsBlockTag(name) ? name : null;
    }

    private static Boolean IsBlockTag(String name)
    {
        return name switch
        {
            "h1" or "h2" or "h3" or "h4" or "h5" or "h6" => true,
            "p" => true,
            "pre" => true,
            "blockquote" => true,
            "ul" or "ol" => true,
            "table" => true,
            "hr" => true,
            "div" => true,
            "section" => true,
            "article" => true,
            "header" => true,
            "footer" => true,
            "nav" => true,
            "main" => true,
            "aside" => true,
            "figure" => true,
            "form" => true,
            "fieldset" => true,
            "dl" or "dt" or "dd" => true,
            "details" or "summary" => true,
            _ => false,
        };
    }
    #endregion

    #region 辅助方法
    private void SkipWhitespace()
    {
        while (_pos < _html.Length && Char.IsWhiteSpace(_html[_pos]) && _html[_pos] != '\n' && _html[_pos] != '\r')
            _pos++;
    }

    private void SkipWhitespaceAndNewlines()
    {
        while (_pos < _html.Length && Char.IsWhiteSpace(_html[_pos]))
            _pos++;
    }

    private Boolean IsAt(String target, StringComparison comparison = StringComparison.Ordinal)
    {
        if (_pos + target.Length > _html.Length) return false;
        return String.Compare(_html, _pos, target, 0, target.Length, comparison) == 0;
    }

    private String ReadUntil(Char ch)
    {
        var start = _pos;
        while (_pos < _html.Length && _html[_pos] != ch)
            _pos++;
        return _html[start.._pos];
    }

    private static Boolean IsTagNameChar(Char c)
    {
        return Char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == ':';
    }

    private static Boolean IsAttrNameChar(Char c)
    {
        return Char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == ':' || c == '.';
    }

    private Boolean IsAllowedScheme(String uri)
    {
        if (uri.StartsWith("#") || uri.StartsWith("/") || uri.StartsWith("./") || uri.StartsWith("../"))
            return true;

        var colonIdx = uri.IndexOf(':');
        if (colonIdx <= 0) return true; // 相对路径

        var scheme = uri[..colonIdx];
        return _allowedSchemes.Contains(scheme);
    }

    private static String ExtractLanguageFromClass(String? classAttr)
    {
        if (String.IsNullOrEmpty(classAttr)) return String.Empty;

        // 匹配 language-xxx 格式
        var parts = classAttr!.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            if (p.StartsWith("language-", StringComparison.OrdinalIgnoreCase))
                return p["language-".Length..];
            if (p.StartsWith("lang-", StringComparison.OrdinalIgnoreCase))
                return p["lang-".Length..];
        }

        // 常见语言类名直接匹配（检查第一个 className 部分）
        var firstPart = parts.Length > 0 ? parts[0] : String.Empty;
        return firstPart switch
        {
            "csharp" or "cs" => "csharp",
            "javascript" or "js" => "javascript",
            "python" or "py" => "python",
            "java" => "java",
            "cpp" or "c++" => "cpp",
            "html" => "html",
            "css" => "css",
            "sql" => "sql",
            "xml" => "xml",
            "json" => "json",
            "yaml" or "yml" => "yaml",
            "bash" or "sh" => "bash",
            "powershell" or "ps1" => "powershell",
            "markdown" or "md" => "markdown",
            "rust" or "rs" => "rust",
            "go" or "golang" => "go",
            _ => String.Empty,
        };
    }

    private static String GetAttributesString(HtmlTag tag)
    {
        if (tag.Attributes.Count == 0) return String.Empty;
        var sb = new StringBuilder();
        foreach (var kv in tag.Attributes)
        {
            sb.Append(' ').Append(kv.Key);
            if (kv.Key != kv.Value)
                sb.Append("=\"").Append(kv.Value).Append('"');
        }
        return sb.ToString();
    }

    private static String EscapeHtml(String text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    /// <summary>解码 HTML 实体（具名 + 数字字符引用）</summary>
    /// <param name="text">HTML 文本</param>
    /// <returns>解码后文本</returns>
    /// <remarks>与 <see cref="MarkdownParser"/> 共享 <see cref="EntityDecoder"/>，行为一致（HTML5 实体全表 + 代理对）</remarks>
    public static String DecodeHtmlEntities(String text)
    {
        if (String.IsNullOrEmpty(text) || text.IndexOf('&') < 0) return text;

        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c != '&') { sb.Append(c); continue; }

            var semi = text.IndexOf(';', i + 1);
            if (semi > i + 1 && semi - i <= 32)
            {
                var entity = text.Substring(i + 1, semi - i - 1);
                var decoded = EntityDecoder.Decode(entity);
                if (decoded != null)
                {
                    sb.Append(decoded);
                    i = semi;
                    continue;
                }
            }
            sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>规范化行内元素列表，合并相邻文本节点</summary>
    private static List<MarkdownInline> NormalizeInlines(List<MarkdownInline> inlines)
    {
        if (inlines.Count <= 1) return inlines;

        var result = new List<MarkdownInline>();
        var sb = new StringBuilder();

        foreach (var inline in inlines)
        {
            if (inline.Type == MarkdownInlineType.Text && inline.Children.Count == 0)
            {
                sb.Append(inline.Text);
                continue;
            }

            if (sb.Length > 0)
            {
                result.Add(MarkdownInline.CreateText(sb.ToString()));
                sb.Clear();
            }

            result.Add(inline);
        }

        if (sb.Length > 0)
            result.Add(MarkdownInline.CreateText(sb.ToString()));

        return result;
    }

    private static String InlinesToPlainText(List<MarkdownInline> inlines)
    {
        var sb = new StringBuilder();
        foreach (var inline in inlines)
        {
            sb.Append(inline.GetPlainText());
        }
        return sb.ToString();
    }
    #endregion

    #region 内部类型
    private sealed class HtmlTag
    {
        public String Name { get; }
        public Boolean IsClosing { get; }
        public Boolean IsSelfClosing { get; }
        public Dictionary<String, String> Attributes { get; }
        public String RawText { get; }

        public HtmlTag(String name, Boolean isClosing, Boolean isSelfClosing,
            Dictionary<String, String> attributes, String rawText)
        {
            Name = name;
            IsClosing = isClosing;
            IsSelfClosing = isSelfClosing;
            Attributes = attributes;
            RawText = rawText;
        }

        public String GetAttribute(String name)
        {
            return Attributes.TryGetValue(name, out var value) ? value : String.Empty;
        }
    }
    #endregion
}
