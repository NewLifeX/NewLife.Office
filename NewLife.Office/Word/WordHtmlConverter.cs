using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office.Word;

/// <summary>Word docx 转 HTML 转换器</summary>
/// <remarks>
/// 将 docx 文件解析为语义化 HTML，支持标题/段落/表格/超链接/文字格式等核心元素。
/// 无需任何外部依赖，直接操作 Open XML（ZIP+XML）内容。
/// <para>示例：</para>
/// <code>
/// var converter = new WordHtmlConverter { FullPage = true };
/// var html = converter.ConvertFromFile("doc.docx");
/// </code>
/// </remarks>
public sealed class WordHtmlConverter
{
    #region 属性
    /// <summary>是否将图片嵌入为 base64 Data URI（默认 false）</summary>
    public Boolean EmbedImages { get; set; }

    /// <summary>是否输出完整 HTML 页面（含 DOCTYPE/head/style），默认 true</summary>
    public Boolean FullPage { get; set; } = true;

    /// <summary>完整页面时的文档标题，默认 "Document"</summary>
    public String PageTitle { get; set; } = "Document";
    #endregion

    #region 公开方法
    /// <summary>从文件路径转换</summary>
    /// <param name="path">docx 文件路径</param>
    /// <returns>HTML 字符串</returns>
    public String ConvertFromFile(String path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert(fs);
    }

    /// <summary>从流转换</summary>
    /// <param name="stream">包含 docx 内容的可读流</param>
    /// <returns>HTML 字符串</returns>
    public String Convert(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var rels = LoadRelationships(zip);
        var images = EmbedImages ? LoadImages(zip) : [];
        var numFmts = LoadNumIdFormats(zip);
        var doc = LoadDocumentXml(zip);
        var body = RenderDocument(doc, rels, images, numFmts);
        return FullPage ? BuildFullPage(body) : body;
    }
    #endregion

    #region 渲染
    private static String RenderDocument(XmlDocument doc, Dictionary<String, String> rels, Dictionary<String, String> images, Dictionary<Int32, String> numFmts)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);
        ns.AddNamespace("r", R);
        ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        ns.AddNamespace("v", "urn:schemas-microsoft-com:vml");

        var body = doc.SelectSingleNode("//w:body", ns);
        if (body == null) return String.Empty;

        var sb = new StringBuilder();
        var openList = String.Empty; // 当前打开的列表标签（"" = 未打开）

        foreach (XmlNode node in body.ChildNodes)
        {
            if (node is not XmlElement el) continue;

            if (el.LocalName == "p")
            {
                // 列表段落：连续同类列表包裹 <ul>/<ol><li>
                var listTag = GetListTag(el, ns, numFmts);
                if (listTag != null)
                {
                    if (openList != listTag)
                    {
                        if (openList.Length > 0) sb.AppendLine($"</{openList}>");
                        sb.AppendLine($"<{listTag}>");
                        openList = listTag;
                    }
                    sb.Append("  <li");
                    sb.Append(GetAlignAttr(el, ns));
                    sb.Append(">");
                    RenderParagraphContent(sb, el, rels, images, ns);
                    sb.AppendLine("</li>");
                    continue;
                }
                if (openList.Length > 0)
                {
                    sb.AppendLine($"</{openList}>");
                    openList = String.Empty;
                }
                RenderParagraph(sb, el, rels, images, ns);
            }
            else if (el.LocalName == "tbl")
            {
                if (openList.Length > 0) { sb.AppendLine($"</{openList}>"); openList = String.Empty; }
                RenderTable(sb, el, rels, images, ns);
            }
        }
        if (openList.Length > 0) sb.AppendLine($"</{openList}>");
        return sb.ToString();
    }

    /// <summary>解析 numbering.xml 的 numId → numFmt（ilvl 0）映射，用于列表类型判定</summary>
    private static Dictionary<Int32, String> LoadNumIdFormats(ZipArchive zip)
    {
        var map = new Dictionary<Int32, String>();
        var entry = zip.GetEntry("word/numbering.xml");
        if (entry == null) return map;
        try
        {
            var doc = new XmlDocument();
            using (var s = entry.Open()) doc.Load(s);
            const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", W);

            var numToAbs = new Dictionary<Int32, Int32>();
            foreach (XmlElement numEl in doc.SelectNodes("//w:num", ns)!)
            {
                if (!Int32.TryParse(numEl.GetAttribute("w:numId"), out var numId)) continue;
                var absEl = numEl.SelectSingleNode("w:abstractNumId", ns) as XmlElement;
                if (absEl != null && Int32.TryParse(absEl.GetAttribute("w:val"), out var absId))
                    numToAbs[numId] = absId;
            }
            foreach (var kv in numToAbs)
            {
                var absEl = doc.SelectSingleNode($"//w:abstractNum[@w:abstractNumId='{kv.Value}']", ns) as XmlElement;
                if (absEl == null) continue;
                var lvl0 = absEl.SelectSingleNode("w:lvl[@w:ilvl='0']/w:numFmt", ns) as XmlElement;
                var fmt = lvl0?.GetAttribute("w:val");
                if (fmt != null) map[kv.Key] = fmt;
            }
        }
        catch { /* 忽略损坏的 numbering.xml */ }
        return map;
    }

    /// <summary>判定段落是否为列表，返回列表标签（"ul"/"ol"），非列表返回 null</summary>
    private static String? GetListTag(XmlElement para, XmlNamespaceManager ns, Dictionary<Int32, String> numFmts)
    {
        var numPr = para.SelectSingleNode("w:pPr/w:numPr", ns) as XmlElement;
        if (numPr == null) return null;
        var numIdEl = numPr.SelectSingleNode("w:numId", ns) as XmlElement;
        var numIdStr = numIdEl?.GetAttribute("w:val");
        if (numIdStr == "0") return null;
        if (Int32.TryParse(numIdStr, out var numId))
        {
            if (numFmts.TryGetValue(numId, out var fmt))
                return fmt == "bullet" ? "ul" : "ol";
        }
        return "ul"; // 无映射时回退无序列表
    }

    /// <summary>生成段落对齐的 HTML 属性（如 style="text-align:center"）</summary>
    private static String GetAlignAttr(XmlElement para, XmlNamespaceManager ns)
    {
        var jcEl = para.SelectSingleNode("w:pPr/w:jc", ns) as XmlElement;
        var align = jcEl?.GetAttribute("w:val") ?? String.Empty;
        return align switch
        {
            "center" => " style=\"text-align:center\"",
            "right" => " style=\"text-align:right\"",
            "both" => " style=\"text-align:justify\"",
            _ => String.Empty,
        };
    }

    private static void RenderParagraph(StringBuilder sb, XmlElement para,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        // 分页符
        var pageBreak = para.SelectSingleNode("w:r/w:br[@w:type='page']", ns) != null;
        if (pageBreak)
        {
            sb.AppendLine("<div style=\"page-break-after:always\"></div>");
            // 分页符段若仅含分页符则不再输出内容
            var hasText = para.SelectSingleNode(".//w:t", ns) != null;
            var hasDrawing = para.SelectSingleNode(".//w:drawing", ns) != null;
            if (!hasText && !hasDrawing) return;
        }

        // 检测标题级别
        var styleEl = para.SelectSingleNode("w:pPr/w:pStyle", ns) as XmlElement;
        var styleVal = styleEl?.GetAttribute("w:val") ?? String.Empty;
        var level = GetHeadingLevel(styleVal);

        var tag = level > 0 ? $"h{level}" : "p";
        sb.Append($"<{tag}");
        sb.Append(GetAlignAttr(para, ns));
        sb.Append(">");
        RenderParagraphContent(sb, para, rels, images, ns);
        sb.AppendLine($"</{tag}>");
    }

    /// <summary>渲染段落内容（run/超链接/图片/制表符），不含外层标签</summary>
    private static void RenderParagraphContent(StringBuilder sb, XmlElement para,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        // 遍历子节点（run/hyperlink/drawing/bookmarkStart 等）
        foreach (XmlNode child in para.ChildNodes)
        {
            if (child is not XmlElement childEl) continue;
            if (childEl.LocalName == "r")
                RenderRun(sb, childEl, ns, rels, images);
            else if (childEl.LocalName == "hyperlink")
                RenderHyperlink(sb, childEl, rels, images, ns);
            else if (childEl.LocalName == "drawing")
                RenderDrawing(sb, childEl, rels, images, ns);
            else if (childEl.LocalName == "pict")
                RenderVmlDrawing(sb, childEl, rels, images, ns);
        }
    }

    private static void RenderRun(StringBuilder sb, XmlElement run, XmlNamespaceManager ns,
        Dictionary<String, String> rels, Dictionary<String, String> images)
    {
        // run 内嵌 DrawingML 图片（w:r/w:drawing）
        var drawing = run.SelectSingleNode("w:drawing", ns) as XmlElement;
        if (drawing != null)
        {
            RenderDrawing(sb, drawing, rels, images, ns);
            return;
        }
        var pict = run.SelectSingleNode("w:pict", ns) as XmlElement;
        if (pict != null)
        {
            RenderVmlDrawing(sb, pict, rels, images, ns);
            return;
        }

        // 读取格式属性
        var rPr = run.SelectSingleNode("w:rPr", ns) as XmlElement;
        var bold = rPr?.SelectSingleNode("w:b", ns) != null;
        var italic = rPr?.SelectSingleNode("w:i", ns) != null;
        var underline = rPr?.SelectSingleNode("w:u", ns) != null;
        var strike = rPr?.SelectSingleNode("w:strike", ns) != null;
        var vertAlign = rPr?.SelectSingleNode("w:vertAlign", ns) as XmlElement;
        var vaVal = vertAlign?.GetAttribute("w:val") ?? String.Empty;
        var highlightEl = rPr?.SelectSingleNode("w:highlight", ns) as XmlElement;
        var highlight = highlightEl?.GetAttribute("w:val") ?? String.Empty;
        var colorEl = rPr?.SelectSingleNode("w:color", ns) as XmlElement;
        var color = colorEl?.GetAttribute("w:val") ?? String.Empty;
        var szEl = rPr?.SelectSingleNode("w:sz", ns) as XmlElement;
        var szVal = szEl?.GetAttribute("w:val") ?? String.Empty;

        // 提取文本（w:t 节点，处理 xml:space="preserve"；w:tab 转制表符）
        var textSb = new StringBuilder();
        foreach (XmlNode child in run.ChildNodes)
        {
            if (child is XmlElement el && el.LocalName == "t")
                textSb.Append(el.InnerText);
            else if (child is XmlElement brEl && brEl.LocalName == "br")
                textSb.Append('\n');
            else if (child is XmlElement tabEl && tabEl.LocalName == "tab")
                textSb.Append('\t');
        }
        var text = textSb.ToString();
        if (text.Length == 0) return;

        var encoded = HtmlEncode(text).Replace("\n", "<br />").Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;");

        // 构建内联样式
        var spanStyle = new StringBuilder();
        if (!String.IsNullOrEmpty(color) && color != "auto" && color != "000000")
            spanStyle.Append($"color:#{color};");
        if (!String.IsNullOrEmpty(szVal) && Int32.TryParse(szVal, out var sz))
        {
            // w:sz 单位是半磅
            var pt = sz / 2.0;
            spanStyle.Append($"font-size:{pt}pt;");
        }

        var content = encoded;
        if (spanStyle.Length > 0)
            content = $"<span style=\"{spanStyle}\">{content}</span>";
        if (underline) content = $"<u>{content}</u>";
        if (vaVal == "superscript") content = $"<sup>{content}</sup>";
        else if (vaVal == "subscript") content = $"<sub>{content}</sub>";
        if (strike) content = $"<s>{content}</s>";
        if (!String.IsNullOrEmpty(highlight) && highlight != "none")
            content = $"<mark>{content}</mark>";
        if (italic) content = $"<em>{content}</em>";
        if (bold) content = $"<strong>{content}</strong>";

        sb.Append(content);
    }

    /// <summary>渲染 DrawingML 图片（wp:inline / wp:anchor 中的 a:blip）；无图片时为文本框则渲染文本</summary>
    private static void RenderDrawing(StringBuilder sb, XmlElement drawing,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        var blip = drawing.SelectSingleNode(".//a:blip", ns) as XmlElement;
        var rId = blip?.GetAttribute("r:embed");
        if (rId != null)
        {
            if (images.TryGetValue(rId, out var src))
            {
                sb.Append($"<img src=\"{src}\" alt=\"image\" />");
            }
            else if (rels.TryGetValue(rId, out var target))
            {
                sb.Append($"<img src=\"word/{HtmlAttrEncode(target)}\" alt=\"image\" />");
            }
            return;
        }

        // 文本框（wps:txbx/w:txbxContent）：渲染其中段落文本（W22，与 Word 导出行为一致）
        foreach (XmlElement txbxContent in drawing.SelectNodes(".//w:txbxContent", ns)!)
        {
            foreach (XmlElement p in txbxContent.SelectNodes("w:p", ns)!)
                RenderParagraph(sb, p, rels, images, ns);
        }
    }

    /// <summary>渲染 VML 图片（w:pict/v:shape/v:imagedata）</summary>
    private static void RenderVmlDrawing(StringBuilder sb, XmlElement pict,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        var vmlData = pict.SelectSingleNode(".//v:imagedata", ns) as XmlElement;
        var rId = vmlData?.GetAttribute("r:id") ?? vmlData?.GetAttribute("id");
        if (rId == null) return;

        if (images.TryGetValue(rId, out var src))
        {
            sb.Append($"<img src=\"{src}\" alt=\"image\" />");
        }
        else if (rels.TryGetValue(rId, out var target))
        {
            sb.Append($"<img src=\"word/{HtmlAttrEncode(target)}\" alt=\"image\" />");
        }
    }

    private static void RenderHyperlink(StringBuilder sb, XmlElement hyperlink,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        var relId = hyperlink.GetAttribute("r:id");
        var url = String.Empty;
        if (!String.IsNullOrEmpty(relId))
            rels.TryGetValue(relId, out url!);

        if (!String.IsNullOrEmpty(url))
            sb.Append($"<a href=\"{HtmlAttrEncode(url)}\">");

        foreach (XmlNode child in hyperlink.ChildNodes)
        {
            if (child is XmlElement el && el.LocalName == "r")
                RenderRun(sb, el, ns, rels, images);
        }

        if (!String.IsNullOrEmpty(url))
            sb.Append("</a>");
    }

    private static void RenderTable(StringBuilder sb, XmlElement tbl,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        sb.AppendLine("<table border=\"1\" style=\"border-collapse:collapse\">");

        // 收集所有行与单元格
        var rows = new List<(XmlElement RowEl, List<XmlElement> Cells)>();
        foreach (XmlNode rowNode in tbl.ChildNodes)
        {
            if (rowNode is not XmlElement rowEl || rowEl.LocalName != "tr") continue;
            var cells = new List<XmlElement>();
            foreach (XmlNode cellNode in rowEl.ChildNodes)
            {
                if (cellNode is XmlElement cellEl && cellEl.LocalName == "tc")
                    cells.Add(cellEl);
            }
            rows.Add((rowEl, cells));
        }

        // 表头判定：存在 tblHeader 行则按标记，否则首行作表头
        var hasAnyHeader = rows.Any(r => r.RowEl.SelectSingleNode("w:trPr/w:tblHeader", ns) != null);

        for (var ri = 0; ri < rows.Count; ri++)
        {
            var (rowEl, cells) = rows[ri];
            var isHeader = hasAnyHeader
                ? rowEl.SelectSingleNode("w:trPr/w:tblHeader", ns) != null
                : ri == 0;
            var cellTag = isHeader ? "th" : "td";
            sb.AppendLine("<tr>");

            for (var ci = 0; ci < cells.Count; ci++)
            {
                var cellEl = cells[ci];
                var attrs = new StringBuilder();
                // gridSpan → colspan
                var gs = cellEl.SelectSingleNode("w:tcPr/w:gridSpan", ns) as XmlElement;
                if (gs != null && Int32.TryParse(gs.GetAttribute("w:val"), out var cs) && cs > 1)
                    attrs.Append($" colspan=\"{cs}\"");
                // vMerge → rowspan（扫描下方连续 continue 单元格）
                var vm = cellEl.SelectSingleNode("w:tcPr/w:vMerge", ns) as XmlElement;
                if (vm != null)
                {
                    var vVal = vm.GetAttribute("w:val");
                    if (vVal == "restart")
                    {
                        var rs = 1;
                        for (var r2 = ri + 1; r2 < rows.Count; r2++)
                        {
                            if (ci >= rows[r2].Cells.Count) break;
                            var vm2 = rows[r2].Cells[ci].SelectSingleNode("w:tcPr/w:vMerge", ns) as XmlElement;
                            if (vm2 != null && String.IsNullOrEmpty(vm2.GetAttribute("w:val")))
                                rs++;
                            else
                                break;
                        }
                        if (rs > 1) attrs.Append($" rowspan=\"{rs}\"");
                    }
                    else
                    {
                        continue; // 继续合并单元格：由 restart 行覆盖，不重复输出
                    }
                }

                sb.Append($"<{cellTag}{attrs}>");
                foreach (XmlNode pNode in cellEl.ChildNodes)
                {
                    if (pNode is not XmlElement pEl) continue;
                    if (pEl.LocalName == "p")
                        RenderParagraph(sb, pEl, rels, images, ns);
                    else if (pEl.LocalName == "tbl")
                        RenderTable(sb, pEl, rels, images, ns); // 嵌套表格递归渲染
                }
                sb.AppendLine($"</{cellTag}>");
            }
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</table>");
    }
    #endregion

    #region 辅助方法
    /// <summary>检测段落样式值对应的标题级别（1-6），非标题返回 0</summary>
    /// <param name="styleVal">w:pStyle 的 w:val 属性值</param>
    /// <returns>标题级别 1-6，普通段落返回 0</returns>
    private static Int32 GetHeadingLevel(String styleVal)
    {
        if (String.IsNullOrEmpty(styleVal)) return 0;

        // 直接数字 "1".."6"
        if (styleVal.Length == 1 && styleVal[0] >= '1' && styleVal[0] <= '6')
            return styleVal[0] - '0';

        // 标准化为小写无空格比较
        var normalized = styleVal.ToLowerInvariant().Replace(" ", String.Empty).Replace("-", String.Empty);

        // "heading1".."heading6" 或 "überschrift1" 等
        for (var i = 1; i <= 6; i++)
        {
            if (normalized == "heading" + i) return i;
        }

        // 纯数字尾
        if (normalized.Length > 1)
        {
            var lastChar = normalized[normalized.Length - 1];
            if (lastChar >= '1' && lastChar <= '6')
            {
                var prefix = normalized[..^1];
                if (prefix == "heading" || prefix == "h" || prefix == "\u6807\u9898")
                    return lastChar - '0';
            }
        }

        return 0;
    }

    /// <summary>HTML 文本转义（&amp; &lt; &gt; " '）</summary>
    /// <param name="text">原始文本</param>
    /// <returns>转义后文本</returns>
    private static String HtmlEncode(String text)
    {
        if (String.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length + 16);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&#39;"); break;
                default: sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>HTML 属性值转义（&amp; &lt; &gt; "）</summary>
    /// <param name="value">原始属性值</param>
    /// <returns>转义后属性值</returns>
    private static String HtmlAttrEncode(String value)
    {
        if (String.IsNullOrEmpty(value)) return value;
        var sb = new StringBuilder(value.Length + 8);
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                default: sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>从 ZipArchive 加载 word/document.xml</summary>
    /// <param name="zip">已打开的 docx ZIP 归档</param>
    /// <returns>解析后的 XmlDocument</returns>
    private static XmlDocument LoadDocumentXml(ZipArchive zip)
    {
        var entry = zip.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("无效的 docx 文件：缺少 word/document.xml");
        var doc = new XmlDocument();
        using var s = entry.Open();
        doc.Load(s);
        return doc;
    }

    /// <summary>加载 word/_rels/document.xml.rels，建立 relId → 目标 URL 映射</summary>
    /// <param name="zip">已打开的 docx ZIP 归档</param>
    /// <returns>关系 ID 到 URL 的字典</returns>
    private static Dictionary<String, String> LoadRelationships(ZipArchive zip)
    {
        var result = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        var entry = zip.GetEntry("word/_rels/document.xml.rels");
        if (entry == null) return result;

        var doc = new XmlDocument();
        using (var s = entry.Open())
            doc.Load(s);

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

        foreach (XmlElement relEl in doc.SelectNodes("//rel:Relationship", ns)!)
        {
            var id = relEl.GetAttribute("Id");
            var target = relEl.GetAttribute("Target");
            if (!String.IsNullOrEmpty(id) && !String.IsNullOrEmpty(target))
                result[id] = target;
        }
        return result;
    }

    /// <summary>从 word/media/* 加载图片，返回 partName → base64 Data URI 映射</summary>
    /// <summary>加载媒体图片并建立关系ID → Data URI 映射（EmbedImages 内嵌用）</summary>
    /// <param name="zip">已打开的 docx ZIP 归档</param>
    /// <returns>关系ID 到 Data URI 的字典（key 为 document.xml.rels 中的 rId）</returns>
    private static Dictionary<String, String> LoadImages(ZipArchive zip)
    {
        // 1. 媒体文件名 → Data URI
        var media = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                continue;
            var ext = Path.GetExtension(entry.Name).TrimStart('.').ToLowerInvariant();
            var mime = ext switch
            {
                "png" => "image/png",
                "jpg" or "jpeg" => "image/jpeg",
                "gif" => "image/gif",
                "bmp" => "image/bmp",
                "webp" => "image/webp",
                _ => "image/octet-stream",
            };
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            var b64 = System.Convert.ToBase64String(ms.ToArray());
            media[entry.Name] = $"data:{mime};base64,{b64}";
        }
        if (media.Count == 0) return [];

        // 2. 从 document.xml.rels 建立 rId → 目标文件名 → Data URI
        var result = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        var relsEntry = zip.GetEntry("word/_rels/document.xml.rels");
        if (relsEntry == null) return result;
        var relsDoc = new XmlDocument();
        using (var s = relsEntry.Open())
            relsDoc.Load(s);
        var ns = new XmlNamespaceManager(relsDoc.NameTable);
        ns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (XmlElement rel in relsDoc.SelectNodes("//rel:Relationship", ns)!)
        {
            var type = rel.GetAttribute("Type");
            if (type == null || !type.EndsWith("/image", StringComparison.Ordinal)) continue;
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            if (String.IsNullOrEmpty(id) || String.IsNullOrEmpty(target)) continue;
            var fileName = Path.GetFileName(target.Replace('\\', '/'));
            if (media.TryGetValue(fileName, out var uri))
                result[id] = uri;
        }
        return result;
    }

    /// <summary>构建完整 HTML 页面（含 DOCTYPE/head/style/body）</summary>
    /// <returns>完整 HTML 字符串</returns>
    private String BuildFullPage(String body)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine($"<meta charset=\"utf-8\" />");
        sb.AppendLine($"<title>{HtmlEncode(PageTitle)}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: Arial, sans-serif; margin: 2em; line-height: 1.5; }");
        sb.AppendLine("h1,h2,h3,h4,h5,h6 { margin-top: 1em; margin-bottom: 0.3em; }");
        sb.AppendLine("p { margin: 0.4em 0; }");
        sb.AppendLine("table { border-collapse: collapse; margin: 1em 0; width: 100%; }");
        sb.AppendLine("th, td { border: 1px solid #ccc; padding: 4px 8px; text-align: left; }");
        sb.AppendLine("th { background: #f0f0f0; font-weight: bold; }");
        sb.AppendLine("a { color: #0563C1; }");
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.Append(body);
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }
    #endregion
}
