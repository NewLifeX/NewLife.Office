using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

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
        var doc = LoadDocumentXml(zip);
        var body = RenderDocument(doc, rels, images);
        return FullPage ? BuildFullPage(body) : body;
    }
    #endregion

    #region 渲染
    private static String RenderDocument(XmlDocument doc, Dictionary<String, String> rels, Dictionary<String, String> images)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);
        ns.AddNamespace("r", R);

        var body = doc.SelectSingleNode("//w:body", ns);
        if (body == null) return String.Empty;

        var sb = new StringBuilder();
        foreach (XmlNode node in body.ChildNodes)
        {
            if (node is not XmlElement el) continue;
            if (el.LocalName == "p")
                RenderParagraph(sb, el, rels, images, ns);
            else if (el.LocalName == "tbl")
                RenderTable(sb, el, rels, images, ns);
        }
        return sb.ToString();
    }

    private static void RenderParagraph(StringBuilder sb, XmlElement para,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        // 检测标题级别
        var styleEl = para.SelectSingleNode("w:pPr/w:pStyle", ns) as XmlElement;
        var styleVal = styleEl?.GetAttribute("w:val") ?? String.Empty;
        var level = GetHeadingLevel(styleVal);

        // 检测对齐
        var jcEl = para.SelectSingleNode("w:pPr/w:jc", ns) as XmlElement;
        var align = jcEl?.GetAttribute("w:val") ?? String.Empty;
        var styleAttr = align switch
        {
            "center" => " style=\"text-align:center\"",
            "right" => " style=\"text-align:right\"",
            "both" => " style=\"text-align:justify\"",
            _ => String.Empty,
        };

        var tag = level > 0 ? $"h{level}" : "p";
        sb.Append($"<{tag}{styleAttr}>");

        // 遍历子节点（run/hyperlink/bookmarkStart 等）
        foreach (XmlNode child in para.ChildNodes)
        {
            if (child is not XmlElement childEl) continue;
            if (childEl.LocalName == "r")
                RenderRun(sb, childEl, ns);
            else if (childEl.LocalName == "hyperlink")
                RenderHyperlink(sb, childEl, rels, ns);
        }

        sb.AppendLine($"</{tag}>");
    }

    private static void RenderRun(StringBuilder sb, XmlElement run, XmlNamespaceManager ns)
    {
        // 读取格式属性
        var rPr = run.SelectSingleNode("w:rPr", ns) as XmlElement;
        var bold = rPr?.SelectSingleNode("w:b", ns) != null;
        var italic = rPr?.SelectSingleNode("w:i", ns) != null;
        var underline = rPr?.SelectSingleNode("w:u", ns) != null;
        var colorEl = rPr?.SelectSingleNode("w:color", ns) as XmlElement;
        var color = colorEl?.GetAttribute("w:val") ?? String.Empty;
        var szEl = rPr?.SelectSingleNode("w:sz", ns) as XmlElement;
        var szVal = szEl?.GetAttribute("w:val") ?? String.Empty;

        // 提取文本（w:t 节点，处理 xml:space="preserve"）
        var textSb = new StringBuilder();
        foreach (XmlNode child in run.ChildNodes)
        {
            if (child is XmlElement el && el.LocalName == "t")
                textSb.Append(el.InnerText);
            else if (child is XmlElement brEl && brEl.LocalName == "br")
                textSb.Append('\n');
        }
        var text = textSb.ToString();
        if (text.Length == 0) return;

        var encoded = HtmlEncode(text).Replace("\n", "<br />");

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
        if (italic) content = $"<em>{content}</em>";
        if (bold) content = $"<strong>{content}</strong>";

        sb.Append(content);
    }

    private static void RenderHyperlink(StringBuilder sb, XmlElement hyperlink,
        Dictionary<String, String> rels, XmlNamespaceManager ns)
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
                RenderRun(sb, el, ns);
        }

        if (!String.IsNullOrEmpty(url))
            sb.Append("</a>");
    }

    private static void RenderTable(StringBuilder sb, XmlElement tbl,
        Dictionary<String, String> rels, Dictionary<String, String> images, XmlNamespaceManager ns)
    {
        sb.AppendLine("<table border=\"1\" style=\"border-collapse:collapse\">");
        var rowIndex = 0;
        foreach (XmlNode rowNode in tbl.ChildNodes)
        {
            if (rowNode is not XmlElement rowEl || rowEl.LocalName != "tr") continue;
            sb.AppendLine("<tr>");
            var isHeader = rowIndex == 0;
            var cellTag = isHeader ? "th" : "td";
            foreach (XmlNode cellNode in rowEl.ChildNodes)
            {
                if (cellNode is not XmlElement cellEl || cellEl.LocalName != "tc") continue;
                sb.Append($"<{cellTag}>");
                foreach (XmlNode pNode in cellEl.ChildNodes)
                {
                    if (pNode is XmlElement pEl && pEl.LocalName == "p")
                        RenderParagraph(sb, pEl, rels, images, ns);
                }
                sb.AppendLine($"</{cellTag}>");
            }
            sb.AppendLine("</tr>");
            rowIndex++;
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

    /// <summary>HTML 文本转义（& < > " '）</summary>
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

    /// <summary>HTML 属性值转义（& < > "）</summary>
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
    /// <param name="zip">已打开的 docx ZIP 归档</param>
    /// <returns>图片名称到 Data URI 的字典</returns>
    private static Dictionary<String, String> LoadImages(ZipArchive zip)
    {
        var result = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
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
            result[entry.Name] = $"data:{mime};base64,{b64}";
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
