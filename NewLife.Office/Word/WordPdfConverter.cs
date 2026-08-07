using System.IO.Compression;
using System.Xml;
using NewLife.Office;

using NewLife.Office.Pdf;

namespace NewLife.Office.Word;

/// <summary>Word docx 转 PDF 转换器（低保真内容映射型）</summary>
/// <remarks>
/// 将 docx 文件解析为 PDF，将段落映射为文本块（标题使用较大字号），
/// 表格映射为 PDF 表格，不依赖 Office/LibreOffice 等外部组件。
/// <para>示例：</para>
/// <code>
/// var converter = new WordPdfConverter();
/// converter.ConvertToFile("document.docx", "output.pdf");
/// // 或
/// var pdfBytes = converter.ConvertToBytes(stream);
/// </code>
/// </remarks>
public sealed class WordPdfConverter
{
    #region 属性

    /// <summary>正文字号，默认 11pt</summary>
    public Single BodyFontSize { get; set; } = 11f;

    /// <summary>H1 字号，默认 22pt</summary>
    public Single H1FontSize { get; set; } = 22f;

    /// <summary>H2 字号，默认 18pt</summary>
    public Single H2FontSize { get; set; } = 18f;

    /// <summary>H3 字号，默认 15pt</summary>
    public Single H3FontSize { get; set; } = 15f;

    /// <summary>H4 字号，默认 13pt</summary>
    public Single H4FontSize { get; set; } = 13f;

    /// <summary>H5/H6 字号，默认与正文一致</summary>
    public Single H56FontSize { get; set; } = 11f;

    #endregion

    #region 公开方法

    /// <summary>从文件路径转换为 PDF 字节</summary>
    /// <param name="path">docx 文件路径</param>
    /// <returns>PDF 字节数组</returns>
    public Byte[] ConvertToBytes(String path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return ConvertToBytes(fs);
    }

    /// <summary>从流转换为 PDF 字节</summary>
    /// <param name="stream">包含 docx 内容的可读流</param>
    /// <returns>PDF 字节数组</returns>
    public Byte[] ConvertToBytes(Stream stream)
    {
        using var pdf = new PdfDocumentBuilder();
        Render(stream, pdf);
        return pdf.ToBytes();
    }

    /// <summary>从文件路径转换，输出到目标路径</summary>
    /// <param name="inputPath">docx 文件路径</param>
    /// <param name="outputPath">输出 PDF 路径</param>
    public void ConvertToFile(String inputPath, String outputPath)
    {
        using var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        ConvertToFile(fs, outputPath);
    }

    /// <summary>从流转换，输出到目标路径</summary>
    /// <param name="stream">包含 docx 内容的可读流</param>
    /// <param name="outputPath">输出 PDF 路径</param>
    public void ConvertToFile(Stream stream, String outputPath)
    {
        using var pdf = new PdfDocumentBuilder();
        Render(stream, pdf);
        pdf.Save(outputPath);
    }

    /// <summary>将 docx 每页渲染为图片（PNG/JPEG）</summary>
    /// <remarks>
    /// TODO: 本方法尚未实现。将文档渲染为光栅图片需要引入渲染引擎（如 SkiaSharp + Docnet.Core）
    /// 将 PDF 中间格式解码为帧位图，当前版本不依赖外部库，故此功能暂不支持。
    /// 如需此功能，建议先调用 ConvertToBytes/ConvertToFile 获得 PDF，再用 PdfDocument.RenderToImages 处理。
    /// </remarks>
    /// <param name="stream">包含 docx 内容的可读流</param>
    /// <param name="dpi">输出分辨率（DPI）</param>
    /// <returns>每页图片字节（PNG 格式）</returns>
    /// <exception cref="NotSupportedException">当前版本始终抛出，待引入渲染库后实现</exception>
    public IEnumerable<Byte[]> ConvertToImages(Stream stream, Int32 dpi = 150)
    {
        throw new NotSupportedException(
            "渲染为图片需要引入 SkiaSharp/Docnet.Core 等渲染库，当前版本不支持。" +
            "建议先转换为 PDF，再使用 PdfDocument.RenderToImages 处理。");
    }

    #endregion

    #region 渲染核心

    /// <summary>解析 docx 并将内容写入 PDF 文档</summary>
    /// <param name="stream">docx 流</param>
    /// <param name="pdf">目标 FluentDocument</param>
    private void Render(Stream stream, PdfDocumentBuilder pdf)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 加载 word/document.xml
        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return;

        var doc = new XmlDocument();
        using (var s = entry.Open())
            doc.Load(s);

        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);
        ns.AddNamespace("r", R);
        ns.AddNamespace("a", A);

        // 关系映射（图片）+ 编号格式映射（列表）
        var rels = LoadRels(zip);
        var numFmts = LoadNumIdFormats(zip);

        var body = doc.SelectSingleNode("//w:body", ns);
        if (body == null) return;

        foreach (XmlNode node in body.ChildNodes)
        {
            if (node is not XmlElement el) continue;
            if (el.LocalName == "p")
                RenderParagraph(el, ns, pdf, zip, rels, numFmts);
            else if (el.LocalName == "tbl")
                RenderTable(el, ns, pdf);
        }
    }

    /// <summary>读取 word/_rels/document.xml.rels 建立 relId → Target 映射</summary>
    private static Dictionary<String, String> LoadRels(ZipArchive zip)
    {
        var map = new Dictionary<String, String>();
        var entry = zip.GetEntry("word/_rels/document.xml.rels");
        if (entry == null) return map;
        var doc = new XmlDocument();
        using (var s = entry.Open()) doc.Load(s);
        const String Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("rel", Rel);
        foreach (XmlElement rel in doc.SelectNodes("//rel:Relationship", ns)!)
        {
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            if (id != null && target != null) map[id] = target;
        }
        return map;
    }

    /// <summary>解析 numbering.xml 的 numId → numFmt（ilvl 0）映射，用于列表前缀</summary>
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
                var fmt = (absEl.SelectSingleNode("w:lvl[@w:ilvl='0']/w:numFmt", ns) as XmlElement)?.GetAttribute("w:val");
                if (fmt != null) map[kv.Key] = fmt;
            }
        }
        catch { /* 忽略损坏的 numbering.xml */ }
        return map;
    }

    /// <summary>将段落节点渲染为 PDF（标题/列表前缀/粗斜体/分页符/图片）</summary>
    private void RenderParagraph(XmlElement para, XmlNamespaceManager ns, PdfDocumentBuilder pdf,
        ZipArchive zip, Dictionary<String, String> rels, Dictionary<Int32, String> numFmts)
    {
        // 分页符
        if (para.SelectSingleNode("w:r/w:br[@w:type='page']", ns) != null)
            pdf.PageBreak();

        // 段落内图片
        foreach (XmlElement drawing in para.SelectNodes(".//w:drawing", ns)!)
            RenderImage(drawing, ns, pdf, zip, rels);

        // 提取文本 + 检测粗体（首个 run 含粗体则整段加粗）
        var textBuilder = new System.Text.StringBuilder();
        var isBold = false;
        var isItalic = false;
        foreach (XmlNode child in para.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            if (el.LocalName == "r")
            {
                foreach (XmlElement t in el.SelectNodes("w:t", ns)!)
                    textBuilder.Append(t.InnerText);
                var rPr = el.SelectSingleNode("w:rPr", ns) as XmlElement;
                if (rPr?.SelectSingleNode("w:b", ns) != null) isBold = true;
                if (rPr?.SelectSingleNode("w:i", ns) != null) isItalic = true;
            }
            else if (el.LocalName == "hyperlink")
            {
                foreach (XmlElement t in el.SelectNodes(".//w:t", ns)!)
                    textBuilder.Append(t.InnerText);
            }
        }

        // 文本框内容并入（wps:txbx/w:txbxContent，W22，与 Word 导出行为一致）
        foreach (XmlElement txbxContent in para.SelectNodes(".//w:txbxContent", ns)!)
        {
            foreach (XmlElement t in txbxContent.SelectNodes(".//w:t", ns)!)
                textBuilder.Append(t.InnerText);
        }

        var text = textBuilder.ToString();

        // 列表前缀
        var prefix = GetListPrefix(para, ns, numFmts);
        if (prefix != null) text = prefix + text;

        if (text.Length == 0) return;

        // 标题级别 → 字号
        var styleVal = para.SelectSingleNode("w:pPr/w:pStyle/@w:val", ns)?.Value ?? String.Empty;
        var headingLevel = GetHeadingLevel(styleVal);
        var fontSize = headingLevel switch
        {
            1 => H1FontSize,
            2 => H2FontSize,
            3 => H3FontSize,
            4 => H4FontSize,
            5 or 6 => H56FontSize,
            _ => BodyFontSize
        };

        if (headingLevel > 0)
            pdf.AddEmptyLine(8f);

        // 粗斜体 → 使用对应 PDF 标准字体
        PdfFont? font = null;
        if (isBold && isItalic) font = pdf.CreateFont("Helvetica-BoldOblique");
        else if (isBold) font = pdf.CreateFont("Helvetica-Bold");
        else if (isItalic) font = pdf.CreateFont("Helvetica-Oblique");
        pdf.AddText(text, fontSize, font);
    }

    /// <summary>渲染段落内 DrawingML 图片到 PDF</summary>
    private static void RenderImage(XmlElement drawing, XmlNamespaceManager ns, PdfDocumentBuilder pdf,
        ZipArchive zip, Dictionary<String, String> rels)
    {
        var blip = drawing.SelectSingleNode(".//a:blip", ns) as XmlElement;
        var rId = blip?.GetAttribute("r:embed");
        if (rId == null || !rels.TryGetValue(rId, out var target)) return;

        var entry = zip.GetEntry($"word/{target}");
        if (entry == null) return;
        using var ms = new MemoryStream();
        using (var es = entry.Open()) es.CopyTo(ms);
        if (ms.Length == 0) return;

        // 尺寸：EMU → 磅（1pt = 12700 EMU）
        var extent = drawing.SelectSingleNode(".//*[local-name()='extent']", ns) as XmlElement;
        Double w = 0, h = 0;
        if (extent != null)
        {
            Double.TryParse(extent.GetAttribute("cx"), out w);
            Double.TryParse(extent.GetAttribute("cy"), out h);
        }
        if (w <= 0 || h <= 0) return;
        pdf.AddImage(ms.ToArray(), (Single)(w / 12700), (Single)(h / 12700));
        pdf.AddEmptyLine(2f);
    }

    /// <summary>根据 numPr 与 numFmts 返回列表前缀（bullet="• "，编号=十进制+"."），非列表返回 null</summary>
    private static String? GetListPrefix(XmlElement para, XmlNamespaceManager ns, Dictionary<Int32, String> numFmts)
    {
        var numPr = para.SelectSingleNode("w:pPr/w:numPr", ns) as XmlElement;
        if (numPr == null) return null;
        var numIdEl = numPr.SelectSingleNode("w:numId", ns) as XmlElement;
        var numIdStr = numIdEl?.GetAttribute("w:val");
        if (numIdStr == "0") return null;
        if (Int32.TryParse(numIdStr, out var numId) && numFmts.TryGetValue(numId, out var fmt))
            return fmt == "bullet" ? "• " : "1. ";
        return "• ";
    }

    /// <summary>将表格节点渲染为 PDF 表格</summary>
    /// <param name="tbl">表格 XML 元素</param>
    /// <param name="ns">命名空间管理器</param>
    /// <param name="pdf">目标 PDF 文档</param>
    private static void RenderTable(XmlElement tbl, XmlNamespaceManager ns, PdfDocumentBuilder pdf)
    {
        var rows = new List<String[]>();
        foreach (XmlElement tr in tbl.SelectNodes("w:tr", ns)!)
        {
            var cells = new List<String>();
            foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
            {
                var sb = new System.Text.StringBuilder();
                foreach (XmlElement t in tc.SelectNodes(".//w:t", ns)!)
                {
                    sb.Append(t.InnerText);
                }
                cells.Add(sb.ToString());
            }
            if (cells.Count > 0)
                rows.Add(cells.ToArray());
        }
        if (rows.Count == 0) return;

        pdf.AddEmptyLine(4f);
        pdf.AddTable(rows, firstRowHeader: true, columnWidths: null);
        pdf.AddEmptyLine(4f);
    }

    /// <summary>解析样式名称返回标题级别（1-6），0 表示正文</summary>
    /// <param name="styleVal">段落样式值（如 "Heading1"/"heading 1"/"1" 等变体）</param>
    /// <returns>标题级别 1-6，或 0</returns>
    private static Int32 GetHeadingLevel(String styleVal)
    {
        if (String.IsNullOrEmpty(styleVal)) return 0;
        var v = styleVal.Trim().ToLowerInvariant().Replace(" ", String.Empty);

        if (v.StartsWith("heading") || v.StartsWith("标题"))
        {
            var suffix = v.TrimStart('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
                'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z');
            if (suffix.Length == 1 && suffix[0] >= '1' && suffix[0] <= '6')
                return suffix[0] - '0';
        }
        else if (v.Length == 1 && v[0] >= '1' && v[0] <= '6')
        {
            return v[0] - '0';
        }
        return 0;
    }

    #endregion
}
