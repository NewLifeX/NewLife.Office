using System.IO.Compression;
using System.Xml;
using NewLife.Office;

namespace NewLife.Office;

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
        using var pdf = new PdfFluentDocument();
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
        using var pdf = new PdfFluentDocument();
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
    /// <param name="pdf">目标 PdfFluentDocument</param>
    private void Render(Stream stream, PdfFluentDocument pdf)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 加载 word/document.xml
        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return;

        var doc = new XmlDocument();
        using (var s = entry.Open())
            doc.Load(s);

        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        var body = doc.SelectSingleNode("//w:body", ns);
        if (body == null) return;

        foreach (XmlNode node in body.ChildNodes)
        {
            if (node is not XmlElement el) continue;
            if (el.LocalName == "p")
                RenderParagraph(el, ns, pdf);
            else if (el.LocalName == "tbl")
                RenderTable(el, ns, pdf);
        }
    }

    /// <summary>将段落节点渲染为 PDF 文本行</summary>
    /// <param name="para">段落 XML 元素</param>
    /// <param name="ns">命名空间管理器</param>
    /// <param name="pdf">目标 PDF 文档</param>
    private void RenderParagraph(XmlElement para, XmlNamespaceManager ns, PdfFluentDocument pdf)
    {
        // 提取段落文本
        var textBuilder = new System.Text.StringBuilder();
        foreach (XmlElement t in para.SelectNodes(".//w:t", ns)!)
            textBuilder.Append(t.InnerText);
        var text = textBuilder.ToString();
        if (text.Length == 0) return;

        // 检测段落样式（标题级别）
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

        // 标题前加空行
        if (headingLevel > 0)
            pdf.AddEmptyLine(8f);

        pdf.AddText(text, fontSize);
    }

    /// <summary>将表格节点渲染为 PDF 表格</summary>
    /// <param name="tbl">表格 XML 元素</param>
    /// <param name="ns">命名空间管理器</param>
    /// <param name="pdf">目标 PDF 文档</param>
    private static void RenderTable(XmlElement tbl, XmlNamespaceManager ns, PdfFluentDocument pdf)
    {
        var rows = new List<String[]>();
        foreach (XmlElement tr in tbl.SelectNodes("w:tr", ns)!)
        {
            var cells = new List<String>();
            foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
            {
                var sb = new System.Text.StringBuilder();
                foreach (XmlElement t in tc.SelectNodes(".//w:t", ns)!)
                    sb.Append(t.InnerText);
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
