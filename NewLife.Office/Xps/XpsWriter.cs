using System.IO.Compression;
using System.Text;

namespace NewLife.Office;

/// <summary>XPS（XML Paper Specification）文档写入器</summary>
/// <remarks>
/// 生成符合 XPS 规范的 .xps 文件（ZIP + XML）。
/// 支持多页文本写入、元数据、简单图片嵌入。
/// 不嵌入实际字体文件，生成的文件结构合法，但在部分渲染器中字体可能回退。
/// <para>用法示例：</para>
/// <code>
/// var writer = new XpsWriter();
/// writer.SetProperties(new XpsProperties { Title = "My Document" });
/// writer.AddPage(816, 1056, new[] { ("Hello, World!", 96.0, 96.0, 16.0) });
/// writer.Save("output.xps");
/// </code>
/// </remarks>
public class XpsWriter
{
    private const String XpsNs = "http://schemas.microsoft.com/xps/2005/06";

    #region 内部状态

    private readonly List<XpsPage> _pages = [];
    private readonly List<(String Text, Double OriginX, Double OriginY, Double FontSize)[]> _pageGlyphs = [];
    private XpsProperties _properties = new();
    private readonly List<(String Name, Byte[] Data, String MimeType)> _images = [];

    #endregion

    #region 构建方法

    /// <summary>设置文档属性（标题、作者等）</summary>
    /// <param name="props">属性对象</param>
    public void SetProperties(XpsProperties props)
    {
        _properties = props ?? new XpsProperties();
    }

    /// <summary>添加一个文本页</summary>
    /// <param name="widthDip">页宽（device-independent pixel，1/96 英寸）</param>
    /// <param name="heightDip">页高</param>
    /// <param name="glyphs">文本片段列表：（文本, X, Y, 字号）</param>
    public void AddPage(Double widthDip, Double heightDip,
        IEnumerable<(String Text, Double OriginX, Double OriginY, Double FontSize)> glyphs)
    {
        var page = new XpsPage { Width = widthDip, Height = heightDip };
        foreach (var g in glyphs)
        {
            page.Glyphs.Add(g.Text);
            page.Text += g.Text;
        }
        _pages.Add(page);
        _pageGlyphs.Add(glyphs.ToArray());
    }

    /// <summary>嵌入图片资源</summary>
    /// <param name="name">资源名称（如 image1.png）</param>
    /// <param name="data">图片字节数据</param>
    /// <param name="mimeType">MIME 类型，默认 image/png</param>
    public void AddImage(String name, Byte[] data, String mimeType = "image/png")
    {
        _images.Add((name, data, mimeType));
    }

    #endregion

    #region 保存

    /// <summary>保存 XPS 文件到路径</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        Save(fs);
    }

    /// <summary>保存 XPS 文件到流</summary>
    /// <param name="stream">可写输出流</param>
    public void Save(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteContentTypes(zip);
        WriteRootRels(zip);
        WriteFixedDocumentSequence(zip);
        WriteFixedDocument(zip);
        WritePages(zip);
        WriteDocProps(zip);
        foreach (var (name, data, _) in _images)
        {
            WriteEntry(zip, $"Resources/Images/{name}", data);
        }
    }

    /// <summary>获取 XPS 字节数组</summary>
    /// <returns>XPS 文件字节数组</returns>
    public Byte[] ToBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }

    #endregion

    #region 内部写入

    private static void WriteEntry(ZipArchive zip, String fullName, String content, String encoding = "utf-8")
    {
        var data = Encoding.UTF8.GetBytes(content);
        WriteEntry(zip, fullName, data);
    }

    private static void WriteEntry(ZipArchive zip, String fullName, Byte[] data)
    {
        var entry = zip.CreateEntry(fullName, CompressionLevel.Optimal);
        using var s = entry.Open();
        s.Write(data, 0, data.Length);
    }

    private static void WriteContentTypes(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.AppendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.AppendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.AppendLine("  <Default Extension=\"fdseq\" ContentType=\"application/vnd.ms-package.xps-fixeddocumentsequence+xml\"/>");
        sb.AppendLine("  <Default Extension=\"fds\" ContentType=\"application/vnd.ms-package.xps-fixeddocument+xml\"/>");
        sb.AppendLine("  <Default Extension=\"fpage\" ContentType=\"application/vnd.ms-package.xps-fixedpage+xml\"/>");
        sb.AppendLine("  <Default Extension=\"png\" ContentType=\"image/png\"/>");
        sb.AppendLine("  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>");
        sb.AppendLine("  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        sb.AppendLine("</Types>");
        WriteEntry(zip, "[Content_Types].xml", sb.ToString());
    }

    private static void WriteRootRels(ZipArchive zip)
    {
        const String TypeFds = "http://schemas.microsoft.com/xps/2005/06/fixedrepresentation";
        const String TypeCoreProps = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.AppendLine("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.AppendLine($"  <Relationship Id=\"R1\" Type=\"{TypeFds}\" Target=\"FixedDocumentSequence.fdseq\"/>");
        sb.AppendLine($"  <Relationship Id=\"R2\" Type=\"{TypeCoreProps}\" Target=\"docProps/core.xml\"/>");
        sb.AppendLine("</Relationships>");
        WriteEntry(zip, "_rels/.rels", sb.ToString());
    }

    private static void WriteFixedDocumentSequence(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.AppendLine($"<FixedDocumentSequence xmlns=\"{XpsNs}\">");
        sb.AppendLine("  <DocumentReference Source=\"Documents/1/1.fds\"/>");
        sb.AppendLine("</FixedDocumentSequence>");
        WriteEntry(zip, "FixedDocumentSequence.fdseq", sb.ToString());

        // rels for the sequence
        var rels = new StringBuilder();
        rels.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        rels.AppendLine("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        rels.AppendLine("  <Relationship Id=\"R1\" Type=\"http://schemas.microsoft.com/xps/2005/06/fixedrepresentation\" Target=\"Documents/1/1.fds\"/>");
        rels.AppendLine("</Relationships>");
        WriteEntry(zip, "_rels/FixedDocumentSequence.fdseq.rels", rels.ToString());
    }

    private void WriteFixedDocument(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.AppendLine($"<FixedDocument xmlns=\"{XpsNs}\">");
        for (var i = 0; i < _pages.Count; i++)
        {
            sb.AppendLine($"  <PageContent Source=\"Pages/{i + 1}.fpage\"/>");
        }
        sb.AppendLine("</FixedDocument>");
        WriteEntry(zip, "Documents/1/1.fds", sb.ToString());
    }

    private void WritePages(ZipArchive zip)
    {
        for (var i = 0; i < _pages.Count; i++)
        {
            var page = _pages[i];
            var glyphs = _pageGlyphs[i];
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine($"<FixedPage Width=\"{page.Width}\" Height=\"{page.Height}\" xmlns=\"{XpsNs}\">");
            foreach (var (text, ox, oy, fs) in glyphs)
            {
                var escaped = text
                    .Replace("&", "&amp;")
                    .Replace("\"", "&quot;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;");
                sb.AppendLine($"  <Glyphs OriginX=\"{ox}\" OriginY=\"{oy}\" FontUri=\"/Resources/Fonts/Arial.ttf\"");
                sb.AppendLine($"          FontRenderingEmSize=\"{fs}\" Fill=\"#FF000000\"");
                sb.AppendLine($"          UnicodeString=\"{escaped}\"/>");
            }
            sb.AppendLine("</FixedPage>");
            WriteEntry(zip, $"Documents/1/Pages/{i + 1}.fpage", sb.ToString());
        }
    }

    private void WriteDocProps(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
        sb.AppendLine("<cp:coreProperties");
        sb.AppendLine("  xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"");
        sb.AppendLine("  xmlns:dc=\"http://purl.org/dc/elements/1.1/\"");
        sb.AppendLine("  xmlns:dcterms=\"http://purl.org/dc/terms/\"");
        sb.AppendLine("  xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        if (!String.IsNullOrEmpty(_properties.Title))
            sb.AppendLine($"  <dc:title>{Escape(_properties.Title)}</dc:title>");
        if (!String.IsNullOrEmpty(_properties.Creator))
            sb.AppendLine($"  <dc:creator>{Escape(_properties.Creator)}</dc:creator>");
        if (!String.IsNullOrEmpty(_properties.Subject))
            sb.AppendLine($"  <dc:subject>{Escape(_properties.Subject)}</dc:subject>");
        if (!String.IsNullOrEmpty(_properties.Description))
            sb.AppendLine($"  <dc:description>{Escape(_properties.Description)}</dc:description>");
        sb.AppendLine("</cp:coreProperties>");
        WriteEntry(zip, "docProps/core.xml", sb.ToString());
        WriteEntry(zip, "_rels/docProps/.rels", "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
    }

    private static String Escape(String input) =>
        input.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    #endregion
}
