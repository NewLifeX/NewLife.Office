using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>XPS 文档页面信息</summary>
public class XpsPage
{
    /// <summary>页宽（1/96 英寸单位）</summary>
    public Double Width { get; set; }

    /// <summary>页高（1/96 英寸单位）</summary>
    public Double Height { get; set; }

    /// <summary>本页提取的文本（Glyphs UnicodeString 拼接）</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>本页 Glyphs 文本片段列表</summary>
    public List<String> Glyphs { get; set; } = [];
}

/// <summary>XPS 文档属性</summary>
public class XpsProperties
{
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Creator { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述</summary>
    public String? Description { get; set; }
}

/// <summary>XPS（XML Paper Specification）文档读取器</summary>
/// <remarks>
/// XPS 是基于 Open Packaging Conventions 的 ZIP+XML 格式。
/// 本类实现从 .xps 文件中提取页数、页面尺寸、文本（Glyphs 元素）和图片资源。
/// <para>用法示例：</para>
/// <code>
/// var reader = new XpsReader();
/// var doc = reader.Read("document.xps");
/// Console.WriteLine($"{doc.Count} 页，全文：{string.Join("", doc.Select(p => p.Text))}");
/// </code>
/// </remarks>
public class XpsReader
{
    private const String XpsNs = "http://schemas.microsoft.com/xps/2005/06";

    #region 读取

    /// <summary>从文件路径读取 XPS 文档</summary>
    /// <param name="path">XPS 文件路径</param>
    /// <returns>页面列表</returns>
    public List<XpsPage> Read(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Open, FileAccess.Read, FileShare.Read);
        return Read(fs);
    }

    /// <summary>从流读取 XPS 文档</summary>
    /// <param name="stream">包含 XPS 内容的流</param>
    /// <returns>页面列表</returns>
    public List<XpsPage> Read(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var pageUris = DiscoverPageUris(zip);
        var pages = new List<XpsPage>();
        foreach (var uri in pageUris)
        {
            var page = ParsePage(zip, uri);
            if (page != null) pages.Add(page);
        }
        return pages;
    }

    /// <summary>读取 XPS 文档元数据</summary>
    /// <param name="path">XPS 文件路径</param>
    /// <returns>文档属性</returns>
    public XpsProperties ReadProperties(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Open, FileAccess.Read, FileShare.Read);
        return ReadProperties(fs);
    }

    /// <summary>从流读取 XPS 文档元数据</summary>
    /// <param name="stream">包含 XPS 内容的流</param>
    /// <returns>文档属性</returns>
    public XpsProperties ReadProperties(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        return ReadDocProps(zip);
    }

    /// <summary>提取所有嵌入图片（资源路径 → 字节内容）</summary>
    /// <param name="path">XPS 文件路径</param>
    /// <returns>（相对路径, 字节数组）序列</returns>
    public IEnumerable<(String Path, Byte[] Data)> ExtractImages(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Open, FileAccess.Read, FileShare.Read);
        return ExtractImages(fs).ToList();
    }

    /// <summary>从流提取所有嵌入图片</summary>
    /// <param name="stream">包含 XPS 内容的流</param>
    /// <returns>（相对路径, 字节数组）序列</returns>
    public IEnumerable<(String Path, Byte[] Data)> ExtractImages(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var results = new List<(String, Byte[])>();
        foreach (var entry in zip.Entries)
        {
            var name = entry.Name.ToLowerInvariant();
            if (!name.EndsWith(".png") && !name.EndsWith(".jpg") && !name.EndsWith(".jpeg")) continue;
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            results.Add((entry.FullName, ms.ToArray()));
        }
        return results;
    }

    #endregion

    #region 内部解析

    private static List<String> DiscoverPageUris(ZipArchive zip)
    {
        // 1. 读取 _rels/.rels 找 FixedDocumentSequence
        var fdsUri = FindFixedDocumentSequenceUri(zip);
        if (fdsUri == null) return [];

        // 2. 读取 FixedDocumentSequence，找 FixedDocument 列表
        var fdUris = ParseFixedDocumentSequence(zip, fdsUri);

        // 3. 读取每个 FixedDocument，收集 Page 路径
        var pageUris = new List<String>();
        foreach (var fdUri in fdUris)
            pageUris.AddRange(ParseFixedDocument(zip, fdUri));

        return pageUris;
    }

    private static String? FindFixedDocumentSequenceUri(ZipArchive zip)
    {
        var entry = zip.GetEntry("_rels/.rels");
        if (entry == null) return null;

        var doc = LoadXml(entry);
        if (doc.DocumentElement == null) return null;
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

        // 查找 FixedDocumentSequence 类型的关系
        const String TypeFds = "http://schemas.microsoft.com/xps/2005/06/fixedrepresentation";
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var type = rel.GetAttribute("Type");
            if (type.EndsWith("/fixedrepresentation") || type == TypeFds)
                return NormalizeUri(rel.GetAttribute("Target"));
        }

        // 回退：取第一个非空 Target
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var target = rel.GetAttribute("Target");
            if (!String.IsNullOrEmpty(target))
                return NormalizeUri(target);
        }
        return null;
    }

    private static List<String> ParseFixedDocumentSequence(ZipArchive zip, String fdsUri)
    {
        var entry = zip.GetEntry(fdsUri);
        if (entry == null) return [];

        var doc = LoadXml(entry);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("x", XpsNs);

        var result = new List<String>();
        foreach (XmlElement dr in doc.SelectNodes("//x:DocumentReference | //DocumentReference", ns)!)
        {
            var src = dr.GetAttribute("Source");
            if (!String.IsNullOrEmpty(src))
                result.Add(ResolveUri(fdsUri, src));
        }
        return result;
    }

    private static List<String> ParseFixedDocument(ZipArchive zip, String fdUri)
    {
        var entry = zip.GetEntry(fdUri);
        if (entry == null) return [];

        var doc = LoadXml(entry);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("x", XpsNs);

        var result = new List<String>();
        foreach (XmlElement pc in doc.SelectNodes("//x:PageContent | //PageContent", ns)!)
        {
            var src = pc.GetAttribute("Source");
            if (!String.IsNullOrEmpty(src))
                result.Add(ResolveUri(fdUri, src));
        }
        return result;
    }

    private static XpsPage? ParsePage(ZipArchive zip, String pageUri)
    {
        var entry = zip.GetEntry(pageUri);
        if (entry == null) return null;

        var xmlDoc = LoadXml(entry);
        var root = xmlDoc.DocumentElement;
        if (root == null) return null;

        var page = new XpsPage();
        if (Double.TryParse(root.GetAttribute("Width"), out var w)) page.Width = w;
        if (Double.TryParse(root.GetAttribute("Height"), out var h)) page.Height = h;

        var ns = new XmlNamespaceManager(xmlDoc.NameTable);
        ns.AddNamespace("x", XpsNs);

        var sb = new System.Text.StringBuilder();
        foreach (XmlElement g in xmlDoc.SelectNodes("//x:Glyphs | //Glyphs", ns)!)
        {
            var unicode = g.GetAttribute("UnicodeString");
            if (!String.IsNullOrEmpty(unicode))
            {
                page.Glyphs.Add(unicode);
                sb.Append(unicode);
            }
        }
        page.Text = sb.ToString();
        return page;
    }

    private static XpsProperties ReadDocProps(ZipArchive zip)
    {
        var props = new XpsProperties();
        var entry = zip.GetEntry("docProps/core.xml");
        if (entry == null) return props;

        var doc = LoadXml(entry);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
        ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");

        props.Title       = doc.SelectSingleNode("//dc:title", ns)?.InnerText;
        props.Creator     = doc.SelectSingleNode("//dc:creator", ns)?.InnerText;
        props.Subject     = doc.SelectSingleNode("//dc:subject", ns)?.InnerText;
        props.Description = doc.SelectSingleNode("//dc:description", ns)?.InnerText;
        return props;
    }

    private static XmlDocument LoadXml(ZipArchiveEntry entry)
    {
        var doc = new XmlDocument { XmlResolver = null };
        using var s = entry.Open();
        doc.Load(s);
        return doc;
    }

    private static String NormalizeUri(String uri) =>
        uri.TrimStart('/');

    private static String ResolveUri(String baseUri, String relative)
    {
        if (relative.StartsWith("/")) return relative.TrimStart('/');
        var dir = baseUri.Contains('/') ? baseUri[..(baseUri.LastIndexOf('/') + 1)] : String.Empty;
        return dir + relative;
    }

    #endregion
}
