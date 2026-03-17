using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>EPUB 电子书读取器，支持 EPUB 2/3</summary>
public class EpubReader
{
    #region 读取方法

    /// <summary>从文件路径读取 EPUB</summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>EPUB 文档</returns>
    public EpubDocument Read(String filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Read(fs);
    }

    /// <summary>从流读取 EPUB</summary>
    /// <param name="stream">输入流</param>
    /// <returns>EPUB 文档</returns>
    public EpubDocument Read(Stream stream)
    {
        var doc = new EpubDocument();
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // 1. 读取 META-INF/container.xml 获取 OPF 路径
        var opfPath = GetOpfPath(zip);
        if (opfPath == null) return doc;

        // 2. 解析 OPF
        var opfDir = GetDirectory(opfPath);
        var opfEntry = zip.GetEntry(opfPath);
        if (opfEntry == null) return doc;

        ParseOpf(doc, zip, opfEntry, opfDir);
        return doc;
    }

    #endregion

    #region 内部解析

    private static String? GetOpfPath(ZipArchive zip)
    {
        var container = zip.GetEntry("META-INF/container.xml");
        if (container == null) return null;

        using var stream = container.Open();
        var xml = new XmlDocument();
        xml.Load(stream);

        var ns = new XmlNamespaceManager(xml.NameTable);
        ns.AddNamespace("c", "urn:oasis:names:tc:opendocument:xmlns:container");

        var node = xml.SelectSingleNode("//c:rootfile/@full-path", ns)
                ?? xml.SelectSingleNode("//*[local-name()='rootfile']/@full-path");
        return node?.Value;
    }

    private static void ParseOpf(EpubDocument doc, ZipArchive zip, ZipArchiveEntry opfEntry, String opfDir)
    {
        using var stream = opfEntry.Open();
        var xml = new XmlDocument();
        xml.Load(stream);

        var ns = new XmlNamespaceManager(xml.NameTable);
        ns.AddNamespace("opf", "http://www.idpf.org/2007/opf");
        ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");

        // 元数据
        doc.Title = GetNodeText(xml, "//dc:title", ns);
        doc.Author = GetNodeText(xml, "//dc:creator", ns);
        doc.Language = GetNodeText(xml, "//dc:language", ns);
        doc.Publisher = GetNodeText(xml, "//dc:publisher", ns);
        doc.Description = GetNodeText(xml, "//dc:description", ns);
        doc.Identifier = GetNodeText(xml, "//dc:identifier", ns);
        doc.PublishDate = GetNodeText(xml, "//dc:date", ns);

        // manifest: 建立 id -> href 映射
        var manifest = new Dictionary<String, String>();
        var mediaTypes = new Dictionary<String, String>();
        var manifestNodes = xml.SelectNodes("//*[local-name()='item']");
        if (manifestNodes != null)
        {
            foreach (XmlNode item in manifestNodes)
            {
                var id = item.Attributes?["id"]?.Value;
                var href = item.Attributes?["href"]?.Value;
                var mt = item.Attributes?["media-type"]?.Value ?? "";
                if (id != null && href != null)
                {
                    manifest[id] = href;
                    mediaTypes[id] = mt;
                }
            }
        }

        // 封面
        var coverId = FindCoverId(xml);
        if (coverId != null && manifest.ContainsKey(coverId))
        {
            var coverPath = CombinePath(opfDir, manifest[coverId]);
            var coverEntry = zip.GetEntry(coverPath);
            if (coverEntry != null)
            {
                using var cs = coverEntry.Open();
                using var ms = new MemoryStream();
                cs.CopyTo(ms);
                doc.Cover = ms.ToArray();
                if (mediaTypes.ContainsKey(coverId))
                    doc.CoverMediaType = mediaTypes[coverId];
            }
        }

        // spine: 读取内容顺序
        var spineNodes = xml.SelectNodes("//*[local-name()='itemref']");
        if (spineNodes == null) return;

        var chapterIdx = 0;
        foreach (XmlNode itemref in spineNodes)
        {
            var idref = itemref.Attributes?["idref"]?.Value;
            if (idref == null || !manifest.ContainsKey(idref)) continue;

            var href = manifest[idref];
            var filePath = CombinePath(opfDir, href);
            var entry = zip.GetEntry(filePath);
            if (entry == null) continue;

            String content;
            using (var cs = entry.Open())
            using (var sr = new StreamReader(cs, Encoding.UTF8))
                content = sr.ReadToEnd();

            var title = ExtractTitle(content);
            if (String.IsNullOrEmpty(title)) title = "Chapter " + (++chapterIdx);

            doc.Chapters.Add(new EpubChapter
            {
                FileName = Path.GetFileName(href),
                Title = title,
                Content = content,
            });
        }
    }

    private static String? FindCoverId(XmlDocument xml)
    {
        // EPUB3 cover via properties
        var nodes = xml.SelectNodes("//*[local-name()='item']");
        if (nodes != null)
        {
            foreach (XmlNode n in nodes)
            {
                var props = n.Attributes?["properties"]?.Value ?? "";
                if (props.Contains("cover-image"))
                    return n.Attributes?["id"]?.Value;
            }
        }

        // EPUB2 cover via meta name="cover"
        var metaNodes = xml.SelectNodes("//*[local-name()='meta']");
        if (metaNodes != null)
        {
            foreach (XmlNode n in metaNodes)
            {
                if ((n.Attributes?["name"]?.Value ?? "").Equals("cover", StringComparison.OrdinalIgnoreCase))
                    return n.Attributes?["content"]?.Value;
            }
        }

        return null;
    }

    private static String GetNodeText(XmlDocument xml, String xpath, XmlNamespaceManager ns)
    {
        var node = xml.SelectSingleNode(xpath, ns);
        return node?.InnerText ?? String.Empty;
    }

    private static String ExtractTitle(String html)
    {
        var start = html.IndexOf("<title>", StringComparison.OrdinalIgnoreCase);
        if (start < 0) start = html.IndexOf("<h1", StringComparison.OrdinalIgnoreCase);
        if (start < 0) return String.Empty;

        var end = html.IndexOf('<', start + 7);
        if (end < 0) return String.Empty;

        var tagEnd = html.IndexOf('>', start);
        if (tagEnd < 0 || tagEnd >= end) return String.Empty;

        return html.Substring(tagEnd + 1, end - tagEnd - 1).Trim();
    }

    private static String GetDirectory(String path)
    {
        var idx = path.LastIndexOf('/');
        return idx < 0 ? String.Empty : path[..(idx + 1)];
    }

    private static String CombinePath(String dir, String href)
    {
        if (String.IsNullOrEmpty(dir)) return href;
        // handle ../ relative paths
        var parts = (dir + href).Split('/');
        var stack = new List<String>();
        foreach (var p in parts)
        {
            if (p == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); }
            else if (p != ".") stack.Add(p);
        }
        return String.Join("/", stack.ToArray());
    }

    #endregion
}
