using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>Word docx 读取器</summary>
/// <remarks>
/// 直接解析 Open XML（ZIP+XML）提取文本、表格、图片等内容。
/// </remarks>
public class WordReader : IDisposable
{
    #region 属性
    /// <summary>源文件路径（从文件构造时有效）</summary>
    public String? FilePath { get; private set; }
    #endregion

    #region 私有字段
    private readonly ZipArchive _zip;
    private Boolean _disposed;
    #endregion

    #region 构造
    /// <summary>从文件路径打开</summary>
    /// <param name="path">docx 文件路径</param>
    public WordReader(String path)
    {
        FilePath = path.GetFullPath();
        _zip = ZipFile.OpenRead(FilePath);
    }

    /// <summary>从流打开</summary>
    /// <param name="stream">包含 docx 内容的流</param>
    public WordReader(Stream stream)
    {
        _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
    }

    /// <summary>释放资源</summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _zip.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
    #endregion

    #region 读取方法
    /// <summary>读取所有段落文本</summary>
    /// <returns>段落字符串序列</returns>
    public IEnumerable<String> ReadParagraphs()
    {
        var doc = LoadDocumentXml();
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        foreach (XmlElement para in doc.SelectNodes("//w:p", ns)!)
        {
            var sb = new StringBuilder();
            foreach (XmlElement t in para.SelectNodes(".//w:t", ns)!)
            {
                sb.Append(t.InnerText);
            }
            var text = sb.ToString();
            if (text.Length > 0)
                yield return text;
        }
    }

    /// <summary>读取全文（段落间用换行分隔）</summary>
    /// <returns>完整文本</returns>
    public String ReadFullText() => String.Join(Environment.NewLine, ReadParagraphs());

    /// <summary>读取所有表格数据</summary>
    /// <returns>每个表格是 string[][] 的序列</returns>
    public IEnumerable<String[][]> ReadTables()
    {
        var doc = LoadDocumentXml();
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        foreach (XmlElement tbl in doc.SelectNodes("//w:tbl", ns)!)
        {
            var rows = new List<String[]>();
            foreach (XmlElement tr in tbl.SelectNodes("w:tr", ns)!)
            {
                var cells = new List<String>();
                foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
                {
                    var sb = new StringBuilder();
                    foreach (XmlElement t in tc.SelectNodes(".//w:t", ns)!)
                    {
                        sb.Append(t.InnerText);
                    }
                    cells.Add(sb.ToString());
                }
                if (cells.Count > 0)
                    rows.Add(cells.ToArray());
            }
            if (rows.Count > 0)
                yield return rows.ToArray();
        }
    }

    /// <summary>提取所有图片数据</summary>
    /// <returns>（扩展名, 字节数据）序列</returns>
    public IEnumerable<(String Extension, Byte[] Data)> ExtractImages()
    {
        foreach (var entry in _zip.Entries)
        {
            if (!entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                continue;
            var ext = Path.GetExtension(entry.Name).TrimStart('.').ToLowerInvariant();
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            yield return (ext, ms.ToArray());
        }
    }

    /// <summary>获取文档属性</summary>
    /// <returns>属性对象</returns>
    public WordProperties GetProperties()
    {
        var props = new WordProperties();
        var entry = _zip.GetEntry("docProps/core.xml");
        if (entry == null) return props;

        var doc = new XmlDocument();
        using (var s = entry.Open())
            doc.Load(s);

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
        ns.AddNamespace("dcterms", "http://purl.org/dc/terms/");
        ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");

        props.Title = doc.SelectSingleNode("//dc:title", ns)?.InnerText;
        props.Author = doc.SelectSingleNode("//dc:creator", ns)?.InnerText;
        props.Subject = doc.SelectSingleNode("//dc:subject", ns)?.InnerText;
        props.Description = doc.SelectSingleNode("//dc:description", ns)?.InnerText;
        var createdText = doc.SelectSingleNode("//dcterms:created", ns)?.InnerText;
        if (DateTime.TryParse(createdText, out var dt))
            props.Created = dt;

        return props;
    }

    /// <summary>读取对象集合（将第一行表格映射到属性）</summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <returns>对象序列</returns>
    public IEnumerable<T> ReadObjects<T>() where T : class, new()
    {
        var props = typeof(T).GetProperties();
        foreach (var tbl in ReadTables())
        {
            if (tbl.Length < 2) continue;
            var headers = tbl[0];
            for (var ri = 1; ri < tbl.Length; ri++)
            {
                var row = tbl[ri];
                var obj = new T();
                for (var ci = 0; ci < Math.Min(headers.Length, row.Length); ci++)
                {
                    var hdr = headers[ci].Trim();
                    var prop = props.FirstOrDefault(p =>
                        p.Name.Equals(hdr, StringComparison.OrdinalIgnoreCase) ||
                        p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                         .OfType<System.ComponentModel.DisplayNameAttribute>().Any(a => a.DisplayName == hdr));
                    if (prop == null) continue;
                    try
                    {
                        var value = row[ci];
                        if (prop.PropertyType == typeof(String))
                            prop.SetValue(obj, value);
                        else
                            prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
                    }
                    catch { /* skip conversion errors */ }
                }
                yield return obj;
            }
        }
    }
    #endregion

    #region 私有方法
    private XmlDocument LoadDocumentXml()
    {
        var entry = _zip.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("无效的 docx 文件：缺少 word/document.xml");
        var doc = new XmlDocument();
        using var s = entry.Open();
        doc.Load(s);
        return doc;
    }
    #endregion
}
