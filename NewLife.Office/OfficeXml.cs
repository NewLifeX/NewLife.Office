using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace NewLife.Office;

/// <summary>安全的 XML 读取与转义辅助</summary>
/// <remarks>
/// 统一禁止 DTD/外部实体、限制实体展开与文档字符数；写入侧清理 XML 1.0 非法字符与孤立代理项。
/// </remarks>
internal static class OfficeXml
{
    /// <summary>创建安全读取设置（复制调用方设置后强制覆盖安全项）</summary>
    /// <param name="settings">调用方设置，可为 null</param>
    /// <returns>安全设置</returns>
    public static XmlReaderSettings CreateSettings(XmlReaderSettings? settings = null)
    {
        var result = settings?.Clone() ?? new XmlReaderSettings();
        result.DtdProcessing = DtdProcessing.Ignore;
        result.XmlResolver = null;
        result.MaxCharactersFromEntities = 0;
        if (result.MaxCharactersInDocument == 0) result.MaxCharactersInDocument = OfficeLimits.MaxXmlCharacters;
        return result;
    }

    /// <summary>从流创建安全 XML 读取器</summary>
    /// <param name="stream">输入流</param>
    /// <param name="settings">调用方设置，可为 null</param>
    /// <returns>XML 读取器</returns>
    public static XmlReader CreateReader(Stream stream, XmlReaderSettings? settings = null) => XmlReader.Create(stream, CreateSettings(settings));

    /// <summary>从字符串创建安全 XML 读取器</summary>
    /// <param name="xml">XML 文本</param>
    /// <param name="settings">调用方设置，可为 null</param>
    /// <returns>XML 读取器</returns>
    public static XmlReader CreateReader(String xml, XmlReaderSettings? settings = null) => XmlReader.Create(new StringReader(xml), CreateSettings(settings));

    /// <summary>安全加载 XML 文档</summary>
    /// <param name="stream">输入流</param>
    /// <returns>XML 文档</returns>
    public static XmlDocument Load(Stream stream)
    {
        var doc = new XmlDocument { XmlResolver = null };
        using var reader = CreateReader(stream);
        doc.Load(reader);
        return doc;
    }

    /// <summary>从 ZIP 条目安全加载 XML 文档</summary>
    /// <param name="entry">ZIP 条目</param>
    /// <returns>XML 文档</returns>
    public static XmlDocument Load(ZipArchiveEntry entry)
    {
        using var stream = OfficeStreams.OpenEntry(entry);
        return Load(stream);
    }

    /// <summary>安全加载 XML 文本</summary>
    /// <param name="xml">XML 文本</param>
    /// <returns>XML 文档</returns>
    public static XmlDocument LoadXml(String xml)
    {
        var doc = new XmlDocument { XmlResolver = null };
        using var reader = CreateReader(xml);
        doc.Load(reader);
        return doc;
    }

    /// <summary>把流内容安全加载到已有 XML 文档</summary>
    /// <param name="doc">目标文档</param>
    /// <param name="stream">输入流</param>
    public static void LoadInto(XmlDocument doc, Stream stream)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        doc.XmlResolver = null;
        using var reader = CreateReader(stream);
        doc.Load(reader);
    }

    /// <summary>把 XML 文本安全加载到已有 XML 文档</summary>
    /// <param name="doc">目标文档</param>
    /// <param name="xml">XML 文本</param>
    public static void LoadXmlInto(XmlDocument doc, String xml)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        doc.XmlResolver = null;
        using var reader = CreateReader(xml);
        doc.Load(reader);
    }

    /// <summary>安全加载 LINQ to XML 文档</summary>
    /// <param name="stream">输入流</param>
    /// <returns>XML 文档</returns>
    public static XDocument LoadX(Stream stream)
    {
        using var reader = CreateReader(stream);
        return XDocument.Load(reader);
    }

    /// <summary>从 ZIP 条目安全加载 LINQ to XML 文档</summary>
    /// <param name="entry">ZIP 条目</param>
    /// <returns>XML 文档</returns>
    public static XDocument LoadX(ZipArchiveEntry entry)
    {
        using var stream = OfficeStreams.OpenEntry(entry);
        return LoadX(stream);
    }

    /// <summary>清理 XML 1.0 非法字符与孤立代理项</summary>
    /// <param name="text">原始文本</param>
    /// <returns>可安全写入 XML 的文本</returns>
    public static String Sanitize(String? text)
    {
        if (String.IsNullOrEmpty(text)) return String.Empty;

        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (Char.IsHighSurrogate(ch))
            {
                if (i + 1 < text.Length && Char.IsLowSurrogate(text[i + 1]))
                {
                    sb.Append(ch).Append(text[i + 1]);
                    i++;
                }
                else
                {
                    sb.Append('\uFFFD');
                }
                continue;
            }
            if (Char.IsLowSurrogate(ch))
            {
                sb.Append('\uFFFD');
                continue;
            }
            if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20 && ch <= 0xD7FF || ch >= 0xE000 && ch <= 0xFFFD)
                sb.Append(ch);
            else
                sb.Append('\uFFFD');
        }
        return sb.ToString();
    }

    /// <summary>清理并转义 XML 文本/属性值</summary>
    /// <param name="text">原始文本</param>
    /// <returns>转义后的文本</returns>
    public static String Escape(String? text)
    {
        var clean = Sanitize(text);
        if (clean.Length == 0) return String.Empty;

        return clean
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
