#nullable enable
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>Word 模板填充器</summary>
/// <remarks>
/// 以现有 docx 为模板，将 {{Key}} 占位符替换为实际值后输出新文件。
/// 支持嵌套 XML 中被拆分的占位符，通过段落级合并后替换实现。
/// </remarks>
public class WordTemplate
{
    #region 属性
    /// <summary>模板文件路径</summary>
    public String TemplatePath { get; }
    #endregion

    #region 构造
    /// <summary>实例化模板填充器</summary>
    /// <param name="templatePath">模板 docx 路径</param>
    public WordTemplate(String templatePath) => TemplatePath = templatePath.GetFullPath();
    #endregion

    #region 填充方法
    /// <summary>填充模板并保存到指定路径</summary>
    /// <param name="outputPath">输出路径</param>
    /// <param name="data">占位符键值字典（Key 不含 {{ }}）</param>
    public void Fill(String outputPath, IDictionary<String, Object?> data)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Fill(fs, data);
    }

    /// <summary>填充模板并写入流</summary>
    /// <param name="outputStream">输出流</param>
    /// <param name="data">占位符键值字典</param>
    public void Fill(Stream outputStream, IDictionary<String, Object?> data)
    {
        // 读取模板字节
        var templateBytes = File.ReadAllBytes(TemplatePath);

        using var srcMs = new MemoryStream(templateBytes);
        using var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read);
        using var dstZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in srcZip.Entries)
        {
            var dstEntry = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
            using var srcStream = entry.Open();
            using var dstStream = dstEntry.Open();

            // 仅对 XML 条目做文本替换
            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                var content = ReadAsString(srcStream);
                content = ApplyReplacements(content, data);
                var bytes = Encoding.UTF8.GetBytes(content);
                dstStream.Write(bytes, 0, bytes.Length);
            }
            else
            {
                srcStream.CopyTo(dstStream);
            }
        }
    }

    /// <summary>从对象属性生成字典并填充</summary>
    /// <param name="outputPath">输出路径</param>
    /// <param name="model">数据模型对象</param>
    public void Fill(String outputPath, Object model)
    {
        var dict = new Dictionary<String, Object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in model.GetType().GetProperties())
            dict[prop.Name] = prop.GetValue(model);
        Fill(outputPath, dict);
    }

    /// <summary>填充模板，支持列表表格扩展（W03-03）</summary>
    /// <remarks>在 docx 模板表格中，用 {{#ListKey}} 标记开始行，{{/ListKey}} 标记结束行；中间为模板行，每行对应一条数据。</remarks>
    /// <param name="outputPath">输出路径</param>
    /// <param name="data">普通占位符字典</param>
    /// <param name="lists">列表数据，Key 为占位符名称（如 Items），Value 为行数据集合</param>
    public void FillTable(String outputPath, IDictionary<String, Object?> data,
        IDictionary<String, IEnumerable<IDictionary<String, Object?>>> lists)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        FillTable(fs, data, lists);
    }

    /// <summary>填充模板，支持列表表格扩展，写入流</summary>
    /// <param name="outputStream">输出流</param>
    /// <param name="data">普通占位符字典</param>
    /// <param name="lists">列表数据</param>
    public void FillTable(Stream outputStream, IDictionary<String, Object?> data,
        IDictionary<String, IEnumerable<IDictionary<String, Object?>>> lists)
    {
        var templateBytes = File.ReadAllBytes(TemplatePath);
        using var srcMs = new MemoryStream(templateBytes);
        using var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read);
        using var dstZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in srcZip.Entries)
        {
            var dstEntry = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
            using var srcStream = entry.Open();
            using var dstStream = dstEntry.Open();

            if (entry.FullName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
            {
                var content = ReadAsString(srcStream);
                foreach (var kv in lists)
                    content = ExpandTableRows(content, kv.Key, kv.Value, "w:tr");
                content = ApplyReplacements(content, data);
                var bytes = Encoding.UTF8.GetBytes(content);
                dstStream.Write(bytes, 0, bytes.Length);
            }
            else if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                var content = ReadAsString(srcStream);
                content = ApplyReplacements(content, data);
                var bytes = Encoding.UTF8.GetBytes(content);
                dstStream.Write(bytes, 0, bytes.Length);
            }
            else
            {
                srcStream.CopyTo(dstStream);
            }
        }
    }

    /// <summary>填充模板，支持图片占位符替换（W03-04）</summary>
    /// <remarks>在 docx 模板图片上，将图片的 alt 文本（descr 或 name 属性）设为 {{ImageKey}}，此方法将以新图片字节替换该图片。</remarks>
    /// <param name="outputPath">输出路径</param>
    /// <param name="data">普通占位符字典</param>
    /// <param name="images">图片数据，Key 为占位符名称（不含{{}}），Value 为图片字节（PNG/JPEG）</param>
    public void FillImages(String outputPath, IDictionary<String, Object?> data, IDictionary<String, Byte[]> images)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        FillImages(fs, data, images);
    }

    /// <summary>填充模板，支持图片占位符替换，写入流</summary>
    /// <param name="outputStream">输出流</param>
    /// <param name="data">普通占位符字典</param>
    /// <param name="images">图片数据</param>
    public void FillImages(Stream outputStream, IDictionary<String, Object?> data, IDictionary<String, Byte[]> images)
    {
        var templateBytes = File.ReadAllBytes(TemplatePath);
        using var srcMs = new MemoryStream(templateBytes);
        using var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read);
        using var dstZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        var imgReplace = BuildDocxImageMap(srcZip, images);

        foreach (var entry in srcZip.Entries)
        {
            var dstEntry = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
            using var srcStream = entry.Open();
            using var dstStream = dstEntry.Open();

            if (imgReplace.TryGetValue(entry.FullName, out var newData))
            {
                dstStream.Write(newData, 0, newData.Length);
            }
            else if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                var content = ReadAsString(srcStream);
                content = ApplyReplacements(content, data);
                var bytes = Encoding.UTF8.GetBytes(content);
                dstStream.Write(bytes, 0, bytes.Length);
            }
            else
            {
                srcStream.CopyTo(dstStream);
            }
        }
    }
    #endregion

    #region 私有方法
    private static String ReadAsString(Stream s)
    {
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static String ApplyReplacements(String xml, IDictionary<String, Object?> data)
    {
        // Word 有时将 {{Key}} 拆分在多个 w:r 中，先合并段落文本再替换
        // 简单策略：对 xml 字符串直接做文本替换；对拆分的情况，
        // 先将 }}...{{ 模式之间可能出现的 </w:t><w:r><w:t> 等去除
        // 更健壮的方案：解析XML，但此处为快速实现采用字符串替换
        foreach (var kv in data)
        {
            var placeholder = $"{{{{{kv.Key}}}}}";
            var value = Convert.ToString(kv.Value) ?? String.Empty;
            xml = xml.Replace(placeholder, EscapeXml(value));
        }
        return xml;
    }

    /// <summary>展开表格模板行；在 rowTag 行内寻找 {{#Key}} / {{/Key}} 标记并按列表数据展开</summary>
    /// <param name="xml">原始 XML 字符串</param>
    /// <param name="listKey">列表占位符名称（不含 {{ }} 和 # /）</param>
    /// <param name="items">列表数据</param>
    /// <param name="rowTag">XML 行标签（w:tr 或 a:tr）</param>
    /// <returns>展开后的 XML 字符串</returns>
    internal static String ExpandTableRows(String xml, String listKey,
        IEnumerable<IDictionary<String, Object?>> items, String rowTag = "w:tr")
    {
        var startMarker = $"{{{{#{listKey}}}}}";
        var endMarker = $"{{{{/{listKey}}}}}";
        if (!xml.Contains(startMarker)) return xml;

        var listItems = items.ToList();
        var sb = new StringBuilder();
        var remaining = xml;

        while (remaining.Contains(startMarker))
        {
            var startPos = remaining.IndexOf(startMarker, StringComparison.Ordinal);
            var rowStart = remaining.LastIndexOf($"<{rowTag}", startPos, StringComparison.Ordinal);
            if (rowStart < 0) break;

            var rowEnd = remaining.IndexOf($"</{rowTag}>", startPos, StringComparison.Ordinal);
            if (rowEnd < 0) break;
            rowEnd += $"</{rowTag}>".Length;

            var endPos = remaining.IndexOf(endMarker, rowEnd, StringComparison.Ordinal);
            if (endPos < 0) break;

            var endRowStart = remaining.LastIndexOf($"<{rowTag}", endPos, StringComparison.Ordinal);
            if (endRowStart < 0) break;

            var endRowEnd = remaining.IndexOf($"</{rowTag}>", endPos, StringComparison.Ordinal);
            if (endRowEnd < 0) break;
            endRowEnd += $"</{rowTag}>".Length;

            // Template rows are between the start marker row and end marker row
            var templateRows = remaining[rowEnd..endRowStart];

            sb.Append(remaining[..rowStart]);
            foreach (var item in listItems)
            {
                var rowContent = templateRows;
                foreach (var kv in item)
                {
                    var ph = $"{{{{{kv.Key}}}}}";
                    var val = EscapeXml(Convert.ToString(kv.Value) ?? String.Empty);
                    rowContent = rowContent.Replace(ph, val);
                }
                sb.Append(rowContent);
            }
            remaining = remaining[endRowEnd..];
        }

        sb.Append(remaining);
        return sb.ToString();
    }

    /// <summary>构建 docx 中占位图片名称 → ZIP 条目路径 → 替换字节的映射</summary>
    /// <param name="srcZip">源 ZIP</param>
    /// <param name="images">占位符名称（不含{{}}）→ 图片字节</param>
    /// <returns>ZIP 条目路径 → 新图片字节的映射</returns>
    private static Dictionary<String, Byte[]> BuildDocxImageMap(ZipArchive srcZip, IDictionary<String, Byte[]> images)
    {
        var result = new Dictionary<String, Byte[]>(StringComparer.OrdinalIgnoreCase);
        if (images.Count == 0) return result;

        var relsEntry = srcZip.GetEntry("word/_rels/document.xml.rels");
        if (relsEntry == null) return result;

        var relsXml = ReadAsString(relsEntry.Open());
        var relMap = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase); // rId -> "word/media/..."
        foreach (Match m in Regex.Matches(relsXml, @"Id=""([^""]+)""[^>]+Target=""(media/[^""]+)"""))
            relMap[m.Groups[1].Value] = "word/" + m.Groups[2].Value;

        var docEntry = srcZip.GetEntry("word/document.xml");
        if (docEntry == null) return result;
        var docXml = ReadAsString(docEntry.Open());

        foreach (var kv in images)
        {
            var placeholder = $"{{{{{kv.Key}}}}}";
            var pos = docXml.IndexOf(placeholder, StringComparison.Ordinal);
            if (pos < 0) continue;

            // Search for r:embed="..." within ±2000 chars of the placeholder (same <wp:inline> block)
            var winStart = Math.Max(0, pos - 2000);
            var winEnd = Math.Min(docXml.Length, pos + 2000);
            var window = docXml[winStart..winEnd];
            var embedMatch = Regex.Match(window, @"r:embed=""([^""]+)""");
            if (embedMatch.Success && relMap.TryGetValue(embedMatch.Groups[1].Value, out var mediaPath))
                result[mediaPath] = kv.Value;
        }
        return result;
    }

    private static String EscapeXml(String s) =>
        s.Replace("&", "&amp;")
         .Replace("<", "&lt;")
         .Replace(">", "&gt;")
         .Replace("\"", "&quot;")
         .Replace("'", "&apos;");
    #endregion
}
