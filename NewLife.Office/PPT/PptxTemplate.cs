using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>PPT 模板填充器</summary>
/// <remarks>
/// 以现有 pptx 为模板，将幻灯片中的 {{Key}} 占位符替换为实际值后输出新文件。
/// 支持表格行扩展（{{#ListKey}} / {{/ListKey}}）和图片占位符替换。
/// </remarks>
public class PptxTemplate
{
    #region 属性
    /// <summary>模板文件路径</summary>
    public String TemplatePath { get; }
    #endregion

    #region 构造
    /// <summary>实例化模板填充器</summary>
    /// <param name="templatePath">模板 pptx 路径</param>
    public PptxTemplate(String templatePath) => TemplatePath = templatePath.GetFullPath();
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
        var templateBytes = File.ReadAllBytes(TemplatePath);
        using var srcMs = new MemoryStream(templateBytes);
        using var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read);
        using var dstZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in srcZip.Entries)
        {
            var dstEntry = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
            using var srcStream = entry.Open();
            using var dstStream = dstEntry.Open();

            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                using var ms = new MemoryStream();
                srcStream.CopyTo(ms);
                var content = Encoding.UTF8.GetString(ms.ToArray());
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

    /// <summary>填充模板，支持表格行扩展（S03-03）</summary>
    /// <remarks>在 pptx 模板表格的 a:tr 行中放置 {{#ListKey}} / {{/ListKey}} 标记，中间行作为模板行按数据展开。</remarks>
    /// <param name="outputPath">输出路径</param>
    /// <param name="data">普通占位符字典</param>
    /// <param name="lists">列表数据，Key 为占位符名称，Value 为行数据集合</param>
    public void FillTable(String outputPath, IDictionary<String, Object?> data,
        IDictionary<String, IEnumerable<IDictionary<String, Object?>>> lists)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        FillTable(fs, data, lists);
    }

    /// <summary>填充模板，支持表格行扩展，写入流</summary>
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

            // Apply table expansion only on slide XML files
            if (entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                using var ms = new MemoryStream();
                srcStream.CopyTo(ms);
                var content = Encoding.UTF8.GetString(ms.ToArray());
                foreach (var kv in lists)
                    content = WordTemplate.ExpandTableRows(content, kv.Key, kv.Value, "a:tr");
                content = ApplyReplacements(content, data);
                var bytes = Encoding.UTF8.GetBytes(content);
                dstStream.Write(bytes, 0, bytes.Length);
            }
            else if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                using var ms = new MemoryStream();
                srcStream.CopyTo(ms);
                var content = Encoding.UTF8.GetString(ms.ToArray());
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

    /// <summary>填充模板，支持图片占位符替换（S03-04）</summary>
    /// <remarks>在 pptx 模板图片上，将 p:cNvPr 的 descr 或 name 属性设为 {{ImageKey}}，此方法将以新图片字节替换该图片。</remarks>
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

        // Build map: "ppt/media/imageN.xxx" -> new bytes
        var imgReplace = BuildPptxImageMap(srcZip, images);

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
                using var ms = new MemoryStream();
                srcStream.CopyTo(ms);
                var content = Encoding.UTF8.GetString(ms.ToArray());
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
    private static String ApplyReplacements(String xml, IDictionary<String, Object?> data)
    {
        foreach (var kv in data)
        {
            var placeholder = $"{{{{{kv.Key}}}}}";
            var value = Convert.ToString(kv.Value) ?? String.Empty;
            xml = xml.Replace(placeholder, EscapeXml(value));
        }
        return xml;
    }

    /// <summary>构建 pptx 中占位图片名称 → ZIP 条目路径 → 替换字节的映射</summary>
    /// <param name="srcZip">源 ZIP</param>
    /// <param name="images">占位符名称（不含{{}}）→ 图片字节</param>
    /// <returns>ZIP 条目路径 → 新图片字节的映射</returns>
    private static Dictionary<String, Byte[]> BuildPptxImageMap(ZipArchive srcZip, IDictionary<String, Byte[]> images)
    {
        var result = new Dictionary<String, Byte[]>(StringComparer.OrdinalIgnoreCase);
        if (images.Count == 0) return result;

        // For each slide, find images with placeholder alt text and map to media files
        foreach (var slideEntry in srcZip.Entries.Where(e =>
            e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
            && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            String slideXml;
            using (var ms = new MemoryStream()) { slideEntry.Open().CopyTo(ms); slideXml = Encoding.UTF8.GetString(ms.ToArray()); }

            // Find corresponding rels file to resolve media paths
            var slideNum = slideEntry.FullName; // e.g. ppt/slides/slide1.xml
            var relsPath = slideNum.Replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
            var relsEntry = srcZip.GetEntry(relsPath);
            if (relsEntry == null) continue;

            String relsXml;
            using (var ms = new MemoryStream()) { relsEntry.Open().CopyTo(ms); relsXml = Encoding.UTF8.GetString(ms.ToArray()); }

            var relMap = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase); // rId -> "ppt/media/..."
            foreach (Match m in Regex.Matches(relsXml, @"Id=""([^""]+)""[^>]+Target=""(\.\./media/[^""]+)"""))
                relMap[m.Groups[1].Value] = "ppt/media/" + m.Groups[2].Value["../media/".Length..];

            foreach (var kv in images)
            {
                var placeholder = $"{{{{{kv.Key}}}}}";
                var pos = slideXml.IndexOf(placeholder, StringComparison.Ordinal);
                if (pos < 0) continue;

                // Search for r:embed="..." within ±2000 chars of the placeholder
                var winStart = Math.Max(0, pos - 2000);
                var winEnd = Math.Min(slideXml.Length, pos + 2000);
                var window = slideXml[winStart..winEnd];
                var embedMatch = Regex.Match(window, @"r:embed=""([^""]+)""");
                if (embedMatch.Success && relMap.TryGetValue(embedMatch.Groups[1].Value, out var mediaPath))
                    result[mediaPath] = kv.Value;
            }
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
