#nullable enable
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>PPT 幻灯片母版信息（S04-01）</summary>
public class PptMasterInfo
{
    #region 属性
    /// <summary>母版索引（0起始）</summary>
    public Int32 Index { get; set; }

    /// <summary>母版文件名（不含扩展名）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>母版背景色（16进制 RGB），null 表示未设置</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>关联版式 ID 列表</summary>
    public List<String> LayoutIds { get; } = [];

    /// <summary>关联主题名称</summary>
    public String ThemeRef { get; set; } = String.Empty;
    #endregion
}

/// <summary>PPT 幻灯片版式信息（S04-02）</summary>
public class PptLayoutInfo
{
    #region 属性
    /// <summary>版式索引（0起始）</summary>
    public Int32 Index { get; set; }

    /// <summary>版式文件名（不含扩展名）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>版式类型（如 blank、title、twoContent 等）</summary>
    public String LayoutType { get; set; } = String.Empty;

    /// <summary>版式显示名称</summary>
    public String DisplayName { get; set; } = String.Empty;
    #endregion
}

/// <summary>PPT 幻灯片文本形状</summary>
public class PptShape{
    #region 属性
    /// <summary>形状ID</summary>
    public Int32 Id { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>形状类型（如 textBox, rect, ellipse, roundRect, triangle, diamond 等）</summary>
    public String ShapeType { get; set; } = String.Empty;

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>填充色（16进制 RGB），null 表示无填充（写入时使用）</summary>
    public String? FillColor { get; set; }

    /// <summary>线条颜色（16进制 RGB），null 表示无线条（写入时使用）</summary>
    public String? LineColor { get; set; }

    /// <summary>线宽（EMU，12700=1pt，写入时使用）</summary>
    public Int32 LineWidth { get; set; } = 12700;

    /// <summary>文字字号（磅，写入时使用）</summary>
    public Int32 FontSize { get; set; } = 14;

    /// <summary>文字颜色（16进制 RGB，写入时使用）</summary>
    public String? FontColor { get; set; }

    /// <summary>文字粗体（写入时使用）</summary>
    public Boolean Bold { get; set; }
    #endregion
}

/// <summary>PPT 幻灯片摘要</summary>
public class PptSlideSummary
{
    #region 属性
    /// <summary>幻灯片索引（0起始）</summary>
    public Int32 Index { get; set; }

    /// <summary>幻灯片文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>形状集合</summary>
    public List<PptShape> Shapes { get; } = [];
    #endregion
}

/// <summary>PowerPoint pptx 读取器</summary>
/// <remarks>
/// 直接解析 Open XML（ZIP+XML）提取幻灯片文本、形状等内容。
/// </remarks>
public class PptxReader : IDisposable
{
    #region 属性
    /// <summary>源文件路径</summary>
    public String? FilePath { get; private set; }
    #endregion

    #region 私有字段
    private readonly ZipArchive _zip;
    private Boolean _disposed;
    #endregion

    #region 构造
    /// <summary>从文件路径打开</summary>
    /// <param name="path">pptx 文件路径</param>
    public PptxReader(String path)
    {
        FilePath = path.GetFullPath();
        _zip = ZipFile.OpenRead(FilePath);
    }

    /// <summary>从流打开</summary>
    /// <param name="stream">包含 pptx 内容的流</param>
    public PptxReader(Stream stream)
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
    /// <summary>获取幻灯片总数</summary>
    /// <returns>幻灯片数量</returns>
    public Int32 GetSlideCount()
    {
        var count = 0;
        foreach (var entry in _zip.Entries)
        {
            if (IsSlideEntry(entry.FullName))
                count++;
        }
        return count;
    }

    /// <summary>获取指定幻灯片的文本内容</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <returns>文本内容</returns>
    public String GetSlideText(Int32 slideIndex)
    {
        var entry = _zip.GetEntry($"ppt/slides/slide{slideIndex + 1}.xml");
        if (entry == null) return String.Empty;
        return ExtractTextFromXml(entry);
    }

    /// <summary>读取全部幻灯片文本（每页用分页符分隔）</summary>
    /// <returns>完整文本</returns>
    public String ReadAllText()
    {
        var count = GetSlideCount();
        if (count == 0) return String.Empty;
        var sb = new StringBuilder();
        for (var i = 0; i < count; i++)
        {
            if (i > 0) sb.AppendLine("--- 幻灯片分隔 ---");
            sb.AppendLine(GetSlideText(i));
        }
        return sb.ToString();
    }

    /// <summary>读取所有幻灯片摘要</summary>
    /// <returns>幻灯片摘要序列</returns>
    public IEnumerable<PptSlideSummary> ReadSlides()
    {
        var count = GetSlideCount();
        for (var i = 0; i < count; i++)
        {
            var entry = _zip.GetEntry($"ppt/slides/slide{i + 1}.xml");
            if (entry == null) continue;

            var summary = new PptSlideSummary { Index = i };
            var doc = LoadXml(entry);
            const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("a", A);

            var textSb = new StringBuilder();
            foreach (XmlElement para in doc.SelectNodes("//a:p", ns)!)
            {
                var lineSb = new StringBuilder();
                foreach (XmlElement t in para.SelectNodes(".//a:t", ns)!)
                    lineSb.Append(t.InnerText);
                var line = lineSb.ToString();
                if (line.Length > 0)
                    textSb.AppendLine(line);
            }
            summary.Text = textSb.ToString().TrimEnd();

            // shapes
            foreach (XmlElement sp in doc.SelectNodes("//*[local-name()='sp']")!)
            {
                var id = sp.SelectSingleNode(".//*[local-name()='cNvPr']")?.Attributes?["id"]?.Value ?? "0";
                var spTypAttr = sp.SelectSingleNode(".//*[local-name()='prstGeom']")?.Attributes?["prst"]?.Value ?? "textBox";
                var shapeTextSb = new StringBuilder();
                foreach (XmlElement t in sp.SelectNodes(".//*[local-name()='t']")!)
                    shapeTextSb.Append(t.InnerText);

                var xfrm = sp.SelectSingleNode(".//*[local-name()='xfrm']");
                var off = xfrm?.SelectSingleNode(".//*[local-name()='off']");
                var ext = xfrm?.SelectSingleNode(".//*[local-name()='ext']");
                summary.Shapes.Add(new PptShape
                {
                    Id = Int32.TryParse(id, out var idNum) ? idNum : 0,
                    ShapeType = spTypAttr,
                    Text = shapeTextSb.ToString(),
                    Left = Int64.TryParse(off?.Attributes?["x"]?.Value, out var x) ? x : 0,
                    Top = Int64.TryParse(off?.Attributes?["y"]?.Value, out var y) ? y : 0,
                    Width = Int64.TryParse(ext?.Attributes?["cx"]?.Value, out var cx) ? cx : 0,
                    Height = Int64.TryParse(ext?.Attributes?["cy"]?.Value, out var cy) ? cy : 0,
                });
            }

            yield return summary;
        }
    }

    /// <summary>提取所有图片</summary>
    /// <returns>（扩展名, 字节数据）序列</returns>
    public IEnumerable<(String Extension, Byte[] Data)> ExtractImages()
    {
        foreach (var entry in _zip.Entries)
        {
            if (!entry.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
                continue;
            var ext = Path.GetExtension(entry.Name).TrimStart('.').ToLowerInvariant();
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            yield return (ext, ms.ToArray());
        }
    }

    /// <summary>读取幻灯片母版信息（S04-01）</summary>
    /// <remarks>
    /// 解析 ppt/slideMasters/*.xml，返回每个母版的背景色及关联版式列表索引。
    /// 对生成工具创建的 pptx 文件，通常只有一个母版（slideMaster1.xml）。
    /// </remarks>
    /// <returns>母版信息列表</returns>
    public IEnumerable<PptMasterInfo> ReadSlideMasters()
    {
        ThrowIfDisposed();
        var masters = _zip.Entries
            .Where(e => e.FullName.StartsWith("ppt/slideMasters/", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && !e.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(e => e.FullName)
            .ToList();

        var idx = 0;
        foreach (var entry in masters)
        {
            var doc = LoadXml(entry);
            var mi = new PptMasterInfo { Index = idx++, Name = Path.GetFileNameWithoutExtension(entry.Name) };

            // 背景色
            var bgNode = doc.SelectSingleNode("//*[local-name()='bg']//*[local-name()='srgbClr']") as XmlElement;
            mi.BackgroundColor = bgNode?.GetAttribute("val");

            // 版式列表（sldLayoutId）
            var layoutIds = doc.SelectNodes("//*[local-name()='sldLayoutId']");
            if (layoutIds != null)
            {
                foreach (XmlElement lid in layoutIds)
                    mi.LayoutIds.Add(lid.GetAttribute("id") ?? String.Empty);
            }

            // 主题引用
            mi.ThemeRef = (doc.SelectSingleNode("//*[local-name()='theme']") as XmlElement)
                ?.GetAttribute("name") ?? String.Empty;

            yield return mi;
        }
    }

    /// <summary>读取幻灯片版式列表（S04-02）</summary>
    /// <remarks>
    /// 解析 ppt/slideLayouts/*.xml，返回版式名称及类型。
    /// </remarks>
    /// <returns>版式信息列表</returns>
    public IEnumerable<PptLayoutInfo> ReadSlideLayouts()
    {
        ThrowIfDisposed();
        var layouts = _zip.Entries
            .Where(e => e.FullName.StartsWith("ppt/slideLayouts/", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && !e.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(e => e.FullName)
            .ToList();

        var idx = 0;
        foreach (var entry in layouts)
        {
            var doc = LoadXml(entry);
            var root = doc.DocumentElement;
            var li = new PptLayoutInfo
            {
                Index = idx++,
                Name = Path.GetFileNameWithoutExtension(entry.Name),
                LayoutType = root?.GetAttribute("type") ?? String.Empty,
                DisplayName = root?.GetAttribute("showMasterSp") == "0" ? String.Empty
                    : (doc.SelectSingleNode("//*[local-name()='cSld']") as XmlElement)?.GetAttribute("name") ?? String.Empty,
            };
            yield return li;
        }
    }
    #endregion

    #region 私有方法
    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(PptxReader));
    }

    private static Boolean IsSlideEntry(String name) =>
        name.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
        && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        && !name.Contains("_rels", StringComparison.OrdinalIgnoreCase);

    private static String ExtractTextFromXml(ZipArchiveEntry entry)
    {
        var doc = LoadXml(entry);
        var sb = new StringBuilder();
        foreach (XmlElement t in doc.SelectNodes("//*[local-name()='t']")!)
        {
            var text = t.InnerText;
            if (text.Length > 0) sb.AppendLine(text);
        }
        return sb.ToString().TrimEnd();
    }

    private static XmlDocument LoadXml(ZipArchiveEntry entry)
    {
        var doc = new XmlDocument();
        using var s = entry.Open();
        doc.Load(s);
        return doc;
    }
    #endregion
}
