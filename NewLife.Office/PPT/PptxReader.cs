using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>PowerPoint pptx 读取器</summary>
/// <remarks>
/// 直接解析 Open XML（ZIP+XML）提取幻灯片文本、形状等内容。
/// </remarks>
public class PptxReader : IDisposable, ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>源文件路径</summary>
    public String? FilePath { get; private set; }

    /// <summary>幻灯片宽度（EMU），从 presentation.xml 的 sldSz 解析</summary>
    public Int64 SlideWidth => _slideWidth ??= ParsePresentationInfo().width;

    /// <summary>幻灯片高度（EMU），从 presentation.xml 的 sldSz 解析</summary>
    public Int64 SlideHeight => _slideHeight ??= ParsePresentationInfo().height;

    /// <summary>主题强调色（6个），从 theme1.xml 的 clrScheme 解析</summary>
    public String[] AccentColors => _accentColors ??= ParseThemeColors();
    #endregion

    #region 私有字段
    private readonly ZipArchive _zip;
    private Boolean _disposed;

    // 延迟加载字段
    private Int64? _slideWidth;
    private Int64? _slideHeight;
    private String[]? _accentColors;
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
                {
                    lineSb.Append(t.InnerText);
                }
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
                {
                    shapeTextSb.Append(t.InnerText);
                }

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
                {
                    mi.LayoutIds.Add(lid.GetAttribute("id") ?? String.Empty);
                }
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
    /// <summary>读取指定幻灯片关联的图表数据（S06-04）</summary>
    /// <remarks>
    /// 通过幻灯片关系文件定位图表 XML，解析 c:ser 中的分类和数值缓存。
    /// 仅读取 numCache/strCache 中的缓存数据，不依赖内嵌 Excel。
    /// </remarks>
    /// <param name="slideIndex">幻灯片索引（0 起始）</param>
    /// <returns>该页所有图表的数据集合</returns>
    public IEnumerable<PptChartInfo> ReadChartData(Int32 slideIndex)
    {
        ThrowIfDisposed();
        var relsEntry = _zip.GetEntry($"ppt/slides/_rels/slide{slideIndex + 1}.xml.rels");
        if (relsEntry == null) yield break;
        var relsDoc = LoadXml(relsEntry);
        const String PKGNS = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(relsDoc.NameTable);
        ns.AddNamespace("r", PKGNS);
        var chartNum = 0;
        foreach (XmlElement rel in relsDoc.SelectNodes("//r:Relationship", ns)!)
        {
            var target = rel.GetAttribute("Target");
            var type = rel.GetAttribute("Type");
            if (!type.Contains("chart", StringComparison.OrdinalIgnoreCase)) continue;
            // target 形如 ../charts/chart1.xml
            var chartPath = "ppt/" + target.TrimStart('.').TrimStart('/');
            var chartEntry = _zip.GetEntry(chartPath);
            if (chartEntry == null) continue;
            var chartDoc = LoadXml(chartEntry);
            const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var cns = new XmlNamespaceManager(chartDoc.NameTable);
            cns.AddNamespace("c", C);
            var info = new PptChartInfo { ChartNumber = ++chartNum };
            // 图表类型
            var chartTypeNode = chartDoc.SelectSingleNode("//*[substring(local-name(), string-length(local-name())-4) = 'Chart'][@*]", null);
            info.ChartType = chartTypeNode?.LocalName?.Replace("Chart", String.Empty) ?? "bar";
            // 第一个系列的分类
            var firstCatNode = chartDoc.SelectSingleNode("//c:ser[1]/c:cat//c:strCache", cns)
                            ?? chartDoc.SelectSingleNode("//c:ser[1]/c:cat//c:numCache", cns);
            if (firstCatNode != null)
            {
                var cats = new List<String>();
                foreach (XmlElement pt in firstCatNode.SelectNodes("c:pt/c:v", cns)!)
                {
                    cats.Add(pt.InnerText);
                }
                info.Categories = cats.ToArray();
            }
            // 所有系列
            foreach (XmlElement ser in chartDoc.SelectNodes("//c:ser", cns)!)
            {
                var serName = ser.SelectSingleNode(".//c:tx//c:v", cns)?.InnerText ?? String.Empty;
                var vals = new List<Double>();
                foreach (XmlElement v in ser.SelectNodes(".//c:val//c:numCache/c:pt/c:v", cns)!)
                {
                    if (Double.TryParse(v.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                        vals.Add(d);
                }
                info.Series.Add(new PptChartSeriesData { Name = serName, Values = vals.ToArray() });
            }
            yield return info;
        }
    }

    /// <summary>获取指定母版的原始 XML 内容（S04-Master）</summary>
    /// <param name="index">母版索引（0起始）</param>
    /// <returns>母版 XML 字符串，不存在则返回 null</returns>
    public String? GetSlideMasterXml(Int32 index)
    {
        ThrowIfDisposed();
        var entry = _zip.GetEntry($"ppt/slideMasters/slideMaster{index + 1}.xml");
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>获取指定幻灯片的原始 XML 内容</summary>
    /// <param name="index">幻灯片索引（0起始）</param>
    /// <returns>幻灯片 XML 字符串，不存在则返回 null</returns>
    public String? GetSlideXml(Int32 index)
    {
        ThrowIfDisposed();
        var entry = _zip.GetEntry($"ppt/slides/slide{index + 1}.xml");
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>获取指定版式的原始 XML 内容（S04-Master）</summary>
    /// <param name="index">版式索引（0起始）</param>
    /// <returns>版式 XML 字符串，不存在则返回 null</returns>
    public String? GetSlideLayoutXml(Int32 index)
    {
        ThrowIfDisposed();
        var entry = _zip.GetEntry($"ppt/slideLayouts/slideLayout{index + 1}.xml");
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>获取主题原始 XML 内容（S04-Master）</summary>
    /// <returns>主题 XML 字符串，不存在则返回 null</returns>
    public String? GetThemeXml()
    {
        ThrowIfDisposed();
        var entry = _zip.GetEntry("ppt/theme/theme1.xml")
            ?? _zip.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase)
                && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>完整读取指定幻灯片为内存对象（含文本框/形状/图片/表格/图表/组/备注/切换动画/超链接/视频）</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <returns>完整的幻灯片对象</returns>
    public PptSlide ReadSlide(Int32 slideIndex)
    {
        ThrowIfDisposed();
        var entry = _zip.GetEntry($"ppt/slides/slide{slideIndex + 1}.xml");
        if (entry == null) throw new ArgumentOutOfRangeException(nameof(slideIndex), $"幻灯片 {slideIndex} 不存在");

        var slide = new PptSlide();
        var doc = LoadXml(entry);
        var rels = ReadSlideRels(slideIndex);

        // 解析幻灯片隐藏属性（S12-04）
        var show = doc.DocumentElement?.GetAttribute("show");
        if (show == "0") slide.Hidden = true;

        // 解析版式索引：从 slide rels 中查找 slideLayout 关系的 Target（如 ../slideLayouts/slideLayout{N}.xml），提取 N-1
        ParseLayoutIndex(slide, rels);

        // 读取版式的 lstStyle 字体默认值（占位符文本继承自版式，幻灯片自身 lstStyle 通常为空）
        var layoutDefaults = ParseLayoutLstStyleDefaults(slide.LayoutIndex);

        ParseSlideBackground(doc, slide, slideIndex, rels);

        var spTree = doc.SelectSingleNode("//*[local-name()='spTree']");
        if (spTree != null)
        {
            foreach (XmlNode child in spTree.ChildNodes)
            {
                if (child is not XmlElement el) continue;
                switch (el.LocalName)
                {
                    case "sp": ParseShapeOrTextBox(el, slide, rels, layoutDefaults); break;
                    case "pic": ParsePicture(el, slide, rels); break;
                    case "graphicFrame": ParseGraphicFrame(el, slide, rels); break;
                    case "grpSp": ParseGroup(el, slide, rels, layoutDefaults); break;
                    case "cxnSp": ParseConnector(el, slide); break;
                }
            }
        }

        ParseNotes(doc, slide);
        ParseTransition(doc, slide);
        ParseSlideComments(slideIndex, slide);
        ParseSlideAnimations(doc, slide);

        return slide;
    }

    /// <summary>完整读取所有幻灯片为内存对象</summary>
    /// <returns>幻灯片对象序列</returns>
    public IEnumerable<PptSlide> ReadAllSlides()
    {
        var count = GetSlideCount();
        for (var i = 0; i < count; i++)
            yield return ReadSlide(i);
    }

    /// <summary>读取整个演示文稿为 PptDocument 数据模型</summary>
    /// <returns>包含全部幻灯片和文档属性的 PptDocument</returns>
    public PptDocument ReadDocument()
    {
        var doc = new PptDocument
        {
            SlideWidth  = SlideWidth,
            SlideHeight = SlideHeight,
            AccentColors = [.. AccentColors],
        };
        var count = GetSlideCount();
        for (var i = 0; i < count; i++)
            doc.Slides.Add(ReadSlide(i));
        ParseDocProps(doc);
        ParseHeaderFooter(doc);
        doc.Sections = ParseSections();
        return doc;
    }
    #endregion

    #region 私有方法
    /// <summary>从幻灯片关系文件解析所用版式索引</summary>
    /// <param name="slide">目标幻灯片</param>
    /// <param name="slideRels">幻灯片关系映射</param>
    private static void ParseLayoutIndex(PptSlide slide, Dictionary<String, (String Target, String Type)> slideRels)
    {
        foreach (var kv in slideRels)
        {
            var rel = kv.Value;
            if (!rel.Type.Contains("slideLayout", StringComparison.OrdinalIgnoreCase)) continue;
            // Target 形如 ../slideLayouts/slideLayout{N}.xml
            var name = Path.GetFileNameWithoutExtension(rel.Target);
            if (name.StartsWith("slideLayout") && Int32.TryParse(name.Replace("slideLayout", ""), out var layoutNum))
            {
                slide.LayoutIndex = Math.Max(0, layoutNum - 1);
                return;
            }
        }
    }

    /// <summary>从 presentation.xml 解析幻灯片尺寸（cx/cy）</summary>
    /// <returns>(宽度, 高度) EMU</returns>
    private (Int64 width, Int64 height) ParsePresentationInfo()
    {
        var width = 12192000L; // 默认 16:9
        var height = 6858000L;
        var entry = _zip.GetEntry("ppt/presentation.xml");
        if (entry == null) return (width, height);
        var doc = LoadXml(entry);
        var sldSz = doc.SelectSingleNode("//*[local-name()='sldSz']") as XmlElement;
        if (sldSz != null)
        {
            if (Int64.TryParse(sldSz.GetAttribute("cx"), out var cx)) width = cx;
            if (Int64.TryParse(sldSz.GetAttribute("cy"), out var cy)) height = cy;
        }
        _slideWidth = width;
        _slideHeight = height;
        return (width, height);
    }

    /// <summary>从 presentation.xml 解析节（Section）信息</summary>
    private List<PptSection>? ParseSections()
    {
        var entry = _zip.GetEntry("ppt/presentation.xml");
        if (entry == null) return null;

        var doc = LoadXml(entry);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        var sectionNodes = doc.SelectNodes("//p14:section", ns);
        if (sectionNodes == null || sectionNodes.Count == 0) return null;

        var sections = new List<PptSection>();
        foreach (XmlElement secEl in sectionNodes)
        {
            var name = secEl.GetAttribute("name");
            if (String.IsNullOrEmpty(name)) name = "默认节";
            var indices = new List<Int32>();

            var sldNodes = secEl.SelectNodes("p14:sldIdLst/p14:sldId", ns);
            if (sldNodes != null)
            {
                foreach (XmlElement sldEl in sldNodes)
                {
                    if (Int32.TryParse(sldEl.GetAttribute("id"), out var id) && id >= 256)
                        indices.Add(id - 256);
                }
            }
            sections.Add(new PptSection { Name = name, SlideIndices = indices });
        }
        return sections.Count > 0 ? sections : null;
    }

    /// <summary>从 theme1.xml 的 clrScheme 解析六个强调色</summary>
    /// <returns>6个 accent 颜色（6位十六进制无#）</returns>
    private String[] ParseThemeColors()
    {
        var colors = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        var entry = _zip.GetEntry("ppt/theme/theme1.xml")
            ?? _zip.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase)
                && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        if (entry == null) return colors;
        var doc = LoadXml(entry);
        for (var i = 1; i <= 6; i++)
        {
            var accentNode = doc.SelectSingleNode($"//*[local-name()='accent{i}']/*[local-name()='srgbClr']") as XmlElement;
            if (accentNode != null)
            {
                var val = accentNode.GetAttribute("val");
                if (val.Length > 0) colors[i - 1] = val;
            }
        }
        _accentColors = colors;
        return colors;
    }

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

    /// <summary>解析幻灯片关系文件，返回 Id→(Target, Type) 映射</summary>
    private Dictionary<String, (String Target, String Type)> ReadSlideRels(Int32 slideIndex)
    {
        var map = new Dictionary<String, (String, String)>();
        var relsEntry = _zip.GetEntry($"ppt/slides/_rels/slide{slideIndex + 1}.xml.rels");
        if (relsEntry == null) return map;
        var doc = LoadXml(relsEntry);
        const String PKGNS = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", PKGNS);
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            var type = rel.GetAttribute("Type");
            if (id.Length > 0) map[id] = (target, type);
        }
        return map;
    }

    /// <summary>解析幻灯片背景（纯色或图片），三级回退：slide rels → layout rels → master rels</summary>
    private void ParseSlideBackground(XmlDocument doc, PptSlide slide, Int32 slideIndex,
        Dictionary<String, (String Target, String Type)> slideRels)
    {
        var bg = doc.SelectSingleNode("//*[local-name()='bg']") as XmlElement;
        if (bg == null) return;

        // 纯色背景（仅 bgPr 的直接子 solidFill，排除嵌套在 gradFill 中的 srgbClr）
        var solidFill = bg.SelectSingleNode("*[local-name()='bgPr']/*[local-name()='solidFill']") as XmlElement;
        var bgClr = solidFill?.SelectSingleNode("*[local-name()='srgbClr']") as XmlElement;
        var val = bgClr?.GetAttribute("val");
        if (!String.IsNullOrEmpty(val)) { slide.BackgroundColor = val; return; }

        // 渐变背景（S15-04）
        var gradFill = bg.SelectSingleNode(".//*[local-name()='gradFill']") as XmlElement;
        if (gradFill != null)
        {
            var stops = gradFill.SelectNodes(".//*[local-name()='gs']");
            if (stops != null && stops.Count >= 2)
            {
                var c1 = (stops[0] as XmlElement)?.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                var c2 = (stops[1] as XmlElement)?.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                if (c1 != null && c2 != null)
                {
                    slide.BackgroundGradientColor1 = c1.GetAttribute("val");
                    slide.BackgroundGradientColor2 = c2.GetAttribute("val");
                    slide.BackgroundGradientType = "linear";
                }
            }
            return;
        }

        // 图片背景
        var blip = bg.SelectSingleNode(".//*[local-name()='blipFill']/*[local-name()='blip']") as XmlElement;
        var embedId = blip?.GetAttribute("r:embed");
        if (embedId == null) return;

        // 三级回退查找背景图片关系：slide rels → layout rels → master rels
        if (!TryResolveBgImage(embedId, slideRels, out var rel))
        {
            // 尝试从 slide layout rels 查找
            var layoutRels = TryReadLayoutRelsFromSlide(slideRels);
            if (layoutRels != null && !TryResolveBgImage(embedId, layoutRels, out rel))
            {
                // 尝试从 slide master rels 查找（最后一个回退级别）
                var masterRels = ReadSlideMasterRels(0);
                if (masterRels != null)
                    TryResolveBgImage(embedId, masterRels, out rel);
            }
        }

        if (rel.Target == null) return;
        var mediaName = Path.GetFileName(rel.Target);
        var mediaEntry = _zip.GetEntry($"ppt/media/{mediaName}");
        if (mediaEntry == null) return;
        var ext = Path.GetExtension(mediaName).TrimStart('.').ToLowerInvariant();
        Byte[] data;
        using (var ms = new MemoryStream())
        using (var es = mediaEntry.Open()) { es.CopyTo(ms); data = ms.ToArray(); }
        slide.BackgroundImage = new PptImage { Data = data, Extension = ext, IsSvg = ext == "svg" };
    }

    private static Boolean TryResolveBgImage(String embedId,
        Dictionary<String, (String Target, String Type)> rels,
        out (String Target, String Type) rel)
        => rels.TryGetValue(embedId, out rel);

    /// <summary>从 slide rels 中提取 layout 引用，读取对应 layout 的 rels</summary>
    private Dictionary<String, (String Target, String Type)>? TryReadLayoutRelsFromSlide(
        Dictionary<String, (String Target, String Type)> slideRels)
    {
        foreach (var kv in slideRels)
        {
            if (kv.Value.Type.Contains("slideLayout", StringComparison.OrdinalIgnoreCase))
            {
                var target = kv.Value.Target; // 如 "../slideLayouts/slideLayout2.xml"
                var name = Path.GetFileNameWithoutExtension(target); // "slideLayout2"
                if (name.StartsWith("slideLayout") && Int32.TryParse(name.Replace("slideLayout", ""), out var layoutNum))
                    return ReadLayoutRels(layoutNum - 1);
            }
        }
        return null;
    }

    /// <summary>读取幻灯片版式关系文件</summary>
    private Dictionary<String, (String Target, String Type)> ReadLayoutRels(Int32 layoutIndex)
    {
        var map = new Dictionary<String, (String, String)>();
        var relsEntry = _zip.GetEntry($"ppt/slideLayouts/_rels/slideLayout{layoutIndex + 1}.xml.rels");
        if (relsEntry == null) return map;
        var doc = LoadXml(relsEntry);
        const String PKGNS = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", PKGNS);
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            var type = rel.GetAttribute("Type");
            if (id.Length > 0) map[id] = (target, type);
        }
        return map;
    }

    /// <summary>读取幻灯片母版关系文件</summary>
    private Dictionary<String, (String Target, String Type)>? ReadSlideMasterRels(Int32 masterIndex)
    {
        var map = new Dictionary<String, (String, String)>();
        var relsEntry = _zip.GetEntry($"ppt/slideMasters/_rels/slideMaster{masterIndex + 1}.xml.rels");
        if (relsEntry == null) return null;
        var doc = LoadXml(relsEntry);
        const String PKGNS = "http://schemas.openxmlformats.org/package/2006/relationships";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", PKGNS);
        foreach (XmlElement rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            var type = rel.GetAttribute("Type");
            if (id.Length > 0) map[id] = (target, type);
        }
        return map;
    }

    /// <summary>解析形状或文本框</summary>
    private void ParseShapeOrTextBox(XmlElement sp, PptSlide slide, Dictionary<String, (String Target, String Type)> rels,
        Dictionary<String, DefaultRunProps>? layoutDefaults = null)
    {
        var hasTxBody = sp.SelectSingleNode(".//*[local-name()='txBody']") != null;
        var (left, top, width, height) = ParseXfrm(sp.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        var rotation = ParseRotation(sp.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        var shapeType = sp.SelectSingleNode(".//*[local-name()='prstGeom']")?.Attributes?["prst"]?.Value ?? "textBox";
        // Alt Text（cNvPr descr）
        var cNvPr = sp.SelectSingleNode(".//*[local-name()='cNvPr']") as XmlElement;
        var altText = cNvPr?.GetAttribute("descr");
        // Corner radius（仅 roundRect）
        Int64 cornerRadius = 0;
        if (shapeType == "roundRect")
        {
            var gd = sp.SelectSingleNode(".//*[local-name()='prstGeom']/*[local-name()='avLst']/*[local-name()='gd']") as XmlElement;
            if (gd != null)
            {
                var fmla = gd.GetAttribute("fmla") ?? "";
                var valStr = fmla.StartsWith("val ") ? fmla[4..] : fmla;
                if (Int64.TryParse(valStr, out var adjVal) && adjVal > 0)
                    cornerRadius = adjVal * width / 50000; // reverse the adj value back to EMU
            }
        }
        var spPr = sp.SelectSingleNode(".//*[local-name()='spPr']") as XmlElement;
        var fillColor = ParseFillColor(spPr);
        var (lineColor, lineWidth) = ParseLine(spPr);

        if (hasTxBody)
        {
            var tb = new PptTextBox { Left = left, Top = top, Width = width, Height = height, Rotation = rotation, AltText = altText };
            if (fillColor != null) tb.BackgroundColor = fillColor;

            var txBody = sp.SelectSingleNode(".//*[local-name()='txBody']") as XmlElement;
            if (txBody == null) { slide.TextBoxes.Add(tb); return; }

            // bodyPr：auto-fit 模式 + 内边距 + anchor
            var bodyPr = txBody.SelectSingleNode("*[local-name()='bodyPr']") as XmlElement;
            if (bodyPr != null)
            {
                if (bodyPr.SelectSingleNode("*[local-name()='spAutoFit']") != null) tb.AutoFit = 1;
                else if (bodyPr.SelectSingleNode("*[local-name()='noAutofit']") != null) tb.AutoFit = 2;
                else tb.AutoFit = 0; // normAutofit
                // 内边距（EMU）
                var lIns = bodyPr.GetAttribute("lIns"); if (lIns.Length > 0 && Int32.TryParse(lIns, out var li)) tb.LeftInset = li;
                var rIns = bodyPr.GetAttribute("rIns"); if (rIns.Length > 0 && Int32.TryParse(rIns, out var ri)) tb.RightInset = ri;
                var tIns = bodyPr.GetAttribute("tIns"); if (tIns.Length > 0 && Int32.TryParse(tIns, out var ti)) tb.TopInset = ti;
                var bIns = bodyPr.GetAttribute("bIns"); if (bIns.Length > 0 && Int32.TryParse(bIns, out var bi)) tb.BottomInset = bi;
                // 垂直锁定方式
                var anchor = bodyPr.GetAttribute("anchor");
                if (anchor.Length > 0) tb.Anchor = anchor;
                // 文本方向（vert 属性）
                var vert = bodyPr.GetAttribute("vert");
                if (vert.Length > 0) tb.TextDirection = vert;
            }

            var lstDefaults = ParseListStyleDefaults(txBody);
            // 版式回退：幻灯片自身 lstStyle 为空时（占位符常见），使用版式定义的字体默认值
            if (lstDefaults.Count == 0 && layoutDefaults != null && layoutDefaults.Count > 0)
                lstDefaults = layoutDefaults;

            var allText = new StringBuilder();
            var hasAnyRuns = false;
            var paraNodes = txBody.SelectNodes(".//*[local-name()='p']");
            if (paraNodes != null)
            {
                foreach (XmlElement para in paraNodes)
                {
                    var pp = new PptParagraph();
                    var pPr = para.SelectSingleNode(".//*[local-name()='pPr']") as XmlElement;
                    var lvl = pPr?.GetAttribute("lvl") ?? "0";
                    if (Int32.TryParse(lvl, out var lv)) pp.Level = lv;
                    var pDefRPr = pPr?.SelectSingleNode(".//*[local-name()='defRPr']") as XmlElement;

                    // 段落级属性
                    if (pPr != null)
                    {
                        var algn = pPr.GetAttribute("algn");
                        if (algn.Length > 0) pp.Alignment = algn;
                        // 行间距（spcPct 优先，其次 spcPts）
                        var lnSpcPct = pPr.SelectSingleNode("*[local-name()='lnSpc']/*[local-name()='spcPct']") as XmlElement;
                        if (lnSpcPct != null && Int32.TryParse(lnSpcPct.GetAttribute("val"), out var lsp)) pp.LineSpacingPct = lsp;
                        var lnSpcPts = pPr.SelectSingleNode("*[local-name()='lnSpc']/*[local-name()='spcPts']") as XmlElement;
                        if (lnSpcPts != null && Int32.TryParse(lnSpcPts.GetAttribute("val"), out var lpt)) pp.LineSpacingPts = lpt;
                        // 段前距
                        var spcBef = pPr.SelectSingleNode("*[local-name()='spcBef']/*[local-name()='spcPts']") as XmlElement;
                        if (spcBef != null && Int32.TryParse(spcBef.GetAttribute("val"), out var sbv)) pp.SpaceBeforePt = sbv / 100;
                        // 段后距
                        var spcAft = pPr.SelectSingleNode("*[local-name()='spcAft']/*[local-name()='spcPts']") as XmlElement;
                        if (spcAft != null && Int32.TryParse(spcAft.GetAttribute("val"), out var sav)) pp.SpaceAfterPt = sav / 100;
                        // 项目符号
                        var buChar = pPr.SelectSingleNode("*[local-name()='buChar']") as XmlElement;
                        if (buChar != null) pp.BulletChar = buChar.GetAttribute("char");
                        var buNone = pPr.SelectSingleNode("*[local-name()='buNone']") as XmlElement;
                        if (buNone != null) pp.BulletNone = true;
                    }

                    var runs = para.SelectNodes(".//*[local-name()='r']");
                    if (runs == null || runs.Count == 0)
                    {
                        var pt = new StringBuilder();
                        foreach (XmlElement t in para.SelectNodes(".//*[local-name()='t']")!)
                            pt.Append(t.InnerText);
                        var text = pt.ToString();
                        if (text.Length > 0)
                        {
                            var dr = CreateDefaultRun(text, lstDefaults, lvl, pDefRPr);
                            pp.Runs.Add(dr);
                            tb.Runs.Add(dr);  // 向后兼容
                            allText.Append(text);
                            hasAnyRuns = true;
                        }
                    }
                    else
                    {
                        foreach (XmlElement r in runs)
                        {
                            var run = ParseTextRun(r, rels, lstDefaults, lvl, pDefRPr);
                            if (run != null)
                            {
                                pp.Runs.Add(run);
                                tb.Runs.Add(run);  // 向后兼容
                                allText.Append(run.Text);
                                hasAnyRuns = true;
                            }
                        }
                    }

                    // 始终添加段落（包括空段，保留段落结构）
                    tb.Paragraphs.Add(pp);
                }
            }

            // 无段落时的回退逻辑
            if (tb.Paragraphs.Count == 0)
            {
                var ft = new StringBuilder();
                foreach (XmlElement t in txBody.SelectNodes(".//*[local-name()='t']")!)
                    ft.Append(t.InnerText);
                var text = ft.ToString();
                if (text.Length > 0)
                {
                    var pp = new PptParagraph();
                    var dr = new PptTextRun { Text = text, FontSize = 18 };
                    if (lstDefaults.TryGetValue("0", out var ld) || lstDefaults.TryGetValue("", out ld))
                    { if (ld.FontSize > 0) dr.FontSize = ld.FontSize; dr.LatinFontName = ld.LatinFontName; dr.EastAsianFontName = ld.EastAsianFontName; }
                    pp.Runs.Add(dr);
                    tb.Runs.Add(dr);
                    tb.Paragraphs.Add(pp);
                    allText.Append(text);
                    hasAnyRuns = true;
                }
            }

            // 向后兼容：从首段首 Run 提取 TextBox 级字体属性
            if (hasAnyRuns)
            {
                var fr = tb.Runs[0];
                tb.FontSize = fr.FontSize > 0 ? fr.FontSize : 18;
                tb.Bold = fr.Bold;
                tb.FontColor = fr.FontColor;
                tb.LatinFontName = fr.LatinFontName;
                tb.EastAsianFontName = fr.EastAsianFontName;
                tb.ComplexScriptFontName = fr.ComplexScriptFontName;
                tb.SymbolFontName = fr.SymbolFontName;
            }
            // 向后兼容：从首段提取 TextBox 级段落属性
            if (tb.Paragraphs.Count > 0)
            {
                var fp = tb.Paragraphs[0];
                tb.Alignment = fp.Alignment;
                tb.LineSpacingPct = fp.LineSpacingPct;
                tb.SpaceBeforePt = fp.SpaceBeforePt;
            }
            tb.Text = allText.ToString();
            tb.AltText = altText;
            slide.TextBoxes.Add(tb);
        }
        else
        {
            slide.Shapes.Add(new PptShape
            {
                ShapeType = shapeType, Left = left, Top = top, Width = width, Height = height,
                FillColor = fillColor, LineColor = lineColor, LineWidth = (Int32)lineWidth,
                AltText = altText, CornerRadius = cornerRadius,
            });
        }
    }

    /// <summary>解析单个文本片段（a:r），可接收段落级默认值</summary>
    private static PptTextRun? ParseTextRun(XmlElement r, Dictionary<String, (String Target, String Type)> rels,
        Dictionary<String, DefaultRunProps>? lstDefaults = null, String lvl = "0", XmlElement? pDefRPr = null)
    {
        var t = r.SelectSingleNode(".//*[local-name()='t']") as XmlElement;
        var text = t?.InnerText;
        if (text == null) return null;
        var run = new PptTextRun { Text = text };
        DefaultRunProps? defaults = null;
        if (lstDefaults != null && lstDefaults.TryGetValue(lvl, out var ld)) defaults = ld;
        var rPr = r.SelectSingleNode(".//*[local-name()='rPr']") as XmlElement;
        var boldExplicit = false;
        if (rPr != null)
        {
            var sz = rPr.GetAttribute("sz");
            if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) run.FontSize = sv / 100;
            var bAttr = rPr.GetAttribute("b");
            if (bAttr.Length > 0) { run.Bold = bAttr == "1"; boldExplicit = true; }
            var iAttr = rPr.GetAttribute("i");
            if (iAttr.Length > 0) run.Italic = iAttr == "1";
            // 下划线
            var uAttr = rPr.GetAttribute("u");
            if (uAttr.Length > 0 && uAttr != "none") run.Underline = true;
            // solidFill 颜色（srgbClr 或 schemeClr）
            var sf = rPr.SelectSingleNode("*[local-name()='solidFill']") as XmlElement;
            var fc = sf?.SelectSingleNode("*[local-name()='srgbClr']") as XmlElement;
            if (fc != null) run.FontColor = fc.GetAttribute("val");
            else
            {
                var sc = sf?.SelectSingleNode("*[local-name()='schemeClr']") as XmlElement;
                if (sc != null) run.FontColor = "scheme:" + sc.GetAttribute("val");
            }
            // 渐变色（取各停靠点的 srgbClr，存为 hex 数组）
            if (run.FontColor == null)
            {
                var gf = rPr.SelectSingleNode("*[local-name()='gradFill']") as XmlElement;
                if (gf != null)
                {
                    var stops = new List<String>();
                    var ang = gf.SelectSingleNode(".//*[local-name()='lin']") as XmlElement;
                    var angVal = ang?.GetAttribute("ang") ?? "0";
                    Int32.TryParse(angVal, out var angle);
                    run.GradAngle = angle;
                    foreach (XmlElement gs in gf.SelectNodes(".//*[local-name()='gs']")!)
                    {
                        var gsClr = gs.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                        if (gsClr != null) stops.Add(gsClr.GetAttribute("val"));
                    }
                    if (stops.Count >= 2)
                    {
                        run.GradFillColors = stops.ToArray();
                        run.FontColor = stops[0]; // 保留第一色作主色
                    }
                }
            }
            // 上标/下标（S15-06）：<a:rPr baseline="30000"> 为上标，baseline="-25000" 为下标
            var baselineAttr = rPr.GetAttribute("baseline");
            if (!String.IsNullOrEmpty(baselineAttr) && Int32.TryParse(baselineAttr, out var bv))
            {
                if (bv >= 20000) run.Superscript = true;
                else if (bv <= -15000) run.Subscript = true;
            }
            var latin = rPr.SelectSingleNode(".//*[local-name()='latin']") as XmlElement;
            var ea = rPr.SelectSingleNode(".//*[local-name()='ea']") as XmlElement;
            var cs = rPr.SelectSingleNode(".//*[local-name()='cs']") as XmlElement;
            var sym = rPr.SelectSingleNode(".//*[local-name()='sym']") as XmlElement;
            run.LatinFontName = latin?.GetAttribute("typeface");
            run.EastAsianFontName = ea?.GetAttribute("typeface");
            run.ComplexScriptFontName = cs?.GetAttribute("typeface");
            run.SymbolFontName = sym?.GetAttribute("typeface");
            var hlink = rPr.SelectSingleNode(".//*[local-name()='hlinkClick']") as XmlElement;
            if (hlink != null)
            {
                var rId = hlink.GetAttribute("r:id");
                if (rId.Length > 0 && rels.TryGetValue(rId, out var rel)) run.HyperlinkUrl = rel.Target;
            }
        }
        ApplyDefaults(run, pDefRPr, boldExplicit);
        if (defaults != null) ApplyDefaults(run, defaults.Value, boldExplicit);
        if (run.FontSize <= 0) run.FontSize = 18;
        return run;
    }

    private static void ApplyDefaults(PptTextRun run, XmlElement? defRPr, Boolean boldExplicit = false)
    {
        if (defRPr == null) return;
        if (run.FontSize <= 0)
        {
            var sz = defRPr.GetAttribute("sz");
            if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) run.FontSize = sv / 100;
        }
        if (!boldExplicit && !run.Bold) run.Bold = defRPr.GetAttribute("b") == "1";
        if (run.LatinFontName == null)
        {
            var l = defRPr.SelectSingleNode(".//*[local-name()='latin']") as XmlElement;
            run.LatinFontName = l?.GetAttribute("typeface");
        }
        if (run.EastAsianFontName == null)
        {
            var e = defRPr.SelectSingleNode(".//*[local-name()='ea']") as XmlElement;
            run.EastAsianFontName = e?.GetAttribute("typeface");
        }
        if (run.ComplexScriptFontName == null)
        {
            var c = defRPr.SelectSingleNode(".//*[local-name()='cs']") as XmlElement;
            run.ComplexScriptFontName = c?.GetAttribute("typeface");
        }
        if (run.SymbolFontName == null)
        {
            var s = defRPr.SelectSingleNode(".//*[local-name()='sym']") as XmlElement;
            run.SymbolFontName = s?.GetAttribute("typeface");
        }
    }

    private static void ApplyDefaults(PptTextRun run, DefaultRunProps def, Boolean boldExplicit = false)
    {
        if (run.FontSize <= 0 && def.FontSize > 0) run.FontSize = def.FontSize;
        if (!boldExplicit && !run.Bold) run.Bold = def.Bold;
        if (run.LatinFontName == null) run.LatinFontName = def.LatinFontName;
        if (run.EastAsianFontName == null) run.EastAsianFontName = def.EastAsianFontName;
        if (run.ComplexScriptFontName == null) run.ComplexScriptFontName = def.ComplexScriptFontName;
        if (run.SymbolFontName == null) run.SymbolFontName = def.SymbolFontName;
        if (run.FontColor == null) run.FontColor = def.FontColor;
    }

    private struct DefaultRunProps
    {
        public Int32 FontSize;
        public Boolean Bold;
        public String? FontColor;
        public String? LatinFontName;
        public String? EastAsianFontName;
        public String? ComplexScriptFontName;
        public String? SymbolFontName;
        /// <summary>兼容属性：getter 返回 EastAsianFontName ?? LatinFontName</summary>
        public readonly String? FontName => EastAsianFontName ?? LatinFontName;
    }

    private static Dictionary<String, DefaultRunProps> ParseListStyleDefaults(XmlElement txBody)
    {
        var r = new Dictionary<String, DefaultRunProps>();
        var lst = txBody.SelectSingleNode(".//*[local-name()='lstStyle']") as XmlElement;
        if (lst == null) return r;
        var gd = lst.SelectSingleNode(".//*[local-name()='defPPr']/*[local-name()='defRPr']") as XmlElement;
        if (gd != null) r[""] = PropsFromDefRPr(gd);
        for (var i = 1; i <= 9; i++)
        {
            var ld = lst.SelectSingleNode($".//*[local-name()='lvl{i}pPr']/*[local-name()='defRPr']") as XmlElement;
            if (ld != null) r[(i - 1).ToString()] = PropsFromDefRPr(ld);
        }
        return r;
    }

    /// <summary>从版式 XML 读取 lstStyle 字体默认值（占位符继承用）</summary>
    /// <param name="layoutIndex">版式索引（0起始）</param>
    /// <returns>lstStyle 默认值字典，版式不存在或没有 lstStyle 时返回空字典</returns>
    private Dictionary<String, DefaultRunProps> ParseLayoutLstStyleDefaults(Int32 layoutIndex)
    {
        var r = new Dictionary<String, DefaultRunProps>();
        var entry = _zip.GetEntry($"ppt/slideLayouts/slideLayout{layoutIndex + 1}.xml");
        if (entry == null) return r;
        var doc = LoadXml(entry);
        var txBody = doc.SelectSingleNode("//*[local-name()='txBody']") as XmlElement;
        if (txBody == null)
        {
            // 版式可能没有 txBody 或有多个，尝试取 cSld/spTree 下第一个含 lstStyle 的 txBody
            var spTree = doc.SelectSingleNode("//*[local-name()='spTree']");
            if (spTree != null)
            {
                foreach (XmlNode child in spTree.ChildNodes)
                {
                    if (child is not XmlElement el) continue;
                    txBody = el.SelectSingleNode(".//*[local-name()='txBody']") as XmlElement;
                    if (txBody != null && txBody.SelectSingleNode(".//*[local-name()='lstStyle']") != null) break;
                    txBody = null;
                }
            }
        }
        if (txBody == null) return r;
        return ParseListStyleDefaults(txBody);
    }

    private static DefaultRunProps PropsFromDefRPr(XmlElement defRPr)
    {
        var p = new DefaultRunProps();
        var sz = defRPr.GetAttribute("sz");
        if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) p.FontSize = sv / 100;
        p.Bold = defRPr.GetAttribute("b") == "1";
        var sf = defRPr.SelectSingleNode("*[local-name()='solidFill']") as XmlElement;
        var fc = sf?.SelectSingleNode("*[local-name()='srgbClr']") as XmlElement;
        p.FontColor = fc?.GetAttribute("val");
        var l = defRPr.SelectSingleNode(".//*[local-name()='latin']") as XmlElement;
        var e = defRPr.SelectSingleNode(".//*[local-name()='ea']") as XmlElement;
        var c = defRPr.SelectSingleNode(".//*[local-name()='cs']") as XmlElement;
        var s = defRPr.SelectSingleNode(".//*[local-name()='sym']") as XmlElement;
        p.LatinFontName = l?.GetAttribute("typeface");
        p.EastAsianFontName = e?.GetAttribute("typeface");
        p.ComplexScriptFontName = c?.GetAttribute("typeface");
        p.SymbolFontName = s?.GetAttribute("typeface");
        return p;
    }

    private static PptTextRun CreateDefaultRun(String text, Dictionary<String, DefaultRunProps> lstDefaults, String lvl, XmlElement? pDefRPr)
    {
        var run = new PptTextRun { Text = text, FontSize = 18 };
        ApplyDefaults(run, pDefRPr);
        if (lstDefaults.TryGetValue(lvl, out var ld)) ApplyDefaults(run, ld);
        else if (lstDefaults.TryGetValue("", out var gd2)) ApplyDefaults(run, gd2);
        return run;
    }

    /// <summary>解析图片或视频（p:pic）</summary>
    private void ParsePicture(XmlElement pic, PptSlide slide, Dictionary<String, (String Target, String Type)> rels)
    {
        var xfrmEl = pic.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement;
        var (left, top, width, height) = ParseXfrm(xfrmEl);
        var rotation = ParseRotation(xfrmEl);
        var videoFile = pic.SelectSingleNode(".//*[local-name()='videoFile']") as XmlElement;

        // 读取缩略图/海报帧（图片用 blip embed，视频用 blip embed 作海报帧）
        Byte[]? thumbData = null;
        String? thumbExt = null;
        var blip = pic.SelectSingleNode(".//*[local-name()='blip']") as XmlElement;
        var embedId = blip?.GetAttribute("r:embed") ?? blip?.GetAttribute("embed");
        if (embedId != null && rels.TryGetValue(embedId, out var thumbRel))
        {
            var thumbName = Path.GetFileName(thumbRel.Target);
            var thumbEntry = _zip.GetEntry($"ppt/media/{thumbName}");
            if (thumbEntry != null)
            {
                thumbExt = Path.GetExtension(thumbName).TrimStart('.').ToLowerInvariant();
                using (var ms = new MemoryStream())
                using (var es = thumbEntry.Open()) { es.CopyTo(ms); thumbData = ms.ToArray(); }
            }
        }

        if (videoFile != null)
        {
            var vlId = videoFile.GetAttribute("r:link");
            if (vlId.Length > 0 && rels.TryGetValue(vlId, out var vr))
            {
                var mn = Path.GetFileName(vr.Target);
                var me = _zip.GetEntry($"ppt/media/{mn}");
                if (me != null)
                {
                    var ext = Path.GetExtension(mn).TrimStart('.').ToLowerInvariant();
                    Byte[] data;
                    using (var ms = new MemoryStream())
                    using (var es = me.Open()) { es.CopyTo(ms); data = ms.ToArray(); }
                    var vid = new PptVideo { Data = data, Extension = ext, Left = left, Top = top, Width = width, Height = height };
                    if (thumbData != null)
                    {
                        vid.ThumbnailData = thumbData;
                        vid.ThumbnailExtension = thumbExt!;
                    }
                    slide.Videos.Add(vid);
                }
            }
            return;
        }

        // 纯图片
        if (embedId == null || !rels.TryGetValue(embedId, out var rel)) return;
        var iname = Path.GetFileName(rel.Target);
        var ientry = _zip.GetEntry($"ppt/media/{iname}");
        if (ientry == null) return;
        var iext = Path.GetExtension(iname).TrimStart('.').ToLowerInvariant();
        Byte[] idata;
        using (var ms = new MemoryStream())
        using (var es = ientry.Open()) { es.CopyTo(ms); idata = ms.ToArray(); }

        var isSvg = iext == "svg";

        // 检测 asvg:svgBlip（PowerPoint 2016+ SVG 标准格式），使用 SVG 数据覆盖 PNG 缩略图
        var svgBlip = blip.SelectSingleNode("*[local-name()='svgBlip']") as XmlElement;
        if (svgBlip != null)
        {
            var svgEmbedId = svgBlip.GetAttribute("r:embed") ?? svgBlip.GetAttribute("embed");
            if (svgEmbedId != null && rels.TryGetValue(svgEmbedId, out var svgRel))
            {
                var svgName = Path.GetFileName(svgRel.Target);
                var svgEntry = _zip.GetEntry($"ppt/media/{svgName}");
                if (svgEntry != null)
                {
                    using var svgMs = new MemoryStream();
                    using var svgEs = svgEntry.Open();
                    svgEs.CopyTo(svgMs);
                    idata = svgMs.ToArray();
                    iext = "svg";
                    isSvg = true;
                }
            }
        }

        slide.Images.Add(new PptImage { Data = idata, Extension = iext, Left = left, Top = top, Width = width, Height = height, IsSvg = isSvg, Rotation = rotation });
    }

    private void ParseGraphicFrame(XmlElement gf, PptSlide slide, Dictionary<String, (String Target, String Type)> rels)
    {
        var (left, top, width, height) = ParseXfrm(gf.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        var tbl = gf.SelectSingleNode(".//*[local-name()='tbl']") as XmlElement;
        if (tbl != null) { ParseTable(tbl, left, top, width, height, slide); return; }
        var cr = gf.SelectSingleNode(".//*[local-name()='chart']") as XmlElement;
        if (cr != null)
        {
            var rId = cr.GetAttribute("r:id");
            if (rId.Length > 0 && rels.TryGetValue(rId, out var rel))
                ParseChartFromRel(rel.Target, left, top, width, height, slide);
        }
    }

    private static void ParseTable(XmlElement tbl, Int64 left, Int64 top, Int64 width, Int64 height, PptSlide slide)
    {
        var table = new PptTable { Left = left, Top = top, Width = width, Height = height };
        var gcs = tbl.SelectNodes(".//*[local-name()='gridCol']");
        if (gcs != null)
        {
            var cw = new List<Int64>();
            foreach (XmlElement gc in gcs)
            { var w = gc.GetAttribute("w"); cw.Add(w.Length > 0 && Int64.TryParse(w, out var v) ? v : width / Math.Max(1, gcs.Count)); }
            table.ColWidths = cw.ToArray();
        }
        var tpr = tbl.SelectSingleNode(".//*[local-name()='tblPr']") as XmlElement;
        if (tpr != null) table.FirstRowHeader = tpr.GetAttribute("firstRow") == "1";
        var rows = tbl.SelectNodes(".//*[local-name()='tr']");
        if (rows != null)
        {
            var ri = 0;
            foreach (XmlElement tr in rows)
            {
                var cells = tr.SelectNodes(".//*[local-name()='tc']");
                if (cells == null) { ri++; continue; }
                var rd = new String[cells.Count];
                var ci = 0;
                foreach (XmlElement tc in cells)
                {
                    var ct = new StringBuilder();
                    foreach (XmlElement t in tc.SelectNodes(".//*[local-name()='t']")!) ct.Append(t.InnerText);
                    rd[ci] = ct.ToString();
                    // 解析合并单元格（S11-01 Reader 侧）
                    var gs = tc.GetAttribute("gridSpan");
                    var rs = tc.GetAttribute("rowSpan");
                    var vm = tc.GetAttribute("vMerge");
                    var colSpan = gs.Length > 0 && Int32.TryParse(gs, out var csv) ? csv : 1;
                    var rowSpan = rs.Length > 0 && Int32.TryParse(rs, out var rsv) ? rsv : 1;
                    if (colSpan > 1 || rowSpan > 1 || vm == "1")
                    {
                        if (vm == "1")
                        {
                            for (var pr = ri - 1; pr >= 0; pr--)
                            {
                                if (table.MergedCells.TryGetValue((pr, ci), out var prev))
                                {
                                    table.MergedCells[(pr, ci)] = (prev.ColSpan, prev.RowSpan + 1);
                                    break;
                                }
                            }
                        }
                        else
                            table.MergedCells[(ri, ci)] = (colSpan, rowSpan);
                    }
                    var tpr2 = tc.SelectSingleNode(".//*[local-name()='tcPr']") as XmlElement;
                    var rpr2 = tc.SelectSingleNode(".//*[local-name()='rPr']") as XmlElement;
                    if (tpr2 != null || rpr2 != null)
                    {
                        var cs = new PptCellStyle();
                        var bg = tpr2?.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                        if (bg != null) cs.BackgroundColor = bg.GetAttribute("val");
                        if (rpr2 != null)
                        {
                            var sz = rpr2.GetAttribute("sz");
                            if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) cs.FontSize = sv / 100;
                            cs.Bold = rpr2.GetAttribute("b") == "1";
                            var fc = rpr2.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                            if (fc != null) cs.FontColor = fc.GetAttribute("val");
                        }
                        table.CellStyles[(ri, ci)] = cs;
                    }
                    // 解析表格边框（S11-02 Reader 侧）
                    if (tpr2 != null)
                    {
                        var border = new PptCellBorder();
                        var hasBorder = false;
                        var lnL = tpr2.SelectSingleNode("*[local-name()='lnL']") as XmlElement;
                        var lnR = tpr2.SelectSingleNode("*[local-name()='lnR']") as XmlElement;
                        var lnT = tpr2.SelectSingleNode("*[local-name()='lnT']") as XmlElement;
                        var lnB = tpr2.SelectSingleNode("*[local-name()='lnB']") as XmlElement;
                        void ParseOneBorder(XmlElement ln, Action<String?> setColor, Action<Int32> setWidth)
                        {
                            var wv = ln.GetAttribute("w");
                            if (wv.Length > 0 && Int32.TryParse(wv, out var w)) setWidth(w);
                            var clr = ln.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
                            setColor(clr?.GetAttribute("val"));
                        }
                        if (lnL != null) { ParseOneBorder(lnL, c => border.LeftColor = c, w => border.LeftWidth = w); hasBorder = true; }
                        if (lnR != null) { ParseOneBorder(lnR, c => border.RightColor = c, w => border.RightWidth = w); hasBorder = true; }
                        if (lnT != null) { ParseOneBorder(lnT, c => border.TopColor = c, w => border.TopWidth = w); hasBorder = true; }
                        if (lnB != null) { ParseOneBorder(lnB, c => border.BottomColor = c, w => border.BottomWidth = w); hasBorder = true; }
                        if (hasBorder) table.CellBorders[(ri, ci)] = border;
                    }
                    ci++;
                }
                table.Rows.Add(rd);
                ri++;
            }
        }
        slide.Tables.Add(table);
    }

    private void ParseChartFromRel(String target, Int64 left, Int64 top, Int64 width, Int64 height, PptSlide slide)
    {
        var cp = "ppt/" + target.TrimStart('.').TrimStart('/');
        var ce = _zip.GetEntry(cp);
        if (ce == null) return;
        var cd = LoadXml(ce);
        const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var cns = new XmlNamespaceManager(cd.NameTable);
        cns.AddNamespace("c", C);
        var chart = new PptChart { Left = left, Top = top, Width = width, Height = height };
        var ctn = cd.SelectSingleNode("//*[substring(local-name(), string-length(local-name())-4) = 'Chart'][@*]", null);
        chart.ChartType = ctn?.LocalName?.Replace("Chart", String.Empty) ?? "bar";
        var tn = cd.SelectSingleNode("//c:title//c:rich//a:t", cns) as XmlElement;
        chart.Title = tn?.InnerText;
        var fcn = cd.SelectSingleNode("//c:ser[1]/c:cat//c:strCache", cns) ?? cd.SelectSingleNode("//c:ser[1]/c:cat//c:numCache", cns);
        if (fcn != null)
        {
            var cats = new List<String>();
            foreach (XmlElement pt in fcn.SelectNodes("c:pt/c:v", cns)!) cats.Add(pt.InnerText);
            chart.Categories = cats.ToArray();
        }
        foreach (XmlElement ser in cd.SelectNodes("//c:ser", cns)!)
        {
            var sn = ser.SelectSingleNode(".//c:tx//c:v", cns)?.InnerText ?? String.Empty;
            var vals = new List<Double>();
            foreach (XmlElement v in ser.SelectNodes(".//c:val//c:numCache/c:pt/c:v", cns)!)
            { if (Double.TryParse(v.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) vals.Add(d); }
            chart.Series.Add(new PptChartSeries { Name = sn, Values = vals.ToArray() });
        }
        slide.Charts.Add(chart);
    }

    private void ParseGroup(XmlElement grpSp, PptSlide slide, Dictionary<String, (String Target, String Type)> rels,
        Dictionary<String, DefaultRunProps>? layoutDefaults = null)
    {
        var (left, top, width, height) = ParseXfrm(grpSp.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        var g = new PptGroup { Left = left, Top = top, Width = width, Height = height };
        foreach (XmlNode child in grpSp.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            switch (el.LocalName)
            {
                case "sp": ParseShapeOrTextBoxInGroup(el, g, rels, layoutDefaults); break;
                case "pic": ParsePictureInGroup(el, g, rels); break;
                case "grpSp": ParseGroupInGroup(el, g, rels, layoutDefaults); break;
            }
        }
        slide.Groups.Add(g);
    }

    private void ParseShapeOrTextBoxInGroup(XmlElement sp, PptGroup g, Dictionary<String, (String Target, String Type)> rels,
        Dictionary<String, DefaultRunProps>? layoutDefaults = null)
    {
        var isTB = sp.SelectSingleNode(".//*[local-name()='cNvSpPr']")?.Attributes?["txBox"]?.Value == "1";
        var (l, t, w, h) = ParseXfrm(sp.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        var st = sp.SelectSingleNode(".//*[local-name()='prstGeom']")?.Attributes?["prst"]?.Value ?? "textBox";
        var fc = ParseFillColor(sp.SelectSingleNode(".//*[local-name()='spPr']") as XmlElement);
        var txt = new StringBuilder();
        foreach (XmlElement tt in sp.SelectNodes(".//*[local-name()='t']")!) txt.Append(tt.InnerText);
        var text = txt.ToString();

        // 解析字体属性：从第一个 rPr 提取字号/粗体/颜色/字体名称
        var fontSize = 0;
        var bold = false;
        String? fontColor = null;
        String? latinFn = null, eaFn = null, csFn = null, symFn = null;
        var firstRPr = sp.SelectSingleNode(".//*[local-name()='rPr']") as XmlElement;
        if (firstRPr != null)
        {
            var sz = firstRPr.GetAttribute("sz");
            if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) fontSize = sv / 100;
            bold = firstRPr.GetAttribute("b") == "1";
            var sf = firstRPr.SelectSingleNode("*[local-name()='solidFill']") as XmlElement;
            var sfc = sf?.SelectSingleNode("*[local-name()='srgbClr']") as XmlElement;
            if (sfc != null) fontColor = sfc.GetAttribute("val");
            else
            {
                var sc = sf?.SelectSingleNode("*[local-name()='schemeClr']") as XmlElement;
                if (sc != null) fontColor = "scheme:" + sc.GetAttribute("val");
            }
            var latin = firstRPr.SelectSingleNode(".//*[local-name()='latin']") as XmlElement;
            var ea = firstRPr.SelectSingleNode(".//*[local-name()='ea']") as XmlElement;
            var cs = firstRPr.SelectSingleNode(".//*[local-name()='cs']") as XmlElement;
            var sym = firstRPr.SelectSingleNode(".//*[local-name()='sym']") as XmlElement;
            latinFn = latin?.GetAttribute("typeface");
            eaFn = ea?.GetAttribute("typeface");
            csFn = cs?.GetAttribute("typeface");
            symFn = sym?.GetAttribute("typeface");
        }

        // 若无显式 rPr，尝试从 lstStyle/defRPr 继承默认字体
        if (firstRPr == null)
        {
            var txBody = sp.SelectSingleNode(".//*[local-name()='txBody']") as XmlElement;
            var defRPr = txBody?.SelectSingleNode(".//*[local-name()='lstStyle']/*[local-name()='defPPr']/*[local-name()='defRPr']") as XmlElement;
            if (defRPr != null)
            {
                var sz = defRPr.GetAttribute("sz");
                if (sz.Length > 0 && Int32.TryParse(sz, out var sv)) fontSize = sv / 100;
                bold = defRPr.GetAttribute("b") == "1";
                var ln = defRPr.SelectSingleNode(".//*[local-name()='latin']") as XmlElement;
                var e = defRPr.SelectSingleNode(".//*[local-name()='ea']") as XmlElement;
                var c = defRPr.SelectSingleNode(".//*[local-name()='cs']") as XmlElement;
                var s = defRPr.SelectSingleNode(".//*[local-name()='sym']") as XmlElement;
                latinFn = ln?.GetAttribute("typeface");
                eaFn = e?.GetAttribute("typeface");
                csFn = c?.GetAttribute("typeface");
                symFn = s?.GetAttribute("typeface");
            }
        }

        if (isTB)
            g.TextBoxes.Add(new PptTextBox
            {
                Left = l, Top = t, Width = w, Height = h, Text = text, BackgroundColor = fc,
                FontSize = fontSize > 0 ? fontSize : 18, Bold = bold, FontColor = fontColor,
                LatinFontName = latinFn, EastAsianFontName = eaFn,
                ComplexScriptFontName = csFn, SymbolFontName = symFn,
            });
        else
            g.Shapes.Add(new PptShape
            {
                Left = l, Top = t, Width = w, Height = h, ShapeType = st, Text = text, FillColor = fc,
                FontSize = fontSize > 0 ? fontSize : 14, Bold = bold, FontColor = fontColor,
                LatinFontName = latinFn, EastAsianFontName = eaFn,
                ComplexScriptFontName = csFn, SymbolFontName = symFn,
            });
    }

    private void ParsePictureInGroup(XmlElement pic, PptGroup g, Dictionary<String, (String Target, String Type)> rels)
    {
        var (l, t, w, h) = ParseXfrm(pic.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement);
        g.Shapes.Add(new PptShape { Left = l, Top = t, Width = w, Height = h, ShapeType = "rect", Text = "[图片]" });
    }

    private void ParseGroupInGroup(XmlElement inner, PptGroup pg, Dictionary<String, (String Target, String Type)> rels,
        Dictionary<String, DefaultRunProps>? layoutDefaults = null)
    {
        foreach (XmlNode child in inner.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            switch (el.LocalName)
            {
                case "sp": ParseShapeOrTextBoxInGroup(el, pg, rels, layoutDefaults); break;
                case "pic": ParsePictureInGroup(el, pg, rels); break;
                case "grpSp": ParseGroupInGroup(el, pg, rels, layoutDefaults); break;
            }
        }
    }

    private static void ParseNotes(XmlDocument doc, PptSlide slide)
    {
        var n = doc.SelectSingleNode("//*[local-name()='notes']");
        if (n == null) return;
        var sb = new StringBuilder();
        foreach (XmlElement t in n.SelectNodes(".//*[local-name()='t']")!) sb.Append(t.InnerText);
        var text = sb.ToString();
        if (text.Length > 0) slide.Notes = text;
    }

    private static void ParseTransition(XmlDocument doc, PptSlide slide)
    {
        var trans = doc.SelectSingleNode("//*[local-name()='transition']") as XmlElement;
        if (trans == null) return;
        var dur = trans.GetAttribute("dur");
        var ac = trans.GetAttribute("advClick");
        var tr = new PptTransition
        {
            DurationMs = dur.Length > 0 && Int32.TryParse(dur, out var d) ? d : 500,
            AdvanceOnClick = ac != "0",
        };
        foreach (var tp in new[] { "fade", "push", "wipe", "zoom", "split", "cut" })
        {
            var ch = trans.SelectSingleNode($"*[local-name()='{tp}']") as XmlElement;
            if (ch == null) continue;
            tr.Type = tp;
            var dir = ch.GetAttribute("dir");
            if (dir.Length > 0) tr.Direction = dir;
            break;
        }
        slide.Transition = tr;
    }

    /// <summary>解析连接器（S13-02 Reader 侧）</summary>
    private static void ParseConnector(XmlElement cxnSp, PptSlide slide)
    {
        var xfrm = cxnSp.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement;
        var (l, t, w, h) = ParseXfrm(xfrm);
        var prstGeom = cxnSp.SelectSingleNode(".//*[local-name()='prstGeom']") as XmlElement;
        var prst = prstGeom?.GetAttribute("prst") ?? "straightConnector1";
        // 去掉 Connector1 后缀还原为类型名
        var ct = prst.Replace("Connector1", "").Replace("1", "");
        var (lc, lw) = ParseLine(cxnSp.SelectSingleNode(".//*[local-name()='spPr']") as XmlElement);
        var ln = cxnSp.SelectSingleNode(".//*[local-name()='ln']") as XmlElement;
        var tailEnd = ln?.SelectSingleNode("*[local-name()='tailEnd']") as XmlElement;
        var headEnd = ln?.SelectSingleNode("*[local-name()='headEnd']") as XmlElement;
        var prstDash = ln?.SelectSingleNode("*[local-name()='prstDash']") as XmlElement;
        slide.Connectors.Add(new PptConnector
        {
            Left = l, Top = t, Width = w, Height = h,
            ConnectorType = ct.Length > 0 ? ct : "straight",
            LineColor = lc,
            LineWidth = lw > 0 ? (Int32)lw : 9525,
            StartArrow = tailEnd?.GetAttribute("type"),
            EndArrow = headEnd?.GetAttribute("type"),
            DashStyle = prstDash?.GetAttribute("val"),
        });
    }

    /// <summary>解析幻灯片批注（S13-01 Reader 侧）</summary>
    private void ParseSlideComments(Int32 slideIndex, PptSlide slide)
    {
        var commentsEntry = _zip.GetEntry($"ppt/comments/comment{slideIndex + 1}.xml");
        if (commentsEntry == null) return;
        var doc = LoadXml(commentsEntry);
        var authorsDoc = TryLoadCommentsAuthors();
        var authorMap = new Dictionary<String, String>();
        if (authorsDoc != null)
        {
            foreach (XmlElement author in authorsDoc.SelectNodes("//*[local-name()='cmAuthor']")!)
            {
                var aId = author.GetAttribute("id");
                var aName = author.GetAttribute("name");
                var aUid = author.GetAttribute("uid");
                if (aId.Length > 0) authorMap[aId] = aName;
            }
        }
        var idx = 0;
        foreach (XmlElement cm in doc.SelectNodes("//*[local-name()='cm']")!)
        {
            var c = new PptComment { Index = ++idx };
            c.Author = authorMap.TryGetValue(cm.GetAttribute("authorId"), out var name) ? name : cm.GetAttribute("authorId");
            var dtStr = cm.GetAttribute("dt");
            if (DateTime.TryParse(dtStr, out var dt)) c.Date = dt;
            var pos = cm.SelectSingleNode("*[local-name()='pos']") as XmlElement;
            if (pos != null)
            {
                if (Int32.TryParse(pos.GetAttribute("x"), out var px) && SlideWidth > 0) c.X = (Single)px / SlideWidth;
                if (Int32.TryParse(pos.GetAttribute("y"), out var py) && SlideHeight > 0) c.Y = (Single)py / SlideHeight;
            }
            c.Text = cm.SelectSingleNode("*[local-name()='text']")?.InnerText;
            slide.Comments.Add(c);
        }
    }

    /// <summary>加载批注作者列表</summary>
    private XmlDocument? TryLoadCommentsAuthors()
    {
        var entry = _zip.GetEntry("ppt/comments/commentAuthors.xml");
        return entry != null ? LoadXml(entry) : null;
    }

    /// <summary>解析幻灯片元素动画（S12 Reader 侧）</summary>
    private static void ParseSlideAnimations(XmlDocument sldDoc, PptSlide slide)
    {
        var timing = sldDoc.SelectSingleNode("//*[local-name()='timing']") as XmlElement;
        if (timing == null) return;
        // 遍历 <p:par>（顺序容器）下的动画节点
        var order = 0;
        foreach (XmlElement par in timing.SelectNodes(".//*[local-name()='par']")!)
        {
            foreach (XmlNode child in par.ChildNodes)
            {
                if (child is not XmlElement el) continue;
                var localName = el.LocalName;
                // 识别动画类型：animEffect(进入)/animEmph(强调)/animMotion(路径)
                var category = localName switch
                {
                    "animEffect" => PptAnimationCategory.Entrance,
                    "animEmph" => PptAnimationCategory.Emphasis,
                    "animMotion" => PptAnimationCategory.MotionPath,
                    _ => PptAnimationCategory.Entrance,
                };
                var cBhvr = el.SelectSingleNode("*[local-name()='cBhvr']") as XmlElement;
                if (cBhvr == null) continue;
                var anim = new PptAnimation
                {
                    Order = order++,
                    Category = category,
                    Effect = el.GetAttribute("filter") ?? el.LocalName,
                };
                // 目标形状
                var tgt = cBhvr.SelectSingleNode("*[local-name()='tgtEl']/*[local-name()='spTgt']") as XmlElement;
                if (tgt != null && Int32.TryParse(tgt.GetAttribute("spid"), out var spid))
                    anim.TargetIndex = spid;
                // 时长
                var durEl = cBhvr.SelectSingleNode("*[local-name()='cTn']") as XmlElement;
                if (durEl != null)
                {
                    var durVal = durEl.GetAttribute("dur");
                    if (durVal.Length > 0) anim.DurationMs = (Int32)(ParseDuration(durVal) / 1000);
                }
                slide.Animations.Add(anim);
            }
        }
    }

    /// <summary>解析动画时长（如 "indefinite" 或毫秒数）</summary>
    private static Int64 ParseDuration(String val) => Int64.TryParse(val, out var ms) ? ms : 0;

    /// <summary>解析文档属性（S14-03 Reader 侧）</summary>
    private void ParseDocProps(PptDocument doc)
    {
        // core.xml
        var coreEntry = _zip.GetEntry("docProps/core.xml");
        if (coreEntry != null)
        {
            var coreDoc = LoadXml(coreEntry);
            doc.Properties.Title = coreDoc.SelectSingleNode("//*[local-name()='title']")?.InnerText;
            doc.Properties.Author = coreDoc.SelectSingleNode("//*[local-name()='creator']")?.InnerText;
            doc.Properties.Subject = coreDoc.SelectSingleNode("//*[local-name()='subject']")?.InnerText;
            doc.Properties.Description = coreDoc.SelectSingleNode("//*[local-name()='description']")?.InnerText;
        }
        // app.xml
        var appEntry = _zip.GetEntry("docProps/app.xml");
        if (appEntry != null)
        {
            var appDoc = LoadXml(appEntry);
            var company = appDoc.SelectSingleNode("//*[local-name()='Company']")?.InnerText;
            if (!company.IsNullOrEmpty() && doc.Properties.Author.IsNullOrEmpty())
                doc.Properties.Author = company;
        }
    }

    /// <summary>解析全局页眉页脚（S13-03 Reader 侧）</summary>
    private void ParseHeaderFooter(PptDocument doc)
    {
        var presEntry = _zip.GetEntry("ppt/presentation.xml");
        if (presEntry == null) return;
        var presDoc = LoadXml(presEntry);
        var hf = presDoc.SelectSingleNode("//*[local-name()='hf']") as XmlElement;
        if (hf == null) return;
        var footer = hf.GetAttribute("footer");
        var showSlideNum = hf.GetAttribute("showSlideNum");
        var dt = hf.GetAttribute("dt");
        var fdt = hf.GetAttribute("fdt");
        var dfmt = hf.GetAttribute("dfmt");
        doc.HeaderFooter = new PptHeaderFooter
        {
            ShowFooter = footer.Length > 0,
            FooterText = footer.Length > 0 ? footer : null,
            ShowPageNumber = showSlideNum == "1",
            ShowDate = dt == "1",
            DateAutomatic = fdt.Length == 0,
            FixedDate = fdt.Length > 0 ? fdt : null,
            DateFormat = dfmt.Length > 0 ? dfmt : null,
        };
    }

    private static String? ParseFillColor(XmlElement? parent)
    {
        if (parent == null) return null;
        var clr = parent.SelectSingleNode(".//*[local-name()='solidFill']/*[local-name()='srgbClr']") as XmlElement;
        return clr?.GetAttribute("val");
    }

    private static (String? Color, Int64 Width) ParseLine(XmlElement? parent)
    {
        if (parent == null) return (null, 0);
        var ln = parent.SelectSingleNode(".//*[local-name()='ln']") as XmlElement;
        if (ln == null) return (null, 0);
        var w = ln.GetAttribute("w");
        var width = w.Length > 0 && Int64.TryParse(w, out var wv) ? wv : 0L;
        var clr = ln.SelectSingleNode(".//*[local-name()='srgbClr']") as XmlElement;
        return (clr?.GetAttribute("val"), width);
    }

    private static (Int64 Left, Int64 Top, Int64 Width, Int64 Height) ParseXfrm(XmlElement? xfrm)
    {
        if (xfrm == null) return (0, 0, 0, 0);
        var off = xfrm.SelectSingleNode("*[local-name()='off']") as XmlElement;
        var ext = xfrm.SelectSingleNode("*[local-name()='ext']") as XmlElement;
        return (
            off != null && Int64.TryParse(off.GetAttribute("x"), out var x) ? x : 0,
            off != null && Int64.TryParse(off.GetAttribute("y"), out var y) ? y : 0,
            ext != null && Int64.TryParse(ext.GetAttribute("cx"), out var cx) ? cx : 0,
            ext != null && Int64.TryParse(ext.GetAttribute("cy"), out var cy) ? cy : 0
        );
    }

    /// <summary>解析旋转角度（S15-02 Reader 侧）</summary>
    private static Int32 ParseRotation(XmlElement? xfrm)
    {
        if (xfrm == null) return 0;
        var rot = xfrm.GetAttribute("rot");
        return rot.Length > 0 && Int32.TryParse(rot, out var r) ? r : 0;
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（幻灯片间换行分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText() => ReadAllText();

    /// <summary>提取 Markdown 格式（每页用标题分隔）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        var count = GetSlideCount();
        if (count == 0) return null;

        var sb = new StringBuilder();
        for (var i = 0; i < count; i++)
        {
            if (i > 0) sb.AppendLine();
            sb.AppendLine($"## 幻灯片 {i + 1}");
            sb.AppendLine();
            sb.AppendLine(GetSlideText(i));
        }
        return sb.ToString();
    }
    #endregion
}
