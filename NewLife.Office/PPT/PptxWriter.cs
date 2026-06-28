using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using NewLife.Buffers;

namespace NewLife.Office;

/// <summary>PowerPoint pptx 写入器</summary>
/// <remarks>
/// 直接操作 Open XML（ZIP+XML）生成 .pptx 文件。
/// 支持文本框/表格/图片/背景/备注等核心功能。
/// 坐标使用 EMU（英制单位，914400 EMU = 1 英寸，360000 EMU = 1 cm）。
/// </remarks>
public partial class PptxWriter : IDisposable
{
    #region 属性
    /// <summary>幻灯片宽度（EMU，默认 16:9 = 12192000）</summary>
    public Int64 SlideWidth { get; set; } = 12192000;

    /// <summary>幻灯片高度（EMU，默认 16:9 = 6858000）</summary>
    public Int64 SlideHeight { get; set; } = 6858000;

    /// <summary>幻灯片集合</summary>
    public List<PptSlide> Slides { get; } = [];

    /// <summary>主题强调色（6个，默认 Office 配色，可通过 SetAccentColors 修改）</summary>
    public String[] AccentColors { get; set; } = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"];

    /// <summary>全局页眉页脚设置（S13-03），null 表示不加页脚</summary>
    public PptHeaderFooter? HeaderFooter { get; set; }

    /// <summary>节（Section）列表（S13-04），按节组织幻灯片，null 表示不使用节</summary>
    public List<PptSection>? Sections { get; set; }
    #endregion

    #region 私有字段
    private Int32 _imgGlobal = 1;
    private Int32 _chartGlobal = 1;
    private Int32 _hlinkGlobal = 1;
    private Int32 _mediaGlobal = 1;
    private Int32 _videoGlobal = 1;
    private String? _protectionHash;
    private String? _protectionSalt;
    // 跨文件复制的原始幻灯片（S10-04）：(幻灯片XML, rels XML)
    private readonly List<(String SlideXml, String RelsXml)> _rawSlides = [];
    // 跨文件复制的媒体文件：(文件名, 字节数据)
    private readonly List<(String Name, Byte[] Data)> _rawSlideMedia = [];
    // 从模板加载的母版/版式/主题（S04-Master）
    private readonly List<PptPartContent> _masterContents = [];
    private readonly List<PptPartContent> _layoutContents = [];
    private readonly List<String> _layoutNames = [];
    private String? _templateThemeXml;
    private readonly Dictionary<String, Byte[]> _infraMedia = [];
    // 编程式创建的母版（Phase 5：无需模板文件）
    private readonly List<PptMaster> _progMasters = [];
    // 嵌入字体：文件名→字节数据（ppt/fonts/*.fntdata）
    private readonly Dictionary<String, Byte[]> _embeddedFonts = [];

    // 文档属性（docProps，S14-01/S14-02），从 PptDocument.Properties 委托  
    private PptDocumentProperties? _documentProperties;

    /// <summary>最小有效 PNG（1×1 黑色像素，67 字节），用作视频无缩略图时的占位</summary>
    private static readonly Byte[] MinimalPng =
    [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00,
        0x00, 0x00, 0x04, 0x00, 0x01, 0x27, 0x34, 0x27, 0x0A, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
        0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    #endregion

    #region 构造
    /// <summary>实例化写入器（默认 16:9 比例）</summary>
    public PptxWriter() { }

    /// <summary>从模板 pptx 文件实例化写入器，复用模板的母版/版式/主题（S04-Master）</summary>
    /// <param name="templatePath">模板 pptx 文件路径</param>
    public PptxWriter(String templatePath) => LoadMaster(templatePath);

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 幻灯片方法
    /// <summary>添加新幻灯片</summary>
    /// <param name="layoutIndex">版式索引（0起始），超出范围时自动修正到末尾版式</param>
    /// <returns>新幻灯片对象</returns>
    public PptSlide AddSlide(Int32 layoutIndex = 0)
    {
        var maxIdx = Math.Max(0, GetLayoutCount() - 1);
        var slide = new PptSlide { LayoutIndex = Math.Min(Math.Max(0, layoutIndex), maxIdx) };
        Slides.Add(slide);
        return slide;
    }

    /// <summary>向幻灯片添加文本框</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="text">文本内容</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <param name="fontSize">字号（磅）</param>
    /// <param name="bold">粗体</param>
    /// <returns>文本框对象</returns>
    public PptTextBox AddTextBox(Int32 slideIndex, String text, Double leftCm, Double topCm,
        Double widthCm, Double heightCm, Int32 fontSize = 18, Boolean bold = false)
    {
        var slide = EnsureSlide(slideIndex);
        var tb = new PptTextBox
        {
            Text = text,
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
            FontSize = fontSize,
            Bold = bold,
        };
        slide.TextBoxes.Add(tb);
        return tb;
    }

    /// <summary>向幻灯片添加图片</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="imageData">图片字节</param>
    /// <param name="extension">扩展名</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图片对象</returns>
    public PptImage AddImage(Int32 slideIndex, Byte[] imageData, String extension,
        Double leftCm, Double topCm, Double widthCm, Double heightCm)
    {
        var slide = EnsureSlide(slideIndex);
        var img = new PptImage
        {
            Data = imageData,
            Extension = extension.TrimStart('.').ToLowerInvariant(),
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
            RelId = $"rImg{_imgGlobal++}",
        };
        slide.Images.Add(img);
        return img;
    }

    /// <summary>替换幻灯片中已有图片的字节数据（S15-01）</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="imageIndex">图片在幻灯片中的索引（0起始）</param>
    /// <param name="newData">新图片字节</param>
    /// <param name="newExtension">新扩展名（如 "png"/"jpg"），null 表示保持原扩展名</param>
    public void ReplaceImage(Int32 slideIndex, Int32 imageIndex, Byte[] newData, String? newExtension = null)
    {
        var slide = EnsureSlide(slideIndex);
        if (imageIndex < 0 || imageIndex >= slide.Images.Count)
            throw new ArgumentOutOfRangeException(nameof(imageIndex), $"图片索引 {imageIndex} 超出范围（共 {slide.Images.Count} 张）");
        var img = slide.Images[imageIndex];
        img.Data = newData;
        if (newExtension != null && newExtension.Length > 0)
            img.Extension = newExtension.TrimStart('.').ToLowerInvariant();
    }

    /// <summary>向幻灯片添加视频/音频</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="mediaData">媒体字节</param>
    /// <param name="extension">扩展名</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>视频对象</returns>
    public PptVideo AddVideo(Int32 slideIndex, Byte[] mediaData, String extension,
        Double leftCm, Double topCm, Double widthCm, Double heightCm)
    {
        var slide = EnsureSlide(slideIndex);
        var vid = new PptVideo
        {
            Data = mediaData,
            Extension = extension.TrimStart('.').ToLowerInvariant(),
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
            RelId = $"rVid{_videoGlobal}",
            ThumbnailRelId = $"rVidThumb{_videoGlobal}",
        };
        _videoGlobal++;
        slide.Videos.Add(vid);
        return vid;
    }

    /// <summary>向幻灯片添加表格</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="rows">行列数据</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="firstRowHeader">首行表头</param>
    /// <returns>表格对象</returns>
    public PptTable AddTable(Int32 slideIndex, IEnumerable<String[]> rows,
        Double leftCm = 1, Double topCm = 2, Double widthCm = 22, Boolean firstRowHeader = true)
    {
        var slide = EnsureSlide(slideIndex);
        var tbl = new PptTable
        {
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            FirstRowHeader = firstRowHeader,
        };
        tbl.Rows.AddRange(rows);
        // 计算高度（每行约 0.8 cm）
        tbl.Height = CmToEmu(tbl.Rows.Count * 0.8 + 0.2);
        slide.Tables.Add(tbl);
        return tbl;
    }

    /// <summary>设置幻灯片背景色</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="colorHex">颜色（16进制 RGB，如 "1F497D"）</param>
    public void SetBackground(Int32 slideIndex, String colorHex)
        => EnsureSlide(slideIndex).BackgroundColor = colorHex;

    /// <summary>设置演讲者备注</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="notes">备注文本</param>
    public void SetNotes(Int32 slideIndex, String notes)
        => EnsureSlide(slideIndex).Notes = notes;

    /// <summary>设置幻灯片切换动画</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="type">切换类型（fade/push/wipe/zoom/split/cut）</param>
    /// <param name="durationMs">时长（毫秒）</param>
    public void SetTransition(Int32 slideIndex, String type = "fade", Int32 durationMs = 500)
        => EnsureSlide(slideIndex).Transition = new PptTransition { Type = type, DurationMs = durationMs };

    /// <summary>设置幻灯片尺寸（厘米）</summary>
    /// <param name="widthCm">宽度（厘米，默认 16:9 = 33.87cm）</param>
    /// <param name="heightCm">高度（厘米，默认 16:9 = 19.05cm）</param>
    public void SetSlideSize(Double widthCm, Double heightCm)
    {
        SlideWidth = CmToEmu(widthCm);
        SlideHeight = CmToEmu(heightCm);
    }

    /// <summary>移除幻灯片</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    public void RemoveSlide(Int32 slideIndex)
    {
        if (slideIndex >= 0 && slideIndex < Slides.Count)
            Slides.RemoveAt(slideIndex);
    }

    /// <summary>移动幻灯片位置</summary>
    /// <param name="fromIndex">源索引</param>
    /// <param name="toIndex">目标索引</param>
    public void MoveSlide(Int32 fromIndex, Int32 toIndex)
    {
        if (fromIndex < 0 || fromIndex >= Slides.Count) return;
        toIndex = Math.Max(0, Math.Min(toIndex, Slides.Count - 1));
        var slide = Slides[fromIndex];
        Slides.RemoveAt(fromIndex);
        Slides.Insert(toIndex, slide);
    }

    /// <summary>复制幻灯片（浅拷贝，图片/文本框引用相同）</summary>
    /// <param name="sourceIndex">源幻灯片索引</param>
    /// <returns>新幻灯片对象</returns>
    public PptSlide CloneSlide(Int32 sourceIndex)
    {
        var src = EnsureSlide(sourceIndex);
        var clone = new PptSlide { BackgroundColor = src.BackgroundColor, Notes = src.Notes };
        foreach (var tb in src.TextBoxes) clone.TextBoxes.Add(tb);
        foreach (var img in src.Images) clone.Images.Add(img);
        foreach (var tbl in src.Tables) clone.Tables.Add(tbl);
        foreach (var sp in src.Shapes) clone.Shapes.Add(sp);
        Slides.Add(clone);
        return clone;
    }

    /// <summary>向幻灯片添加基本图形</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="shapeType">几何类型（rect/ellipse/roundRect/triangle/diamond/arrow 等）</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <param name="fillColor">填充色（16进制 RGB）</param>
    /// <returns>图形对象</returns>
    public PptShape AddShape(Int32 slideIndex, String shapeType,
        Double leftCm, Double topCm, Double widthCm, Double heightCm,
        String? fillColor = null)
    {
        var slide = EnsureSlide(slideIndex);
        var sp = new PptShape
        {
            ShapeType = shapeType,
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
            FillColor = fillColor,
        };
        slide.Shapes.Add(sp);
        return sp;
    }

    /// <summary>克隆指定幻灯片上的形状，返回新形状对象</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="shapeIndex">形状索引</param>
    /// <returns>克隆后的形状</returns>
    public PptShape DuplicateShape(Int32 slideIndex, Int32 shapeIndex)
    {
        var slide = EnsureSlide(slideIndex);
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentOutOfRangeException(nameof(shapeIndex));

        var src = slide.Shapes[shapeIndex];
        var clone = new PptShape
        {
            Id = 0, // 写入时自动分配新 ID
            Text = src.Text,
            ShapeType = src.ShapeType,
            Left = src.Left,
            Top = src.Top + CmToEmu(1.0), // 默认向下偏移 1cm 以可见区别
            Width = src.Width,
            Height = src.Height,
            FillColor = src.FillColor,
            LineColor = src.LineColor,
            LineWidth = src.LineWidth,
            FontSize = src.FontSize,
            FontColor = src.FontColor,
            Bold = src.Bold,
            LatinFontName = src.LatinFontName,
            EastAsianFontName = src.EastAsianFontName,
            ComplexScriptFontName = src.ComplexScriptFontName,
            SymbolFontName = src.SymbolFontName,
            Rotation = src.Rotation,
            AltText = src.AltText,
            CornerRadius = src.CornerRadius,
        };
        slide.Shapes.Add(clone);
        return clone;
    }

    /// <summary>向幻灯片添加柱状图</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">分类轴标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddBarChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "bar", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加折线图</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">分类轴标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddLineChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "line", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加饼图</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">分类标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddPieChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "pie", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加散点图（S06 扩展）</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">X 轴分类标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddScatterChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "scatter", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加气泡图（S06 扩展）</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">X 轴分类标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddBubbleChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "bubble", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加雷达图</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="categories">轴分类标签</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <returns>图表对象</returns>
    public PptChart AddRadarChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "radar", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加股价图（K线/OHLC）</summary>
    public PptChart AddStockChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "stock", categories, leftCm, topCm, widthCm, heightCm);

    /// <summary>向幻灯片添加曲面图</summary>
    public PptChart AddSurfaceChart(Int32 slideIndex, String[] categories,
        Double leftCm = 2, Double topCm = 2, Double widthCm = 18, Double heightCm = 12)
        => AddChart(slideIndex, "surface", categories, leftCm, topCm, widthCm, heightCm);

    private PptChart AddChart(Int32 slideIndex, String chartType, String[] categories,
        Double leftCm, Double topCm, Double widthCm, Double heightCm)
    {
        var slide = EnsureSlide(slideIndex);
        var chartNum = _chartGlobal++;
        var chart = new PptChart
        {
            ChartType = chartType,
            Categories = categories,
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
            RelId = $"rChart{chartNum}",
            ChartNumber = chartNum,
        };
        slide.Charts.Add(chart);
        return chart;
    }

    /// <summary>将对象集合写入幻灯片表格</summary>
    /// <param name="slideIndex">幻灯片索引</param>
    /// <param name="data">对象集合</param>
    /// <param name="leftCm">左边距（厘米）</param>
    /// <param name="topCm">上边距（厘米）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    public void WriteObjects<T>(Int32 slideIndex, IEnumerable<T> data,
        Double leftCm = 1, Double topCm = 2, Double widthCm = 22) where T : class
    {
        var props = typeof(T).GetProperties();
        var headers = props.Select(p =>
        {
            var dn = p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                      .OfType<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault()?.DisplayName;
            return dn ?? p.Name;
        }).ToArray();
        var rows = new List<String[]> { headers };
        foreach (var item in data)
        {
            rows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());
        }
        AddTable(slideIndex, rows, leftCm, topCm, widthCm, firstRowHeader: true);
    }

    /// <summary>设置演示文稿修改密码保护（S07-04）</summary>
    /// <remarks>
    /// 设置后保存的 pptx 文件在 Word/PowerPoint 中打开时需要输入密码才能修改。
    /// 传入 null 可清除保护。基于 SHA-512 算法，符合 OOXML 标准。
    /// </remarks>
    /// <param name="password">修改密码，null 表示清除保护</param>
    public void SetProtection(String? password = null)
    {
        if (password == null) { _protectionHash = null; _protectionSalt = null; return; }
        var salt = new Byte[16];
        using (var rng = RandomNumberGenerator.Create())
            rng.GetBytes(salt);
        _protectionSalt = Convert.ToBase64String(salt);

        var pwd = Encoding.UTF8.GetBytes(password);
        using var sha = SHA512.Create();
        var buf = new Byte[salt.Length + pwd.Length];
        var bw = new SpanWriter(buf, 0, buf.Length);
        bw.Write(salt);
        bw.Write(pwd);
        var hash = sha.ComputeHash(buf);
        var iter = new Byte[hash.Length + 4]; // SHA-512 = 64 bytes，复用缓冲区
        for (var i = 0; i < 100000; i++)
        {
            var iw = new SpanWriter(iter, 0, iter.Length);
            iw.Write(hash);
            iw.Write(i);
            hash = sha.ComputeHash(iter);
        }
        _protectionHash = Convert.ToBase64String(hash);
    }

    /// <summary>设置演示文稿主题强调色（S07-03）</summary>
    /// <remarks>
    /// 修改主题的 6 个强调色（accent1~accent6），影响图表默认配色和内置主题样式。
    /// 所有幻灯片使用同一主题，修改后保存即生效。
    /// </remarks>
    /// <param name="hexColors">最多 6 个颜色（16进制 RGB，可带或不带 # 前缀）</param>
    /// <returns>自身，支持链式调用</returns>
    public PptxWriter SetAccentColors(params String[] hexColors)
    {
        for (var i = 0; i < Math.Min(hexColors.Length, AccentColors.Length); i++)
        {
            AccentColors[i] = hexColors[i].TrimStart('#');
        }
        return this;
    }

    /// <summary>向幻灯片添加形状组（S07-02）</summary>
    /// <remarks>
    /// 组内的形状和文本框将以 <c>&lt;p:grpSp&gt;</c> 元素组合，
    /// 可作为一个整体移动/缩放。返回 PptGroup 对象，向其 Shapes/TextBoxes 属性添加元素即可。
    /// </remarks>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="leftCm">组左边距（厘米）</param>
    /// <param name="topCm">组上边距（厘米）</param>
    /// <param name="widthCm">组宽度（厘米）</param>
    /// <param name="heightCm">组高度（厘米）</param>
    /// <returns>形状组对象</returns>
    public PptGroup GroupShapes(Int32 slideIndex, Double leftCm, Double topCm,
        Double widthCm, Double heightCm)
    {
        var slide = EnsureSlide(slideIndex);
        var group = new PptGroup
        {
            Left = CmToEmu(leftCm),
            Top = CmToEmu(topCm),
            Width = CmToEmu(widthCm),
            Height = CmToEmu(heightCm),
        };
        slide.Groups.Add(group);
        return group;
    }

    /// <summary>将指定形状移至最上层（BringToFront）</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="shapeIndex">形状在 Shapes 列表中的索引</param>
    public void BringToFront(Int32 slideIndex, Int32 shapeIndex)
    {
        var slide = EnsureSlide(slideIndex);
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count) return;
        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.RemoveAt(shapeIndex);
        slide.Shapes.Add(shape);
    }

    /// <summary>将指定形状移至最底层（SendToBack）</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="shapeIndex">形状在 Shapes 列表中的索引</param>
    public void SendToBack(Int32 slideIndex, Int32 shapeIndex)
    {
        var slide = EnsureSlide(slideIndex);
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count) return;
        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.RemoveAt(shapeIndex);
        slide.Shapes.Insert(0, shape);
    }

    /// <summary>将指定形状上移一层（BringForward）</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="shapeIndex">形状在 Shapes 列表中的索引</param>
    public void BringForward(Int32 slideIndex, Int32 shapeIndex)
    {
        var slide = EnsureSlide(slideIndex);
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count - 1) return;
        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.RemoveAt(shapeIndex);
        slide.Shapes.Insert(shapeIndex + 1, shape);
    }

    /// <summary>将指定形状下移一层（SendBackward）</summary>
    /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="shapeIndex">形状在 Shapes 列表中的索引</param>
    public void SendBackward(Int32 slideIndex, Int32 shapeIndex)
    {
        var slide = EnsureSlide(slideIndex);
        if (shapeIndex < 1 || shapeIndex >= slide.Shapes.Count) return;
        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.RemoveAt(shapeIndex);
        slide.Shapes.Insert(shapeIndex - 1, shape);
    }

    /// <summary>为幻灯片添加页脚文本和/或页码（S04-05）</summary>
    /// /// <param name="slideIndex">幻灯片索引（0起始）</param>
    /// <param name="footerText">页脚文本，null 表示不显示</param>
    /// <param name="showSlideNumber">是否在右下角显示幻灯片序号</param>
    public void SetSlideFooter(Int32 slideIndex, String? footerText = null, Boolean showSlideNumber = false)
    {
        if (slideIndex < 0 || slideIndex >= Slides.Count)
            throw new ArgumentOutOfRangeException(nameof(slideIndex));
        var slide = Slides[slideIndex];

        if (footerText != null)
            slide.TextBoxes.Add(new PptTextBox
            {
                Text = footerText,
                Left = CmToEmu(1.3),
                Top = SlideHeight - CmToEmu(1.3),
                Width = CmToEmu(10),
                Height = CmToEmu(1),
                FontSize = 11,
                FontColor = "808080",
            });

        if (showSlideNumber)
            slide.TextBoxes.Add(new PptTextBox
            {
                Text = (slideIndex + 1).ToString(),
                Left = SlideWidth - CmToEmu(2),
                Top = SlideHeight - CmToEmu(1.3),
                Width = CmToEmu(1.5),
                Height = CmToEmu(1),
                FontSize = 11,
                FontColor = "808080",
                Alignment = "r",
            });
    }

    /// <summary>从另一个 pptx 文件复制单张幻灯片（S10-04）</summary>
    /// <remarks>
    /// 在 ZIP 层面直接复制幻灯片 XML 及其引用的媒体文件，并重命名以避免冲突。
    /// 复制的幻灯片追加在所有普通幻灯片之后，调用 Save 时一并写出。
    /// </remarks>
    /// <param name="sourcePath">源 pptx 文件路径</param>
    /// <param name="slideIndex">源文件中的幻灯片索引（0 起始）</param>
    /// <returns>新幻灯片在目标文档中的索引（0 起始）</returns>
    public Int32 CopySlideFrom(String sourcePath, Int32 slideIndex)
        => CopySlideFrom(File.ReadAllBytes(sourcePath.GetFullPath()), slideIndex);

    /// <summary>从另一个 pptx 字节数数据复制单张幻灯片（S10-04）</summary>
    /// <param name="sourceData">源 pptx 字节数据</param>
    /// <param name="slideIndex">源文件中的幻灯片索引（0 起始）</param>
    /// <returns>新幻灯片在目标文档中的索引（0 起始）</returns>
    public Int32 CopySlideFrom(Byte[] sourceData, Int32 slideIndex)
    {
        using var ms = new MemoryStream(sourceData);
        using var srcZip = new ZipArchive(ms, ZipArchiveMode.Read);
        var srcSlideNum = slideIndex + 1;
        var slideEntry = srcZip.GetEntry($"ppt/slides/slide{srcSlideNum}.xml")
            ?? throw new ArgumentOutOfRangeException(nameof(slideIndex), $"源文件中不存在第 {slideIndex} 张幻灯片");
        String slideXml;
        using (var sr = new StreamReader(slideEntry.Open())) slideXml = sr.ReadToEnd();
        String relsXml;
        var relsEntry = srcZip.GetEntry($"ppt/slides/_rels/slide{srcSlideNum}.xml.rels");
        if (relsEntry != null)
        {
            using var sr = new StreamReader(relsEntry.Open());
            relsXml = sr.ReadToEnd();
        }
        else
        {
            relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>";
        }
        // 复制媒体文件并重命名
        foreach (var entry in srcZip.Entries)
        {
            if (!entry.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)) continue;
            var baseName = entry.Name;
            if (relsXml.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) < 0
                && slideXml.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) < 0) continue;
            var ext = Path.GetExtension(baseName);
            var newName = $"m{_mediaGlobal++}{ext}";
            relsXml = relsXml.Replace($"../media/{baseName}", $"../media/{newName}");
            slideXml = slideXml.Replace($"../media/{baseName}", $"../media/{newName}");
            var buf = new MemoryStream();
            using (var es = entry.Open()) es.CopyTo(buf);
            _rawSlideMedia.Add((newName, buf.ToArray()));
        }
        _rawSlides.Add((slideXml, relsXml));
        return Slides.Count + _rawSlides.Count - 1;
    }

    /// <summary>修改 pptx 文件中指定图表的系列数据（S10-03）</summary>
    /// <remarks>
    /// 直接替换图表 XML 中所有 c:numCache 的数值缓存，不修改内嵌 Excel。
    /// 如果系列/数据点数量与原图表不匹配，按最小公集处理。
    /// </remarks>
    /// <param name="sourcePath">源 pptx 文件路径</param>
    /// <param name="chartNumber">图表编号（ppt/charts/chart{N}.xml 的 N，从 1 开始）</param>
    /// <param name="series">新系列数据</param>
    /// <param name="outputPath">输出路径，null 时覆盖源文件</param>
    public static void UpdateChartData(String sourcePath, Int32 chartNumber, IEnumerable<PptChartSeries> series, String? outputPath = null)
    {
        var data = File.ReadAllBytes(sourcePath.GetFullPath());
        var result = UpdateChartData(data, chartNumber, series);
        var dst = (outputPath ?? sourcePath).GetFullPath();
        File.WriteAllBytes(dst, result);
    }

    /// <summary>修改 pptx 字节中指定图表的系列数据并返回新的字节数据（S10-03）</summary>
    /// <param name="pptxData">源 pptx 字节数据</param>
    /// <param name="chartNumber">图表编号（1 起始）</param>
    /// <param name="series">新系列数据</param>
    /// <returns>更新后的 pptx 字节数据</returns>
    public static Byte[] UpdateChartData(Byte[] pptxData, Int32 chartNumber, IEnumerable<PptChartSeries> series)
    {
        var serList = series.ToList();
        var chartPath = $"ppt/charts/chart{chartNumber}.xml";
        using var srcMs = new MemoryStream(pptxData);
        using var dstMs = new MemoryStream();
        using (var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read))
        using (var dstZip = new ZipArchive(dstMs, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in srcZip.Entries)
            {
                if (!entry.FullName.Equals(chartPath, StringComparison.OrdinalIgnoreCase))
                {
                    // Copy as-is
                    var dst = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                    using var ss = entry.Open();
                    using var ds = dst.Open();
                    ss.CopyTo(ds);
                    continue;
                }
                // Rewrite chart XML with new series data
                String chartXml;
                using (var sr = new StreamReader(entry.Open())) chartXml = sr.ReadToEnd();
                chartXml = PatchChartXml(chartXml, serList);
                WriteZipEntryText(dstZip, entry.FullName, chartXml);
            }
        }
        return dstMs.ToArray();
    }

    /// <summary>更新图表 XML 中每个系列的 numCache 数值</summary>
    /// <param name="xml">原始图表 XML 字符串</param>
    /// <param name="series">新系列数据列表</param>
    /// <returns>更新后的 XML 字符串</returns>
    private static String PatchChartXml(String xml, List<PptChartSeries> series)
    {
        var doc = new System.Xml.XmlDocument();
        doc.LoadXml(xml);
        const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var ns = new System.Xml.XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("c", C);
        var serNodes = doc.SelectNodes("//c:ser", ns);
        if (serNodes == null) return xml;
        for (var si = 0; si < serNodes.Count && si < series.Count; si++)
        {
            var ser = series[si];
            var serNode = serNodes[si]!;
            // Update series name
            var txV = serNode.SelectSingleNode(".//c:tx//c:v", ns);
            if (txV != null) txV.InnerText = ser.Name;
            // Update numCache
            var numCache = serNode.SelectSingleNode(".//c:val//c:numCache", ns);
            if (numCache == null) continue;
            var ptCountNode = numCache.SelectSingleNode("c:ptCount", ns);
            if (ptCountNode != null)
                ((System.Xml.XmlElement)ptCountNode).SetAttribute("val", ser.Values.Length.ToString());
            // Remove old pt nodes
            foreach (var old in numCache.SelectNodes("c:pt", ns)!.Cast<System.Xml.XmlNode>().ToList())
            {
                numCache.RemoveChild(old);
            }
            // Add new pt nodes
            for (var vi = 0; vi < ser.Values.Length; vi++)
            {
                var pt = doc.CreateElement("c:pt", C);
                pt.SetAttribute("idx", vi.ToString());
                var v = doc.CreateElement("c:v", C);
                v.InnerText = ser.Values[vi].ToString(CultureInfo.InvariantCulture);
                pt.AppendChild(v);
                numCache.AppendChild(pt);
            }
        }
        var sb = new StringBuilder();
        using var sw = new System.IO.StringWriter(sb);
        doc.Save(sw);
        return sb.ToString();
    }

    /// <summary>合并多个 pptx 文件为一个（S05-02）</summary>
    /// <param name="sourcePaths">源文件路径集合</param>
    /// <param name="outputPath">输出文件路径</param>
    public static void Merge(IEnumerable<String> sourcePaths, String outputPath)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write);
        Merge(sourcePaths.Select(p => File.ReadAllBytes(p.GetFullPath())), fs);
    }

    /// <summary>合并多个 pptx 字节数组为一个，写入流（S05-02）</summary>
    /// <param name="sourceDatas">源 pptx 字节数组集合</param>
    /// <param name="outputStream">输出流</param>
    public static void Merge(IEnumerable<Byte[]> sourceDatas, Stream outputStream)
    {
        var sources = sourceDatas.ToList();
        if (sources.Count == 0) return;
        if (sources.Count == 1) { outputStream.Write(sources[0], 0, sources[0].Length); return; }

        using var dstZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
        var slideTotal = 0;
        var mediaTotal = 0;

        for (var fi = 0; fi < sources.Count; fi++)
        {
            using var srcMs = new MemoryStream(sources[fi]);
            using var srcZip = new ZipArchive(srcMs, ZipArchiveMode.Read);
            var mediaRename = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase); // oldFilename -> newFilename

            // Copy infrastructure entries from the first file only
            if (fi == 0)
            {
                foreach (var entry in srcZip.Entries)
                {
                    var n = entry.FullName;
                    if (n.StartsWith("ppt/slides/") || n.StartsWith("ppt/media/")
                        || n == "ppt/presentation.xml" || n == "ppt/_rels/presentation.xml.rels"
                        || n == "[Content_Types].xml") continue;

                    var dst = dstZip.CreateEntry(n, CompressionLevel.Fastest);
                    using var ss = entry.Open();
                    using var ds = dst.Open();
                    ss.CopyTo(ds);
                }
            }

            // Copy media files with sequential renaming for uniqueness
            foreach (var entry in srcZip.Entries.Where(e => e.FullName.StartsWith("ppt/media/")))
            {
                mediaTotal++;
                var ext = Path.GetExtension(entry.Name);
                var newFilename = $"m{mediaTotal}{ext}";
                mediaRename[entry.Name] = newFilename;
                var dst = dstZip.CreateEntry($"ppt/media/{newFilename}", CompressionLevel.Fastest);
                using var ss = entry.Open();
                using var ds = dst.Open();
                ss.CopyTo(ds);
            }

            // Copy slides with renamed IDs and updated media refs
            var slideEntries = srcZip.Entries
                .Where(e => System.Text.RegularExpressions.Regex.IsMatch(e.FullName, @"^ppt/slides/slide\d+\.xml$"))
                .OrderBy(e =>
                {
                    var m = System.Text.RegularExpressions.Regex.Match(e.FullName, @"slide(\d+)\.xml");
                    return m.Success ? Int32.Parse(m.Groups[1].Value) : 0;
                })
                .ToList();

            foreach (var slideEntry in slideEntries)
            {
                slideTotal++;
                var oldNum = Int32.Parse(System.Text.RegularExpressions.Regex.Match(slideEntry.FullName, @"slide(\d+)\.xml").Groups[1].Value);

                String slideXml;
                using (var sr = new StreamReader(slideEntry.Open())) slideXml = sr.ReadToEnd();
                foreach (var kv in mediaRename)
                {
                    slideXml = slideXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
                }
                WriteZipEntryText(dstZip, $"ppt/slides/slide{slideTotal}.xml", slideXml);

                var relsEntry = srcZip.GetEntry($"ppt/slides/_rels/slide{oldNum}.xml.rels");
                String relsXml;
                if (relsEntry != null)
                {
                    using var sr = new StreamReader(relsEntry.Open()); relsXml = sr.ReadToEnd();
                    foreach (var kv in mediaRename)
                    {
                        relsXml = relsXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
                    }
                }
                else
                {
                    relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                        "</Relationships>";
                }
                WriteZipEntryText(dstZip, $"ppt/slides/_rels/slide{slideTotal}.xml.rels", relsXml);
            }
        }

        // Write merged presentation.xml
        var presSb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        presSb.Append("<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"");
        presSb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        presSb.Append("<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rMaster1\"/></p:sldMasterIdLst>");
        presSb.Append("<p:sldIdLst>");
        for (var i = 0; i < slideTotal; i++)
        {
            presSb.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rSlide{i + 1}\"/>");
        }
        presSb.Append("</p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>");
        WriteZipEntryText(dstZip, "ppt/presentation.xml", presSb.ToString());

        // Write merged ppt/_rels/presentation.xml.rels
        var relsSb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (var i = 1; i <= slideTotal; i++)
        {
            relsSb.Append($"<Relationship Id=\"rSlide{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i}.xml\"/>");
        }
        relsSb.Append("<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>");
        relsSb.Append("</Relationships>");
        WriteZipEntryText(dstZip, "ppt/_rels/presentation.xml.rels", relsSb.ToString());

        // Write [Content_Types].xml
        var ctSb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        ctSb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        ctSb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        ctSb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        ctSb.Append("<Default Extension=\"png\" ContentType=\"image/png\"/>");
        ctSb.Append("<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>");
        ctSb.Append("<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>");
        ctSb.Append("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
        for (var i = 1; i <= slideTotal; i++)
        {
            ctSb.Append($"<Override PartName=\"/ppt/slides/slide{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        }
        ctSb.Append("<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>");
        ctSb.Append("<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
        ctSb.Append("<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        ctSb.Append("</Types>");
        WriteZipEntryText(dstZip, "[Content_Types].xml", ctSb.ToString());
    }
    #endregion

    #region 母版

    /// <summary>从现有 pptx 文件加载母版、版式和主题（S04-Master）</summary>
    /// <remarks>
    /// 以原始 XML 形式加载母版/版式/主题，保存时保留完整品牌设计。
    /// 调用此方法会清除已加载的母版/版式内容，但不影响已添加的幻灯片。
    /// 注意：加载模板后 SetAccentColors 对主题颜色不再生效，颜色由模板主题决定。
    /// </remarks>
    /// <param name="templatePath">模板 pptx 文件路径</param>
    public void LoadMaster(String templatePath)
        => LoadMaster(File.ReadAllBytes(templatePath.GetFullPath()), keepTemplateSlides: false);

    /// <summary>从现有 pptx 文件加载母版、版式和主题，可选保留模板原有幻灯片</summary>
    /// <param name="templatePath">模板 pptx 文件路径</param>
    /// <param name="keepTemplateSlides">true 时保留模板原有幻灯片（追加在程序化幻灯片之后）</param>
    public void LoadMaster(String templatePath, Boolean keepTemplateSlides)
        => LoadMaster(File.ReadAllBytes(templatePath.GetFullPath()), keepTemplateSlides);

    /// <summary>从 pptx 字节数据加载母版、版式和主题（S04-Master）</summary>
    /// <remarks>
    /// 可与 <see cref="PptxWriter(String)"/> 构造函数等效，也可在已有 PptxWriter 实例上调用。
    /// </remarks>
    /// <param name="templateBytes">模板 pptx 字节数据</param>
    public void LoadMaster(Byte[] templateBytes) => LoadMaster(templateBytes, keepTemplateSlides: false);

    /// <summary>从 pptx 字节数据加载母版、版式和主题，可选保留模板原有幻灯片</summary>
    /// <param name="templateBytes">模板 pptx 字节数据</param>
    /// <param name="keepTemplateSlides">true 时保留模板原有幻灯片</param>
    public void LoadMaster(Byte[] templateBytes, Boolean keepTemplateSlides)
    {
        _masterContents.Clear();
        _layoutContents.Clear();
        _layoutNames.Clear();
        _infraMedia.Clear();
        _templateThemeXml = null;

        using var ms = new MemoryStream(templateBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        // 加载母版（按文件名排序）
        var masterEntries = zip.Entries
            .Where(e => e.FullName.StartsWith("ppt/slideMasters/", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && !e.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(e => e.FullName)
            .ToList();

        foreach (var entry in masterEntries)
        {
            var content = new PptPartContent();
            using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                content.Xml = sr.ReadToEnd();
            var relsEntry = zip.GetEntry(GetRelsEntryPath(entry.FullName));
            if (relsEntry != null)
            {
                using var rsr = new StreamReader(relsEntry.Open(), Encoding.UTF8);
                content.RelsXml = rsr.ReadToEnd();
            }
            _masterContents.Add(content);
        }

        // 加载版式（按文件名排序）
        var layoutEntries = zip.Entries
            .Where(e => e.FullName.StartsWith("ppt/slideLayouts/", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && !e.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(e => e.FullName)
            .ToList();

        foreach (var entry in layoutEntries)
        {
            var content = new PptPartContent();
            using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                content.Xml = sr.ReadToEnd();
            var relsEntry = zip.GetEntry(GetRelsEntryPath(entry.FullName));
            if (relsEntry != null)
            {
                using var rsr = new StreamReader(relsEntry.Open(), Encoding.UTF8);
                content.RelsXml = rsr.ReadToEnd();
            }
            _layoutNames.Add(ExtractLayoutDisplayName(content.Xml));
            _layoutContents.Add(content);
        }

        // 加载主题（优先 theme1.xml，否则取第一个）
        var themeEntry = zip.GetEntry("ppt/theme/theme1.xml")
            ?? zip.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase)
                && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        if (themeEntry != null)
        {
            using var sr = new StreamReader(themeEntry.Open(), Encoding.UTF8);
            _templateThemeXml = sr.ReadToEnd();
        }

        // 加载母版/版式 rels 引用的媒体文件
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)) continue;
            var fileName = entry.Name;
            var isReferenced = _masterContents.Any(c => c.RelsXml.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0)
                            || _layoutContents.Any(c => c.RelsXml.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0);
            if (!isReferenced) continue;
            using var buf = new MemoryStream();
            using (var es = entry.Open()) es.CopyTo(buf);
            _infraMedia[fileName] = buf.ToArray();
        }

        // 加载嵌入字体（ppt/fonts/*.fntdata）
        _embeddedFonts.Clear();
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("ppt/fonts/", StringComparison.OrdinalIgnoreCase)) continue;
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue; // 跳过目录条目
            using var buf = new MemoryStream();
            using (var es = entry.Open()) es.CopyTo(buf);
            _embeddedFonts[entry.Name] = buf.ToArray();
        }

        // 可选：保留模板原有幻灯片
        if (keepTemplateSlides)
        {
            CopyTemplateSlides(zip);
        }
    }

    /// <summary>将模板中的幻灯片复制到 _rawSlides（供 keepTemplateSlides 使用）</summary>
    private void CopyTemplateSlides(ZipArchive zip)
    {
        // 收集模板中所有媒体文件的重命名映射
        var mediaRename = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)) continue;
            var newName = $"mt{_mediaGlobal++}{Path.GetExtension(entry.Name)}";
            mediaRename[entry.Name] = newName;
            using var buf = new MemoryStream();
            using (var es = entry.Open()) es.CopyTo(buf);
            _rawSlideMedia.Add((newName, buf.ToArray()));
        }

        // 按幻灯片编号顺序复制
        var slideEntries = zip.Entries
            .Where(e => System.Text.RegularExpressions.Regex.IsMatch(e.FullName, @"^ppt/slides/slide\d+\.xml$"))
            .OrderBy(e =>
            {
                var m = System.Text.RegularExpressions.Regex.Match(e.FullName, @"slide(\d+)\.xml");
                return m.Success ? Int32.Parse(m.Groups[1].Value) : 0;
            })
            .ToList();

        foreach (var slideEntry in slideEntries)
        {
            var slideNum = Int32.Parse(
                System.Text.RegularExpressions.Regex.Match(slideEntry.FullName, @"slide(\d+)\.xml").Groups[1].Value);

            String slideXml;
            using (var sr = new StreamReader(slideEntry.Open())) slideXml = sr.ReadToEnd();
            foreach (var kv in mediaRename)
            {
                slideXml = slideXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
            }

            String relsXml;
            var relsEntry = zip.GetEntry($"ppt/slides/_rels/slide{slideNum}.xml.rels");
            if (relsEntry != null)
            {
                using var sr = new StreamReader(relsEntry.Open()); relsXml = sr.ReadToEnd();
                foreach (var kv in mediaRename)
                {
                    relsXml = relsXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
                }
            }
            else
            {
                relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                    "</Relationships>";
            }

            _rawSlides.Add((slideXml, relsXml));
        }
    }

    /// <summary>获取已加载的版式数量</summary>
    /// <returns>版式数量（未加载模板且未编程创建时为 1）</returns>
    public Int32 GetLayoutCount()
    {
        if (_layoutContents.Count > 0) return _layoutContents.Count;
        var progCount = _progMasters.Sum(m => m.Layouts.Count);
        return Math.Max(1, progCount);
    }

    /// <summary>获取指定版式的显示名称</summary>
    /// <param name="layoutIndex">版式索引（0起始）</param>
    /// <returns>版式名称；超出范围返回空字符串</returns>
    public String GetLayoutName(Int32 layoutIndex)
    {
        if (_layoutNames.Count > 0)
        {
            if (layoutIndex < 0 || layoutIndex >= _layoutNames.Count) return String.Empty;
            return _layoutNames[layoutIndex];
        }
        var total = 0;
        foreach (var m in _progMasters)
        {
            foreach (var l in m.Layouts)
            {
                if (total == layoutIndex) return l.Name;
                total++;
            }
        }
        return total == 0 && layoutIndex == 0 ? "blank" : String.Empty;
    }

    /// <summary>编程式创建一个新母版（Phase 5：无需外部模板文件）</summary>
    /// <remarks>
    /// 创建后可向母版添加形状（如公司 Logo），并通过 <see cref="PptMaster.AddLayout"/> 添加自定义版式。
    /// 保存时自动生成符合 OOXML 规范的母版和版式 XML。
    /// <para>注意：调用此方法会与 <see cref="LoadMaster(String)"/> 互斥——后者通过模板文件加载，前者纯编程构建。</para>
    /// </remarks>
    /// <returns>新创建的母版对象</returns>
    public PptMaster CreateMaster()
    {
        var master = new PptMaster();
        _progMasters.Add(master);
        return master;
    }

    /// <summary>从另一个 pptx 文件复制母版、版式和主题（S04-Master）</summary>
    /// <remarks>语义等同于 <see cref="LoadMaster(String)"/>，命名更明确地表达跨文件母版复制的意图。</remarks>
    /// <param name="sourcePath">源 pptx 文件路径</param>
    public void CopyMasterFrom(String sourcePath) => LoadMaster(sourcePath);

    #endregion

}
