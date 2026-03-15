#nullable enable
using System.IO.Compression;
using System.Text;

namespace NewLife.Office;

/// <summary>PPT 富文本片段（S10-01）</summary>
/// <remarks>
/// 支持每个片段独立设置字拉、复体、斜体、颜色、超链接。
/// 将多个 <see cref="PptTextRun"/> 添加到 <see cref="PptTextBox.Runs"/> 即可实现富文本效果。
/// </remarks>
public class PptTextRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字拉7（磅），0 表示继承文本框默认字拉</summary>
    public Int32 FontSize { get; set; }

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>文字颜色（16进制 RGB），null 表示继承文本框设置</summary>
    public String? FontColor { get; set; }

    /// <summary>超链接 URL，不为 null 时点击该片段跳转</summary>
    public String? HyperlinkUrl { get; set; }
    #endregion
}

/// <summary>PPT 表格单元格样式（S10-02）</summary>
public class PptCellStyle
{
    #region 属性
    /// <summary>单元格背景色（16进制 RGB），null 表示表格默认色</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>字体颜色（16进制 RGB），null 表示继承</summary>
    public String? FontColor { get; set; }

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>字拉（磅），0 表示继承</summary>
    public Int32 FontSize { get; set; }
    #endregion
}

/// <summary>PPT 幻灯片文本框</summary>
public class PptTextBox
{
    #region 属性
    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字号（磅）</summary>
    public Int32 FontSize { get; set; } = 18;

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>文字颜色（16进制 RGB，如 "000000"）</summary>
    public String? FontColor { get; set; }

    /// <summary>对齐（l/ctr/r）</summary>
    public String Alignment { get; set; } = "l";

    /// <summary>背景色（16进制 RGB），null 表示透明</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>超链接 URL，不为 null 时点击文字跳转</summary>
    public String? HyperlinkUrl { get; set; }

    /// <summary>富文本片段集合；非空时优先使用，忽略 Text/FontSize/Bold/FontColor 等单一格式属性</summary>
    public List<PptTextRun> Runs { get; } = [];
    #endregion
}

/// <summary>幻灯片切换动画</summary>
public class PptTransition
{
    #region 属性
    /// <summary>切换类型（fade/push/wipe/zoom/split/cut）</summary>
    public String Type { get; set; } = "fade";

    /// <summary>切换时长（毫秒）</summary>
    public Int32 DurationMs { get; set; } = 500;

    /// <summary>切换方向（l/r/u/d，部分类型使用）</summary>
    public String Direction { get; set; } = "l";

    /// <summary>是否单击时自动切换</summary>
    public Boolean AdvanceOnClick { get; set; } = true;
    #endregion
}

/// <summary>PPT 图表系列数据</summary>
public class PptChartSeries
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据点值</summary>
    public Double[] Values { get; set; } = [];
    #endregion
}

/// <summary>PPT 嵌入图表</summary>
public class PptChart
{
    #region 属性
    /// <summary>图表类型（bar/line/pie/area/scatter）</summary>
    public String ChartType { get; set; } = "bar";

    /// <summary>图表标题，null 表示不显示</summary>
    public String? Title { get; set; }

    /// <summary>分类轴标签</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>系列集合</summary>
    public List<PptChartSeries> Series { get; } = [];

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 6000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 4000000;

    /// <summary>图表关系ID（由写入器内部设置）</summary>
    public String RelId { get; set; } = String.Empty;

    /// <summary>图表文件编号（由写入器内部设置）</summary>
    internal Int32 ChartNumber { get; set; }
    #endregion
}

/// <summary>PPT 幻灯片图片元素</summary>
public class PptImage
{
    #region 属性
    /// <summary>图片字节数据</summary>
    public Byte[] Data { get; set; } = [];

    /// <summary>扩展名（png/jpg）</summary>
    public String Extension { get; set; } = "png";

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 3000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 2000000;

    /// <summary>关系ID（内部用）</summary>
    public String RelId { get; set; } = String.Empty;
    #endregion
}

/// <summary>PPT 幻灯片表格</summary>
public class PptTable
{
    #region 属性
    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 8000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 3000000;

    /// <summary>行列数据</summary>
    public List<String[]> Rows { get; } = [];

    /// <summary>首行是否表头</summary>
    public Boolean FirstRowHeader { get; set; } = true;

    /// <summary>各列宽度（EMU），数组长度等于列数；空时按总宽平均分配</summary>
    public Int64[] ColWidths { get; set; } = [];

    /// <summary>单元格样式字典，键为 (行索引, 列索引)，优先级高于行级默认样式</summary>
    public Dictionary<(Int32 Row, Int32 Col), PptCellStyle> CellStyles { get; } = [];
    #endregion
}

/// <summary>PPT 幻灯片</summary>
public class PptSlide
{
    #region 属性
    /// <summary>文本框集合</summary>
    public List<PptTextBox> TextBoxes { get; } = [];

    /// <summary>图片集合</summary>
    public List<PptImage> Images { get; } = [];

    /// <summary>表格集合</summary>
    public List<PptTable> Tables { get; } = [];

    /// <summary>基本图形集合</summary>
    public List<PptShape> Shapes { get; } = [];

    /// <summary>图表集合</summary>
    public List<PptChart> Charts { get; } = [];

    /// <summary>背景色（16进制 RGB），null 表示白色</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>演讲者备注</summary>
    public String? Notes { get; set; }

    /// <summary>幻灯片切换动画，null 表示不设置</summary>
    public PptTransition? Transition { get; set; }

    /// <summary>图片关系计数器</summary>
    internal Int32 ImageCounter { get; set; } = 1;

    /// <summary>形状组集合（S07-02 组合形状）</summary>
    public List<PptGroup> Groups { get; } = [];
    #endregion
}

/// <summary>PPT 形状组（S07-02）</summary>
/// <remarks>将多个形状组合为一个组，使用 <c>&lt;p:grpSp&gt;</c> 元素生成。</remarks>
public class PptGroup
{
    #region 属性
    /// <summary>组左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>组上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>组宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>组高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>组内形状</summary>
    public List<PptShape> Shapes { get; } = [];

    /// <summary>组内文本框</summary>
    public List<PptTextBox> TextBoxes { get; } = [];
    #endregion
}

/// <summary>PowerPoint pptx 写入器</summary>
/// <remarks>
/// 直接操作 Open XML（ZIP+XML）生成 .pptx 文件。
/// 支持文本框/表格/图片/背景/备注等核心功能。
/// 坐标使用 EMU（英制单位，914400 EMU = 1 英寸，360000 EMU = 1 cm）。
/// </remarks>
public class PptxWriter : IDisposable
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
    #endregion

    #region 私有字段
    private Int32 _imgGlobal = 1;
    private Int32 _chartGlobal = 1;
    private Int32 _hlinkGlobal = 1;
    private Int32 _mediaGlobal = 1;
    private String? _protectionHash;
    private String? _protectionSalt;
    // 跨文件复制的原始幻灯片（S10-04）：(幻灯片XML, rels XML)
    private readonly List<(String SlideXml, String RelsXml)> _rawSlides = [];
    // 跨文件复制的媒体文件：(文件名, 字节数据)
    private readonly List<(String Name, Byte[] Data)> _rawSlideMedia = [];
    #endregion

    #region 构造
    /// <summary>实例化写入器（默认 16:9 比例）</summary>
    public PptxWriter() { }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 幻灯片方法
    /// <summary>添加新幻灯片</summary>
    /// <returns>新幻灯片对象</returns>
    public PptSlide AddSlide()
    {
        var slide = new PptSlide();
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
            rows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());
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
        using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
            rng.GetBytes(salt);
        _protectionSalt = Convert.ToBase64String(salt);

        var pwd = Encoding.UTF8.GetBytes(password);
        using var sha = System.Security.Cryptography.SHA512.Create();
        var buf = new Byte[salt.Length + pwd.Length];
        salt.CopyTo(buf, 0);
        pwd.CopyTo(buf, salt.Length);
        var hash = sha.ComputeHash(buf);
        for (var i = 0; i < 100000; i++)
        {
            var iter = new Byte[hash.Length + 4];
            hash.CopyTo(iter, 0);
            iter[hash.Length] = (Byte)(i & 0xFF);
            iter[hash.Length + 1] = (Byte)((i >> 8) & 0xFF);
            iter[hash.Length + 2] = (Byte)((i >> 16) & 0xFF);
            iter[hash.Length + 3] = (Byte)((i >> 24) & 0xFF);
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
            AccentColors[i] = hexColors[i].TrimStart('#');
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

    /// <summary>为幻灯片添加页脚文本和/或页码（S04-05）</summary>    /// <param name="slideIndex">幻灯片索引（0起始）</param>
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
                numCache.RemoveChild(old);
            // Add new pt nodes
            for (var vi = 0; vi < ser.Values.Length; vi++)
            {
                var pt = doc.CreateElement("c:pt", C);
                ((System.Xml.XmlElement)pt).SetAttribute("idx", vi.ToString());
                var v = doc.CreateElement("c:v", C);
                v.InnerText = ser.Values[vi].ToString(System.Globalization.CultureInfo.InvariantCulture);
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
                    slideXml = slideXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
                WriteZipEntryText(dstZip, $"ppt/slides/slide{slideTotal}.xml", slideXml);

                var relsEntry = srcZip.GetEntry($"ppt/slides/_rels/slide{oldNum}.xml.rels");
                String relsXml;
                if (relsEntry != null)
                {
                    using var sr = new StreamReader(relsEntry.Open()); relsXml = sr.ReadToEnd();
                    foreach (var kv in mediaRename)
                        relsXml = relsXml.Replace($"../media/{kv.Key}", $"../media/{kv.Value}");
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
            presSb.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rSlide{i + 1}\"/>");
        presSb.Append("</p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>");
        WriteZipEntryText(dstZip, "ppt/presentation.xml", presSb.ToString());

        // Write merged ppt/_rels/presentation.xml.rels
        var relsSb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (var i = 1; i <= slideTotal; i++)
            relsSb.Append($"<Relationship Id=\"rSlide{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i}.xml\"/>");
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
            ctSb.Append($"<Override PartName=\"/ppt/slides/slide{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        ctSb.Append("<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>");
        ctSb.Append("<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
        ctSb.Append("<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        ctSb.Append("</Types>");
        WriteZipEntryText(dstZip, "[Content_Types].xml", ctSb.ToString());
    }
    #endregion

    #region 保存方法
    /// <summary>保存到文件</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        using var za = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteContentTypes(za);
        WriteRels(za);
        WritePresentation(za);
        WritePresentationRels(za);
        WriteSlideLayout(za);
        WriteSlideMaster(za);
        for (var i = 0; i < Slides.Count; i++)
            WriteSlide(za, i, Slides[i]);
        // 写入跨文件复制的原始幻灯片（S10-04）
        var totalSlides = Slides.Count;
        for (var ri = 0; ri < _rawSlides.Count; ri++)
        {
            var rawIdx = totalSlides + ri;
            WriteZipEntryText(za, $"ppt/slides/slide{rawIdx + 1}.xml", _rawSlides[ri].SlideXml);
            WriteZipEntryText(za, $"ppt/slides/_rels/slide{rawIdx + 1}.xml.rels", _rawSlides[ri].RelsXml);
        }
        // 写入原始幻灯片的媒体文件
        foreach (var (name, data) in _rawSlideMedia)
        {
            var entry = za.CreateEntry($"ppt/media/{name}", CompressionLevel.Fastest);
            using var es = entry.Open();
            es.Write(data, 0, data.Length);
        }
        WriteTheme(za);
    }
    #endregion

    #region 私有方法
    private PptSlide EnsureSlide(Int32 idx)
    {
        while (Slides.Count <= idx)
            Slides.Add(new PptSlide());
        return Slides[idx];
    }

    /// <summary>厘米转换为 EMU（English Metric Units）</summary>
    /// <param name="cm">厘米值</param>
    /// <returns>EMU 值（1 cm = 360000 EMU）</returns>
    public static Int64 CmToEmu(Double cm) => (Int64)(cm * 360000);

    /// <summary>EMU 转换为厘米</summary>
    /// <param name="emu">EMU 值</param>
    /// <returns>厘米值（1 cm = 360000 EMU）</returns>
    public static Double EmuToCm(Int64 emu) => emu / 360000.0;

    /// <summary>磅（点/pt）转换为 EMU</summary>
    /// <param name="pt">磅值</param>
    /// <returns>EMU 值（1 pt = 12700 EMU）</returns>
    public static Int64 PtToEmu(Double pt) => (Int64)(pt * 12700);

    /// <summary>EMU 转换为磅（点/pt）</summary>
    /// <param name="emu">EMU 值</param>
    /// <returns>磅值（1 pt = 12700 EMU）</returns>
    public static Double EmuToPt(Int64 emu) => emu / 12700.0;

    private void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private static void WriteZipEntryText(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private void WriteContentTypes(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
        sb.Append("<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
        sb.Append("<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>");
        sb.Append("<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        for (var i = 0; i < Slides.Count; i++)
            sb.Append($"<Override PartName=\"/ppt/slides/slide{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        // 原始幻灯片内容类型（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
            sb.Append($"<Override PartName=\"/ppt/slides/slide{Slides.Count + i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        // chart types
        foreach (var slide in Slides)
            foreach (var chart in slide.Charts)
                sb.Append($"<Override PartName=\"/ppt/charts/chart{chart.ChartNumber}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
        // image types
        var addedExt = new HashSet<String>();
        foreach (var slide in Slides)
        {
            foreach (var img in slide.Images)
            {
                if (addedExt.Add(img.Extension))
                {
                    var ct = img.Extension is "jpg" or "jpeg" ? "image/jpeg" : "image/png";
                    sb.Append($"<Default Extension=\"{img.Extension}\" ContentType=\"{ct}\"/>");
                }
            }
        }
        sb.Append("</Types>");
        WriteEntry(za, "[Content_Types].xml", sb.ToString());
    }

    private void WriteRels(ZipArchive za) =>
        WriteEntry(za, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>" +
            "</Relationships>");

    private void WritePresentation(ZipArchive za)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<p:presentation xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\" saveSubsetFonts=\"1\">");
        sb.Append($"<p:sldSz cx=\"{SlideWidth}\" cy=\"{SlideHeight}\"/>");
        sb.Append("<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rMaster1\"/></p:sldMasterIdLst>");
        sb.Append("<p:sldIdLst>");
        for (var i = 0; i < Slides.Count; i++)
            sb.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rSlide{i + 1}\"/>");
        // 原始幻灯片（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
            sb.Append($"<p:sldId id=\"{256 + Slides.Count + i}\" r:id=\"rSlide{Slides.Count + i + 1}\"/>");
        sb.Append("</p:sldIdLst>");
        // 演示文稿保护（S07-04）
        if (_protectionHash != null)
            sb.Append($"<p:modifyVerifier algorithmName=\"SHA-512\" hashData=\"{_protectionHash}\" saltData=\"{_protectionSalt}\" spinCount=\"100000\"/>");
        sb.Append("</p:presentation>");
        WriteEntry(za, "ppt/presentation.xml", sb.ToString());
    }

    private void WritePresentationRels(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.Append("<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>");
        sb.Append("<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
        for (var i = 0; i < Slides.Count; i++)
            sb.Append($"<Relationship Id=\"rSlide{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i + 1}.xml\"/>");
        // 原始幻灯片关系（S10-04）
        for (var i = 0; i < _rawSlides.Count; i++)
            sb.Append($"<Relationship Id=\"rSlide{Slides.Count + i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{Slides.Count + i + 1}.xml\"/>");
        sb.Append("</Relationships>");
        WriteEntry(za, "ppt/_rels/presentation.xml.rels", sb.ToString());
    }

    private void WriteSlide(ZipArchive za, Int32 idx, PptSlide slide)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var shapeId = 2;
        // 收集超链接 relId → url（用于 rels 文件）
        var hlinkMap = new Dictionary<String, String>();
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">");

        // background
        if (slide.BackgroundColor != null)
        {
            sb.Append("<p:bg><p:bgPr>");
            sb.Append($"<a:solidFill><a:srgbClr val=\"{slide.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
            sb.Append("<a:effectLst/></p:bgPr></p:bg>");
        }

        sb.Append("<p:cSld><p:spTree>");
        sb.Append("<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
        sb.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");

        // text boxes
        foreach (var tb in slide.TextBoxes)
        {
            // 处理超链接
            String? hlRelId = null;
            if (tb.HyperlinkUrl != null)
            {
                hlRelId = $"rHlk{_hlinkGlobal++}";
                hlinkMap[hlRelId] = tb.HyperlinkUrl;
            }
            sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"TextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{tb.Left}\" y=\"{tb.Top}\"/><a:ext cx=\"{tb.Width}\" cy=\"{tb.Height}\"/></a:xfrm>");
            sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            if (tb.BackgroundColor != null)
                sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.BackgroundColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:noFill/>");
            sb.Append("</p:spPr>");
            sb.Append("<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr><a:lstStyle/>");
            sb.Append($"<a:p><a:pPr algn=\"{tb.Alignment}\"/>");
            if (tb.Runs.Count > 0)
            {
                foreach (var run in tb.Runs)
                {
                    String? runHlRelId = null;
                    if (run.HyperlinkUrl != null)
                    {
                        runHlRelId = $"rHlk{_hlinkGlobal++}";
                        hlinkMap[runHlRelId] = run.HyperlinkUrl;
                    }
                    var runSz = run.FontSize > 0 ? run.FontSize : tb.FontSize;
                    var runFc = run.FontColor ?? tb.FontColor;
                    sb.Append("<a:r>");
                    sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{runSz * 100}\"{(run.Bold ? " b=\"1\"" : "")}{(run.Italic ? " i=\"1\"" : "")} dirty=\"0\">");
                    if (runFc != null)
                        sb.Append($"<a:solidFill><a:srgbClr val=\"{runFc.TrimStart('#')}\"/></a:solidFill>");
                    if (runHlRelId != null)
                        sb.Append($"<a:hlinkClick r:id=\"{runHlRelId}\"/>");
                    sb.Append("</a:rPr>");
                    sb.Append($"<a:t>{EscXml(run.Text)}</a:t>");
                    sb.Append("</a:r>");
                }
            }
            else
            {
                sb.Append("<a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                if (tb.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.FontColor.TrimStart('#')}\"/></a:solidFill>");
                if (hlRelId != null)
                    sb.Append($"<a:hlinkClick r:id=\"{hlRelId}\"/>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
                sb.Append("</a:r>");
            }
            sb.Append("</a:p></p:txBody></p:sp>");
        }

        // shapes（基本图形）
        foreach (var sp in slide.Shapes)
        {
            sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Shape\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/></a:xfrm>");
            sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
            if (sp.FillColor != null)
                sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
            else
                sb.Append("<a:noFill/>");
            if (sp.LineColor != null)
                sb.Append($"<a:ln w=\"{sp.LineWidth}\"><a:solidFill><a:srgbClr val=\"{sp.LineColor.TrimStart('#')}\"/></a:solidFill></a:ln>");
            else
                sb.Append("<a:ln><a:noFill/></a:ln>");
            sb.Append("</p:spPr>");
            if (sp.Text != null)
            {
                sb.Append("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\"{(sp.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                if (sp.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FontColor.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(sp.Text)}</a:t>");
                sb.Append("</a:r></a:p></p:txBody>");
            }
            sb.Append("</p:sp>");
        }

        // images
        foreach (var img in slide.Images)
        {
            sb.Append($"<p:pic><p:nvPicPr><p:cNvPr id=\"{shapeId++}\" name=\"Image\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>");
            sb.Append("<p:blipFill>");
            sb.Append($"<a:blip r:embed=\"{img.RelId}\"/>");
            sb.Append("<a:stretch><a:fillRect/></a:stretch></p:blipFill>");
            sb.Append("<p:spPr>");
            sb.Append($"<a:xfrm><a:off x=\"{img.Left}\" y=\"{img.Top}\"/><a:ext cx=\"{img.Width}\" cy=\"{img.Height}\"/></a:xfrm>");
            sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
        }

        // tables
        foreach (var tbl in slide.Tables)
            BuildPptTableXml(sb, tbl, ref shapeId);

        // charts
        foreach (var chart in slide.Charts)
        {
            sb.Append($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{shapeId++}\" name=\"Chart\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>");
            sb.Append($"<p:xfrm><a:off x=\"{chart.Left}\" y=\"{chart.Top}\"/><a:ext cx=\"{chart.Width}\" cy=\"{chart.Height}\"/></p:xfrm>");
            sb.Append($"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
            sb.Append($"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{chart.RelId}\"/>");
            sb.Append("</a:graphicData></a:graphic></p:graphicFrame>");
        }

        // groups（形状组，S07-02）
        foreach (var grp in slide.Groups)
        {
            sb.Append($"<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"{shapeId++}\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
            sb.Append("<p:grpSpPr>");
            sb.Append($"<a:xfrm><a:off x=\"{grp.Left}\" y=\"{grp.Top}\"/><a:ext cx=\"{grp.Width}\" cy=\"{grp.Height}\"/>");
            sb.Append($"<a:chOff x=\"{grp.Left}\" y=\"{grp.Top}\"/><a:chExt cx=\"{grp.Width}\" cy=\"{grp.Height}\"/></a:xfrm>");
            sb.Append("</p:grpSpPr>");
            // shapes inside group
            foreach (var sp in grp.Shapes)
            {
                sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"GrpShape\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>");
                sb.Append("<p:spPr>");
                sb.Append($"<a:xfrm><a:off x=\"{sp.Left}\" y=\"{sp.Top}\"/><a:ext cx=\"{sp.Width}\" cy=\"{sp.Height}\"/></a:xfrm>");
                sb.Append($"<a:prstGeom prst=\"{sp.ShapeType}\"><a:avLst/></a:prstGeom>");
                if (sp.FillColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{sp.FillColor.TrimStart('#')}\"/></a:solidFill>");
                else
                    sb.Append("<a:noFill/>");
                sb.Append("</p:spPr>");
                if (sp.Text != null)
                {
                    sb.Append("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>");
                    sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{sp.FontSize * 100}\" dirty=\"0\"/>");
                    sb.Append($"<a:t>{EscXml(sp.Text)}</a:t>");
                    sb.Append("</a:r></a:p></p:txBody>");
                }
                sb.Append("</p:sp>");
            }
            // text boxes inside group
            foreach (var tb in grp.TextBoxes)
            {
                sb.Append($"<p:sp><p:nvSpPr><p:cNvPr id=\"{shapeId++}\" name=\"GrpTextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
                sb.Append("<p:spPr>");
                sb.Append($"<a:xfrm><a:off x=\"{tb.Left}\" y=\"{tb.Top}\"/><a:ext cx=\"{tb.Width}\" cy=\"{tb.Height}\"/></a:xfrm>");
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/>");
                sb.Append("</p:spPr>");
                sb.Append("<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr><a:lstStyle/>");
                sb.Append($"<a:p><a:pPr algn=\"{tb.Alignment}\"/><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" sz=\"{tb.FontSize * 100}\"{(tb.Bold ? " b=\"1\"" : "")} dirty=\"0\">");
                if (tb.FontColor != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{tb.FontColor.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(tb.Text)}</a:t>");
                sb.Append("</a:r></a:p></p:txBody></p:sp>");
            }
            sb.Append("</p:grpSp>");
        }

        sb.Append("</p:spTree></p:cSld>");

        // notes
        if (slide.Notes != null)
        {
            sb.Append("<p:notes><p:cSld><p:spTree>");
            sb.Append("<p:sp><p:nvSpPr><p:cNvPr id=\"1\" name=\"notes\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\"/></p:nvPr></p:nvSpPr>");
            sb.Append("<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:rPr lang=\"zh-CN\" dirty=\"0\"/><a:t>{EscXml(slide.Notes)}</a:t></a:r></a:p>");
            sb.Append("</p:txBody></p:sp></p:spTree></p:cSld></p:notes>");
        }

        // 转场动画
        if (slide.Transition != null)
        {
            var t = slide.Transition;
            sb.Append($"<p:transition dur=\"{t.DurationMs}\" {(t.AdvanceOnClick ? "advClick=\"1\"" : "advClick=\"0\"")}>");
            sb.Append(t.Type switch
            {
                "fade" => "<p:fade/>",
                "push" => $"<p:push dir=\"{t.Direction}\"/>",
                "wipe" => $"<p:wipe dir=\"{t.Direction}\"/>",
                "zoom" => "<p:zoom/>",
                "split" => "<p:split/>",
                "cut" => "<p:cut/>",
                _ => "<p:fade/>",
            });
            sb.Append("</p:transition>");
        }

        sb.Append("</p:sld>");
        WriteEntry(za, $"ppt/slides/slide{idx + 1}.xml", sb.ToString());

        // slide rels
        var relsSb = new StringBuilder();
        relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        relsSb.Append("<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>");
        foreach (var img in slide.Images)
            relsSb.Append($"<Relationship Id=\"{img.RelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/{img.RelId}.{img.Extension}\"/>");
        foreach (var hlEntry in hlinkMap)
            relsSb.Append($"<Relationship Id=\"{hlEntry.Key}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{EscXml(hlEntry.Value)}\" TargetMode=\"External\"/>");
        foreach (var chart in slide.Charts)
            relsSb.Append($"<Relationship Id=\"{chart.RelId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart{chart.ChartNumber}.xml\"/>");
        relsSb.Append("</Relationships>");
        WriteEntry(za, $"ppt/slides/_rels/slide{idx + 1}.xml.rels", relsSb.ToString());

        // write image media
        foreach (var img in slide.Images)
        {
            using var entry = za.CreateEntry($"ppt/media/{img.RelId}.{img.Extension}").Open();
            entry.Write(img.Data, 0, img.Data.Length);
        }

        // write chart XMLs
        foreach (var chart in slide.Charts)
            WriteChartXml(za, chart);
    }

    private void WriteChartXml(ZipArchive za, PptChart chart)
    {
        const String C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<c:chartSpace xmlns:c=\"{C}\" xmlns:a=\"{A}\">");
        sb.Append("<c:date1904 val=\"0\"/>");
        sb.Append("<c:chart>");
        if (chart.Title != null)
        {
            sb.Append("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>");
            sb.Append($"<a:p><a:r><a:t>{EscXml(chart.Title)}</a:t></a:r></a:p>");
            sb.Append("</c:rich></c:tx><c:overlay val=\"0\"/></c:title>");
        }
        sb.Append("<c:autoTitleDeleted val=\"0\"/>");
        sb.Append("<c:plotArea>");

        var chartElem = chart.ChartType switch
        {
            "line" => "lineChart",
            "pie" => "pieChart",
            "area" => "areaChart",
            _ => "barChart",
        };
        sb.Append($"<c:{chartElem}>");
        if (chart.ChartType == "bar")
            sb.Append("<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>");

        var serColors = new[] { "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646" };
        for (var si = 0; si < chart.Series.Count; si++)
        {
            var ser = chart.Series[si];
            var color = serColors[si % serColors.Length];
            sb.Append("<c:ser>");
            sb.Append($"<c:idx val=\"{si}\"/><c:order val=\"{si}\"/>");
            sb.Append($"<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{EscXml(ser.Name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>");
            sb.Append($"<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>");
            // categories
            if (chart.Categories.Length > 0)
            {
                sb.Append("<c:cat><c:strRef><c:f/><c:strCache>");
                sb.Append($"<c:ptCount val=\"{chart.Categories.Length}\"/>");
                for (var ci = 0; ci < chart.Categories.Length; ci++)
                    sb.Append($"<c:pt idx=\"{ci}\"><c:v>{EscXml(chart.Categories[ci])}</c:v></c:pt>");
                sb.Append("</c:strCache></c:strRef></c:cat>");
            }
            // values
            sb.Append("<c:val><c:numRef><c:f/><c:numCache>");
            sb.Append($"<c:ptCount val=\"{ser.Values.Length}\"/>");
            for (var vi = 0; vi < ser.Values.Length; vi++)
                sb.Append($"<c:pt idx=\"{vi}\"><c:v>{ser.Values[vi]}</c:v></c:pt>");
            sb.Append("</c:numCache></c:numRef></c:val>");
            sb.Append("</c:ser>");
        }
        if (chart.ChartType != "pie")
        {
            sb.Append("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
            sb.Append($"</c:{chartElem}>");
            // category axis
            sb.Append("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>");
            // value axis
            sb.Append("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>");
        }
        else
        {
            sb.Append($"</c:{chartElem}>");
        }
        sb.Append("</c:plotArea>");
        sb.Append("<c:legend><c:legendPos val=\"b\"/></c:legend>");
        sb.Append("</c:chart></c:chartSpace>");
        WriteEntry(za, $"ppt/charts/chart{chart.ChartNumber}.xml", sb.ToString());
    }

    private static void BuildPptTableXml(StringBuilder sb, PptTable tbl, ref Int32 shapeId)
    {
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        sb.Append($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{shapeId++}\" name=\"Table\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>");
        sb.Append($"<p:xfrm><a:off x=\"{tbl.Left}\" y=\"{tbl.Top}\"/><a:ext cx=\"{tbl.Width}\" cy=\"{tbl.Height}\"/></p:xfrm>");
        sb.Append($"<a:graphic xmlns:a=\"{A}\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">");
        sb.Append("<a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>");
        // columns
        var colCount = tbl.Rows.Count > 0 ? tbl.Rows[0].Length : 1;
        var autoColW = colCount > 0 ? tbl.Width / colCount : tbl.Width;
        sb.Append("<a:tblGrid>");
        for (var c = 0; c < colCount; c++)
        {
            var cw = tbl.ColWidths.Length > c ? tbl.ColWidths[c] : autoColW;
            sb.Append($"<a:gridCol w=\"{cw}\"/>");
        }
        sb.Append("</a:tblGrid>");
        for (var ri = 0; ri < tbl.Rows.Count; ri++)
        {
            var row = tbl.Rows[ri];
            var isHeaderRow = ri == 0 && tbl.FirstRowHeader;
            sb.Append("<a:tr h=\"370840\">");
            for (var ci = 0; ci < row.Length; ci++)
            {
                tbl.CellStyles.TryGetValue((ri, ci), out var cs);
                var isBold = isHeaderRow || (cs?.Bold ?? false);
                var cellSz = (cs?.FontSize ?? 0) > 0 ? cs!.FontSize : 0;
                var cellFc = cs?.FontColor;
                var cellBg = cs?.BackgroundColor;
                sb.Append("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>");
                sb.Append("<a:p><a:r>");
                sb.Append($"<a:rPr lang=\"zh-CN\" altLang=\"en-US\"{(isBold ? " b=\"1\"" : "")}{(cellSz > 0 ? $" sz=\"{cellSz * 100}\"" : "")} dirty=\"0\">");
                if (cellFc != null)
                    sb.Append($"<a:solidFill><a:srgbClr val=\"{cellFc.TrimStart('#')}\"/></a:solidFill>");
                sb.Append("</a:rPr>");
                sb.Append($"<a:t>{EscXml(row[ci])}</a:t>");
                sb.Append("</a:r></a:p></a:txBody>");
                if (cellBg != null)
                    sb.Append($"<a:tcPr><a:solidFill><a:srgbClr val=\"{cellBg.TrimStart('#')}\"/></a:solidFill></a:tcPr>");
                else
                    sb.Append("<a:tcPr/>");
                sb.Append("</a:tc>");
            }
            sb.Append("</a:tr>");
        }
        sb.Append("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    }

    private void WriteSlideLayout(ZipArchive za)
    {
        WriteEntry(za, "ppt/slideLayouts/slideLayout1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" type=\"blank\" preserve=\"1\">" +
            "<p:cSld name=\"Blank\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
            "</p:spTree></p:cSld></p:sldLayout>");
        WriteEntry(za, "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rMaster1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>" +
            "</Relationships>");
    }

    private void WriteSlideMaster(ZipArchive za)
    {
        const String P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        WriteEntry(za, "ppt/slideMasters/slideMaster1.xml",
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<p:sldMaster xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">" +
            "<p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>" +
            "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
            "</p:spTree></p:cSld>" +
            "<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>" +
            "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rLayout1\"/></p:sldLayoutIdLst>" +
            "</p:sldMaster>");
        WriteEntry(za, "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rTheme1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>" +
            "<Relationship Id=\"rLayout1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
            "</Relationships>");
    }

    private void WriteTheme(ZipArchive za) =>
        WriteEntry(za, "ppt/theme/theme1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
            "<a:themeElements><a:clrScheme name=\"Office\">" +
            "<a:dk1><a:sysClr lastClr=\"000000\" val=\"windowText\"/></a:dk1>" +
            "<a:lt1><a:sysClr lastClr=\"FFFFFF\" val=\"window\"/></a:lt1>" +
            "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>" +
            "<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
            $"<a:accent1><a:srgbClr val=\"{AccentColors[0]}\"/></a:accent1>" +
            $"<a:accent2><a:srgbClr val=\"{AccentColors[1]}\"/></a:accent2>" +
            $"<a:accent3><a:srgbClr val=\"{AccentColors[2]}\"/></a:accent3>" +
            $"<a:accent4><a:srgbClr val=\"{AccentColors[3]}\"/></a:accent4>" +
            $"<a:accent5><a:srgbClr val=\"{AccentColors[4]}\"/></a:accent5>" +
            $"<a:accent6><a:srgbClr val=\"{AccentColors[5]}\"/></a:accent6>" +
            "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>" +
            "<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>" +
            "</a:clrScheme>" +
            "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
            "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:noFill/><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:noFill/></a:fillStyleLst>" +
            "<a:lnStyleLst><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
            "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
            "<a:bgFillStyleLst><a:noFill/><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:noFill/></a:bgFillStyleLst>" +
            "</a:fmtScheme></a:themeElements></a:theme>");

    private static String EscXml(String s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
         .Replace("\"", "&quot;").Replace("'", "&apos;");
    #endregion
}
