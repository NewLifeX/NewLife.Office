namespace NewLife.Office;

/// <summary>PPT 图表系列数据（S06-04）</summary>
public class PptChartSeriesData
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据值数组，对应各分类</summary>
    public Double[] Values { get; set; } = [];
    #endregion
}

/// <summary>PPT 图表信息（S06-04）</summary>
public class PptChartInfo
{
    #region 属性
    /// <summary>图表编号（在 ppt/charts/chart{N}.xml 中的序号）</summary>
    public Int32 ChartNumber { get; set; }

    /// <summary>图表类型（bar/line/pie/area/scatter 等）</summary>
    public String ChartType { get; set; } = String.Empty;

    /// <summary>分类标签数组</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>系列数据集合</summary>
    public List<PptChartSeriesData> Series { get; } = [];
    #endregion
}

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

/// <summary>PPT 富文本片段（S10-01）</summary>
/// <remarks>
/// 支持每个片段独立设置字体、粗体、斜体、颜色、超链接。
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
