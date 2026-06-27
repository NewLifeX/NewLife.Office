namespace NewLife.Office;

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

    /// <summary>背景色（16进制 RGB），null 表示白色或使用背景图</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>背景图片，null 表示纯色背景</summary>
    public PptImage? BackgroundImage { get; set; }

    /// <summary>演讲者备注</summary>
    public String? Notes { get; set; }

    /// <summary>幻灯片切换动画，null 表示不设置</summary>
    public PptTransition? Transition { get; set; }

    /// <summary>图片关系计数器</summary>
    internal Int32 ImageCounter { get; set; } = 1;

    /// <summary>视频/音频媒体集合</summary>
    public List<PptVideo> Videos { get; } = [];

    /// <summary>形状组集合（S07-02 组合形状）</summary>
    public List<PptGroup> Groups { get; } = [];

    /// <summary>使用的版式索引（0起始，对应 PptxWriter 加载的版式列表；无模板时只有索引 0）</summary>
    public Int32 LayoutIndex { get; set; }

    /// <summary>LayoutEngine 自动排版布局策略，null 表示不启用自动排版</summary>
    /// <remarks>支持：title_content（默认）/title_only/two_column/chart_only/blank</remarks>
    public String? Layout { get; set; }

    /// <summary>连接线/连接器集合（流程图简头等）</summary>
    public List<PptConnector> Connectors { get; } = [];

    /// <summary>元素动画列表（进入/强调/退出动画）</summary>
    public List<PptAnimation> Animations { get; } = [];

    /// <summary>幻灯片批注列表（审阅注释）</summary>
    public List<PptComment> Comments { get; } = [];

    /// <summary>是否隐藏幻灯片（S12-04），对应 <c>p:sld show="0"</c></summary>
    public Boolean Hidden { get; set; }
    #endregion
}
