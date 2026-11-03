namespace NewLife.Office.Ppt;

/// <summary>PPT 幻灯片</summary>
public class Slide
{
    #region 属性
    /// <summary>文本框集合</summary>
    public List<TextBox> TextBoxes { get; set; } = [];

    /// <summary>图片集合</summary>
    public List<Picture> Images { get; set; } = [];

    /// <summary>表格集合</summary>
    public List<Table> Tables { get; set; } = [];

    /// <summary>基本图形集合</summary>
    public List<Shape> Shapes { get; set; } = [];

    /// <summary>图表集合</summary>
    public List<Chart> Charts { get; set; } = [];

    /// <summary>背景色（16进制 RGB），null 表示白色或使用背景图</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>背景图片，null 表示纯色背景</summary>
    public Picture? BackgroundImage { get; set; }

    /// <summary>背景渐变类型（S15-04），"linear" 或 "radial"，null 表示不使用渐变</summary>
    public String? BackgroundGradientType { get; set; }

    /// <summary>背景渐变起始色（16进制 RGB）</summary>
    public String? BackgroundGradientColor1 { get; set; }

    /// <summary>背景渐变结束色（16进制 RGB）</summary>
    public String? BackgroundGradientColor2 { get; set; }

    /// <summary>演讲者备注</summary>
    public String? Notes { get; set; }

    /// <summary>幻灯片切换动画，null 表示不设置</summary>
    public Transition? Transition { get; set; }

    /// <summary>图片关系计数器</summary>
    internal Int32 ImageCounter { get; set; } = 1;

    /// <summary>视频/音频媒体集合</summary>
    public List<Video> Videos { get; set; } = [];

    /// <summary>形状组集合（S07-02 组合形状）</summary>
    public List<GroupShape> Groups { get; set; } = [];

    /// <summary>使用的版式索引（0起始，对应 PptxWriter 加载的版式列表；无模板时只有索引 0）</summary>
    public Int32 LayoutIndex { get; set; }

    /// <summary>LayoutEngine 自动排版布局策略，null 表示不启用自动排版</summary>
    /// <remarks>支持：title_content（默认）/title_only/two_column/chart_only/blank</remarks>
    public String? Layout { get; set; }

    /// <summary>连接线/连接器集合（流程图简头等）</summary>
    public List<ConnectionShape> Connectors { get; set; } = [];

    /// <summary>元素动画列表（进入/强调/退出动画）</summary>
    public List<Animation> Animations { get; set; } = [];

    /// <summary>幻灯片批注列表（审阅注释）</summary>
    public List<Comment> Comments { get; set; } = [];

    /// <summary>是否隐藏幻灯片（S12-04），对应 <c>p:sld show="0"</c></summary>
    public Boolean Hidden { get; set; }
    #endregion
    /// <summary>浅拷贝（引用类型字段仍共享，供深拷贝器配合集合重建使用）</summary>
    /// <returns>浅拷贝实例</returns>
    internal Object ShallowClone() => MemberwiseClone();

}
