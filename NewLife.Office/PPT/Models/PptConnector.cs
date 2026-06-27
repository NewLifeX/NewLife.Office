namespace NewLife.Office;

/// <summary>PPT 连接线/连接器</summary>
/// <remarks>
/// 用于流程图、组织结构图等需要连接两个形状的场景。
/// 对标 OOXML <c>p:cxnSp</c>（Connection Shape）元素。
/// <example>
/// <code>
/// var connector = new PptConnector
/// {
///     ConnectorType = "elbow",
///     Left = 0, Top = 0, Width = (Int64)(10 * 360000), Height = (Int64)(5 * 360000),
///     LineColor = "404040",
///     LineWidth = 12700,   // 1pt
///     EndArrow = "arrow",
/// };
/// slide.Connectors.Add(connector);
/// </code>
/// </example>
/// </remarks>
public class PptConnector
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

    /// <summary>连接器类型：straight（直线）/ elbow（折线）/ curved（曲线）</summary>
    public String ConnectorType { get; set; } = "straight";

    /// <summary>线条颜色（16进制 RGB，无 # 前缀）</summary>
    public String? LineColor { get; set; }

    /// <summary>线条粗细（EMU，1pt = 12700），默认 9525（0.75pt）</summary>
    public Int32 LineWidth { get; set; } = 9525;

    /// <summary>起始端箭头类型：none/arrow/block/open/oval/diamond/stealth</summary>
    public String? StartArrow { get; set; }

    /// <summary>末端箭头类型：none/arrow/block/open/oval/diamond/stealth</summary>
    public String? EndArrow { get; set; }

    /// <summary>连接线是否为虚线（null=实线）：dash/dot/dashDot/sysDash/sysDot</summary>
    public String? DashStyle { get; set; }

    /// <summary>关系 ID（Reader/Writer 内部使用）</summary>
    public String? RelId { get; set; }
    #endregion
}
