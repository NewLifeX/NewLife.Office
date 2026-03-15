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
