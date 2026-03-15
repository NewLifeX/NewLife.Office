namespace NewLife.Office;

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
