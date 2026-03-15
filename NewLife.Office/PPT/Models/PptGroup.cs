namespace NewLife.Office;

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
