namespace NewLife.Office;

/// <summary>Word 表格行</summary>
/// <remarks>
/// 对应 OOXML <c>w:tr</c> 元素，持有行级属性和单元格集合。
/// </remarks>
public class WordTableRow
{
    #region 属性
    /// <summary>单元格集合</summary>
    public List<WordCell> Cells { get; set; } = [];

    /// <summary>行高（缇，twips，1 twip = 1/20 磅），null 表示自动高度</summary>
    public Int32? Height { get; set; }

    /// <summary>是否作为标题行（跨页时重复显示）</summary>
    public Boolean IsHeader { get; set; }

    /// <summary>是否禁止跨页断行</summary>
    public Boolean CantSplit { get; set; }

    /// <summary>行背景色（16进制 RGB，无 # 前缀），null 表示无背景</summary>
    public String? BackgroundColor { get; set; }
    #endregion
}
