namespace NewLife.Office;

/// <summary>PPT 表格单元格边框样式</summary>
/// <remarks>
/// 四边独立颜色和线宽，对应 OOXML <c>a:lnL</c>/<c>a:lnR</c>/<c>a:lnT</c>/<c>a:lnB</c>。
/// <example>
/// <code>
/// table.CellBorders[(0, 0)] = new PptCellBorder { TopColor = "000000", TopWidth = 12700 };
/// </code>
/// </example>
/// </remarks>
public class PptCellBorder
{
    #region 属性
    /// <summary>左边框颜色（16进制 RGB，无 # 前缀），null 表示无边框</summary>
    public String? LeftColor { get; set; }

    /// <summary>右边框颜色（16进制 RGB，无 # 前缀）</summary>
    public String? RightColor { get; set; }

    /// <summary>上边框颜色（16进制 RGB，无 # 前缀）</summary>
    public String? TopColor { get; set; }

    /// <summary>下边框颜色（16进制 RGB，无 # 前缀）</summary>
    public String? BottomColor { get; set; }

    /// <summary>左边框线宽（EMU，默认 12700 = 1pt）</summary>
    public Int32 LeftWidth { get; set; } = 12700;

    /// <summary>右边框线宽（EMU）</summary>
    public Int32 RightWidth { get; set; } = 12700;

    /// <summary>上边框线宽（EMU）</summary>
    public Int32 TopWidth { get; set; } = 12700;

    /// <summary>下边框线宽（EMU）</summary>
    public Int32 BottomWidth { get; set; } = 12700;
    #endregion
}
