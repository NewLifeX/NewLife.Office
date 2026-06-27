namespace NewLife.Office;

/// <summary>Word 单边边框样式</summary>
/// <remarks>
/// 描述段落、表格或单元格的一条边的边框，包括线型、颜色和粗细。
/// <example>
/// <code>
/// var border = new WordBorder { Style = WordBorderStyle.Single, Color = "000000", Width = 4 };
/// </code>
/// </example>
/// </remarks>
public class WordBorder
{
    #region 属性
    /// <summary>边框线型</summary>
    public WordBorderStyle Style { get; set; } = WordBorderStyle.Single;

    /// <summary>边框颜色（16进制 RGB，无 # 前缀），null 表示自动/继承</summary>
    public String? Color { get; set; }

    /// <summary>边框粗细（八分之一磅），默认 4（= 0.5pt）；常用值：4/8/12/18/24</summary>
    public Int32 Width { get; set; } = 4;

    /// <summary>是否有阴影</summary>
    public Boolean Shadow { get; set; }

    /// <summary>主题颜色名称（如 "accent1"），与 Color 二选一</summary>
    public String? ThemeColor { get; set; }
    #endregion
}
