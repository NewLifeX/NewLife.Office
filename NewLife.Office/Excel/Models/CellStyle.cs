namespace NewLife.Office;

/// <summary>单元格样式</summary>
/// <remarks>
/// 用于控制单元格的字体、填充、边框、对齐、数字格式等外观属性。
/// 在写入时通过 WriteHeader/WriteRow 等方法传入。
/// </remarks>
public class CellStyle
{
    #region 属性
    /// <summary>字体名称</summary>
    public String? FontName { get; set; }

    /// <summary>字体大小（磅）</summary>
    public Double FontSize { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>是否下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>字体颜色（RGB十六进制，如 "FF0000" 表示红色）</summary>
    public String? FontColor { get; set; }

    /// <summary>背景色（RGB十六进制，如 "4472C4" 表示蓝色）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>水平对齐</summary>
    public HorizontalAlignment HAlign { get; set; }

    /// <summary>垂直对齐</summary>
    public VerticalAlignment VAlign { get; set; }

    /// <summary>是否自动换行</summary>
    public Boolean WrapText { get; set; }

    /// <summary>边框样式</summary>
    public CellBorderStyle Border { get; set; }

    /// <summary>边框颜色（RGB十六进制）</summary>
    public String? BorderColor { get; set; }

    /// <summary>自定义数字格式（如 "#,##0.00"、"yyyy-MM-dd"）</summary>
    public String? NumberFormat { get; set; }
    #endregion

    #region 静态
    /// <summary>默认表头样式（粗体）</summary>
    public static CellStyle Header => new() { Bold = true };

    /// <summary>标题样式（粗体、大字、居中）</summary>
    public static CellStyle Title => new() { Bold = true, FontSize = 14, HAlign = HorizontalAlignment.Center };
    #endregion
}
