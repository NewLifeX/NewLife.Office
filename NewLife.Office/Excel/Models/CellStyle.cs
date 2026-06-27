namespace NewLife.Office;

/// <summary>单元格样式</summary>
/// <remarks>
/// 用于控制单元格的字体、填充、边框、对齐、数字格式等外观属性。
/// 在写入时通过 WriteHeader/WriteRow 等方法传入。
/// </remarks>
public class ExcelCellStyle
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
    public ExcelHorizontalAlignment HAlign { get; set; }

    /// <summary>垂直对齐</summary>
    public ExcelVerticalAlignment VAlign { get; set; }

    /// <summary>是否自动换行</summary>
    public Boolean WrapText { get; set; }

    /// <summary>边框样式（快捷方式：同时设置四边，单边属性优先级更高）</summary>
    public ExcelCellBorderStyle Border { get; set; }

    /// <summary>边框颜色（快捷方式：同时设置四边颜色，单边颜色属性优先级更高）</summary>
    public String? BorderColor { get; set; }

    /// <summary>左边框样式（优先级高于 Border）</summary>
    public ExcelCellBorderStyle LeftBorder { get; set; }

    /// <summary>左边框颜色（RGB十六进制，优先级高于 BorderColor）</summary>
    public String? LeftBorderColor { get; set; }

    /// <summary>右边框样式（优先级高于 Border）</summary>
    public ExcelCellBorderStyle RightBorder { get; set; }

    /// <summary>右边框颜色（RGB十六进制，优先级高于 BorderColor）</summary>
    public String? RightBorderColor { get; set; }

    /// <summary>上边框样式（优先级高于 Border）</summary>
    public ExcelCellBorderStyle TopBorder { get; set; }

    /// <summary>上边框颜色（RGB十六进制，优先级高于 BorderColor）</summary>
    public String? TopBorderColor { get; set; }

    /// <summary>下边框样式（优先级高于 Border）</summary>
    public ExcelCellBorderStyle BottomBorder { get; set; }

    /// <summary>下边框颜色（RGB十六进制，优先级高于 BorderColor）</summary>
    public String? BottomBorderColor { get; set; }

    /// <summary>自定义数字格式（如 "#,##0.00"、"yyyy-MM-dd"）</summary>
    public String? NumberFormat { get; set; }

    /// <summary>是否删除线</summary>
    public Boolean Strike { get; set; }

    /// <summary>上下标（"superscript"/"subscript"，null 表示无）</summary>
    public String? VerticalAlign { get; set; }

    /// <summary>文本旋转角度（0-180，0 表示不旋转）</summary>
    public Int32 TextRotation { get; set; }

    /// <summary>水平缩进级别（1 级 ≈ 3 字符宽）</summary>
    public Int32 Indent { get; set; }

    /// <summary>是否缩小以填充（当文字超出列宽时自动缩小字号）</summary>
    public Boolean ShrinkToFit { get; set; }

    /// <summary>富文本段落列表；不为 null 时优先使用富文本，忽略普通文本内容</summary>
    public List<RichTextRun>? RichTextRuns { get; set; }

    /// <summary>渐变类型（"linear"/"radial"），仅 GradientColor1 和 GradientColor2 均非空时生效</summary>
    public String? GradientType { get; set; }

    /// <summary>渐变起始颜色（RGB六位十六进制）</summary>
    public String? GradientColor1 { get; set; }

    /// <summary>渐变结束颜色（RGB六位十六进制）</summary>
    public String? GradientColor2 { get; set; }

    /// <summary>图案填充类型名（如 "darkGray"、"darkGrid"、"lightGrid"、"dotted" 等）</summary>
    public String? PatternType { get; set; }

    /// <summary>图案填充前景色（RGB六位十六进制）</summary>
    public String? PatternFgColor { get; set; }
    #endregion

    #region 静态
    /// <summary>默认表头样式（粗体）</summary>
    public static ExcelCellStyle Header => new() { Bold = true };

    /// <summary>标题样式（粗体、大字、居中）</summary>
    public static ExcelCellStyle Title => new() { Bold = true, FontSize = 14, HAlign = ExcelHorizontalAlignment.Center };
    #endregion
}
