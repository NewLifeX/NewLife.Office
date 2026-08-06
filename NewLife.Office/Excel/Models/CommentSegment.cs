namespace NewLife.Office.Excel;

/// <summary>批注富文本段（M26，对标 EPPlus）</summary>
/// <remarks>
/// 单元格批注支持多段富文本，每段可独立设置字体、字号、加粗、斜体与颜色。
/// 对应 comment XML 中 <c>text</c> 下的多个 <c>r</c> 段。
/// </remarks>
public class CommentSegment
{
    /// <summary>段文本</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字体名称（如 "宋体"），null 表示默认</summary>
    public String? FontName { get; set; }

    /// <summary>字号（磅），0 表示默认</summary>
    public Double FontSize { get; set; }

    /// <summary>是否加粗</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>字体颜色（ARGB 十六进制，如 "FFFF0000"），null 表示默认</summary>
    public String? Color { get; set; }

    /// <summary>创建富文本段</summary>
    /// <param name="text">段文本</param>
    /// <returns>富文本段</returns>
    public static CommentSegment Create(String text) => new() { Text = text };

    /// <summary>创建加粗富文本段</summary>
    /// <param name="text">段文本</param>
    /// <returns>富文本段</returns>
    public static CommentSegment BoldText(String text) => new() { Text = text, Bold = true };

    /// <summary>创建带颜色的富文本段</summary>
    /// <param name="text">段文本</param>
    /// <param name="color">颜色（ARGB 十六进制）</param>
    /// <returns>富文本段</returns>
    public static CommentSegment Colored(String text, String color) => new() { Text = text, Color = color };

    /// <inheritdoc/>
    public override String ToString() => Text;
}
