namespace NewLife.Office.Word;

/// <summary>Word 自选图形（W12，对标 Open XML SDK/GemBox）</summary>
/// <remarks>
/// 表示插入到 Word 文档中的 DrawingML 形状（矩形/椭圆/线条/三角形等）。
/// 形状以 <c>wps:wsp</c>（WordprocessingShape）内联形式嵌入段落，
/// 支持填充色、线条色/线宽、可选文本与旋转。
/// </remarks>
public class WordShape
{
    /// <summary>预设几何类型（prst），如 rect/roundRect/ellipse/line/triangle/diamond</summary>
    public String Type { get; set; } = "rect";

    /// <summary>宽度（厘米）</summary>
    public Double WidthCm { get; set; } = 4;

    /// <summary>高度（厘米）</summary>
    public Double HeightCm { get; set; } = 4;

    /// <summary>填充色（RRGGBB，如 "FF0000"，null 为无填充）</summary>
    public String? FillColor { get; set; }

    /// <summary>线条色（RRGGBB，null 为无边框）</summary>
    public String? LineColor { get; set; }

    /// <summary>线条宽度（磅）</summary>
    public Double LineWidth { get; set; } = 1;

    /// <summary>形状内文本（可选）</summary>
    public String? Text { get; set; }

    /// <summary>旋转角度（度，可选）</summary>
    public Double? Rotation { get; set; }

    /// <summary>快速创建矩形</summary>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <param name="fillColor">填充色 RRGGBB</param>
    /// <returns>形状对象</returns>
    public static WordShape Rect(Double widthCm, Double heightCm, String? fillColor = null) => new() { Type = "rect", WidthCm = widthCm, HeightCm = heightCm, FillColor = fillColor };

    /// <summary>快速创建圆角矩形</summary>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <param name="fillColor">填充色 RRGGBB</param>
    /// <returns>形状对象</returns>
    public static WordShape RoundRect(Double widthCm, Double heightCm, String? fillColor = null) => new() { Type = "roundRect", WidthCm = widthCm, HeightCm = heightCm, FillColor = fillColor };

    /// <summary>快速创建椭圆/圆</summary>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    /// <param name="fillColor">填充色 RRGGBB</param>
    /// <returns>形状对象</returns>
    public static WordShape Ellipse(Double widthCm, Double heightCm, String? fillColor = null) => new() { Type = "ellipse", WidthCm = widthCm, HeightCm = heightCm, FillColor = fillColor };

    /// <summary>快速创建直线</summary>
    /// <param name="widthCm">长度（厘米）</param>
    /// <param name="lineColor">线条色 RRGGBB</param>
    /// <param name="lineWidth">线宽（磅）</param>
    /// <returns>形状对象</returns>
    public static WordShape Line(Double widthCm, String lineColor = "000000", Double lineWidth = 1) => new() { Type = "line", WidthCm = widthCm, HeightCm = 0.1, LineColor = lineColor, LineWidth = lineWidth };
}
