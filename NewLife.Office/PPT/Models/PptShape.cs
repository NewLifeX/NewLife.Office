namespace NewLife.Office;

/// <summary>PPT 幻灯片文本形状</summary>
public class PptShape
{
    #region 属性
    /// <summary>形状ID</summary>
    public Int32 Id { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>形状类型（如 textBox, rect, ellipse, roundRect, triangle, diamond 等）</summary>
    public String ShapeType { get; set; } = String.Empty;

    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>填充色（16进制 RGB），null 表示无填充（写入时使用）</summary>
    public String? FillColor { get; set; }

    /// <summary>图片填充数据（写入时使用），设置后覆盖 FillColor，使用 blipFill 替代 solidFill</summary>
    public Byte[]? FillImage { get; set; }

    /// <summary>图片填充扩展名（默认 "png"），配合 FillImage 使用</summary>
    public String FillImageExt { get; set; } = "png";

    /// <summary>形状图片填充的关系 ID（内部用）</summary>
    public String? ShapeImageRelId { get; set; }

    /// <summary>线条颜色（16进制 RGB），null 表示无线条（写入时使用）</summary>
    public String? LineColor { get; set; }

    /// <summary>线宽（EMU，12700=1pt，写入时使用）</summary>
    public Int32 LineWidth { get; set; } = 12700;

    /// <summary>文字字号（磅，写入时使用）</summary>
    public Int32 FontSize { get; set; } = 14;

    /// <summary>文字颜色（16进制 RGB，写入时使用）</summary>
    public String? FontColor { get; set; }

    /// <summary>文字粗体（写入时使用）</summary>
    public Boolean Bold { get; set; }

    /// <summary>拉丁/西文字体名称（如"Arial"），null 表示使用默认字体</summary>
    public String? LatinFontName { get; set; }

    /// <summary>东亚/中文字体名称（如"微软雅黑"），null 表示使用默认字体</summary>
    public String? EastAsianFontName { get; set; }

    /// <summary>复杂脚本字体名称（如阿拉伯/泰文），null 表示使用默认字体</summary>
    public String? ComplexScriptFontName { get; set; }

    /// <summary>符号字体名称，null 表示使用默认字体</summary>
    public String? SymbolFontName { get; set; }

    /// <summary>旋转角度（S15-02），以 60000 分之一度为单位（如 5400000=90°）</summary>
    public Int32 Rotation { get; set; }

    /// <summary>替换文本/无障碍描述（对应 OOXML descr 属性）</summary>
    public String? AltText { get; set; }

    /// <summary>圆角半径（仅 roundRect 形状有效，EMU）</summary>
    public Int64 CornerRadius { get; set; }
    #endregion
}
