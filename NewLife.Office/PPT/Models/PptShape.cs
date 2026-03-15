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
    #endregion
}
