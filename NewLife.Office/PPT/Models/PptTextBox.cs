namespace NewLife.Office;

/// <summary>PPT 幻灯片文本框</summary>
public class PptTextBox
{
    #region 属性
    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; }

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; }

    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字号（磅）</summary>
    public Int32 FontSize { get; set; } = 18;

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>文字颜色（16进制 RGB，如 "000000"）</summary>
    public String? FontColor { get; set; }

    /// <summary>对齐（l/ctr/r）</summary>
    public String Alignment { get; set; } = "l";

    /// <summary>背景色（16进制 RGB），null 表示透明</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>超链接 URL，不为 null 时点击文字跳转</summary>
    public String? HyperlinkUrl { get; set; }

    /// <summary>富文本片段集合；非空时优先使用，忽略 Text/FontSize/Bold/FontColor 等单一格式属性</summary>
    public List<PptTextRun> Runs { get; } = [];
    #endregion
}
