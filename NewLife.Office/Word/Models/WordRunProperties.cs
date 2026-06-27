namespace NewLife.Office;

/// <summary>文字格式属性</summary>
public class WordRunProperties
{
    #region 属性
    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>前景色（16进制 RGB，如 "FF0000"）</summary>
    public String? ForeColor { get; set; }

    /// <summary>字号（磅）</summary>
    public Single? FontSize { get; set; }

    /// <summary>字体名称</summary>
    public String? FontName { get; set; }

    /// <summary>删除线</summary>
    public Boolean Strikethrough { get; set; }

    /// <summary>上标（与 Subscript 互斥，对应 OOXML w:vertAlign w:val="superscript"）</summary>
    public Boolean Superscript { get; set; }

    /// <summary>下标（与 Superscript 互斥，对应 OOXML w:vertAlign w:val="subscript"）</summary>
    public Boolean Subscript { get; set; }

    /// <summary>下划线样式。设置任意值即自动视为 Underline=true；支持 single/double/dotted/dash/wave/thick/wavyDouble/words 等，见 <see cref="WordUnderlineStyles"/></summary>
    /// <remarks>为 null 且 Underline=true 时 Writer 输出默认 single；为 null 且 Underline=false 时不输出下划线。</remarks>
    public String? UnderlineStyle { get; set; }

    /// <summary>字符间距（缇，twips），正值=加宽，负值=紧缩</summary>
    public Single? CharacterSpacing { get; set; }

    /// <summary>字符缩放百分比（100=正常, 150=宽150%, 80=窄80%）</summary>
    public Int32? CharacterScaling { get; set; }

    /// <summary>发光颜色（16进制 RGB，如 "FFD700"）。设置后自动启用发光效果</summary>
    public String? GlowColor { get; set; }

    /// <summary>发光半径（EMU，默认 254000 = 10pt）</summary>
    public Int64? GlowSize { get; set; }

    /// <summary>阴影颜色（16进制 RGB，如 "808080"）。设置后自动启用阴影效果</summary>
    public String? ShadowColor { get; set; }

    /// <summary>阴影 X 偏移（EMU，正值=右偏移，默认 25400 = 1pt）</summary>
    public Int64? ShadowOffsetX { get; set; }

    /// <summary>阴影 Y 偏移（EMU，正值=下偏移，默认 25400 = 1pt）</summary>
    public Int64? ShadowOffsetY { get; set; }
    #endregion
}
