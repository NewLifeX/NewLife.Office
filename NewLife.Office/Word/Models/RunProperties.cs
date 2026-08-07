namespace NewLife.Office.Word;

/// <summary>文字格式属性</summary>
/// <remarks>
/// 布尔格式属性采用三态（<see cref="Nullable{T}"/>）设计：
/// <c>null</c> = 未设置（继承样式/默认），<c>true</c> = 显式开启，
/// <c>false</c> = 显式关闭（对应 OOXML <c>w:val="0"</c>，用于取消样式继承的格式）。
/// </remarks>
public class RunProperties
{
    #region 属性
    /// <summary>粗体（null=继承，true=加粗，false=显式取消加粗）</summary>
    public Boolean? Bold { get; set; }

    /// <summary>斜体（null=继承，true=斜体，false=显式取消斜体）</summary>
    public Boolean? Italic { get; set; }

    /// <summary>下划线（null=继承，true=下划线，false=显式取消下划线）</summary>
    public Boolean? Underline { get; set; }

    /// <summary>前景色（16进制 RGB，如 "FF0000"）</summary>
    public String? ForeColor { get; set; }

    /// <summary>字号（磅）</summary>
    public Single? FontSize { get; set; }

    /// <summary>西文/复杂文种字体名称（对应 w:rFonts w:ascii / w:hAnsi）</summary>
    public String? FontName { get; set; }

    /// <summary>东亚字体名称（对应 w:rFonts w:eastAsia，如 "宋体"），null 时回退使用 <see cref="FontName"/></summary>
    public String? EastAsiaFontName { get; set; }

    /// <summary>删除线（null=继承，true=删除线，false=显式取消）</summary>
    public Boolean? Strikethrough { get; set; }

    /// <summary>上标（与 Subscript 互斥，对应 OOXML w:vertAlign w:val="superscript"）</summary>
    public Boolean? Superscript { get; set; }

    /// <summary>下标（与 Superscript 互斥，对应 OOXML w:vertAlign w:val="subscript"）</summary>
    public Boolean? Subscript { get; set; }

    /// <summary>高亮颜色（16进制 RGB，如 "FFFF00"；对应 w:highlight w:val）</summary>
    public String? HighlightColor { get; set; }

    /// <summary>小型大写字母（对应 w:smallCaps）</summary>
    public Boolean? SmallCaps { get; set; }

    /// <summary>全大写字母（对应 w:caps）</summary>
    public Boolean? AllCaps { get; set; }

    /// <summary>隐藏文字（对应 w:vanish，Word 中不显示）</summary>
    public Boolean? Hidden { get; set; }

    /// <summary>语言（对应 w:lang w:val，如 "zh-CN"）</summary>
    public String? Language { get; set; }

    /// <summary>下划线样式。设置任意值即自动视为 Underline=true；支持 single/double/dotted/dash/wave/thick/wavyDouble/words 等，见 <see cref="UnderlineStyles"/></summary>
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
