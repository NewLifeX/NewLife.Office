namespace NewLife.Office.Word;

/// <summary>样式定义快照（styles.xml 中 <c>w:style</c> 的可编辑视图）</summary>
/// <remarks>
/// 由 <see cref="WordStyles.GetStyle"/> / <see cref="WordStyles.GetStyles"/> 产出，
/// 修改请走 <see cref="WordStyles.SetStyleFont"/> / <see cref="WordStyles.SetStyleParagraph"/> 等方法（直接改本对象不会写回）。
/// </remarks>
public class WordStyleInfo
{
    #region 属性
    /// <summary>样式 ID（w:styleId，如 Normal、Heading1、MyStyle）</summary>
    public String StyleId { get; set; } = String.Empty;

    /// <summary>样式显示名（w:name）</summary>
    public String? Name { get; set; }

    /// <summary>样式类型：paragraph / character / table（w:type）</summary>
    public String? Type { get; set; }

    /// <summary>基础样式 ID（w:basedOn）</summary>
    public String? BasedOn { get; set; }

    /// <summary>西文字体名（w:rFonts w:ascii）</summary>
    public String? FontName { get; set; }

    /// <summary>字号（磅，来自 w:sz 半磅值折算）</summary>
    public Double? FontSize { get; set; }

    /// <summary>是否加粗（三态）</summary>
    public Boolean? Bold { get; set; }

    /// <summary>是否斜体（三态）</summary>
    public Boolean? Italic { get; set; }

    /// <summary>文字颜色（16 进制 RGB，无 # 前缀）</summary>
    public String? Color { get; set; }

    /// <summary>段落对齐（left/center/right/both）</summary>
    public String? Alignment { get; set; }

    /// <summary>段前间距（twips）</summary>
    public Int32? SpaceBefore { get; set; }

    /// <summary>段后间距（twips）</summary>
    public Int32? SpaceAfter { get; set; }

    /// <summary>行距（百分值，100=单倍）</summary>
    public Int32? LineSpacingPct { get; set; }

    /// <summary>左缩进（twips）</summary>
    public Int32? IndentLeft { get; set; }

    /// <summary>首行缩进（twips，负值=悬挂）</summary>
    public Int32? FirstLineIndent { get; set; }
    #endregion
}
