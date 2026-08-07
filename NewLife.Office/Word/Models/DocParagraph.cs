namespace NewLife.Office.Word;

/// <summary>doc（MS-DOC 二进制）带格式段落</summary>
/// <remarks>
/// 由 <see cref="DocReader.ReadParagraphsWithFormat"/> 产出，基于 CHPX/PAPX 解析段落与文字格式。
/// 覆盖日常高频格式：粗体/斜体/删除线/字号/字体/颜色/对齐/缩进/列表级别。
/// </remarks>
public class DocParagraph
{
    #region 属性
    /// <summary>段落文本</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>删除线</summary>
    public Boolean Strikethrough { get; set; }

    /// <summary>字号（磅）</summary>
    public Single? FontSize { get; set; }

    /// <summary>字体名称</summary>
    public String? FontName { get; set; }

    /// <summary>前景色（16进制 RGB，如 "FF0000"）</summary>
    public String? ForeColor { get; set; }

    /// <summary>对齐：left / center / right / justify（对应 PAPX sprmPJc）</summary>
    public String? Alignment { get; set; }

    /// <summary>左缩进（twips）</summary>
    public Int32? IndentLeft { get; set; }

    /// <summary>首行缩进（twips，负值=悬挂缩进）</summary>
    public Int32? FirstLineIndent { get; set; }

    /// <summary>段前间距（twips）</summary>
    public Int32? SpaceBefore { get; set; }

    /// <summary>段后间距（twips）</summary>
    public Int32? SpaceAfter { get; set; }

    /// <summary>列表级别（-1 或 0 表示非列表，PAPX sprmPIlvl）</summary>
    public Int32 ListLevel { get; set; } = -1;

    /// <summary>编号定义 ID（PAPX sprmPNumId），0 或 -1 表示无编号</summary>
    public Int32 NumberingId { get; set; } = -1;
    #endregion
}
