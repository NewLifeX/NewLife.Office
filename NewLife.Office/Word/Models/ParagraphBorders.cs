namespace NewLife.Office.Word;

/// <summary>段落四边边框</summary>
/// <remarks>
/// 对应 OOXML w:pBdr 元素；复用现有 <see cref="Border"/> 类型描述单边。
/// 赋值到 <see cref="Paragraph.Borders"/> 后，Writer 会在 w:pPr/w:pBdr 中生成对应 XML。
/// <example>
/// <code>
/// var para = writer.AppendParagraph("带边框段落");
/// para.Borders = new ParagraphBorders
/// {
///     Top    = new Border { Style = BorderStyle.Single, Color = "FF0000", Width = 12 },
///     Bottom = new Border { Style = BorderStyle.Double, Color = "0000FF", Width = 8 },
///     Left   = new Border { Style = BorderStyle.Dotted, Color = "00AA00", Width = 4 },
/// };
/// </code>
/// </example>
/// </remarks>
public class ParagraphBorders
{
    #region 属性
    /// <summary>上边框，null 表示无</summary>
    public Border? Top { get; set; }

    /// <summary>下边框，null 表示无</summary>
    public Border? Bottom { get; set; }

    /// <summary>左边框，null 表示无</summary>
    public Border? Left { get; set; }

    /// <summary>右边框，null 表示无</summary>
    public Border? Right { get; set; }
    #endregion
}
