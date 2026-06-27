namespace NewLife.Office;

/// <summary>段落四边边框</summary>
/// <remarks>
/// 对应 OOXML w:pBdr 元素；复用现有 <see cref="WordBorder"/> 类型描述单边。
/// 赋值到 <see cref="WordParagraph.Borders"/> 后，Writer 会在 w:pPr/w:pBdr 中生成对应 XML。
/// <example>
/// <code>
/// var para = writer.AppendParagraph("带边框段落");
/// para.Borders = new WordParagraphBorders
/// {
///     Top    = new WordBorder { Style = WordBorderStyle.Single, Color = "FF0000", Width = 12 },
///     Bottom = new WordBorder { Style = WordBorderStyle.Double, Color = "0000FF", Width = 8 },
///     Left   = new WordBorder { Style = WordBorderStyle.Dotted, Color = "00AA00", Width = 4 },
/// };
/// </code>
/// </example>
/// </remarks>
public class WordParagraphBorders
{
    #region 属性
    /// <summary>上边框，null 表示无</summary>
    public WordBorder? Top { get; set; }

    /// <summary>下边框，null 表示无</summary>
    public WordBorder? Bottom { get; set; }

    /// <summary>左边框，null 表示无</summary>
    public WordBorder? Left { get; set; }

    /// <summary>右边框，null 表示无</summary>
    public WordBorder? Right { get; set; }
    #endregion
}
