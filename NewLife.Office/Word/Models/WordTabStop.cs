namespace NewLife.Office;

/// <summary>Word 制表位</summary>
/// <remarks>
/// 对应 OOXML w:tab 元素，定义段落内的制表位位置、对齐方式和前导符。
/// 赋值到 <see cref="WordParagraph.TabStops"/> 集合后，Writer 会在 w:pPr/w:tabs 中生成对应 XML。
/// <example>
/// <code>
/// var para = writer.AppendParagraph("姓名\t部门\t薪资");
/// para.TabStops = new List&lt;WordTabStop&gt;
/// {
///     new WordTabStop { Position = 3600, Alignment = "left" },
///     new WordTabStop { Position = 7200, Alignment = "right", Leader = "dot" },
/// };
/// </code>
/// </example>
/// </remarks>
public class WordTabStop
{
    #region 属性
    /// <summary>制表位位置（缇，twips），相对于段落左边缘</summary>
    public Int32 Position { get; set; }

    /// <summary>对齐方式：left（默认）/ center / right / decimal / bar</summary>
    public String Alignment { get; set; } = "left";

    /// <summary>前导符：null(无) / dot / hyphen / underscore / heavy / middleDot</summary>
    public String? Leader { get; set; }
    #endregion
}
