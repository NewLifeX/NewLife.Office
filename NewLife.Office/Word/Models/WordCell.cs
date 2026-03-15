namespace NewLife.Office;

/// <summary>表格单元格</summary>
public class WordCell
{
    #region 属性
    /// <summary>段落集合</summary>
    public List<WordParagraph> Paragraphs { get; } = [];

    /// <summary>背景色（16进制 RGB）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>合并列数</summary>
    public Int32 ColSpan { get; set; } = 1;

    /// <summary>合并行数（垂直合并）</summary>
    public Int32 RowSpan { get; set; } = 1;
    #endregion
}
