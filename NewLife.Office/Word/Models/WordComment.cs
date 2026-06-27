namespace NewLife.Office;

/// <summary>Word 批注（审阅注释）</summary>
/// <remarks>
/// 对应 OOXML <c>comments.xml</c> 中的 <c>w:comment</c> 元素。
/// <example>
/// <code>
/// var comment = new WordComment
/// {
///     Id = 1,
///     Author = "张三",
///     Date = DateTime.Now,
///     Text = "此处语义不清，建议修改。",
/// };
/// doc.Comments.Add(comment);
/// </code>
/// </example>
/// </remarks>
public class WordComment
{
    #region 属性
    /// <summary>批注 ID（文档内唯一）</summary>
    public Int32 Id { get; set; }

    /// <summary>批注作者</summary>
    public String? Author { get; set; }

    /// <summary>批注创建时间</summary>
    public DateTime? Date { get; set; }

    /// <summary>作者缩写（通常为姓名首字母）</summary>
    public String? Initials { get; set; }

    /// <summary>批注纯文本内容（快捷属性，与 Paragraphs 二选一）</summary>
    public String? Text { get; set; }

    /// <summary>批注内容段落列表（支持富文本格式），优先级高于 Text</summary>
    public List<WordParagraph> Paragraphs { get; set; } = [];
    #endregion
}
