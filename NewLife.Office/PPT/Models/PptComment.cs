namespace NewLife.Office;

/// <summary>PPT 批注（审阅注释）</summary>
/// <remarks>
/// 对应 OOXML pptx 中的 <c>comments.xml</c>，记录审阅者在幻灯片上添加的注释。
/// 通过 <see cref="PptSlide.Comments"/> 集合关联到幻灯片。
/// <example>
/// <code>
/// var comment = new PptComment
/// {
///     Author = "李四",
///     Text = "建议将此图表改为饼图",
///     X = 0.3f,  // 幻灯片宽度的 30% 处
///     Y = 0.5f,
/// };
/// slide.Comments.Add(comment);
/// </code>
/// </example>
/// </remarks>
public class PptComment
{
    #region 属性
    /// <summary>批注 ID（文档内唯一，从 1 开始）</summary>
    public Int32 Index { get; set; }

    /// <summary>批注作者姓名</summary>
    public String? Author { get; set; }

    /// <summary>作者 ID（通常为邮件地址）</summary>
    public String? AuthorId { get; set; }

    /// <summary>批注创建时间</summary>
    public DateTime? Date { get; set; }

    /// <summary>批注文本内容</summary>
    public String? Text { get; set; }

    /// <summary>批注图钉的 X 坐标（幻灯片宽度的比例，0=左, 1=右）</summary>
    public Single X { get; set; }

    /// <summary>批注图钉的 Y 坐标（幻灯片高度的比例，0=上, 1=下）</summary>
    public Single Y { get; set; }
    #endregion
}
