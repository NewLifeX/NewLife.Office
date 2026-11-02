namespace NewLife.Office.Word;

/// <summary>段落内的批注锚点范围（对应 OOXML 的 <c>w:commentRangeStart</c> / <c>w:commentRangeEnd</c>）</summary>
/// <remarks>
/// 一个批注只能挂在文档里的一段连续文字上，位置用段落内的 Run 索引表达：
/// 起点 Run 之前插入 <c>commentRangeStart</c>，终点 Run 之后插入 <c>commentRangeEnd</c> 与引用 run。
/// <para>通常由 <see cref="WordDocumentEditExtensions.AddComment(Document, Paragraph, String, String?)"/> 生成，无需手工构造。</para>
/// </remarks>
public class CommentRange
{
    #region 属性
    /// <summary>批注 ID，需与 <see cref="Comment.Id"/> 对应且在文档内唯一</summary>
    public Int32 Id { get; set; }

    /// <summary>起始 Run 索引（含）</summary>
    public Int32 StartRunIndex { get; set; }

    /// <summary>结束 Run 索引（含）。等于 <see cref="StartRunIndex"/> 时批注单个 Run</summary>
    public Int32 EndRunIndex { get; set; }
    #endregion
}
