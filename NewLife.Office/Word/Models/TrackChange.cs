namespace NewLife.Office.Word;

/// <summary>修订类型（Track Changes）</summary>
public enum TrackChangeType
{
    /// <summary>插入</summary>
    Insert,

    /// <summary>删除</summary>
    Delete,
}

/// <summary>文档修订记录（Track Changes）</summary>
/// <remarks>
/// 对应 Word 修订追踪中的 <c>w:ins</c>（插入）与 <c>w:del</c>（删除）元素，
/// 记录作者、日期与修订文本，用于审计与 AI 知识库场景。
/// </remarks>
public class TrackChange
{
    /// <summary>修订类型：插入或删除</summary>
    public TrackChangeType Type { get; set; }

    /// <summary>修订作者（w:author）</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>修订日期（w:date，ISO 8601）</summary>
    public String? Date { get; set; }

    /// <summary>修订 ID（w:id）</summary>
    public String? Id { get; set; }

    /// <summary>修订文本（插入取 w:t，删除取 w:delText）</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>所在段落纯文本（上下文）</summary>
    public String? ParagraphText { get; set; }

    /// <inheritdoc/>
    public override String ToString() => $"{Type}: {Text} (作者 {Author})";
}
