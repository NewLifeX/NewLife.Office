namespace NewLife.Office.Word;

/// <summary>文字段（Run）</summary>
public class Run
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>格式属性</summary>
    public RunProperties? Properties { get; set; }

    /// <summary>超链接关系ID（内部用）</summary>
    public String? HyperlinkRelId { get; set; }

    /// <summary>
    /// 域指令（复杂域的 instrText 内容），如 <c>PAGE \* MERGEFORMAT</c>、<c>NUMPAGES</c>、<c>TOC \o "1-3" \h \z \u</c>。
    /// 非空时 Writer 输出完整复杂域（begin/instrText/separate/结果/end），并忽略 <see cref="Text"/>。
    /// </summary>
    public String? FieldCode { get; set; }

    /// <summary>域缓存结果（显示文本）。为 null 时由 Word 打开后更新域填充（需 <see cref="Document.UpdateFields"/>）</summary>
    public String? FieldResult { get; set; }
    #endregion
}
