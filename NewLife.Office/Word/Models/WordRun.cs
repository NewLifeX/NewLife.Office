namespace NewLife.Office;

/// <summary>文字段（Run）</summary>
public class WordRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>格式属性</summary>
    public WordRunProperties? Properties { get; set; }

    /// <summary>超链接关系ID（内部用）</summary>
    public String? HyperlinkRelId { get; set; }
    #endregion
}
