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
    #endregion
}
