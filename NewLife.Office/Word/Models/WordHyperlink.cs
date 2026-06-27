namespace NewLife.Office;

/// <summary>Word 超链接模型</summary>
/// <remarks>
/// 比 <see cref="WordRun.HyperlinkRelId"/> 更完整，持有 URL、显示文本和工具提示。
/// 可作为 <see cref="WordElement"/> 中段落内 Run 的来源，也可独立描述文档级超链接。
/// <example>
/// <code>
/// var link = new WordHyperlink
/// {
///     Url = "https://newlifex.com",
///     DisplayText = "NewLife 官网",
///     Tooltip = "点击访问官网",
/// };
/// </code>
/// </example>
/// </remarks>
public class WordHyperlink
{
    #region 属性
    /// <summary>目标 URL（外部链接），与 BookmarkName 二选一</summary>
    public String? Url { get; set; }

    /// <summary>目标书签名（文档内部跳转），与 Url 二选一</summary>
    public String? BookmarkName { get; set; }

    /// <summary>显示文本，null 时显示 Url</summary>
    public String? DisplayText { get; set; }

    /// <summary>悬停提示文本</summary>
    public String? Tooltip { get; set; }

    /// <summary>显示文本的格式属性</summary>
    public WordRunProperties? TextProperties { get; set; }

    /// <summary>关系 ID（Reader 填充，内部使用）</summary>
    public String? RelId { get; set; }
    #endregion
}
