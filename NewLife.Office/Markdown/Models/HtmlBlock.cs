namespace NewLife.Office.Markdown;

/// <summary>HTML 块</summary>
public sealed class HtmlBlock : MarkdownBlock
{
    /// <summary>HTML 原始文本</summary>
    public String RawText { get; set; } = String.Empty;

    /// <summary>创建 HTML 块</summary>
    /// <param name="html">原始 HTML 内容</param>
    public HtmlBlock(String html)
    {
        Type = MarkdownBlockType.HtmlBlock;
        RawText = html;
    }
}
