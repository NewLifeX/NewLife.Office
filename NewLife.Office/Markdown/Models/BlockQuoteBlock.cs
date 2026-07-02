namespace NewLife.Office.Markdown;

/// <summary>引用块</summary>
public sealed class BlockQuoteBlock : MarkdownBlock
{
    /// <summary>创建引用块</summary>
    /// <param name="children">子块</param>
    public BlockQuoteBlock(IEnumerable<MarkdownBlock> children)
    {
        Type = MarkdownBlockType.BlockQuote;
        Children.AddRange(children);
    }
}
