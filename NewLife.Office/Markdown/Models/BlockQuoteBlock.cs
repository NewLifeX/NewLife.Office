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

    /// <summary>创建引用块（MD25：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="children">子块列表</param>
    internal BlockQuoteBlock(List<MarkdownBlock> children)
    {
        Type = MarkdownBlockType.BlockQuote;
        SetChildren(children);
    }
}
