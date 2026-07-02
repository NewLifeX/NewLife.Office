namespace NewLife.Office.Markdown;

/// <summary>无序列表块</summary>
public sealed class BulletListBlock : MarkdownBlock
{
    /// <summary>创建无序列表</summary>
    /// <param name="items">列表项</param>
    public BulletListBlock(IEnumerable<MarkdownBlock> items)
    {
        Type = MarkdownBlockType.BulletList;
        Children.AddRange(items);
    }
}
