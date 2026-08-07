namespace NewLife.Office.Markdown;

/// <summary>无序列表块</summary>
public sealed class BulletListBlock : MarkdownBlock
{
    /// <summary>是否为松散列表（项间有空白行，CommonMark 项内容渲染为 &lt;p&gt;）</summary>
    public Boolean IsLoose { get; set; }

    /// <summary>创建无序列表</summary>
    /// <param name="items">列表项</param>
    public BulletListBlock(IEnumerable<MarkdownBlock> items)
    {
        Type = MarkdownBlockType.BulletList;
        Children.AddRange(items);
    }

    /// <summary>创建无序列表（MD25：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="items">列表项列表</param>
    internal BulletListBlock(List<MarkdownBlock> items)
    {
        Type = MarkdownBlockType.BulletList;
        SetChildren(items);
    }
}
