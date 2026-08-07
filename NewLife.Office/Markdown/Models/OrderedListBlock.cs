namespace NewLife.Office.Markdown;

/// <summary>有序列表块</summary>
public sealed class OrderedListBlock : MarkdownBlock
{
    /// <summary>起始序号</summary>
    public Int32 OrderedStart { get; set; } = 1;

    /// <summary>是否为松散列表（项间有空白行，CommonMark 项内容渲染为 &lt;p&gt;）</summary>
    public Boolean IsLoose { get; set; }

    /// <summary>创建有序列表</summary>
    /// <param name="items">列表项</param>
    /// <param name="start">起始序号</param>
    public OrderedListBlock(IEnumerable<MarkdownBlock> items, Int32 start = 1)
    {
        Type = MarkdownBlockType.OrderedList;
        OrderedStart = start;
        Children.AddRange(items);
    }

    /// <summary>创建有序列表（MD25：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="items">列表项列表</param>
    /// <param name="start">起始序号</param>
    internal OrderedListBlock(List<MarkdownBlock> items, Int32 start = 1)
    {
        Type = MarkdownBlockType.OrderedList;
        OrderedStart = start;
        SetChildren(items);
    }
}
