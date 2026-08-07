namespace NewLife.Office.Markdown;

/// <summary>列表项块</summary>
public sealed class ListItemBlock : MarkdownBlock
{
    /// <summary>是否为任务项</summary>
    public Boolean IsTaskItem { get; set; }

    /// <summary>任务项是否已勾选</summary>
    public Boolean IsChecked { get; set; }

    /// <summary>创建列表项（简单行内内容）</summary>
    /// <param name="inlines">行内内容</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    public ListItemBlock(IEnumerable<MarkdownInline> inlines, Boolean isTaskItem = false, Boolean isChecked = false)
    {
        Type = MarkdownBlockType.ListItem;
        IsTaskItem = isTaskItem;
        IsChecked = isChecked;
        Inlines.AddRange(inlines);
    }

    /// <summary>创建列表项（MD24：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="inlines">行内列表</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    internal ListItemBlock(List<MarkdownInline> inlines, Boolean isTaskItem = false, Boolean isChecked = false)
    {
        Type = MarkdownBlockType.ListItem;
        IsTaskItem = isTaskItem;
        IsChecked = isChecked;
        SetInlines(inlines);
    }

    /// <summary>创建列表项（嵌套块内容）</summary>
    /// <param name="children">子块</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    public static ListItemBlock CreateWithBlocks(IEnumerable<MarkdownBlock> children,
        Boolean isTaskItem = false, Boolean isChecked = false)
    {
        var b = new ListItemBlock([], isTaskItem, isChecked);
        b.Children.AddRange(children);
        return b;
    }

    /// <summary>创建列表项（嵌套块内容，MD25：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="children">子块列表</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    internal static ListItemBlock CreateWithBlocks(List<MarkdownBlock> children,
        Boolean isTaskItem = false, Boolean isChecked = false)
    {
        var b = new ListItemBlock([], isTaskItem, isChecked);
        b.SetChildren(children);
        return b;
    }
}
