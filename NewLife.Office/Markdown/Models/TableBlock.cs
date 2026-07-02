namespace NewLife.Office.Markdown;

/// <summary>表格块</summary>
public sealed class TableBlock : MarkdownBlock
{
    /// <summary>创建表格块</summary>
    public TableBlock()
    {
        Type = MarkdownBlockType.Table;
    }
}
