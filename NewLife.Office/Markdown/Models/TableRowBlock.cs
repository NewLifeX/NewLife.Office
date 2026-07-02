namespace NewLife.Office.Markdown;

/// <summary>表格行块</summary>
public sealed class TableRowBlock : MarkdownBlock
{
    /// <summary>创建表格行</summary>
    public TableRowBlock()
    {
        Type = MarkdownBlockType.TableRow;
    }
}
