namespace NewLife.Office.Markdown;

/// <summary>表格单元格块</summary>
public sealed class TableCellBlock : MarkdownBlock
{
    /// <summary>是否为表头</summary>
    public Boolean IsHeader { get; set; }

    /// <summary>对齐方式（"left"/"center"/"right"/""）</summary>
    public String Alignment { get; set; } = String.Empty;

    /// <summary>创建表格单元格</summary>
    /// <param name="inlines">行内内容</param>
    /// <param name="isHeader">是否表头</param>
    /// <param name="alignment">对齐方式</param>
    public TableCellBlock(IEnumerable<MarkdownInline> inlines, Boolean isHeader = false, String alignment = "")
    {
        Type = MarkdownBlockType.TableCell;
        IsHeader = isHeader;
        Alignment = alignment ?? String.Empty;
        Inlines.AddRange(inlines);
    }

    /// <summary>创建表格单元格（MD24：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="inlines">行内列表</param>
    /// <param name="isHeader">是否表头</param>
    /// <param name="alignment">对齐方式</param>
    internal TableCellBlock(List<MarkdownInline> inlines, Boolean isHeader = false, String alignment = "")
    {
        Type = MarkdownBlockType.TableCell;
        IsHeader = isHeader;
        Alignment = alignment ?? String.Empty;
        SetInlines(inlines);
    }
}
