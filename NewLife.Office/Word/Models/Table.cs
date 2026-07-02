namespace NewLife.Office.Word;

/// <summary>Word 表格独立模型</summary>
/// <remarks>
/// 比 <c>Element.TableRows</c>（嵌套列表）更丰富：支持行级属性、表格宽度、对齐和四边边框。
/// 在 <see cref="Element"/> 中通过 <see cref="Element.Table"/> 属性使用。
/// <example>
/// <code>
/// var table = new Table
/// {
///     FirstRowHeader = true,
///     Style = new TableStyle { HeaderBgColor = "4472C4", HeaderBold = true },
///     Borders = TableBorders.All(BorderStyle.Single),
/// };
/// table.Rows.Add(new TableRow
/// {
///     IsHeader = true,
///     Cells = [new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "姓名" } } } } },
///              new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "部门" } } } } }],
/// });
/// </code>
/// </example>
/// </remarks>
public class Table
{
    #region 属性
    /// <summary>表格行集合</summary>
    public List<TableRow> Rows { get; set; } = [];

    /// <summary>表格样式（表头背景/斑马纹等）</summary>
    public TableStyle? Style { get; set; }

    /// <summary>四边边框配置，null 表示使用默认样式</summary>
    public TableBorders? Borders { get; set; }

    /// <summary>表格总宽度（缇，twips），null 表示自动（铺满可用宽度）</summary>
    public Int32? Width { get; set; }

    /// <summary>水平对齐方式（left/center/right），null 表示继承</summary>
    public String? Alignment { get; set; }

    /// <summary>首行是否作为表头（影响斑马纹起始行和跨页标题）</summary>
    public Boolean FirstRowHeader { get; set; } = true;

    /// <summary>各列宽度（缇，twips），null 表示等宽分配；数组长度应与最宽行的单元格数一致</summary>
    public Int32[]? ColumnWidths { get; set; }

    /// <summary>原始表格 XML，非空时 Writer 直接写入（完整保留所有未建模属性）</summary>
    public String? RawXml { get; set; }
    #endregion
}
