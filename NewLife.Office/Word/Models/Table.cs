using System.Text;

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

    /// <summary>表格样式引用 ID（w:tblPr/w:tblStyle w:val，如 "TableGrid"），Reader 保留用于写回</summary>
    public String? StyleId { get; set; }

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

    #region 行列操作
    /// <summary>追加一行（空行，可在返回的 <see cref="TableRow"/> 上填充单元格）</summary>
    /// <returns>新增行</returns>
    public TableRow AddRow()
    {
        var row = new TableRow();
        Rows.Add(row);
        return row;
    }

    /// <summary>在指定位置插入一行（空行）</summary>
    /// <param name="index">插入位置（0=最前），超出范围则追加到末尾</param>
    /// <returns>新增行</returns>
    public TableRow InsertRow(Int32 index)
    {
        var row = new TableRow();
        if (index < 0) index = 0;
        if (index > Rows.Count) index = Rows.Count;
        Rows.Insert(index, row);
        return row;
    }

    /// <summary>移除指定行</summary>
    /// <param name="index">行索引</param>
    public void RemoveRow(Int32 index)
    {
        if (index >= 0 && index < Rows.Count) Rows.RemoveAt(index);
    }

    /// <summary>为每行追加一列（空单元格）</summary>
    /// <param name="cellIndex">可选：插入到指定列位置（默认追加到末尾）</param>
    public void AddColumn(Int32? cellIndex = null)
    {
        foreach (var row in Rows)
        {
            var cell = new Cell { Paragraphs = { new Paragraph() } };
            if (cellIndex.HasValue)
            {
                var idx = cellIndex.Value < 0 ? 0 : cellIndex.Value > row.Cells.Count ? row.Cells.Count : cellIndex.Value;
                row.Cells.Insert(idx, cell);
            }
            else
            {
                row.Cells.Add(cell);
            }
        }
    }

    /// <summary>从每行移除一列</summary>
    /// <param name="index">列索引</param>
    public void RemoveColumn(Int32 index)
    {
        foreach (var row in Rows)
        {
            if (index >= 0 && index < row.Cells.Count) row.Cells.RemoveAt(index);
        }
    }
    #endregion

    #region 单元格访问
    /// <summary>获取单元格（行/列越界返回 null）</summary>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <returns>单元格，越界为 null</returns>
    public Cell? GetCell(Int32 row, Int32 col)
    {
        if (row < 0 || row >= Rows.Count) return null;
        var cells = Rows[row].Cells;
        if (col < 0 || col >= cells.Count) return null;
        return cells[col];
    }

    /// <summary>获取单元格文本（拼接段落与 Run，单元格间无分隔）</summary>
    /// <param name="row">行索引</param>
    /// <param name="col">列索引</param>
    /// <returns>单元格纯文本，越界或空为 String.Empty</returns>
    public String GetCellText(Int32 row, Int32 col)
    {
        var cell = GetCell(row, col);
        if (cell == null) return String.Empty;
        var sb = new StringBuilder();
        foreach (var p in cell.Paragraphs)
            foreach (var r in p.Runs)
                sb.Append(r.Text);
        return sb.ToString();
    }

    /// <summary>设置单元格文本（复用首个 Run 保留格式；单元格不存在时自动补齐行/列）</summary>
    /// <param name="row">行索引（超出时追加空行）</param>
    /// <param name="col">列索引（超出时追加空单元格）</param>
    /// <param name="text">文本内容</param>
    /// <returns>目标单元格</returns>
    public Cell SetCellText(Int32 row, Int32 col, String text)
    {
        // 补齐行
        while (Rows.Count <= row) Rows.Add(new TableRow());
        var cells = Rows[row].Cells;
        // 补齐列
        while (cells.Count <= col)
            cells.Add(new Cell { Paragraphs = { new Paragraph() } });
        var cell = cells[col];

        // 设置文本：复用首个 Run（保留其格式），无 Run 则新建
        if (cell.Paragraphs.Count == 0)
            cell.Paragraphs.Add(new Paragraph());
        var para = cell.Paragraphs[0];
        if (para.Runs.Count == 0)
            para.Runs.Add(new Run());
        para.Runs[0].Text = text;
        // 清空其余 Run（避免旧文本残留）
        for (var i = 1; i < para.Runs.Count; i++)
            para.Runs[i].Text = String.Empty;
        return cell;
    }
    #endregion
}
