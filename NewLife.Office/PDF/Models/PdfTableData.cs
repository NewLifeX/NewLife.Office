namespace NewLife.Office.Pdf;

/// <summary>PDF 表格提取行（单元格文本列表）</summary>
public class PdfTableRowData
{
    /// <summary>单元格文本（按列顺序，空单元格为空字符串）</summary>
    public List<String> Cells { get; } = [];

    /// <summary>单元格数量</summary>
    public Int32 CellCount => Cells.Count;
}

/// <summary>PDF 结构化表格提取结果（P09，对标 iText 7/Aspose.PDF）</summary>
/// <remarks>
/// 从 PDF 页面中提取的表格数据：包含页索引与行列文本网格。
/// 通过文本块坐标聚类识别列边界，按行还原单元格文本。
/// </remarks>
public class PdfTableData
{
    /// <summary>页面索引（0 起）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>行列表</summary>
    public List<PdfTableRowData> Rows { get; } = [];

    /// <summary>行数</summary>
    public Int32 RowCount => Rows.Count;

    /// <summary>列数（首行单元格数）</summary>
    public Int32 ColumnCount => Rows.Count > 0 ? Rows[0].CellCount : 0;
}
