namespace NewLife.Office;

/// <summary>PPT 幻灯片表格</summary>
public class PptTable
{
    #region 属性
    /// <summary>左边距（EMU）</summary>
    public Int64 Left { get; set; }

    /// <summary>上边距（EMU）</summary>
    public Int64 Top { get; set; }

    /// <summary>宽度（EMU）</summary>
    public Int64 Width { get; set; } = 8000000;

    /// <summary>高度（EMU）</summary>
    public Int64 Height { get; set; } = 3000000;

    /// <summary>行列数据</summary>
    public List<String[]> Rows { get; } = [];

    /// <summary>首行是否表头</summary>
    public Boolean FirstRowHeader { get; set; } = true;

    /// <summary>各列宽度（EMU），数组长度等于列数；空时按总宽平均分配</summary>
    public Int64[] ColWidths { get; set; } = [];

    /// <summary>单元格样式字典，键为 (行索引, 列索引)，优先级高于行级默认样式</summary>
    public Dictionary<(Int32 Row, Int32 Col), PptCellStyle> CellStyles { get; } = [];

    /// <summary>单元格合并字典（S11-01），键为 (起始行, 起始列)，值为 (跨列数, 跨行数)</summary>
    public Dictionary<(Int32 Row, Int32 Col), (Int32 ColSpan, Int32 RowSpan)> MergedCells { get; } = [];

    /// <summary>单元格边框样式（S11-02），键为 (行索引, 列索引)</summary>
    public Dictionary<(Int32 Row, Int32 Col), PptCellBorder> CellBorders { get; } = [];

    /// <summary>表格样式主题引用 GUID（S11-04），如 "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"（默认中等样式2）</summary>
    public String? TableStyleGuid { get; set; }
    #endregion

    #region 方法
    /// <summary>添加行</summary>
    /// <param name="cells">行中各单元格文本</param>
    public void AddRow(String[] cells)
    {
        Rows.Add(cells);
    }

    /// <summary>删除行</summary>
    /// <param name="index">0基行索引</param>
    public void RemoveRow(Int32 index)
    {
        if (index < 0 || index >= Rows.Count) return;
        Rows.RemoveAt(index);
    }

    /// <summary>在指定位置插入列</summary>
    /// <param name="index">0基列索引（插入位置，原列及之后右移）</param>
    /// <param name="header">列头文本（仅当 FirstRowHeader 时用于首行）</param>
    public void AddColumn(Int32 index, String? header = null)
    {
        for (var r = 0; r < Rows.Count; r++)
        {
            var newRow = new List<String>(Rows[r]);
            var cellText = r == 0 && FirstRowHeader && header != null ? header : String.Empty;
            if (index >= newRow.Count)
                newRow.Add(cellText);
            else
                newRow.Insert(index, cellText);
            Rows[r] = newRow.ToArray();
        }
    }
    #endregion
}
