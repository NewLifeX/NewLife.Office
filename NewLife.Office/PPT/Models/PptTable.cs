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
    #endregion
}
