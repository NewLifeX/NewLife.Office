namespace NewLife.Office.Word;

/// <summary>表格单元格</summary>
public class Cell
{
    #region 属性
    /// <summary>段落集合</summary>
    public List<Paragraph> Paragraphs { get; } = [];

    /// <summary>嵌套表格（w:tc 内的 w:tbl，W46），Writer 在段落前输出</summary>
    public List<Table> NestedTables { get; set; } = [];

    /// <summary>背景色（16进制 RGB）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>单元格四边边框（w:tcBorders），null 表示使用表格默认边框</summary>
    public TableBorders? Borders { get; set; }

    /// <summary>合并列数</summary>
    public Int32 ColSpan { get; set; } = 1;

    /// <summary>
    /// 合并行数（垂直合并）：1=普通（默认），0=继续合并（<c>w:vMerge</c> 无 val），
    /// -1=合并起点（<c>w:vMerge w:val="restart"</c>），&gt;1=程序化指定的实际跨行数。
    /// </summary>
    public Int32 RowSpan { get; set; } = 1;

    /// <summary>单元格宽度（缇，twips，对应 w:tcW），null 表示自动</summary>
    public Int32? Width { get; set; }

    /// <summary>单元格垂直对齐（"top"/"center"/"bottom"），null 表示继承默认</summary>
    public String? VerticalAlignment { get; set; }
    #endregion
}
