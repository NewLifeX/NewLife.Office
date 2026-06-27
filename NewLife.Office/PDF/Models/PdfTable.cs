namespace NewLife.Office;

/// <summary>PDF 表格单元格</summary>
public class PdfTableCell
{
    #region 属性
    /// <summary>单元格文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>单元格专用字体（null 表示继承行/表格默认字体）</summary>
    public OfficeFont? Font { get; set; }

    /// <summary>单元格背景色（null 表示继承行/表格背景）</summary>
    public OfficeColor? BackColor { get; set; }

    /// <summary>文本水平对齐：left/center/right</summary>
    public String Alignment { get; set; } = "left";

    /// <summary>跨列数，默认 1</summary>
    public Int32 ColSpan { get; set; } = 1;

    /// <summary>内边距（磅），null 表示继承表格默认（3pt）</summary>
    public Single? Padding { get; set; }
    #endregion
}

/// <summary>PDF 表格行</summary>
public class PdfTableRow
{
    #region 属性
    /// <summary>单元格集合</summary>
    public List<PdfTableCell> Cells { get; set; } = [];

    /// <summary>行高（磅），null 表示使用表格默认行高</summary>
    public Single? Height { get; set; }

    /// <summary>行背景色（null 表示继承表格背景或斑马纹色）</summary>
    public OfficeColor? BackColor { get; set; }

    /// <summary>是否为标题行（影响字体粗体和背景色）</summary>
    public Boolean IsHeader { get; set; }
    #endregion
}

/// <summary>PDF 表格</summary>
/// <remarks>
/// 描述 PDF 页面中的表格布局，由 PdfWriter 自动计算列宽/行高并绘制边框和文本。
/// <example>
/// <code>
/// var table = new PdfTable
/// {
///     X = 56, Y = 700,
///     Width = 483,                            // A4 可用宽度
///     ColumnWidths = [150f, 150f, 183f],
///     HeaderBackColor = OfficeColor.FromHex("4472C4"),
/// };
/// var header = new PdfTableRow { IsHeader = true };
/// header.Cells.AddRange(["姓名", "部门", "邮箱"].Select(t => new PdfTableCell { Text = t }));
/// table.Rows.Add(header);
/// </code>
/// </example>
/// </remarks>
public class PdfTable
{
    #region 属性
    /// <summary>表格左上角 X 坐标（从页面左边量起，单位磅）</summary>
    public Single X { get; set; }

    /// <summary>表格左上角 Y 坐标（从页面顶部向下量起，单位磅）</summary>
    public Single Y { get; set; }

    /// <summary>表格总宽度（单位磅）</summary>
    public Single Width { get; set; }

    /// <summary>各列宽度数组（单位磅），总和应等于 Width；为空时等宽分配</summary>
    public Single[] ColumnWidths { get; set; } = [];

    /// <summary>默认行高（单位磅），默认 20pt</summary>
    public Single RowHeight { get; set; } = 20f;

    /// <summary>单元格内边距（单位磅），默认 3pt</summary>
    public Single CellPadding { get; set; } = 3f;

    /// <summary>是否绘制边框</summary>
    public Boolean ShowBorders { get; set; } = true;

    /// <summary>边框颜色，null 表示黑色</summary>
    public OfficeColor? BorderColor { get; set; }

    /// <summary>表头行背景色，null 表示不着色</summary>
    public OfficeColor? HeaderBackColor { get; set; }

    /// <summary>表头文字颜色，null 表示继承默认</summary>
    public OfficeColor? HeaderForeColor { get; set; }

    /// <summary>斑马纹交替行背景色，null 表示不着色</summary>
    public OfficeColor? StripeColor { get; set; }

    /// <summary>表格默认字体，null 表示继承 PdfWriter 当前字体</summary>
    public OfficeFont? DefaultFont { get; set; }

    /// <summary>表格行集合</summary>
    public List<PdfTableRow> Rows { get; set; } = [];
    #endregion
}
