namespace NewLife.Office.Excel;

/// <summary>Excel 单个工作表完整数据快照</summary>
/// <remarks>
/// 包含工作表的行数据、单元格样式和全部元数据。
/// 所有行列索引均为 0 基。
/// </remarks>
public class Worksheet
{
    #region 属性
    /// <summary>工作表名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>行数据（每行为对象数组，第一行通常是表头）</summary>
    public List<Object?[]> Rows { get; set; } = [];

    /// <summary>
    /// 每行对应的实际 Excel 行号（0基）。
    /// 当源文件存在跳行（如第13行为空行不存在于 sheetData 中）时，此列表记录每一项在 Rows 里
    /// 对应的真实行号，WriteExcel 使用它还原原始行布局（含空行间隔）。
    /// 若为 null 则表示 Rows 是连续的（行号等于下标），无需特殊处理。
    /// </summary>
    public List<Int32>? ActualRowNumbers { get; set; }

    /// <summary>单元格样式映射（行, 列）→ 样式</summary>
    public Dictionary<(Int32 Row, Int32 Col), CellFormat> CellFormats { get; set; } = [];

    /// <summary>合并单元格区域（起始行, 起始列, 结束行, 结束列），均为0基</summary>
    public List<(Int32 StartRow, Int32 StartCol, Int32 EndRow, Int32 EndCol)> Merges { get; set; } = [];

    /// <summary>冻结窗格（行数, 列数），null 表示未冻结</summary>
    public (Int32 Rows, Int32 Cols)? FreezePane { get; set; }

    /// <summary>自动筛选范围（如 "A1:F1"）</summary>
    public String? AutoFilter { get; set; }

    /// <summary>是否显示网格线（默认 true）</summary>
    public Boolean Gridlines { get; set; } = true;

    /// <summary>视图缩放比例（百分比，默认 100）</summary>
    public Int32 ZoomScale { get; set; } = 100;

    /// <summary>默认行高（磅值），null 表示使用 Excel 默认</summary>
    public Double? DefaultRowHeight { get; set; }

    /// <summary>行高（0基行号 → 磅值）</summary>
    public Dictionary<Int32, Double> RowHeights { get; set; } = [];

    /// <summary>列宽（0基列号 → 字符宽度）</summary>
    public Dictionary<Int32, Double> ColumnWidths { get; set; } = [];

    /// <summary>隐藏行（0基行号集合），源文件 row hidden="1" 的行</summary>
    public HashSet<Int32> HiddenRows { get; set; } = [];

    /// <summary>隐藏列（0基列号集合），源文件 &lt;col hidden="1"&gt; 的列</summary>
    public HashSet<Int32> HiddenColumns { get; set; } = [];

    /// <summary>超链接（0基行, 0基列）→ (URL, 显示文本)</summary>
    public Dictionary<(Int32 Row, Int32 Col), (String Url, String? Display)> Hyperlinks { get; set; } = [];

    /// <summary>图片集合</summary>
    public List<ExcelImage> Images { get; set; } = [];

    /// <summary>页面方向</summary>
    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

    /// <summary>纸张大小</summary>
    public PaperSize PaperSize { get; set; } = PaperSize.Default;

    /// <summary>上边距（英寸）</summary>
    public Double MarginTop { get; set; } = 0.75;

    /// <summary>下边距（英寸）</summary>
    public Double MarginBottom { get; set; } = 0.75;

    /// <summary>左边距（英寸）</summary>
    public Double MarginLeft { get; set; } = 0.7;

    /// <summary>右边距（英寸）</summary>
    public Double MarginRight { get; set; } = 0.7;

    /// <summary>页眉文本（可含 &amp;P 页码等控制符）</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本（可含 &amp;P 页码等控制符）</summary>
    public String? FooterText { get; set; }

    /// <summary>打印标题起始行（1基），0 表示未设置</summary>
    public Int32 PrintTitleStartRow { get; set; }

    /// <summary>打印标题结束行（1基），0 表示未设置</summary>
    public Int32 PrintTitleEndRow { get; set; }

    /// <summary>打印区域（如 "$A$1:$F$20"），null 表示未设置</summary>
    public String? PrintArea { get; set; }

    /// <summary>水平分页符行号（1基），此行开始新页</summary>
    public List<Int32> RowPageBreaks { get; set; } = [];

    /// <summary>垂直分页符列号（1基），此列开始新页</summary>
    public List<Int32> ColumnPageBreaks { get; set; } = [];

    /// <summary>工作表保护密码哈希，null 表示未保护</summary>
    public String? ProtectionPassword { get; set; }

    /// <summary>条件格式集合</summary>
    public List<ConditionalFormatting> ConditionalFormats { get; set; } = [];

    /// <summary>批注（0基行, 0基列）→ 批注模型（含尺寸/可见性）</summary>
    public Dictionary<(Int32 Row, Int32 Col), Comment> Comments { get; set; } = [];

    /// <summary>数据验证集合</summary>
    public List<DataValidation> Validations { get; set; } = [];

    /// <summary>公式（0基行, 0基列）→ 公式文本（不含等号）</summary>
    public Dictionary<(Int32 Row, Int32 Col), String> Formulas { get; set; } = [];

    /// <summary>图表集合（嵌入工作表的图表，由 ExcelWriter.AddChart 写入）</summary>
    public List<ExcelChart> Charts { get; set; } = [];

    /// <summary>结构化表格集合（由 ExcelWriter.AddTable 写入）</summary>
    public List<Table> Tables { get; set; } = [];

    /// <summary>迷你图组集合（由 ExcelWriter.AddSparklineGroup 写入）</summary>
    public List<SparklineGroup> SparklineGroups { get; set; } = [];
    #endregion
}
