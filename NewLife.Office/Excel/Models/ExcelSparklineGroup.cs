namespace NewLife.Office;

/// <summary>迷你图组（对应 OOXML x14:sparklineGroup 元素）</summary>
/// <remarks>
/// Excel 2010+ 迷你图（Sparkline）是单元格中的微型图表。
/// 一个 SparklineGroup 包含相同类型的一组迷你图，共享数据范围和样式。
/// 通过 <see cref="ExcelWriter.AddSparklineGroup"/> 写入，<see cref="ExcelReader.ReadSparklines"/> 读取。
/// </remarks>
public class ExcelSparklineGroup
{
    #region 属性
    /// <summary>类型：line（折线图）、column（柱状图）、stacked（盈亏图）</summary>
    public String Type { get; set; } = "line";

    /// <summary>数据区域（如 "Sheet1!B2:F2"）</summary>
    public String DataRange { get; set; } = String.Empty;

    /// <summary>放置单元格区域（如 "Sheet1!G2"）</summary>
    public String CellRange { get; set; } = String.Empty;

    /// <summary>线条/柱颜色（16进制RGB，如 "FF0000"）</summary>
    public String? LineColor { get; set; }

    /// <summary>是否显示标记点</summary>
    public Boolean ShowMarkers { get; set; }

    /// <summary>迷你图位置列表（每个迷你图对应一个输出单元格）</summary>
    public List<String> Sparklines { get; set; } = [];
    #endregion
}
