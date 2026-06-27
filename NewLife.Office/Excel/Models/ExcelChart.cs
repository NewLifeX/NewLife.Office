namespace NewLife.Office;

/// <summary>Excel 图表定义</summary>
/// <remarks>
/// 描述嵌入工作表中的图表元素，通过 ExcelWriter.AddChart 写入。
/// <example>
/// <code>
/// var chart = new ExcelChart
/// {
///     Type = "bar",
///     Title = "月度销售额",
///     Categories = ["1月", "2月", "3月"],
///     Series =
///     [
///         new ExcelChartSeries { Name = "产品A", Data = [120, 150, 180] },
///         new ExcelChartSeries { Name = "产品B", Data = [80, 95, 110] },
///     ],
/// };
/// writer.AddChart(null, chart);
/// </code>
/// </example>
/// </remarks>
public class ExcelChart
{
    #region 属性
    /// <summary>图表类型：bar（柱状）/ line（折线）/ pie（饼图）</summary>
    public String Type { get; set; } = "bar";

    /// <summary>图表标题（可选）</summary>
    public String? Title { get; set; }

    /// <summary>分类轴标签数组</summary>
    public String[]? Categories { get; set; }

    /// <summary>数据系列集合</summary>
    public List<ExcelChartSeries> Series { get; set; } = [];

    /// <summary>图表起始行（0基，图表放置于该行之后）</summary>
    public Int32 AnchorRow { get; set; }

    /// <summary>图表起始列（0基）</summary>
    public Int32 AnchorCol { get; set; }

    /// <summary>图表宽度（像素，默认 480）</summary>
    public Int32 WidthPx { get; set; } = 480;

    /// <summary>图表高度（像素，默认 300）</summary>
    public Int32 HeightPx { get; set; } = 300;
    #endregion
}
