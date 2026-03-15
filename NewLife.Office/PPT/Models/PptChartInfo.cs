namespace NewLife.Office;

/// <summary>PPT 图表信息（S06-04）</summary>
public class PptChartInfo
{
    #region 属性
    /// <summary>图表编号（在 ppt/charts/chart{N}.xml 中的序号）</summary>
    public Int32 ChartNumber { get; set; }

    /// <summary>图表类型（bar/line/pie/area/scatter 等）</summary>
    public String ChartType { get; set; } = String.Empty;

    /// <summary>分类标签数组</summary>
    public String[] Categories { get; set; } = [];

    /// <summary>系列数据集合</summary>
    public List<PptChartSeriesData> Series { get; } = [];
    #endregion
}
