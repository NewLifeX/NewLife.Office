namespace NewLife.Office;

/// <summary>PPT 图表系列数据</summary>
public class PptChartSeries
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据点值</summary>
    public Double[] Values { get; set; } = [];
    #endregion
}
