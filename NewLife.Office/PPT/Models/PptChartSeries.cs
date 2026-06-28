namespace NewLife.Office;

/// <summary>PPT 图表系列数据</summary>
public class PptChartSeries
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据点值（Y 值）</summary>
    public Double[] Values { get; set; } = [];

    /// <summary>数据点 X 值（散点图/气泡图使用），null 时使用分类索引</summary>
    public Double[]? XValues { get; set; }
    #endregion
}
