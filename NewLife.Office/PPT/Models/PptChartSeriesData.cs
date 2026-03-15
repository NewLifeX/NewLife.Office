namespace NewLife.Office;

/// <summary>PPT 图表系列数据（S06-04）</summary>
public class PptChartSeriesData
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据值数组，对应各分类</summary>
    public Double[] Values { get; set; } = [];
    #endregion
}
