namespace NewLife.Office;

/// <summary>Excel 图表系列数据</summary>
public class ExcelChartSeries
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数值数组，与 Categories 一一对应</summary>
    public Double[] Data { get; set; } = [];
    #endregion
}
