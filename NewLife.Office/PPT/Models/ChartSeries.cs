namespace NewLife.Office.Ppt;

/// <summary>PPT 图表系列数据</summary>
public class ChartSeries
{
    #region 属性
    /// <summary>系列名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>数据点值（Y 值）</summary>
    public Double[] Values { get; set; } = [];

    /// <summary>数据点 X 值（散点图/气泡图使用），null 时使用分类索引</summary>
    public Double[]? XValues { get; set; }

    /// <summary>系列填充色（RRGGBB，null 使用默认配色，S19）</summary>
    public String? Color { get; set; }
    #endregion
    /// <summary>浅拷贝（引用类型字段仍共享，供深拷贝器配合集合重建使用）</summary>
    /// <returns>浅拷贝实例</returns>
    internal Object ShallowClone() => MemberwiseClone();

}
