namespace NewLife.Office;

/// <summary>Excel 数据透视表字段配置</summary>
public class PivotField
{
    #region 属性
    /// <summary>字段名称（对应源数据表头）</summary>
    public String Name { get; set; } = null!;

    /// <summary>是否作为行字段</summary>
    public Boolean IsRowField { get; set; }

    /// <summary>是否作为列字段</summary>
    public Boolean IsColumnField { get; set; }

    /// <summary>是否作为数据字段（汇总计算）</summary>
    public Boolean IsDataField { get; set; }

    /// <summary>是否作为筛选字段（报表筛选区）</summary>
    public Boolean IsFilterField { get; set; }

    /// <summary>数据字段汇总函数（仅 IsDataField=true 时有效）</summary>
    public PivotSummaryFunction SummaryFunction { get; set; } = PivotSummaryFunction.Sum;

    /// <summary>数据字段显示名称（可空，空时自动生成）</summary>
    public String? Caption { get; set; }
    #endregion
}
