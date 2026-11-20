namespace NewLife.Office.Excel;

/// <summary>分组汇总定义</summary>
/// <remarks>
/// 描述“对哪一列、用什么函数、结果叫什么名字”。可直接构造，也可用
/// <see cref="Sum(String, String)"/>、<see cref="Average(String, String)"/> 等工厂方法，
/// 或二/三元素元组简写：
/// <code>
/// using static NewLife.Office.Excel.AggregateSpec;
///
/// var summary = frame.GroupBy("地区", "月份")
///     .Aggregate(
///         Sum("销售额", "销售额合计"),
///         Average("含税金额", "平均含税金额"));
///
/// var compact = frame.GroupBy("地区").Aggregate(
///     ("销售额", AggregateFunction.Sum, "销售额合计"),
///     ("数量", AggregateFunction.Count));
/// </code>
/// </remarks>
public sealed class AggregateSpec
{
    /// <summary>实例化分组汇总定义</summary>
    /// <param name="columnName">汇总列名；仅 <see cref="AggregateFunction.Count"/> 可传 null，此时统计组内行数</param>
    /// <param name="function">汇总函数</param>
    /// <param name="resultName">结果列名；null 时按函数和源列名生成</param>
    public AggregateSpec(String? columnName, AggregateFunction function, String? resultName = null)
    {
        if (columnName == null && function != AggregateFunction.Count)
            throw new ArgumentException(L.T("A column is required for this aggregation."), nameof(columnName));
        if (resultName != null && resultName.Length == 0)
            throw new ArgumentException(L.T("Result column name cannot be empty."), nameof(resultName));

        ColumnName = columnName;
        Function = function;
        ResultName = resultName;
    }

    /// <summary>汇总列名；null 表示统计行数</summary>
    public String? ColumnName { get; }

    /// <summary>汇总函数</summary>
    public AggregateFunction Function { get; }

    /// <summary>结果列名</summary>
    public String? ResultName { get; }

    /// <summary>创建数值求和定义</summary>
    /// <param name="columnName">数值列名</param>
    /// <param name="resultName">结果列名；null 时生成 Sum_列名</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec Sum(String columnName, String? resultName = null) => Create(columnName, AggregateFunction.Sum, resultName);

    /// <summary>创建数值平均值定义</summary>
    /// <param name="columnName">数值列名</param>
    /// <param name="resultName">结果列名；null 时生成 Average_列名</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec Average(String columnName, String? resultName = null) => Create(columnName, AggregateFunction.Average, resultName);

    /// <summary>创建最小值定义</summary>
    /// <param name="columnName">列名</param>
    /// <param name="resultName">结果列名；null 时生成 Min_列名</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec Min(String columnName, String? resultName = null) => Create(columnName, AggregateFunction.Min, resultName);

    /// <summary>创建最大值定义</summary>
    /// <param name="columnName">列名</param>
    /// <param name="resultName">结果列名；null 时生成 Max_列名</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec Max(String columnName, String? resultName = null) => Create(columnName, AggregateFunction.Max, resultName);

    /// <summary>创建非空值数量定义</summary>
    /// <param name="columnName">待统计列名</param>
    /// <param name="resultName">结果列名；null 时生成 Count_列名</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec Count(String columnName, String? resultName = null) => Create(columnName, AggregateFunction.Count, resultName);

    /// <summary>创建组内行数统计定义</summary>
    /// <param name="resultName">结果列名，默认 Count</param>
    /// <returns>汇总定义</returns>
    public static AggregateSpec CountRows(String resultName = "Count") => new(null, AggregateFunction.Count, resultName);

    /// <summary>由“列名 + 汇总函数”元组隐式创建定义</summary>
    /// <param name="value">列名与汇总函数</param>
    public static implicit operator AggregateSpec((String Column, AggregateFunction Function) value)
        => new(value.Column, value.Function);

    /// <summary>由“列名 + 汇总函数 + 结果列名”元组隐式创建定义</summary>
    /// <param name="value">列名、汇总函数与结果列名</param>
    public static implicit operator AggregateSpec((String Column, AggregateFunction Function, String? ResultName) value)
        => new(value.Column, value.Function, value.ResultName);

    private static AggregateSpec Create(String columnName, AggregateFunction function, String? resultName)
    {
        if (String.IsNullOrWhiteSpace(columnName)) throw new ArgumentException(L.T("Column name cannot be empty."), nameof(columnName));

        return new AggregateSpec(columnName, function, resultName);
    }
}
