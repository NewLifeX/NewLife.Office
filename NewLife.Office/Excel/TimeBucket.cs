namespace NewLife.Office.Excel;

/// <summary>时间分桶单位</summary>
/// <remarks>供 <see cref="ExcelDataFrame.Resample(String, TimeBucket, AggregateSpec[])"/> 按日历区间分组。</remarks>
public enum TimeBucket
{
    /// <summary>按分钟</summary>
    Minute,

    /// <summary>按小时</summary>
    Hour,

    /// <summary>按日</summary>
    Day,

    /// <summary>按周，周一为起始日</summary>
    Week,

    /// <summary>按月</summary>
    Month,

    /// <summary>按季度</summary>
    Quarter,

    /// <summary>按年</summary>
    Year,
}
