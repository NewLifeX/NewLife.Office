namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 时间序列重采样操作</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>按时间列的固定日历区间分组并汇总</summary>
    /// <remarks>时间值必须为 DateTime；空时间行会被忽略。周从星期一开始，按时间值自身的日历日期分桶，不转换时区。</remarks>
    /// <param name="timeColumn">时间列名</param>
    /// <param name="bucket">时间分桶单位</param>
    /// <param name="aggregations">每个时间段要计算的指标</param>
    /// <returns>按时间段起点升序排列的汇总表</returns>
    public ExcelDataFrame Resample(String timeColumn, TimeBucket bucket, params AggregateSpec[] aggregations)
    {
        var timeIndex = GetColumnIndex(timeColumn);
        if (bucket is not (TimeBucket.Minute or TimeBucket.Hour or TimeBucket.Day or TimeBucket.Week or TimeBucket.Month or TimeBucket.Quarter or TimeBucket.Year))
            throw new ArgumentOutOfRangeException(nameof(bucket));
        if (aggregations == null) throw new ArgumentNullException(nameof(aggregations));
        if (aggregations.Length == 0) throw new ArgumentException(L.T("At least one aggregation is required."), nameof(aggregations));

        var valueIndexes = new Int32[aggregations.Length];
        var resultColumns = new String[aggregations.Length + 1];
        var names = new HashSet<String>(StringComparer.OrdinalIgnoreCase) { _columns[timeIndex] };
        resultColumns[0] = _columns[timeIndex];
        for (var i = 0; i < aggregations.Length; i++)
        {
            var spec = aggregations[i] ?? throw new ArgumentException(L.T("Aggregation definitions cannot contain null."), nameof(aggregations));
            if (spec.Function is not (AggregateFunction.Count or AggregateFunction.Sum or AggregateFunction.Average or AggregateFunction.Min or AggregateFunction.Max))
                throw new ArgumentOutOfRangeException(nameof(aggregations));

            valueIndexes[i] = spec.ColumnName == null ? -1 : GetColumnIndex(spec.ColumnName);
            var resultName = spec.ResultName ?? (spec.ColumnName == null
                ? spec.Function.ToString()
                : $"{spec.Function}_{_columns[valueIndexes[i]]}");
            if (String.IsNullOrWhiteSpace(resultName))
                throw new ArgumentException(L.T("Result column name cannot be empty."), nameof(aggregations));
            if (!names.Add(resultName))
                throw new ArgumentException(L.T("Duplicate result column: {0}", resultName), nameof(aggregations));
            resultColumns[i + 1] = resultName;
        }

        var groups = new SortedDictionary<DateTime, List<Object?[]>>();
        foreach (var row in _rows)
        {
            var value = row[timeIndex];
            if (IsMissing(value)) continue;
            if (value is not DateTime time)
                throw new ArgumentException(L.T("Time column values must be DateTime: {0}", timeColumn), nameof(timeColumn));

            var periodStart = GetPeriodStart(time, bucket);
            if (!groups.TryGetValue(periodStart, out var group))
            {
                group = [];
                groups.Add(periodStart, group);
            }
            group.Add(row);
        }

        var resultRows = new List<Object?[]>(groups.Count);
        foreach (var group in groups)
        {
            var result = new Object?[resultColumns.Length];
            result[0] = group.Key;
            for (var i = 0; i < aggregations.Length; i++)
                result[i + 1] = valueIndexes[i] < 0
                    ? group.Value.Count
                    : Aggregate(group.Value, valueIndexes[i], aggregations[i].Function, aggregations[i].ColumnName!);
            resultRows.Add(result);
        }

        return new ExcelDataFrame(resultColumns, resultRows);
    }

    private static DateTime GetPeriodStart(DateTime time, TimeBucket bucket)
    {
        return bucket switch
        {
            TimeBucket.Minute => new DateTime(time.Year, time.Month, time.Day, time.Hour, time.Minute, 0, time.Kind),
            TimeBucket.Hour => new DateTime(time.Year, time.Month, time.Day, time.Hour, 0, 0, time.Kind),
            TimeBucket.Day => time.Date,
            TimeBucket.Week => time.Date.AddDays(-((7 + (Int32)time.DayOfWeek - (Int32)DayOfWeek.Monday) % 7)),
            TimeBucket.Month => new DateTime(time.Year, time.Month, 1, 0, 0, 0, time.Kind),
            TimeBucket.Quarter => new DateTime(time.Year, ((time.Month - 1) / 3) * 3 + 1, 1, 0, 0, 0, time.Kind),
            TimeBucket.Year => new DateTime(time.Year, 1, 1, 0, 0, 0, time.Kind),
            _ => throw new ArgumentOutOfRangeException(nameof(bucket)),
        };
    }
}
