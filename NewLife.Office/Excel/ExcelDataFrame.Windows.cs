namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 累计与滚动统计操作</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>按当前行顺序计算累计求和；空值对应结果为空且不影响后续累计</summary>
    /// <param name="columnName">数值列名</param>
    /// <param name="resultColumn">结果列名；默认使用 CumulativeSum_列名</param>
    /// <returns>追加累计结果列后的新数据表</returns>
    public ExcelDataFrame CumulativeSum(String columnName, String? resultColumn = null)
    {
        var index = GetColumnIndex(columnName);
        var name = resultColumn ?? $"CumulativeSum_{_columns[index]}";
        var sum = 0d;
        return AddColumn(name, row =>
        {
            var value = row[columnName];
            if (IsMissing(value)) return null;

            sum += ToNumber(value!, columnName);
            return sum;
        });
    }

    /// <summary>计算固定行数窗口内的滚动平均值</summary>
    /// <param name="columnName">数值列名</param>
    /// <param name="windowSize">窗口行数</param>
    /// <param name="minPeriods">窗口内至少需要的非空数值数；默认要求完整窗口</param>
    /// <param name="resultColumn">结果列名；默认使用 RollingAverage_列名</param>
    /// <returns>追加滚动平均列后的新数据表</returns>
    public ExcelDataFrame RollingAverage(String columnName, Int32 windowSize, Int32? minPeriods = null, String? resultColumn = null)
        => Rolling(columnName, windowSize, minPeriods, resultColumn, average: true);

    /// <summary>计算固定行数窗口内的滚动求和</summary>
    /// <param name="columnName">数值列名</param>
    /// <param name="windowSize">窗口行数</param>
    /// <param name="minPeriods">窗口内至少需要的非空数值数；默认要求完整窗口</param>
    /// <param name="resultColumn">结果列名；默认使用 RollingSum_列名</param>
    /// <returns>追加滚动求和列后的新数据表</returns>
    public ExcelDataFrame RollingSum(String columnName, Int32 windowSize, Int32? minPeriods = null, String? resultColumn = null)
        => Rolling(columnName, windowSize, minPeriods, resultColumn, average: false);

    private ExcelDataFrame Rolling(String columnName, Int32 windowSize, Int32? minPeriods, String? resultColumn, Boolean average)
    {
        if (windowSize <= 0) throw new ArgumentOutOfRangeException(nameof(windowSize));
        var required = minPeriods ?? windowSize;
        if (required <= 0 || required > windowSize)
            throw new ArgumentOutOfRangeException(nameof(minPeriods));

        var index = GetColumnIndex(columnName);
        var name = resultColumn ?? $"{(average ? "RollingAverage" : "RollingSum")}_{_columns[index]}";
        var window = new Queue<Double?>();
        var sum = 0d;
        var count = 0;
        return AddColumn(name, row =>
        {
            var value = row[columnName];
            var number = IsMissing(value) ? (Double?)null : ToNumber(value!, columnName);
            window.Enqueue(number);
            if (number.HasValue)
            {
                sum += number.Value;
                count++;
            }

            if (window.Count > windowSize)
            {
                var removed = window.Dequeue();
                if (removed.HasValue)
                {
                    sum -= removed.Value;
                    count--;
                }
            }

            if (count < required) return null;
            return average ? sum / count : sum;
        });
    }
}
