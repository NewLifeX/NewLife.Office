namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 透视与反透视操作</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>按行键和列键创建透视表</summary>
    /// <param name="indexColumn">结果行键列</param>
    /// <param name="pivotColumn">用于生成结果列名的列</param>
    /// <param name="valueColumn">待汇总的值列</param>
    /// <param name="function">重复行键与列键组合的汇总方式</param>
    /// <returns>透视后的新数据表</returns>
    public ExcelDataFrame Pivot(String indexColumn, String pivotColumn, String valueColumn, AggregateFunction function = AggregateFunction.Sum)
        => Pivot([indexColumn], pivotColumn, valueColumn, function);

    /// <summary>按多个行键和一个列键创建透视表</summary>
    /// <param name="indexColumns">结果行键列</param>
    /// <param name="pivotColumn">用于生成结果列名的列</param>
    /// <param name="valueColumn">待汇总的值列</param>
    /// <param name="function">重复行键与列键组合的汇总方式</param>
    /// <returns>透视后的新数据表</returns>
    public ExcelDataFrame Pivot(String[] indexColumns, String pivotColumn, String valueColumn, AggregateFunction function = AggregateFunction.Sum)
    {
        if (indexColumns == null) throw new ArgumentNullException(nameof(indexColumns));
        if (indexColumns.Length == 0) throw new ArgumentException(L.T("At least one index column is required."), nameof(indexColumns));
        if (function is not (AggregateFunction.Count or AggregateFunction.Sum or AggregateFunction.Average or AggregateFunction.Min or AggregateFunction.Max))
            throw new ArgumentOutOfRangeException(nameof(function));

        var indexIndexes = ResolveJoinKeys(indexColumns, nameof(indexColumns));
        var pivotIndex = GetColumnIndex(pivotColumn);
        var valueIndex = GetColumnIndex(valueColumn);
        var categories = new List<String>();
        var categoryIndexes = new Dictionary<String, Int32>(StringComparer.OrdinalIgnoreCase);
        var groups = new Dictionary<CompositeKey, PivotRow>();
        var orderedRows = new List<PivotRow>();

        foreach (var source in _rows)
        {
            var pivotValue = source[pivotIndex];
            if (IsMissing(pivotValue)) continue;

            var category = Convert.ToString(pivotValue, System.Globalization.CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(category))
                throw new ArgumentException(L.T("Pivot column values cannot be empty."), nameof(pivotColumn));
            if (!categoryIndexes.TryGetValue(category!, out var categoryIndex))
            {
                categoryIndex = categories.Count;
                categories.Add(category!);
                categoryIndexes.Add(category!, categoryIndex);
            }

            var keys = new Object?[indexIndexes.Length];
            for (var i = 0; i < indexIndexes.Length; i++)
                keys[i] = IsMissing(source[indexIndexes[i]]) ? null : source[indexIndexes[i]];
            var key = new CompositeKey(keys);
            if (!groups.TryGetValue(key, out var group))
            {
                group = new PivotRow(keys, categories.Count);
                groups.Add(key, group);
                orderedRows.Add(group);
            }
            group.EnsureCategoryCount(categories.Count);
            group.Values[categoryIndex].Add(source[valueIndex]);
        }

        var resultColumns = new String[indexIndexes.Length + categories.Count];
        var names = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < indexIndexes.Length; i++)
        {
            resultColumns[i] = _columns[indexIndexes[i]];
            names.Add(resultColumns[i]);
        }
        for (var i = 0; i < categories.Count; i++)
        {
            if (!names.Add(categories[i]))
                throw new ArgumentException(L.T("Pivot result column conflicts with an index column: {0}", categories[i]), nameof(pivotColumn));
            resultColumns[indexIndexes.Length + i] = categories[i];
        }

        var resultRows = new List<Object?[]>(orderedRows.Count);
        foreach (var group in orderedRows)
        {
            var row = new Object?[resultColumns.Length];
            Array.Copy(group.Keys, row, group.Keys.Length);
            for (var i = 0; i < categories.Count; i++)
                if (group.Values[i].Count > 0)
                    row[indexIndexes.Length + i] = AggregateValues(group.Values[i], function, valueColumn);
            resultRows.Add(row);
        }

        return new ExcelDataFrame(resultColumns, resultRows);
    }

    /// <summary>将多列值转换为变量和值两列，保留空值</summary>
    /// <param name="idColumns">保留为标识列的列名</param>
    /// <param name="valueColumns">展开为变量和值行的列名</param>
    /// <param name="variableColumn">变量名称列名</param>
    /// <param name="valueColumn">变量值列名</param>
    /// <returns>反透视后的新数据表</returns>
    public ExcelDataFrame Melt(String[] idColumns, String[] valueColumns, String variableColumn = "Variable", String valueColumn = "Value")
    {
        if (idColumns == null) throw new ArgumentNullException(nameof(idColumns));
        if (valueColumns == null) throw new ArgumentNullException(nameof(valueColumns));
        if (valueColumns.Length == 0) throw new ArgumentException(L.T("At least one value column is required."), nameof(valueColumns));
        if (String.IsNullOrWhiteSpace(variableColumn)) throw new ArgumentException(L.T("Variable column name cannot be empty."), nameof(variableColumn));
        if (String.IsNullOrWhiteSpace(valueColumn)) throw new ArgumentException(L.T("Value column name cannot be empty."), nameof(valueColumn));

        var idIndexes = ResolveNamedColumns(idColumns, nameof(idColumns));
        var valueIndexes = ResolveNamedColumns(valueColumns, nameof(valueColumns));
        var names = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        var resultColumns = new String[idColumns.Length + 2];
        for (var i = 0; i < idIndexes.Length; i++)
        {
            resultColumns[i] = _columns[idIndexes[i]];
            names.Add(resultColumns[i]);
        }
        if (!names.Add(variableColumn))
            throw new ArgumentException(L.T("Melt result column conflicts: {0}", variableColumn), nameof(variableColumn));
        if (!names.Add(valueColumn))
            throw new ArgumentException(L.T("Melt result column conflicts: {0}", valueColumn), nameof(valueColumn));
        resultColumns[^2] = variableColumn;
        resultColumns[^1] = valueColumn;

        var rows = new List<Object?[]>();
        foreach (var source in _rows)
        {
            foreach (var valueIndex in valueIndexes)
            {
                var row = new Object?[resultColumns.Length];
                for (var i = 0; i < idIndexes.Length; i++)
                    row[i] = source[idIndexes[i]];
                row[^2] = _columns[valueIndex];
                row[^1] = source[valueIndex];
                rows.Add(row);
            }
        }

        return new ExcelDataFrame(resultColumns, rows);
    }

    private Int32[] ResolveNamedColumns(String[] columns, String parameterName)
    {
        var indexes = new Int32[columns.Length];
        var unique = new HashSet<Int32>();
        for (var i = 0; i < columns.Length; i++)
        {
            indexes[i] = GetColumnIndex(columns[i]);
            if (!unique.Add(indexes[i]))
                throw new ArgumentException(L.T("Duplicate column: {0}", columns[i]), parameterName);
        }
        return indexes;
    }

    private static Object? AggregateValues(List<Object?> values, AggregateFunction function, String columnName)
    {
        var count = 0;
        var sum = 0d;
        Object? minimum = null;
        Object? maximum = null;
        foreach (var value in values)
        {
            if (IsMissing(value)) continue;
            count++;
            if (function is AggregateFunction.Sum or AggregateFunction.Average)
                sum += ToNumber(value!, columnName);
            if (minimum == null || new CellComparer(false).Compare(value, minimum) < 0) minimum = value;
            if (maximum == null || new CellComparer(false).Compare(value, maximum) > 0) maximum = value;
        }

        return function switch
        {
            AggregateFunction.Count => count,
            AggregateFunction.Sum => sum,
            AggregateFunction.Average => count > 0 ? sum / count : null,
            AggregateFunction.Min => minimum,
            AggregateFunction.Max => maximum,
            _ => throw new ArgumentOutOfRangeException(nameof(function)),
        };
    }

    private sealed class PivotRow(Object?[] keys, Int32 categoryCount)
    {
        public Object?[] Keys { get; } = keys;
        public List<List<Object?>> Values { get; } = Enumerable.Range(0, categoryCount).Select(_ => new List<Object?>()).ToList();

        public void EnsureCategoryCount(Int32 count)
        {
            while (Values.Count < count)
                Values.Add([]);
        }
    }
}
