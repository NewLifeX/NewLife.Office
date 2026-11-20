using System.Globalization;

namespace NewLife.Office.Excel;

/// <summary>面向 Excel 表格的轻量数据分析模型</summary>
/// <remarks>
/// 以首行为列名，将工作表数据载入内存，提供筛选、分组汇总、时间重采样和常见统计。
/// 不计算 Excel 公式；公式结果以工作簿中已保存的缓存值为准。
/// <example>
/// <code>
/// using var reader = new ExcelReader("sales.xlsx");
/// var sheet = reader.ReadExcel().GetSheet("销售");
/// var summary = ExcelDataFrame.FromWorksheet(sheet)
///     .Where(row => Convert.ToDouble(row["销售额"]) &gt; 0)
///     .GroupBy("地区")
///     .Sum("销售额");
/// </code>
/// </example>
/// </remarks>
public sealed partial class ExcelDataFrame
{
    #region 属性
    private readonly String[] _columns;
    private readonly List<Object?[]> _rows;
    private readonly Dictionary<String, Int32> _columnIndexes;

    /// <summary>列名集合</summary>
    public IReadOnlyList<String> Columns { get; }

    /// <summary>数据行数（不含表头）</summary>
    public Int32 RowCount => _rows.Count;

    /// <summary>列数</summary>
    public Int32 ColumnCount => _columns.Length;
    #endregion

    #region 构造
    private ExcelDataFrame(String[] columns, List<Object?[]> rows)
    {
        _columns = columns;
        _rows = rows;
        _columnIndexes = new Dictionary<String, Int32>(columns.Length, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Length; i++)
            _columnIndexes.Add(columns[i], i);
        Columns = Array.AsReadOnly(_columns);
    }

    /// <summary>从工作表创建数据表，指定行作为列名</summary>
    /// <param name="sheet">源工作表</param>
    /// <param name="headerRow">列名所在行的 0 基索引，默认第一行</param>
    /// <returns>内存中的数据表</returns>
    public static ExcelDataFrame FromWorksheet(Worksheet sheet, Int32 headerRow = 0)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (headerRow < 0 || headerRow >= sheet.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(headerRow), L.T("Header row is outside the worksheet."));

        var columnCount = 0;
        foreach (var row in sheet.Rows)
            columnCount = Math.Max(columnCount, row.Length);
        if (columnCount == 0)
            throw new ArgumentException(L.T("The worksheet has no columns."), nameof(sheet));

        var header = sheet.Rows[headerRow];
        var columns = new String[columnCount];
        var indexes = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columnCount; i++)
        {
            var headerName = i < header.Length
                ? Convert.ToString(header[i], CultureInfo.InvariantCulture)
                : null;
            var name = String.IsNullOrWhiteSpace(headerName) ? $"Column{i + 1}" : headerName!;
            if (!indexes.Add(name))
                throw new ArgumentException(L.T("Duplicate column name: {0}", name), nameof(sheet));
            columns[i] = name;
        }

        var rows = new List<Object?[]>(Math.Max(0, sheet.Rows.Count - headerRow - 1));
        for (var i = headerRow + 1; i < sheet.Rows.Count; i++)
        {
            var source = sheet.Rows[i];
            var values = new Object?[columnCount];
            Array.Copy(source, values, Math.Min(source.Length, columnCount));
            rows.Add(values);
        }

        return new ExcelDataFrame(columns, rows);
    }

    /// <summary>从行数据创建数据表</summary>
    /// <remarks>
    /// 与 <see cref="FromWorksheet"/> 使用相同的列名规则：空列名生成 Column1..ColumnN，重复列名抛异常。
    /// 各行长度可不等，短行按空单元格补齐。
    /// </remarks>
    /// <param name="rows">行数据，首行可为列名</param>
    /// <param name="hasHeader">首行是否列名；false 时列名为 Column1..ColumnN</param>
    /// <returns>内存中的数据表</returns>
    public static ExcelDataFrame FromRows(IEnumerable<Object?[]> rows, Boolean hasHeader = true)
    {
        if (rows == null) throw new ArgumentNullException(nameof(rows));

        var list = new List<Object?[]>();
        var columnCount = 0;
        foreach (var row in rows)
        {
            if (row == null) continue;
            columnCount = Math.Max(columnCount, row.Length);
            list.Add(row);
        }
        if (columnCount == 0)
            throw new ArgumentException(L.T("The table has no columns."), nameof(rows));

        var header = hasHeader && list.Count > 0 ? list[0] : [];
        var columns = new String[columnCount];
        var indexes = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columnCount; i++)
        {
            var headerName = i < header.Length
                ? Convert.ToString(header[i], CultureInfo.InvariantCulture)
                : null;
            var name = String.IsNullOrWhiteSpace(headerName) ? $"Column{i + 1}" : headerName!;
            if (!indexes.Add(name))
                throw new ArgumentException(L.T("Duplicate column name: {0}", name), nameof(rows));
            columns[i] = name;
        }

        var data = new List<Object?[]>(Math.Max(0, list.Count - (hasHeader ? 1 : 0)));
        for (var i = hasHeader ? 1 : 0; i < list.Count; i++)
        {
            var source = list[i];
            var values = new Object?[columnCount];
            Array.Copy(source, values, Math.Min(source.Length, columnCount));
            data.Add(values);
        }

        return new ExcelDataFrame(columns, data);
    }

    /// <summary>导出为行数据</summary>
    /// <param name="includeHeader">是否包含列名行</param>
    /// <returns>行数据副本，首行为列名（includeHeader 为 true 时）</returns>
    public IEnumerable<Object?[]> ToRows(Boolean includeHeader = true)
    {
        if (includeHeader)
        {
            var header = new Object?[_columns.Length];
            for (var i = 0; i < _columns.Length; i++)
                header[i] = _columns[i];
            yield return header;
        }

        foreach (var row in _rows)
            yield return (Object?[])row.Clone();
    }
    #endregion

    #region 查询与转换
    /// <summary>按行列索引读取单元格值</summary>
    /// <param name="rowIndex">数据行的 0 基索引</param>
    /// <param name="columnName">列名，不区分大小写</param>
    /// <returns>单元格值；空单元格返回 null</returns>
    public Object? GetValue(Int32 rowIndex, String columnName)
    {
        if (rowIndex < 0 || rowIndex >= _rows.Count) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        return _rows[rowIndex][GetColumnIndex(columnName)];
    }

    /// <summary>选择指定列，保留原始列顺序以外的调用方顺序</summary>
    /// <param name="columns">要保留的列名</param>
    /// <returns>仅包含指定列的新数据表</returns>
    public ExcelDataFrame Select(params String[] columns)
    {
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (columns.Length == 0) throw new ArgumentException(L.T("At least one column is required."), nameof(columns));

        var indexes = new Int32[columns.Length];
        var names = new String[columns.Length];
        var unique = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Length; i++)
        {
            indexes[i] = GetColumnIndex(columns[i]);
            names[i] = _columns[indexes[i]];
            if (!unique.Add(names[i]))
                throw new ArgumentException(L.T("Duplicate selected column: {0}", names[i]), nameof(columns));
        }

        var rows = new List<Object?[]>(_rows.Count);
        foreach (var source in _rows)
        {
            var values = new Object?[indexes.Length];
            for (var i = 0; i < indexes.Length; i++)
                values[i] = source[indexes[i]];
            rows.Add(values);
        }

        return new ExcelDataFrame(names, rows);
    }

    /// <summary>按自定义条件筛选数据行</summary>
    /// <param name="predicate">返回 true 表示保留该行的条件</param>
    /// <returns>符合条件的新数据表</returns>
    public ExcelDataFrame Where(Func<Row, Boolean> predicate)
    {
        if (predicate == null) throw new ArgumentNullException(nameof(predicate));

        var rows = new List<Object?[]>();
        foreach (var values in _rows)
            if (predicate(new Row(values, _columnIndexes))) rows.Add(values);

        return new ExcelDataFrame(_columns, rows);
    }

    /// <summary>按指定列排序；空值始终排在末尾</summary>
    /// <param name="columnName">排序列名</param>
    /// <param name="descending">是否降序</param>
    /// <returns>排序后的新数据表</returns>
    public ExcelDataFrame OrderBy(String columnName, Boolean descending = false)
    {
        var index = GetColumnIndex(columnName);
        var comparer = new CellComparer(descending);
        var rows = _rows.OrderBy(row => row[index], comparer).ToList();
        return new ExcelDataFrame(_columns, rows);
    }

    /// <summary>按指定列分组，再计算每组的汇总值</summary>
    /// <param name="columnName">分组列名</param>
    /// <returns>分组汇总操作</returns>
    public Grouping GroupBy(String columnName) => GroupBy([columnName]);

    /// <summary>按多个指定列分组</summary>
    /// <param name="columnNames">分组列名，输出顺序与参数顺序一致</param>
    /// <returns>分组汇总操作</returns>
    public Grouping GroupBy(params String[] columnNames)
    {
        if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
        if (columnNames.Length == 0) throw new ArgumentException(L.T("At least one grouping column is required."), nameof(columnNames));

        var indexes = new Int32[columnNames.Length];
        var unique = new HashSet<Int32>();
        for (var i = 0; i < columnNames.Length; i++)
        {
            indexes[i] = GetColumnIndex(columnNames[i]);
            if (!unique.Add(indexes[i]))
                throw new ArgumentException(L.T("Duplicate grouping column: {0}", columnNames[i]), nameof(columnNames));
        }

        return new Grouping(this, indexes);
    }

    /// <summary>生成数值列的描述统计</summary>
    /// <param name="columnName">数值列名；非空值必须可转换为数字</param>
    /// <returns>包含 count、mean、std、min、四分位数、max 和 missing 的两列表</returns>
    public ExcelDataFrame Describe(String columnName)
    {
        var index = GetColumnIndex(columnName);
        var values = new List<Double>(_rows.Count);
        var missing = 0;
        foreach (var row in _rows)
        {
            var value = row[index];
            if (IsMissing(value))
            {
                missing++;
                continue;
            }
            values.Add(ToNumber(value!, columnName));
        }

        values.Sort();
        var count = values.Count;
        var mean = count > 0 ? values.Average() : (Double?)null;
        Double? std = null;
        if (count > 1 && mean.HasValue)
        {
            var squared = 0d;
            foreach (var value in values)
                squared += (value - mean.Value) * (value - mean.Value);
            std = Math.Sqrt(squared / (count - 1));
        }

        var rows = new List<Object?[]>
        {
            new Object?[] { "count", count },
            new Object?[] { "mean", mean },
            new Object?[] { "std", std },
            new Object?[] { "min", count > 0 ? values[0] : null },
            new Object?[] { "25%", count > 0 ? Quantile(values, 0.25) : null },
            new Object?[] { "50%", count > 0 ? Quantile(values, 0.5) : null },
            new Object?[] { "75%", count > 0 ? Quantile(values, 0.75) : null },
            new Object?[] { "max", count > 0 ? values[^1] : null },
            new Object?[] { "missing", missing },
        };
        return new ExcelDataFrame(["Statistic", "Value"], rows);
    }

    /// <summary>将数据表转换为可写入工作簿的工作表</summary>
    /// <param name="name">工作表名称</param>
    /// <returns>首行为列名的工作表副本</returns>
    public Worksheet ToWorksheet(String name)
    {
        if (name == null) throw new ArgumentNullException(nameof(name));

        var sheet = new Worksheet { Name = name };
        sheet.Rows.Add(_columns.Cast<Object?>().ToArray());
        foreach (var row in _rows)
            sheet.Rows.Add((Object?[])row.Clone());
        return sheet;
    }
    #endregion

    #region 辅助
    private Int32 GetColumnIndex(String columnName)
    {
        if (columnName == null) throw new ArgumentNullException(nameof(columnName));
        if (!_columnIndexes.TryGetValue(columnName, out var index))
            throw new ArgumentException(L.T("Column not found: {0}", columnName), nameof(columnName));
        return index;
    }

    private static Boolean IsMissing(Object? value) => value == null || value == DBNull.Value;

    private static Double ToNumber(Object value, String columnName)
    {
        if (value is Byte or SByte or Int16 or UInt16 or Int32 or UInt32 or Int64 or UInt64 or Single or Double or Decimal)
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);

        if (value is String text && Double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var number))
            return number;

        throw new ArgumentException(L.T("Column '{0}' contains a non-numeric value: {1}", columnName, value));
    }

    private static Double Quantile(List<Double> values, Double probability)
    {
        var position = (values.Count - 1) * probability;
        var lower = (Int32)Math.Floor(position);
        var upper = (Int32)Math.Ceiling(position);
        return values[lower] + (values[upper] - values[lower]) * (position - lower);
    }

    private static Object? Aggregate(List<Object?[]> rows, Int32 columnIndex, AggregateFunction function, String columnName)
    {
        var values = new List<Object?>(rows.Count);
        foreach (var row in rows)
        {
            var value = row[columnIndex];
            if (!IsMissing(value)) values.Add(value);
        }

        if (function == AggregateFunction.Count) return values.Count;
        if (function == AggregateFunction.Sum || function == AggregateFunction.Average)
        {
            var sum = 0d;
            foreach (var value in values)
                sum += ToNumber(value!, columnName);
            return function == AggregateFunction.Sum ? sum : values.Count > 0 ? sum / values.Count : null;
        }

        if (values.Count == 0) return null;
        var ordered = values.OrderBy(value => value, new CellComparer(false)).ToList();
        return function == AggregateFunction.Min ? ordered[0] : ordered[^1];
    }
    #endregion

    #region 嵌套类型
    /// <summary>按列名访问的一行数据，只读</summary>
    public sealed class Row
    {
        private readonly Object?[] _values;
        private readonly Dictionary<String, Int32> _columnIndexes;

        internal Row(Object?[] values, Dictionary<String, Int32> columnIndexes)
        {
            _values = values;
            _columnIndexes = columnIndexes;
        }

        /// <summary>按列名读取单元格值；列名不区分大小写</summary>
        /// <param name="columnName">列名</param>
        /// <returns>单元格值；空单元格返回 null</returns>
        public Object? this[String columnName]
        {
            get
            {
                if (columnName == null) throw new ArgumentNullException(nameof(columnName));
                if (!_columnIndexes.TryGetValue(columnName, out var index))
                    throw new ArgumentException(L.T("Column not found: {0}", columnName), nameof(columnName));
                return _values[index];
            }
        }

        /// <summary>读取数值；空值返回 null，文本数字按固定区域解析</summary>
        /// <remarks>与 <c>Convert.ToDouble(row["列名"])</c> 相比，空值不会变成 0，而是保留为 null。</remarks>
        /// <param name="columnName">数值列名</param>
        /// <returns>数值；空单元格返回 null</returns>
        /// <exception cref="ArgumentException">列不存在或包含非数字值</exception>
        public Double? Number(String columnName)
        {
            var value = this[columnName];
            return IsMissing(value) ? null : ToNumber(value!, columnName);
        }

        /// <summary>读取文本；空值返回 null，其他类型按固定区域转成字符串</summary>
        /// <param name="columnName">列名</param>
        /// <returns>文本；空单元格返回 null</returns>
        public String? Text(String columnName)
        {
            var value = this[columnName];
            return IsMissing(value) ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }

    /// <summary>对分组执行汇总并返回新的数据表</summary>
    public sealed class Grouping
    {
        private readonly ExcelDataFrame _source;
        private readonly Int32[] _keyIndexes;

        internal Grouping(ExcelDataFrame source, Int32[] keyIndexes)
        {
            _source = source;
            _keyIndexes = keyIndexes;
        }

        /// <summary>统计每组行数</summary>
        /// <returns>分组列与 Count 列组成的数据表</returns>
        public ExcelDataFrame Count() => Aggregate(AggregateSpec.CountRows());

        /// <summary>统计每组指定列的非空值数量</summary>
        /// <param name="columnName">待统计列名</param>
        /// <param name="resultName">结果列名；null 时生成 Count_列名</param>
        /// <returns>分组汇总表</returns>
        public ExcelDataFrame Count(String columnName, String? resultName = null) => Aggregate(AggregateSpec.Count(columnName, resultName));

        /// <summary>计算每组指定数值列的总和</summary>
        /// <param name="columnName">数值列名</param>
        /// <param name="resultName">结果列名；null 时生成 Sum_列名</param>
        /// <returns>分组汇总表</returns>
        public ExcelDataFrame Sum(String columnName, String? resultName = null) => Aggregate(AggregateSpec.Sum(columnName, resultName));

        /// <summary>计算每组指定数值列的平均值</summary>
        /// <param name="columnName">数值列名</param>
        /// <param name="resultName">结果列名；null 时生成 Average_列名</param>
        /// <returns>分组汇总表</returns>
        public ExcelDataFrame Average(String columnName, String? resultName = null) => Aggregate(AggregateSpec.Average(columnName, resultName));

        /// <summary>计算每组指定列的最小值</summary>
        /// <param name="columnName">列名</param>
        /// <param name="resultName">结果列名；null 时生成 Min_列名</param>
        /// <returns>分组汇总表</returns>
        public ExcelDataFrame Min(String columnName, String? resultName = null) => Aggregate(AggregateSpec.Min(columnName, resultName));

        /// <summary>计算每组指定列的最大值</summary>
        /// <param name="columnName">列名</param>
        /// <param name="resultName">结果列名；null 时生成 Max_列名</param>
        /// <returns>分组汇总表</returns>
        public ExcelDataFrame Max(String columnName, String? resultName = null) => Aggregate(AggregateSpec.Max(columnName, resultName));

        /// <summary>一次计算多个汇总列</summary>
        /// <param name="aggregations">汇总定义</param>
        /// <returns>包含分组列和所有汇总列的数据表</returns>
        public ExcelDataFrame Aggregate(params AggregateSpec[] aggregations)
        {
            if (aggregations == null) throw new ArgumentNullException(nameof(aggregations));
            if (aggregations.Length == 0) throw new ArgumentException(L.T("At least one aggregation is required."), nameof(aggregations));

            var valueIndexes = new Int32[aggregations.Length];
            var resultColumns = new String[_keyIndexes.Length + aggregations.Length];
            var names = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < _keyIndexes.Length; i++)
            {
                var name = _source._columns[_keyIndexes[i]];
                if (!names.Add(name))
                    throw new ArgumentException(L.T("Duplicate result column: {0}", name), nameof(aggregations));
                resultColumns[i] = name;
            }

            for (var i = 0; i < aggregations.Length; i++)
            {
                var spec = aggregations[i] ?? throw new ArgumentException(L.T("Aggregation definitions cannot contain null."), nameof(aggregations));
                valueIndexes[i] = spec.ColumnName == null ? -1 : _source.GetColumnIndex(spec.ColumnName);
                var resultName = spec.ResultName ?? (spec.ColumnName == null
                    ? spec.Function.ToString()
                    : $"{spec.Function}_{_source._columns[valueIndexes[i]]}");
                if (String.IsNullOrWhiteSpace(resultName))
                    throw new ArgumentException(L.T("Result column name cannot be empty."), nameof(aggregations));
                if (!names.Add(resultName))
                    throw new ArgumentException(L.T("Duplicate result column: {0}", resultName), nameof(aggregations));
                resultColumns[_keyIndexes.Length + i] = resultName;
            }

            var groups = new Dictionary<CompositeKey, List<Object?[]>>();
            var ordered = new List<(Object?[] Keys, List<Object?[]> Rows)>();
            foreach (var row in _source._rows)
            {
                var keyValues = new Object?[_keyIndexes.Length];
                for (var i = 0; i < _keyIndexes.Length; i++)
                    keyValues[i] = IsMissing(row[_keyIndexes[i]]) ? null : row[_keyIndexes[i]];

                var key = new CompositeKey(keyValues);
                if (!groups.TryGetValue(key, out var groupRows))
                {
                    groupRows = [];
                    groups.Add(key, groupRows);
                    ordered.Add((keyValues, groupRows));
                }
                groupRows.Add(row);
            }

            var rows = new List<Object?[]>(ordered.Count);
            foreach (var group in ordered)
            {
                var result = new Object?[resultColumns.Length];
                Array.Copy(group.Keys, result, group.Keys.Length);
                for (var i = 0; i < aggregations.Length; i++)
                    result[_keyIndexes.Length + i] = valueIndexes[i] < 0
                        ? group.Rows.Count
                        : ExcelDataFrame.Aggregate(group.Rows, valueIndexes[i], aggregations[i].Function, aggregations[i].ColumnName!);
                rows.Add(result);
            }

            return new ExcelDataFrame(resultColumns, rows);
        }
    }

    private sealed class CompositeKey(Object?[] values) : IEquatable<CompositeKey>
    {
        private readonly Object?[] _values = values;

        public Boolean Equals(CompositeKey? other)
        {
            if (other == null || _values.Length != other._values.Length) return false;
            for (var i = 0; i < _values.Length; i++)
                if (!ObjectKeyComparer.Instance.Equals(_values[i], other._values[i])) return false;
            return true;
        }

        public override Boolean Equals(Object? obj) => obj is CompositeKey other && Equals(other);

        public override Int32 GetHashCode()
        {
            var hash = 17;
            foreach (var value in _values)
                hash = unchecked(hash * 31 + (value == null ? 0 : ObjectKeyComparer.Instance.GetHashCode(value)));
            return hash;
        }
    }

    private sealed class CellComparer(Boolean descending) : IComparer<Object?>
    {
        public Int32 Compare(Object? left, Object? right)
        {
            if (IsMissing(left)) return IsMissing(right) ? 0 : 1;
            if (IsMissing(right)) return -1;

            var result = CompareValues(left!, right!);
            return descending ? -result : result;
        }

        private static Int32 CompareValues(Object left, Object right)
        {
            if (IsNumeric(left) && IsNumeric(right))
                return Convert.ToDouble(left, CultureInfo.InvariantCulture).CompareTo(Convert.ToDouble(right, CultureInfo.InvariantCulture));
            if (left is DateTime leftDate && right is DateTime rightDate) return leftDate.CompareTo(rightDate);
            if (left is Boolean leftBoolean && right is Boolean rightBoolean) return leftBoolean.CompareTo(rightBoolean);
            if (left is String leftText && right is String rightText) return String.CompareOrdinal(leftText, rightText);
            if (left.GetType() == right.GetType() && left is IComparable comparable)
                return comparable.CompareTo(right);
            return String.Compare(Convert.ToString(left, CultureInfo.InvariantCulture), Convert.ToString(right, CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        private static Boolean IsNumeric(Object value)
            => value is Byte or SByte or Int16 or UInt16 or Int32 or UInt32 or Int64 or UInt64 or Single or Double or Decimal;
    }

    private sealed class ObjectKeyComparer : IEqualityComparer<Object>
    {
        public static readonly ObjectKeyComparer Instance = new();

        public new Boolean Equals(Object? x, Object? y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x == null || y == null) return false;
            if (IsNumeric(x) && IsNumeric(y))
                return Convert.ToDouble(x, CultureInfo.InvariantCulture).Equals(Convert.ToDouble(y, CultureInfo.InvariantCulture));
            return x.Equals(y);
        }

        public Int32 GetHashCode(Object obj)
            => IsNumeric(obj) ? Convert.ToDouble(obj, CultureInfo.InvariantCulture).GetHashCode() : obj.GetHashCode();

        private static Boolean IsNumeric(Object value)
            => value is Byte or SByte or Int16 or UInt16 or Int32 or UInt32 or Int64 or UInt64 or Single or Double or Decimal;
    }

    #endregion
}
