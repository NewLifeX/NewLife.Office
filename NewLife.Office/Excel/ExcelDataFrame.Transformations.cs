namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 列变换与数据清理操作</summary>
public sealed partial class ExcelDataFrame
{
    #region 查询
    /// <summary>按多个列键依次排序；相同排序值保留原行顺序，空值排在末尾</summary>
    /// <param name="sortOrders">排序列与是否降序</param>
    /// <returns>排序后的新数据表</returns>
    public ExcelDataFrame OrderBy(params (String Column, Boolean Descending)[] sortOrders)
    {
        if (sortOrders == null) throw new ArgumentNullException(nameof(sortOrders));
        if (sortOrders.Length == 0) throw new ArgumentException(L.T("At least one sort column is required."), nameof(sortOrders));

        var indexes = new Int32[sortOrders.Length];
        for (var i = 0; i < sortOrders.Length; i++)
        {
            indexes[i] = GetColumnIndex(sortOrders[i].Column);
            for (var j = 0; j < i; j++)
                if (indexes[j] == indexes[i])
                    throw new ArgumentException(L.T("Duplicate sort column: {0}", sortOrders[i].Column), nameof(sortOrders));
        }

        var comparer = new MultiColumnComparer(indexes, sortOrders);
        var rows = _rows.OrderBy(row => row, comparer).ToList();
        return new ExcelDataFrame(_columns, rows);
    }

    /// <summary>按多列排序后取前若干行</summary>
    /// <param name="count">最大返回行数</param>
    /// <param name="sortOrders">排序列与是否降序</param>
    /// <returns>排序后的前若干行</returns>
    public ExcelDataFrame Top(Int32 count, params (String Column, Boolean Descending)[] sortOrders)
    {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        return OrderBy(sortOrders).Take(count);
    }

    /// <summary>取前若干行，不修改源数据</summary>
    /// <param name="count">最大返回行数</param>
    /// <returns>包含前若干行的新数据表</returns>
    public ExcelDataFrame Take(Int32 count)
    {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        return new ExcelDataFrame(_columns, _rows.Take(count).ToList());
    }
    #endregion

    #region 列变换
    /// <summary>根据每行计算并追加一列</summary>
    /// <param name="columnName">新列名</param>
    /// <param name="selector">按行计算新值</param>
    /// <returns>追加新列后的数据表</returns>
    public ExcelDataFrame AddColumn(String columnName, Func<Row, Object?> selector)
    {
        if (String.IsNullOrWhiteSpace(columnName)) throw new ArgumentException(L.T("Column name cannot be empty."), nameof(columnName));
        if (selector == null) throw new ArgumentNullException(nameof(selector));
        if (_columnIndexes.ContainsKey(columnName))
            throw new ArgumentException(L.T("Column already exists: {0}", columnName), nameof(columnName));

        var columns = new String[_columns.Length + 1];
        Array.Copy(_columns, columns, _columns.Length);
        columns[^1] = columnName;

        var rows = new List<Object?[]>(_rows.Count);
        foreach (var source in _rows)
        {
            var values = new Object?[columns.Length];
            Array.Copy(source, values, source.Length);
            values[^1] = selector(new Row(source, _columnIndexes));
            rows.Add(values);
        }

        return new ExcelDataFrame(columns, rows);
    }
    #endregion

    #region 缺失值与重复行
    /// <summary>删除指定列存在空值的行；未指定列时检查所有列</summary>
    /// <param name="columns">要检查的列名</param>
    /// <returns>移除含空值行后的新数据表</returns>
    public ExcelDataFrame DropMissing(params String[] columns)
    {
        var indexes = ResolveColumns(columns);
        var rows = new List<Object?[]>(_rows.Count);
        foreach (var row in _rows)
        {
            var keep = true;
            foreach (var index in indexes)
                if (IsMissing(row[index]))
                {
                    keep = false;
                    break;
                }
            if (keep) rows.Add(row);
        }

        return new ExcelDataFrame(_columns, rows);
    }

    /// <summary>将指定列的空值替换为同一个值；未指定列时处理所有列</summary>
    /// <param name="value">替换值；null 或 DBNull 不会改变空值</param>
    /// <param name="columns">要处理的列名</param>
    /// <returns>替换后的新数据表</returns>
    public ExcelDataFrame FillMissing(Object? value, params String[] columns)
    {
        if (IsMissing(value)) return new ExcelDataFrame(_columns, _rows.ToList());

        var indexes = ResolveColumns(columns);
        var rows = new List<Object?[]>(_rows.Count);
        foreach (var source in _rows)
        {
            var row = (Object?[])source.Clone();
            foreach (var index in indexes)
                if (IsMissing(row[index])) row[index] = value;
            rows.Add(row);
        }

        return new ExcelDataFrame(_columns, rows);
    }

    /// <summary>删除重复行并保留首次出现的行；未指定列时比较所有列</summary>
    /// <param name="columns">用于判断重复的列名</param>
    /// <returns>去重后的新数据表</returns>
    public ExcelDataFrame DropDuplicates(params String[] columns)
    {
        var indexes = ResolveColumns(columns);
        var seen = new HashSet<CompositeKey>();
        var rows = new List<Object?[]>(_rows.Count);
        foreach (var row in _rows)
        {
            var key = new Object?[indexes.Length];
            for (var i = 0; i < indexes.Length; i++)
                key[i] = IsMissing(row[indexes[i]]) ? null : row[indexes[i]];
            if (seen.Add(new CompositeKey(key))) rows.Add(row);
        }

        return new ExcelDataFrame(_columns, rows);
    }

    private Int32[] ResolveColumns(String[]? columns)
    {
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (columns.Length == 0) return Enumerable.Range(0, _columns.Length).ToArray();

        var indexes = new Int32[columns.Length];
        var unique = new HashSet<Int32>();
        for (var i = 0; i < columns.Length; i++)
        {
            indexes[i] = GetColumnIndex(columns[i]);
            if (!unique.Add(indexes[i]))
                throw new ArgumentException(L.T("Duplicate column: {0}", columns[i]), nameof(columns));
        }
        return indexes;
    }
    #endregion

    private sealed class MultiColumnComparer(Int32[] indexes, (String Column, Boolean Descending)[] orders) : IComparer<Object?[]>
    {
        public Int32 Compare(Object?[]? left, Object?[]? right)
        {
            if (ReferenceEquals(left, right)) return 0;
            if (left == null) return -1;
            if (right == null) return 1;

            for (var i = 0; i < indexes.Length; i++)
            {
                var result = new CellComparer(orders[i].Descending).Compare(left[indexes[i]], right[indexes[i]]);
                if (result != 0) return result;
            }
            return 0;
        }
    }
}
