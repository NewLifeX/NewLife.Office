namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 表关联操作</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>按单个键列合并两个数据表</summary>
    /// <param name="right">右侧数据表</param>
    /// <param name="leftKey">左侧关联列</param>
    /// <param name="rightKey">右侧关联列</param>
    /// <param name="joinKind">关联类型</param>
    /// <param name="rightSuffix">右侧重名列的后缀</param>
    /// <returns>合并后的新数据表</returns>
    public ExcelDataFrame Merge(ExcelDataFrame right, String leftKey, String rightKey, JoinKind joinKind = JoinKind.Inner, String rightSuffix = "_right")
        => Merge(right, [leftKey], [rightKey], joinKind, rightSuffix);

    /// <summary>按一个或多个键列合并两个数据表</summary>
    /// <param name="right">右侧数据表</param>
    /// <param name="leftKeys">左侧关联列，顺序对应右侧键列</param>
    /// <param name="rightKeys">右侧关联列，顺序对应左侧键列</param>
    /// <param name="joinKind">关联类型</param>
    /// <param name="rightSuffix">右侧重名列的后缀</param>
    /// <returns>合并后的新数据表</returns>
    public ExcelDataFrame Merge(ExcelDataFrame right, String[] leftKeys, String[] rightKeys, JoinKind joinKind = JoinKind.Inner, String rightSuffix = "_right")
    {
        if (right == null) throw new ArgumentNullException(nameof(right));
        if (leftKeys == null) throw new ArgumentNullException(nameof(leftKeys));
        if (rightKeys == null) throw new ArgumentNullException(nameof(rightKeys));
        if (leftKeys.Length == 0 || leftKeys.Length != rightKeys.Length)
            throw new ArgumentException(L.T("Left and right join keys must have the same non-zero length."));
        if (joinKind is not (JoinKind.Inner or JoinKind.Left or JoinKind.Right or JoinKind.Full))
            throw new ArgumentOutOfRangeException(nameof(joinKind));
        if (String.IsNullOrEmpty(rightSuffix))
            throw new ArgumentException(L.T("Right column suffix cannot be empty."), nameof(rightSuffix));

        var leftIndexes = ResolveJoinKeys(leftKeys, nameof(leftKeys));
        var rightIndexes = right.ResolveJoinKeys(rightKeys, nameof(rightKeys));

        var columns = new List<String>(_columns);
        var rightOutputIndexes = new List<Int32>();
        var rightOutputNames = new HashSet<String>(columns, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < right._columns.Length; i++)
        {
            if (rightIndexes.Contains(i)) continue;

            var name = right._columns[i];
            if (rightOutputNames.Contains(name)) name += rightSuffix;
            if (!rightOutputNames.Add(name))
                throw new ArgumentException(L.T("Merged column name conflicts: {0}", name), nameof(right));
            columns.Add(name);
            rightOutputIndexes.Add(i);
        }

        var rightGroups = new Dictionary<CompositeKey, List<Int32>>();
        for (var i = 0; i < right._rows.Count; i++)
        {
            if (!TryBuildJoinKey(right._rows[i], rightIndexes, out var key)) continue;
            if (!rightGroups.TryGetValue(key!, out var matches))
            {
                matches = [];
                rightGroups.Add(key!, matches);
            }
            matches.Add(i);
        }

        var rows = new List<Object?[]>();
        var matchedRight = new Boolean[right._rows.Count];
        foreach (var leftRow in _rows)
        {
            if (TryBuildJoinKey(leftRow, leftIndexes, out var key) && rightGroups.TryGetValue(key!, out var matches))
            {
                foreach (var rightIndex in matches)
                {
                    matchedRight[rightIndex] = true;
                    rows.Add(CombineRows(leftRow, right._rows[rightIndex], rightOutputIndexes));
                }
            }
            else if (joinKind is JoinKind.Left or JoinKind.Full)
                rows.Add(CombineRows(leftRow, null, rightOutputIndexes));
        }

        if (joinKind is JoinKind.Right or JoinKind.Full)
        {
            for (var i = 0; i < right._rows.Count; i++)
            {
                if (matchedRight[i]) continue;

                var row = CombineRows(null, right._rows[i], rightOutputIndexes);
                for (var k = 0; k < leftIndexes.Length; k++)
                    row[leftIndexes[k]] = right._rows[i][rightIndexes[k]];
                rows.Add(row);
            }
        }

        return new ExcelDataFrame(columns.ToArray(), rows);
    }

    private Int32[] ResolveJoinKeys(String[] keys, String parameterName)
    {
        var indexes = new Int32[keys.Length];
        var unique = new HashSet<Int32>();
        for (var i = 0; i < keys.Length; i++)
        {
            indexes[i] = GetColumnIndex(keys[i]);
            if (!unique.Add(indexes[i]))
                throw new ArgumentException(L.T("Duplicate join key: {0}", keys[i]), parameterName);
        }
        return indexes;
    }

    private static Boolean TryBuildJoinKey(Object?[] row, Int32[] indexes, out CompositeKey? key)
    {
        var values = new Object?[indexes.Length];
        for (var i = 0; i < indexes.Length; i++)
        {
            var value = row[indexes[i]];
            if (IsMissing(value))
            {
                key = null;
                return false;
            }
            values[i] = value;
        }

        key = new CompositeKey(values);
        return true;
    }

    private Object?[] CombineRows(Object?[]? left, Object?[]? right, List<Int32> rightOutputIndexes)
    {
        var result = new Object?[_columns.Length + rightOutputIndexes.Count];
        if (left != null) Array.Copy(left, result, left.Length);
        if (right == null) return result;

        for (var i = 0; i < rightOutputIndexes.Count; i++)
            result[_columns.Length + i] = right[rightOutputIndexes[i]];
        return result;
    }
}
