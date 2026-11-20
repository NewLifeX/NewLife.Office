using System.Globalization;

namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 分类交叉表操作</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>统计两个分类列的交叉频数</summary>
    /// <remarks>
    /// 每个数据行都计数，不依赖值列。行键和列键均按首次出现顺序排列；
    /// null、DBNull 与空白字符串归为同一空类别，行侧结果值为 null，列侧标题为“(空)”。
    /// 若列类别转成标题后与行键列名或其他标题冲突（不区分大小写），则抛出异常。
    /// </remarks>
    /// <param name="rowColumn">作为结果行键的分类列</param>
    /// <param name="columnColumn">作为结果列键的分类列</param>
    /// <returns>第一列为行键，其余列为列类别，单元格为对应组合的行数</returns>
    public ExcelDataFrame Crosstab(String rowColumn, String columnColumn)
    {
        var rowIndex = GetColumnIndex(rowColumn);
        var columnIndex = GetColumnIndex(columnColumn);
        if (rowIndex == columnIndex)
            throw new ArgumentException(L.T("Row and column category columns must be different."), nameof(columnColumn));

        var rowCategories = new List<Object?>();
        var columnCategories = new List<Object?>();
        var rowIndexes = new Dictionary<CompositeKey, Int32>();
        var columnIndexes = new Dictionary<CompositeKey, Int32>();
        var counts = new Dictionary<CompositeKey, Int32>();

        foreach (var source in _rows)
        {
            var rowCategory = NormalizeCrosstabCategory(source[rowIndex]);
            var columnCategory = NormalizeCrosstabCategory(source[columnIndex]);
            var rowCategoryIndex = GetOrAddCrosstabCategory(rowCategory, rowIndexes, rowCategories);
            var columnCategoryIndex = GetOrAddCrosstabCategory(columnCategory, columnIndexes, columnCategories);

            var pair = new CompositeKey([rowCategoryIndex, columnCategoryIndex]);
            if (counts.TryGetValue(pair, out var count))
                counts[pair] = count + 1;
            else
                counts.Add(pair, 1);
        }

        var resultColumns = new String[columnCategories.Count + 1];
        resultColumns[0] = _columns[rowIndex];
        var names = new HashSet<String>(StringComparer.OrdinalIgnoreCase) { resultColumns[0] };
        for (var i = 0; i < columnCategories.Count; i++)
        {
            var name = GetCrosstabCategoryName(columnCategories[i]);
            if (!names.Add(name))
                throw new ArgumentException(L.T("Crosstab result column conflicts: {0}", name), nameof(columnColumn));
            resultColumns[i + 1] = name;
        }

        var resultRows = new List<Object?[]>(rowCategories.Count);
        for (var i = 0; i < rowCategories.Count; i++)
        {
            var row = new Object?[resultColumns.Length];
            row[0] = rowCategories[i];
            for (var j = 0; j < columnCategories.Count; j++)
            {
                row[j + 1] = counts.TryGetValue(new CompositeKey([i, j]), out var count) ? count : 0;
            }
            resultRows.Add(row);
        }

        return new ExcelDataFrame(resultColumns, resultRows);
    }

    private static Object? NormalizeCrosstabCategory(Object? value)
        => IsMissing(value) || value is String text && String.IsNullOrWhiteSpace(text) ? null : value;

    private static Int32 GetOrAddCrosstabCategory(
        Object? value,
        Dictionary<CompositeKey, Int32> indexes,
        List<Object?> categories)
    {
        var key = new CompositeKey([value]);
        if (indexes.TryGetValue(key, out var index)) return index;

        index = categories.Count;
        categories.Add(value);
        indexes.Add(key, index);
        return index;
    }

    private static String GetCrosstabCategoryName(Object? value)
    {
        if (value == null) return "(空)";

        var name = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (String.IsNullOrWhiteSpace(name))
            throw new ArgumentException(L.T("Crosstab category cannot be used as a column name."));

        return name;
    }
}
