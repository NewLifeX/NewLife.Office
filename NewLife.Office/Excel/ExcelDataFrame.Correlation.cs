namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 相关性分析</summary>
public sealed partial class ExcelDataFrame
{
    /// <summary>计算指定数值列的 Pearson 相关矩阵</summary>
    /// <remarks>
    /// 每对列分别使用两列均非空的行计算样本相关系数。有效样本少于 2 个或任一列为常量时，结果为 null；
    /// 非空值必须可转换为数字。
    /// </remarks>
    /// <param name="columns">参与计算的数值列，结果按指定顺序排列</param>
    /// <returns>以列名为行列标签的相关矩阵；标签列名为 Column，遇重名时自动追加数字</returns>
    public ExcelDataFrame CorrelationMatrix(params String[] columns)
    {
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (columns.Length == 0) throw new ArgumentException(L.T("At least one column is required."), nameof(columns));

        var indexes = new Int32[columns.Length];
        var names = new String[columns.Length];
        var uniqueIndexes = new HashSet<Int32>();
        for (var i = 0; i < columns.Length; i++)
        {
            indexes[i] = GetColumnIndex(columns[i]);
            if (!uniqueIndexes.Add(indexes[i]))
                throw new ArgumentException(L.T("Duplicate column: {0}", columns[i]), nameof(columns));
            names[i] = _columns[indexes[i]];
        }

        var labelColumn = "Column";
        var suffix = 2;
        var nameSet = new HashSet<String>(names, StringComparer.OrdinalIgnoreCase);
        while (nameSet.Contains(labelColumn))
            labelColumn = $"Column{suffix++}";

        var resultColumns = new String[names.Length + 1];
        resultColumns[0] = labelColumn;
        Array.Copy(names, 0, resultColumns, 1, names.Length);

        var rows = new List<Object?[]>(names.Length);
        for (var i = 0; i < names.Length; i++)
        {
            var row = new Object?[resultColumns.Length];
            row[0] = names[i];
            for (var j = 0; j < names.Length; j++)
                row[j + 1] = CalculateCorrelation(indexes[i], indexes[j], names[i], names[j]);
            rows.Add(row);
        }

        return new ExcelDataFrame(resultColumns, rows);
    }

    private Double? CalculateCorrelation(Int32 leftIndex, Int32 rightIndex, String leftName, String rightName)
    {
        var count = 0;
        var meanLeft = 0d;
        var meanRight = 0d;
        var sumSquaresLeft = 0d;
        var sumSquaresRight = 0d;
        var sumProducts = 0d;

        foreach (var row in _rows)
        {
            var left = row[leftIndex];
            var right = row[rightIndex];
            if (IsMissing(left) || IsMissing(right)) continue;

            var leftValue = ToNumber(left!, leftName);
            var rightValue = ToNumber(right!, rightName);
            count++;

            var deltaLeft = leftValue - meanLeft;
            meanLeft += deltaLeft / count;
            var deltaRight = rightValue - meanRight;
            meanRight += deltaRight / count;

            sumSquaresLeft += deltaLeft * (leftValue - meanLeft);
            sumSquaresRight += deltaRight * (rightValue - meanRight);
            sumProducts += deltaLeft * (rightValue - meanRight);
        }

        if (count < 2 || sumSquaresLeft <= 0 || sumSquaresRight <= 0) return null;

        var correlation = sumProducts / Math.Sqrt(sumSquaresLeft * sumSquaresRight);
        if (correlation > 1d) return 1d;
        if (correlation < -1d) return -1d;
        return correlation;
    }
}
