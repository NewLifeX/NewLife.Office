using System.Globalization;

namespace NewLife.Office.Excel;

/// <summary>工作表编辑辅助：按 A1 坐标读写单元格、调整行列与样式</summary>
/// <remarks>
/// 所有操作直接作用于 <see cref="Worksheet"/> 模型，保存由 <see cref="ExcelDocumentExtensions.Save(ExcelDocument, String)"/> 负责。
/// 编辑前会把"跳行"布局规整为稠密行（<see cref="Worksheet.ActualRowNumbers"/> 置 null），
/// 因此后续行列索引与 Excel 行号一一对应。
/// </remarks>
public static class ExcelSheetEditor
{
    #region 规整与查询
    /// <summary>把跳行布局规整为稠密行，使行索引与 Excel 行号一一对应</summary>
    /// <param name="sheet">工作表</param>
    public static void Normalize(Worksheet sheet)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (sheet.ActualRowNumbers == null) return;

        var maxRow = -1;
        foreach (var row in sheet.ActualRowNumbers)
            if (row > maxRow) maxRow = row;

        var rows = new List<Object?[]>(maxRow + 1);
        for (var i = 0; i <= maxRow; i++) rows.Add([]);

        for (var i = 0; i < sheet.ActualRowNumbers.Count && i < sheet.Rows.Count; i++)
        {
            var actual = sheet.ActualRowNumbers[i];
            if (actual >= 0 && actual < rows.Count) rows[actual] = sheet.Rows[i];
        }

        sheet.Rows.Clear();
        sheet.Rows.AddRange(rows);
        sheet.ActualRowNumbers = null;
    }

    /// <summary>获取已用行数</summary>
    /// <param name="sheet">工作表</param>
    /// <returns>行数</returns>
    public static Int32 GetRowCount(Worksheet sheet) => sheet.Rows.Count;

    /// <summary>获取已用列数</summary>
    /// <param name="sheet">工作表</param>
    /// <returns>列数</returns>
    public static Int32 GetColumnCount(Worksheet sheet)
    {
        var max = 0;
        foreach (var row in sheet.Rows)
            if (row.Length > max) max = row.Length;
        return max;
    }

    /// <summary>读取单元格值</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <returns>单元格值，超出范围返回 null</returns>
    public static Object? GetValue(Worksheet sheet, Int32 row, Int32 col)
    {
        if (row < 0 || row >= sheet.Rows.Count) return null;
        var cells = sheet.Rows[row];
        return col >= 0 && col < cells.Length ? cells[col] : null;
    }

    /// <summary>读取单元格公式（不含等号）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <returns>公式文本，非公式单元格返回 null</returns>
    public static String? GetFormula(Worksheet sheet, Int32 row, Int32 col)
        => sheet.Formulas.TryGetValue((row, col), out var formula) ? formula : null;

    /// <summary>读取区域内容（公式以 <c>=</c> 前缀返回）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">区域</param>
    /// <returns>逐行值数组</returns>
    public static List<Object?[]> ReadRange(Worksheet sheet, ExcelCellRange range)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (range == null) throw new ArgumentNullException(nameof(range));

        var rows = new List<Object?[]>();
        for (var r = range.StartRow; r <= range.EndRow; r++)
        {
            var cells = new Object?[range.ColumnCount];
            for (var c = range.StartCol; c <= range.EndCol; c++)
            {
                var formula = GetFormula(sheet, r, c);
                cells[c - range.StartCol] = formula != null ? "=" + formula : GetValue(sheet, r, c);
            }
            rows.Add(cells);
        }
        return rows;
    }
    #endregion

    #region 写入
    /// <summary>扩展工作表到指定行列</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    public static void EnsureSize(Worksheet sheet, Int32 row, Int32 col)
    {
        if (row < 0 || col < 0) throw new ArgumentOutOfRangeException(nameof(row), "行列索引不能为负");
        Normalize(sheet);

        while (sheet.Rows.Count <= row) sheet.Rows.Add([]);
        var cells = sheet.Rows[row];
        if (cells.Length <= col)
        {
            var expanded = new Object?[col + 1];
            Array.Copy(cells, expanded, cells.Length);
            sheet.Rows[row] = expanded;
        }
    }

    /// <summary>写入单元格值（清除同位置公式）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <param name="value">值</param>
    public static void SetValue(Worksheet sheet, Int32 row, Int32 col, Object? value)
    {
        EnsureSize(sheet, row, col);
        sheet.Rows[row][col] = value;
        sheet.Formulas.Remove((row, col));
    }

    /// <summary>写入单元格公式</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <param name="formula">公式文本（可带前导 <c>=</c>），null 表示清除公式</param>
    public static void SetFormula(Worksheet sheet, Int32 row, Int32 col, String? formula)
    {
        EnsureSize(sheet, row, col);
        if (formula == null)
        {
            sheet.Formulas.Remove((row, col));
            return;
        }

        var text = formula.Trim();
        if (text.StartsWith("=", StringComparison.Ordinal)) text = text[1..];
        if (text.Length == 0)
        {
            sheet.Formulas.Remove((row, col));
            return;
        }
        sheet.Formulas[(row, col)] = text;
    }

    /// <summary>批量写入二维数组（字符串以 <c>=</c> 开头时写入公式）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">起始区域（按值数组实际尺寸写入，可超出区域）</param>
    /// <param name="values">二维值数组</param>
    /// <returns>实际写入的行列数</returns>
    public static (Int32 Rows, Int32 Columns) WriteMatrix(Worksheet sheet, ExcelCellRange range, IList<Object?[]> values)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (range == null) throw new ArgumentNullException(nameof(range));
        if (values == null) throw new ArgumentNullException(nameof(values));

        var rows = 0;
        var cols = 0;
        for (var r = 0; r < values.Count; r++)
        {
            var row = values[r];
            for (var c = 0; c < row.Length; c++)
            {
                var rowIndex = range.StartRow + r;
                var colIndex = range.StartCol + c;
                if (row[c] is String text && text.StartsWith("=", StringComparison.Ordinal))
                    SetFormula(sheet, rowIndex, colIndex, text);
                else
                    SetValue(sheet, rowIndex, colIndex, row[c]);
            }
            if (row.Length > 0) rows = r + 1;
            if (row.Length > cols) cols = row.Length;
        }
        return (rows, cols);
    }

    /// <summary>追加数据行（从最后一个有内容的行之后开始）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="values">二维值数组</param>
    /// <returns>起始行索引（0 基）</returns>
    public static Int32 AppendRows(Worksheet sheet, IList<Object?[]> values)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (values == null) throw new ArgumentNullException(nameof(values));

        Normalize(sheet);
        var start = 0;
        for (var r = sheet.Rows.Count - 1; r >= 0; r--)
        {
            var hasValue = false;
            foreach (var cell in sheet.Rows[r])
            {
                if (cell != null) { hasValue = true; break; }
            }
            if (hasValue) { start = r + 1; break; }
        }

        var range = new ExcelCellRange(sheet.Name, start, 0, start, 0);
        WriteMatrix(sheet, range, values);
        return start;
    }
    #endregion

    #region 行列调整
    /// <summary>插入空行并平移样式、公式、批注等行级信息</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="at">插入位置（0 基，原行及之后下移）</param>
    /// <param name="count">插入行数</param>
    public static void InsertRows(Worksheet sheet, Int32 at, Int32 count)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (count <= 0) return;

        Normalize(sheet);
        if (at < 0) at = 0;
        if (at > sheet.Rows.Count) at = sheet.Rows.Count;

        for (var i = 0; i < count; i++) sheet.Rows.Insert(at, []);

        ShiftRows(sheet, at, count);
    }

    /// <summary>删除行并平移样式、公式、批注等行级信息</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="start">起始行（0 基）</param>
    /// <param name="count">删除行数</param>
    public static void DeleteRows(Worksheet sheet, Int32 start, Int32 count)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (count <= 0) return;

        Normalize(sheet);
        if (start < 0) start = 0;
        if (start >= sheet.Rows.Count) return;

        var end = Math.Min(start + count - 1, sheet.Rows.Count - 1);
        var actual = end - start + 1;
        sheet.Rows.RemoveRange(start, actual);

        ShiftRows(sheet, start, -actual);
    }

    private static void ShiftRows(Worksheet sheet, Int32 at, Int32 delta)
    {
        sheet.CellFormats = ShiftRowKeys(sheet.CellFormats, at, delta);
        sheet.Formulas = ShiftRowKeys(sheet.Formulas, at, delta);
        sheet.Comments = ShiftRowKeys(sheet.Comments, at, delta);
        sheet.Hyperlinks = ShiftRowKeys(sheet.Hyperlinks, at, delta);
        sheet.RowHeights = ShiftIndexKeys(sheet.RowHeights, at, delta);

        var hidden = new HashSet<Int32>();
        foreach (var row in sheet.HiddenRows)
        {
            if (delta > 0 && row >= at) hidden.Add(row + delta);
            else if (delta < 0 && row >= at && row < at - delta) continue;
            else if (delta < 0 && row >= at - delta) hidden.Add(row + delta);
            else hidden.Add(row);
        }
        sheet.HiddenRows = hidden;

        var merges = new List<(Int32 StartRow, Int32 StartCol, Int32 EndRow, Int32 EndCol)>();
        foreach (var (sr, sc, er, ec) in sheet.Merges)
        {
            if (delta > 0)
            {
                if (sr >= at) merges.Add((sr + delta, sc, er + delta, ec));
                else if (at <= er) merges.Add((sr, sc, er + delta, ec));
                else merges.Add((sr, sc, er, ec));
            }
            else
            {
                var remove = -delta;
                var from = at;
                var to = at + remove - 1;
                if (er < from) merges.Add((sr, sc, er, ec));
                else if (sr > to) merges.Add((sr - remove, sc, er - remove, ec));
                else
                {
                    var newStart = sr < from ? sr : from;
                    var newEnd = er > to ? er - remove : from - 1;
                    if (newEnd > newStart) merges.Add((newStart, sc, newEnd, ec));
                }
            }
        }
        sheet.Merges = merges;
    }

    private static Dictionary<(Int32 Row, Int32 Col), T> ShiftRowKeys<T>(Dictionary<(Int32 Row, Int32 Col), T> source, Int32 at, Int32 delta)
    {
        var result = new Dictionary<(Int32 Row, Int32 Col), T>();
        foreach (var kv in source)
        {
            var row = kv.Key.Row;
            if (delta > 0)
            {
                if (row >= at) row += delta;
            }
            else
            {
                var remove = -delta;
                if (row >= at && row < at + remove) continue;
                if (row >= at + remove) row -= remove;
            }
            result[(row, kv.Key.Col)] = kv.Value;
        }
        return result;
    }

    private static Dictionary<Int32, T> ShiftIndexKeys<T>(Dictionary<Int32, T> source, Int32 at, Int32 delta)
    {
        var result = new Dictionary<Int32, T>();
        foreach (var kv in source)
        {
            var row = kv.Key;
            if (delta > 0)
            {
                if (row >= at) row += delta;
            }
            else
            {
                var remove = -delta;
                if (row >= at && row < at + remove) continue;
                if (row >= at + remove) row -= remove;
            }
            result[row] = kv.Value;
        }
        return result;
    }

    /// <summary>设置行高</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="height">高度（磅）</param>
    public static void SetRowHeight(Worksheet sheet, Int32 row, Double height)
    {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        sheet.RowHeights[row] = height;
    }

    /// <summary>设置列宽</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="col">列索引（0 基）</param>
    /// <param name="width">宽度（字符）</param>
    public static void SetColumnWidth(Worksheet sheet, Int32 col, Double width)
    {
        if (col < 0) throw new ArgumentOutOfRangeException(nameof(col));
        sheet.ColumnWidths[col] = width;
    }
    #endregion

    #region 样式
    /// <summary>读取单元格样式（不存在返回 null）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <returns>样式</returns>
    public static CellFormat? GetStyle(Worksheet sheet, Int32 row, Int32 col)
        => sheet.CellFormats.TryGetValue((row, col), out var format) ? format : null;

    /// <summary>按区域合并应用样式（未设置的属性保持原值）</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">区域</param>
    /// <param name="apply">样式设置委托（作用于该单元格样式副本）</param>
    public static void ApplyStyle(Worksheet sheet, ExcelCellRange range, Action<CellFormat> apply)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (range == null) throw new ArgumentNullException(nameof(range));
        if (apply == null) throw new ArgumentNullException(nameof(apply));

        for (var r = range.StartRow; r <= range.EndRow; r++)
        {
            for (var c = range.StartCol; c <= range.EndCol; c++)
            {
                var format = GetStyle(sheet, r, c)?.Clone() ?? new CellFormat();
                apply(format);
                sheet.CellFormats[(r, c)] = format;
            }
        }
    }

    /// <summary>设置区域数字格式</summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">区域</param>
    /// <param name="format">数字格式（如 <c>#,##0.00</c>）</param>
    public static void SetNumberFormat(Worksheet sheet, ExcelCellRange range, String format)
        => ApplyStyle(sheet, range, f => f.NumberFormat = format);
    #endregion

    #region 辅助
    /// <summary>把对象转成写入用的基础类型（<see cref="Double"/> 保持数值，其余转字符串）</summary>
    /// <param name="value">原始值</param>
    /// <returns>规范化后的值</returns>
    public static Object? NormalizeValue(Object? value)
    {
        if (value == null || value is String || value is Boolean || value is DateTime) return value;
        if (value is Double || value is Single || value is Decimal || value is Byte || value is SByte
            || value is Int16 || value is UInt16 || value is Int32 || value is UInt32 || value is Int64 || value is UInt64)
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }
    #endregion
}
