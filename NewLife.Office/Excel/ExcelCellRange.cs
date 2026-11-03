using System.Globalization;
using System.Text;

namespace NewLife.Office.Excel;

/// <summary>Excel A1 区域引用（如 <c>Sheet1!A1:D20</c>）</summary>
/// <remarks>
/// 支持整列（<c>A:C</c>）、整行（<c>1:3</c>）与单元格区域（<c>A1:D20</c>）形式，行列索引均为 0 基、区间闭区间。
/// 供面向 AI 的区域读写与补丁定位使用。
/// <example>
/// <code>
/// var range = ExcelCellRange.Parse("Sheet1!A1:D20");
/// Console.WriteLine(range.StartRow);   // 0
/// Console.WriteLine(range.RowCount);   // 20
/// </code>
/// </example>
/// </remarks>
public class ExcelCellRange
{
    /// <summary>Excel 最大行数（含）</summary>
    public const Int32 MaxRow = 1048575;

    /// <summary>Excel 最大列数（含）</summary>
    public const Int32 MaxCol = 16383;

    #region 属性
    /// <summary>工作表名称，null 表示使用默认工作表</summary>
    public String? Sheet { get; set; }

    /// <summary>起始行（0 基，含）</summary>
    public Int32 StartRow { get; set; }

    /// <summary>起始列（0 基，含）</summary>
    public Int32 StartCol { get; set; }

    /// <summary>结束行（0 基，含）</summary>
    public Int32 EndRow { get; set; }

    /// <summary>结束列（0 基，含）</summary>
    public Int32 EndCol { get; set; }

    /// <summary>行数</summary>
    public Int32 RowCount => EndRow - StartRow + 1;

    /// <summary>列数</summary>
    public Int32 ColumnCount => EndCol - StartCol + 1;

    /// <summary>是否单单元格</summary>
    public Boolean IsSingleCell => StartRow == EndRow && StartCol == EndCol;

    /// <summary>A1 记法（不含工作表名，如 <c>A1:D20</c>）</summary>
    public String Address => IsSingleCell
        ? $"{ColumnName(StartCol)}{StartRow + 1}"
        : $"{ColumnName(StartCol)}{StartRow + 1}:{ColumnName(EndCol)}{EndRow + 1}";

    /// <summary>带工作表名的完整记法（如 <c>Sheet1!A1:D20</c>）</summary>
    public String FullAddress => Sheet.IsNullOrEmpty() ? Address : $"{QuoteSheet(Sheet!)}!{Address}";
    #endregion

    #region 构造
    /// <summary>实例化空区域</summary>
    public ExcelCellRange() { }

    /// <summary>实例化区域</summary>
    /// <param name="sheet">工作表名称</param>
    /// <param name="startRow">起始行（0 基）</param>
    /// <param name="startCol">起始列（0 基）</param>
    /// <param name="endRow">结束行（0 基）</param>
    /// <param name="endCol">结束列（0 基）</param>
    public ExcelCellRange(String? sheet, Int32 startRow, Int32 startCol, Int32 endRow, Int32 endCol)
    {
        Sheet = sheet;
        StartRow = startRow;
        StartCol = startCol;
        EndRow = endRow;
        EndCol = endCol;
    }
    #endregion

    #region 解析
    /// <summary>解析 A1 记法区域</summary>
    /// <param name="text">区域文本，如 <c>Sheet1!A1:D20</c>、<c>A1</c>、<c>A:C</c>、<c>Sheet1</c></param>
    /// <param name="defaultSheet">文本未带工作表名时使用的工作表</param>
    /// <returns>区域对象，非法输入返回 null</returns>
    public static ExcelCellRange? Parse(String? text, String? defaultSheet = null)
    {
        if (text.IsNullOrWhiteSpace()) return null;

        var s = text!.Trim();
        String? sheet = defaultSheet;
        var bang = s.LastIndexOf('!');
        if (bang >= 0)
        {
            sheet = UnquoteSheet(s[..bang].Trim());
            s = s[(bang + 1)..].Trim();
        }

        if (s.Length == 0)
            return new ExcelCellRange(sheet, 0, 0, MaxRow, MaxCol);

        var colon = s.IndexOf(':');
        if (colon < 0)
        {
            if (TryParseCell(s, out var row, out var col, out var wholeColumn, out var wholeRow))
            {
                if (wholeColumn) return new ExcelCellRange(sheet, 0, col, MaxRow, col);
                if (wholeRow) return new ExcelCellRange(sheet, row, 0, row, MaxCol);
                return new ExcelCellRange(sheet, row, col, row, col);
            }
            return null;
        }

        var left = s[..colon].Trim();
        var right = s[(colon + 1)..].Trim();
        if (!TryParseCell(left, out var r1, out var c1, out var colOnly1, out var rowOnly1)) return null;
        if (!TryParseCell(right, out var r2, out var c2, out var colOnly2, out var rowOnly2)) return null;

        if (colOnly1 && colOnly2) return new ExcelCellRange(sheet, 0, c1, MaxRow, c2);
        if (rowOnly1 && rowOnly2) return new ExcelCellRange(sheet, r1, 0, r2, MaxCol);
        if (colOnly1 || colOnly2 || rowOnly1 || rowOnly2) return null;

        if (r1 > r2) (r1, r2) = (r2, r1);
        if (c1 > c2) (c1, c2) = (c2, c1);
        return new ExcelCellRange(sheet, r1, c1, r2, c2);
    }

    /// <summary>解析单元格引用（支持 $A$1 绝对引用）</summary>
    /// <param name="text">引用文本</param>
    /// <param name="row">行索引（0 基）</param>
    /// <param name="col">列索引（0 基）</param>
    /// <param name="columnOnly">是否整列（如 A）</param>
    /// <param name="rowOnly">是否整行（如 3）</param>
    /// <returns>是否解析成功</returns>
    public static Boolean TryParseCell(String text, out Int32 row, out Int32 col, out Boolean columnOnly, out Boolean rowOnly)
    {
        row = 0;
        col = 0;
        columnOnly = false;
        rowOnly = false;
        if (text == null || String.IsNullOrWhiteSpace(text)) return false;

        var s = text.Trim().Replace("$", String.Empty);
        var i = 0;
        while (i < s.Length && Char.IsLetter(s[i])) i++;
        var letters = s.Substring(0, i);
        var digits = s.Substring(i);

        if (letters.Length > 0 && digits.Length == 0)
        {
            columnOnly = true;
            return TryColumnIndex(letters, out col);
        }
        if (letters.Length == 0 && digits.Length > 0)
        {
            if (!Int32.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var r)) return false;
            if (r < 1 || r > MaxRow + 1) return false;
            rowOnly = true;
            row = r - 1;
            return true;
        }
        if (letters.Length == 0 || digits.Length == 0) return false;

        if (!TryColumnIndex(letters, out col)) return false;
        if (!Int32.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowNo)) return false;
        if (rowNo < 1 || rowNo > MaxRow + 1) return false;
        row = rowNo - 1;
        return true;
    }
    #endregion

    #region 列名
    /// <summary>列索引转列名（0 → A，26 → AA）</summary>
    /// <param name="col">列索引（0 基）</param>
    /// <returns>列名</returns>
    public static String ColumnName(Int32 col)
    {
        if (col < 0) return String.Empty;

        var sb = new StringBuilder();
        var n = col + 1;
        while (n > 0)
        {
            var m = (n - 1) % 26;
            sb.Insert(0, (Char)('A' + m));
            n = (n - 1) / 26;
        }
        return sb.ToString();
    }

    /// <summary>列名转列索引（A → 0，AA → 26）</summary>
    /// <param name="name">列名</param>
    /// <returns>列索引（0 基），非法返回 -1</returns>
    public static Int32 ColumnIndex(String? name)
    {
        return TryColumnIndex(name, out var col) ? col : -1;
    }

    private static Boolean TryColumnIndex(String? name, out Int32 col)
    {
        col = 0;
        if (name.IsNullOrEmpty()) return false;

        foreach (var ch in name!)
        {
            if (ch < 'A' || ch > 'Z') return false;
            col = col * 26 + (ch - 'A' + 1);
        }
        col -= 1;
        return col >= 0 && col <= MaxCol;
    }
    #endregion

    #region 辅助
    /// <summary>把区域裁剪到已用行列范围内</summary>
    /// <param name="rows">已用行数</param>
    /// <param name="columns">已用列数</param>
    /// <returns>裁剪后的区域</returns>
    public ExcelCellRange Clamp(Int32 rows, Int32 columns)
    {
        var maxRow = rows > 0 ? rows - 1 : 0;
        var maxCol = columns > 0 ? columns - 1 : 0;
        return new ExcelCellRange(Sheet, StartRow, StartCol, Math.Min(EndRow, maxRow), Math.Min(EndCol, maxCol));
    }

    /// <summary>工作表名加上引号（含空格或特殊字符时）</summary>
    /// <param name="sheet">工作表名称</param>
    /// <returns>可解析的引用片段</returns>
    public static String QuoteSheet(String sheet)
    {
        if (sheet.IndexOfAny([' ', '\'', '!', '-', '.']) < 0) return sheet;
        return "'" + sheet.Replace("'", "''") + "'";
    }

    private static String UnquoteSheet(String text)
    {
        if (text.Length >= 2 && text[0] == '\'' && text[^1] == '\'')
            return text[1..^1].Replace("''", "'");
        return text;
    }

    /// <inheritdoc/>
    public override String ToString() => FullAddress;
    #endregion
}
