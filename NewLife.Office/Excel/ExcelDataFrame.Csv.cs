using System.Globalization;
using NewLife.IO;
using NewLife.Office.Csv;

namespace NewLife.Office.Excel;

/// <summary>ExcelDataFrame 的 CSV 导入导出</summary>
/// <remarks>
/// CSV 解析与写入由 NewLife.Core 的 <see cref="CsvFile"/> 完成，本文件只做两件事：
/// 把 <see cref="CsvDocument"/> 的行数据转成分析模型，以及在导出前按固定区域格式化字段。
/// <para>示例：</para>
/// <code>
/// var frame = ExcelDataFrame.FromCsv("sales.csv")
///     .Where(row =&gt; Convert.ToDouble(row["金额"]) &gt; 0)
///     .GroupBy("地区")
///     .Sum("金额");
/// frame.ToCsv("summary.csv");
/// </code>
/// </remarks>
public sealed partial class ExcelDataFrame
{
    #region 导入
    /// <summary>从 CSV 文件创建数据表</summary>
    /// <param name="path">CSV 文件路径</param>
    /// <param name="options">导入选项，null 使用默认（自动识别编码与分隔符，首行为表头，推断类型）</param>
    /// <returns>内存中的数据表</returns>
    /// <exception cref="ArgumentNullException">path 为空</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    /// <exception cref="ArgumentException">文件无数据或列名重复</exception>
    public static ExcelDataFrame FromCsv(String path, CsvOptions? options = null)
    {
        if (String.IsNullOrWhiteSpace(path)) throw new ArgumentNullException(nameof(path));

        return FromCsv(CsvDocument.ReadFile(path, options), options);
    }

    /// <summary>从 CSV 数据流创建数据表</summary>
    /// <param name="stream">CSV 数据流，调用方负责释放</param>
    /// <param name="options">导入选项，null 使用默认（自动识别编码与分隔符，首行为表头，推断类型）</param>
    /// <returns>内存中的数据表</returns>
    /// <exception cref="ArgumentNullException">stream 为空</exception>
    /// <exception cref="ArgumentException">流无数据或列名重复</exception>
    public static ExcelDataFrame FromCsv(Stream stream, CsvOptions? options = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        return FromCsv(CsvDocument.Read(stream, options), options);
    }

    private static ExcelDataFrame FromCsv(CsvDocument document, CsvOptions? options)
    {
        var hasHeader = options?.HasHeader ?? true;
        var inferTypes = options?.InferTypes ?? true;
        var rows = document.Rows;
        if (rows.Count == 0) throw new ArgumentException(L.T("The CSV file contains no data."));

        var columnCount = 0;
        foreach (var row in rows)
            columnCount = Math.Max(columnCount, row.Length);

        var types = inferTypes ? InferColumnTypes(rows, columnCount, hasHeader) : null;

        var header = hasHeader ? rows[0] : [];
        var table = new List<Object?[]>(rows.Count)
        {
            BuildHeader(header, columnCount),
        };
        for (var r = hasHeader ? 1 : 0; r < rows.Count; r++)
        {
            var source = rows[r];
            var values = new Object?[columnCount];
            for (var c = 0; c < columnCount; c++)
            {
                var text = c < source.Length ? source[c] : null;
                values[c] = types == null ? NormalizeText(text) : ConvertCell(text, types[c]);
            }
            table.Add(values);
        }

        // 交给 FromRows 统一处理列名去重与长度补齐
        return FromRows(table, true);
    }

    private static Object?[] BuildHeader(String[] header, Int32 columnCount)
    {
        var values = new Object?[columnCount];
        for (var i = 0; i < columnCount; i++)
        {
            var name = i < header.Length ? header[i] : null;
            values[i] = String.IsNullOrWhiteSpace(name) ? $"Column{i + 1}" : name;
        }
        return values;
    }
    #endregion

    #region 导出
    /// <summary>导出为 CSV 文件</summary>
    /// <remarks>数值与日期按固定区域格式化，UTF-8 带 BOM，Excel 可直接打开。</remarks>
    /// <param name="path">输出路径，目录不存在时自动创建</param>
    /// <param name="includeHeader">是否输出列名行</param>
    /// <param name="separator">字段分隔符，默认逗号</param>
    public void ToCsv(String path, Boolean includeHeader = true, Char separator = ',')
    {
        if (String.IsNullOrWhiteSpace(path)) throw new ArgumentNullException(nameof(path));

        var fullPath = path.GetFullPath();
        fullPath.EnsureDirectory(true);
        using var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.Read);
        WriteCsv(fs, includeHeader, separator);
    }

    /// <summary>导出为 CSV 数据流</summary>
    /// <param name="stream">目标流，调用方负责释放</param>
    /// <param name="includeHeader">是否输出列名行</param>
    /// <param name="separator">字段分隔符，默认逗号</param>
    public void ToCsv(Stream stream, Boolean includeHeader = true, Char separator = ',')
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        WriteCsv(stream, includeHeader, separator);
    }

    private void WriteCsv(Stream stream, Boolean includeHeader, Char separator)
    {
        using var csv = new CsvFile(stream, leaveOpen: true) { Separator = separator };

        if (includeHeader)
        {
            var header = new Object[_columns.Length];
            for (var i = 0; i < _columns.Length; i++)
                header[i] = _columns[i];
            csv.WriteLine(header);
        }

        foreach (var row in _rows)
        {
            var values = new Object[row.Length];
            for (var i = 0; i < row.Length; i++)
                values[i] = CsvText.FormatValue(row[i]);
            csv.WriteLine(values);
        }
    }
    #endregion

    #region 类型推断
    private enum CsvColumnType
    {
        Text,
        Int64,
        Double,
        Date,
        Boolean,
    }

    private static readonly String[] _csvDateFormats =
    [
        "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
        "yyyy-M-d", "yyyy/M/d", "yyyy.M.d",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss.fff",
        "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
        "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.fff",
    ];

    private static CsvColumnType[] InferColumnTypes(IReadOnlyList<String[]> rows, Int32 columnCount, Boolean hasHeader)
    {
        var types = new CsvColumnType[columnCount];
        for (var c = 0; c < columnCount; c++)
        {
            var any = false;
            var allInt = true;
            var allDouble = true;
            var allDate = true;
            var allBoolean = true;
            for (var r = hasHeader ? 1 : 0; r < rows.Count; r++)
            {
                var row = rows[r];
                var text = c < row.Length ? row[c] : null;
                if (String.IsNullOrWhiteSpace(text)) continue;

                any = true;
                var value = text!.Trim();
                if (allInt && !IsInt64(value)) allInt = false;
                if (allDouble && !IsDouble(value)) allDouble = false;
                if (allDate && !IsDate(value)) allDate = false;
                if (allBoolean && !IsBoolean(value)) allBoolean = false;
                if (!allInt && !allDouble && !allDate && !allBoolean) break;
            }

            types[c] = !any ? CsvColumnType.Text
                : allInt ? CsvColumnType.Int64
                : allDouble ? CsvColumnType.Double
                : allDate ? CsvColumnType.Date
                : allBoolean ? CsvColumnType.Boolean
                : CsvColumnType.Text;
        }
        return types;
    }

    private static Boolean IsInt64(String text)
    {
        // 前导零（工号/邮编）保持字符串
        if (text.Length > 1 && text[0] == '0') return false;
        if (text.Length > 2 && text[0] == '-' && text[1] == '0') return false;

        var digitCount = text.Length;
        if (digitCount > 0 && (text[0] == '+' || text[0] == '-')) digitCount--;
        if (digitCount > 15) return false;

        return Int64.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _);
    }

    private static Boolean IsDouble(String text)
    {
        // 前导零（工号/邮编/编号）保持字符串；0.5 这类小数不受影响
        if (text.Length > 1 && text[0] == '0' && Char.IsDigit(text[1])) return false;
        if (text.Length > 2 && text[0] == '-' && text[1] == '0' && Char.IsDigit(text[2])) return false;

        if (!Double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var value)) return false;
        if (Double.IsNaN(value) || Double.IsInfinity(value)) return false;

        // 超过 15 位有效数字（订单号、身份证号）保持字符串，避免丢精度
        var digits = 0;
        var seen = false;
        foreach (var ch in text)
        {
            if (ch < '0' || ch > '9') continue;
            if (ch != '0') seen = true;
            if (seen) digits++;
        }
        return digits <= 15;
    }

    private static Boolean IsDate(String text)
        => DateTime.TryParseExact(text, _csvDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

    private static Boolean IsBoolean(String text)
        => String.Equals(text, "true", StringComparison.OrdinalIgnoreCase) || String.Equals(text, "false", StringComparison.OrdinalIgnoreCase);

    private static Object? ConvertCell(String? text, CsvColumnType type)
    {
        if (String.IsNullOrWhiteSpace(text)) return null;

        var value = text!.Trim();
        return type switch
        {
            CsvColumnType.Int64 => Int64.Parse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture),
            CsvColumnType.Double => Double.Parse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture),
            CsvColumnType.Date => DateTime.ParseExact(value, _csvDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None),
            CsvColumnType.Boolean => Boolean.Parse(value),
            _ => text,
        };
    }

    private static Object? NormalizeText(String? text) => String.IsNullOrWhiteSpace(text) ? null : text;
    #endregion
}
