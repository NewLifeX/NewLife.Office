using System.Globalization;
using System.Text;

namespace NewLife.Office.Csv;

/// <summary>CSV 文本格式化：字段转义、行拼接与分隔符识别</summary>
/// <remarks>
/// <c>CsvFile</c> 的写入口用 <c>ToString()</c> 输出字段，结果随当前区域变化（德语下 1.5 写成 1,5 会破坏列结构），
/// 因此写 CSV 一律经 <see cref="FormatValue"/> 转成与区域无关的文本。
/// </remarks>
internal static class CsvText
{
    #region 分隔符
    /// <summary>按首行识别分隔符</summary>
    /// <param name="sample">文本样本（通常为首行）</param>
    /// <returns>分隔符；无法判断时返回逗号</returns>
    public static Char DetectSeparator(String? sample)
    {
        if (String.IsNullOrEmpty(sample)) return ',';

        var comma = 0;
        var semicolon = 0;
        var tab = 0;
        var quoted = false;
        for (var i = 0; i < sample!.Length; i++)
        {
            var ch = sample[i];
            if (ch == '"')
            {
                quoted = !quoted;
                continue;
            }
            if (quoted) continue;

            switch (ch)
            {
                case ',': comma++; break;
                case ';': semicolon++; break;
                case '\t': tab++; break;
            }
        }

        var best = Math.Max(comma, Math.Max(semicolon, tab));
        if (best == 0) return ',';
        if (comma == best) return ',';
        return semicolon == best ? ';' : '\t';
    }
    #endregion

    #region 格式化
    /// <summary>按固定区域格式化字段值</summary>
    /// <param name="value">字段值，null/DBNull 输出空串</param>
    /// <returns>与区域无关的文本；日期输出 yyyy-MM-dd HH:mm:ss</returns>
    public static String FormatValue(Object? value)
    {
        if (value == null || value == DBNull.Value) return "";
        if (value is DateTime date) return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        if (value is DateTimeOffset offset) return offset.ToString("yyyy-MM-dd HH:mm:sszzz", CultureInfo.InvariantCulture);
        if (value is Boolean flag) return flag ? "true" : "false";
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
    }

    /// <summary>按 RFC4180 转义单个字段</summary>
    /// <param name="value">字段文本</param>
    /// <param name="separator">分隔符</param>
    /// <returns>含分隔符/引号/换行时用双引号包裹并转义内部引号</returns>
    public static String Escape(String? value, Char separator)
    {
        if (String.IsNullOrEmpty(value)) return "";
        if (value!.IndexOfAny([separator, '"', '\n', '\r']) >= 0)
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    /// <summary>拼接一行 CSV 文本（不含行尾换行）</summary>
    /// <param name="values">字段集合</param>
    /// <param name="separator">分隔符</param>
    /// <returns>CSV 行文本</returns>
    public static String BuildLine(IEnumerable<String?> values, Char separator)
    {
        var sb = new StringBuilder();
        var first = true;
        foreach (var value in values)
        {
            if (!first) sb.Append(separator);
            sb.Append(Escape(value, separator));
            first = false;
        }
        return sb.ToString();
    }
    #endregion
}
