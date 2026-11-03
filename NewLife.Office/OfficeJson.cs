using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>极简 JSON 输出辅助（零依赖、AOT 友好）</summary>
/// <remarks>
/// 只负责把基础类型写成合法 JSON 片段，供各格式的结构快照手写序列化使用。
/// 解析侧见 <see cref="JsonLite"/>。
/// </remarks>
public static class OfficeJson
{
    /// <summary>转义为 JSON 字符串字面量（含首尾双引号）</summary>
    /// <param name="value">原始字符串，null 输出 null</param>
    /// <returns>JSON 字面量</returns>
    public static String Escape(String? value)
    {
        if (value == null) return "null";

        var sb = new StringBuilder(value.Length + 2);
        sb.Append('"');
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (ch < ' ') sb.Append("\\u").Append(((Int32)ch).ToString("x4"));
                    else sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    /// <summary>把基础类型写成 JSON 值</summary>
    /// <param name="value">值（null/布尔/数值/日期/字符串，其余调用 ToString）</param>
    /// <returns>JSON 片段</returns>
    public static String Value(Object? value) => value switch
    {
        null => "null",
        Boolean b => b ? "true" : "false",
        String s => Escape(s),
        DateTime dt => Escape(dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)),
        DateTimeOffset dto => Escape(dto.ToString("yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture)),
        Single f => f.ToString("R", CultureInfo.InvariantCulture),
        Double d => d.ToString("R", CultureInfo.InvariantCulture),
        Decimal m => m.ToString(CultureInfo.InvariantCulture),
        Byte or SByte or Int16 or UInt16 or Int32 or UInt32 or Int64 or UInt64
            => Convert.ToString(value, CultureInfo.InvariantCulture) ?? "0",
        _ => Escape(Convert.ToString(value, CultureInfo.InvariantCulture)),
    };

    /// <summary>输出二维数组</summary>
    /// <param name="rows">行集合</param>
    /// <returns>JSON 数组</returns>
    public static String Matrix(IEnumerable<Object?[]> rows)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        var first = true;
        foreach (var row in rows)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append('[');
            for (var i = 0; i < row.Length; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(Value(row[i]));
            }
            sb.Append(']');
        }
        sb.Append(']');
        return sb.ToString();
    }

    /// <summary>输出字符串数组</summary>
    /// <param name="items">字符串集合</param>
    /// <returns>JSON 数组</returns>
    public static String Array(IEnumerable<String?> items)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        var first = true;
        foreach (var item in items)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(Escape(item));
        }
        sb.Append(']');
        return sb.ToString();
    }
}
