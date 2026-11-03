using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>极简 JSON 解析器（零依赖、AOT 友好），供补丁文件等结构化输入使用</summary>
/// <remarks>
/// 只做"够用"的范围：对象 → <see cref="Dictionary{TKey,TValue}"/>，数组 → <see cref="List{T}"/>，
/// 数字 → <see cref="Double"/>，字符串/布尔/null → 原生类型。不支持注释、尾逗号。
/// <para>选用自研而非 System.Text.Json：本库需兼容 net45，且坚持零外部依赖。</para>
/// </remarks>
internal static class JsonLite
{
    /// <summary>解析 JSON 文本</summary>
    /// <param name="text">JSON 文本</param>
    /// <returns>根节点（字典/列表/字符串/数字/布尔/null）</returns>
    /// <exception cref="FormatException">JSON 非法</exception>
    public static Object? Parse(String text)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        var pos = 0;
        var value = ParseValue(text, ref pos);
        SkipWhite(text, ref pos);
        if (pos < text.Length) throw new FormatException($"Unexpected trailing characters at {pos}.");
        return value;
    }

    /// <summary>读取字符串字段（不存在或为 null 返回 null）</summary>
    public static String? GetString(Dictionary<String, Object?> obj, String key)
    {
        if (!obj.TryGetValue(key, out var value) || value == null) return null;
        return value as String ?? Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    /// <summary>读取整数字段（不存在或非数字返回 null）</summary>
    public static Int32? GetInt32(Dictionary<String, Object?> obj, String key)
    {
        if (!obj.TryGetValue(key, out var value) || value == null) return null;
        if (value is Double d) return (Int32)d;
        if (value is String s && Double.TryParse(s, out var parsed)) return (Int32)parsed;
        return null;
    }

    /// <summary>读取布尔字段（不存在返回 null）</summary>
    public static Boolean? GetBoolean(Dictionary<String, Object?> obj, String key)
    {
        if (!obj.TryGetValue(key, out var value) || value == null) return null;
        if (value is Boolean b) return b;
        if (value is String s && Boolean.TryParse(s, out var parsed)) return parsed;
        return null;
    }

    #region 解析实现
    private static Object? ParseValue(String text, ref Int32 pos)
    {
        SkipWhite(text, ref pos);
        if (pos >= text.Length) throw new FormatException("Unexpected end of JSON.");

        return text[pos] switch
        {
            '{' => ParseObject(text, ref pos),
            '[' => ParseArray(text, ref pos),
            '"' => ParseString(text, ref pos),
            't' => ParseLiteral(text, ref pos, "true", true),
            'f' => ParseLiteral(text, ref pos, "false", false),
            'n' => ParseLiteral(text, ref pos, "null", null),
            _ => ParseNumber(text, ref pos),
        };
    }

    private static Dictionary<String, Object?> ParseObject(String text, ref Int32 pos)
    {
        var obj = new Dictionary<String, Object?>(StringComparer.Ordinal);
        pos++; // {
        SkipWhite(text, ref pos);
        if (pos < text.Length && text[pos] == '}') { pos++; return obj; }

        while (true)
        {
            SkipWhite(text, ref pos);
            var key = ParseString(text, ref pos);
            SkipWhite(text, ref pos);
            if (pos >= text.Length || text[pos] != ':') throw new FormatException($"Expected ':' at {pos}.");
            pos++;
            obj[key] = ParseValue(text, ref pos);
            SkipWhite(text, ref pos);

            if (pos >= text.Length) throw new FormatException("Unterminated object.");
            if (text[pos] == ',') { pos++; continue; }
            if (text[pos] == '}') { pos++; return obj; }
            throw new FormatException($"Expected ',' or '}}' at {pos}.");
        }
    }

    private static List<Object?> ParseArray(String text, ref Int32 pos)
    {
        var list = new List<Object?>();
        pos++; // [
        SkipWhite(text, ref pos);
        if (pos < text.Length && text[pos] == ']') { pos++; return list; }

        while (true)
        {
            list.Add(ParseValue(text, ref pos));
            SkipWhite(text, ref pos);

            if (pos >= text.Length) throw new FormatException("Unterminated array.");
            if (text[pos] == ',') { pos++; continue; }
            if (text[pos] == ']') { pos++; return list; }
            throw new FormatException($"Expected ',' or ']' at {pos}.");
        }
    }

    private static String ParseString(String text, ref Int32 pos)
    {
        if (pos >= text.Length || text[pos] != '"') throw new FormatException($"Expected string at {pos}.");
        pos++;

        var sb = new StringBuilder();
        while (pos < text.Length)
        {
            var c = text[pos++];
            if (c == '"') return sb.ToString();
            if (c != '\\') { sb.Append(c); continue; }

            if (pos >= text.Length) break;
            var esc = text[pos++];
            switch (esc)
            {
                case '"': sb.Append('"'); break;
                case '\\': sb.Append('\\'); break;
                case '/': sb.Append('/'); break;
                case 'b': sb.Append('\b'); break;
                case 'f': sb.Append('\f'); break;
                case 'n': sb.Append('\n'); break;
                case 'r': sb.Append('\r'); break;
                case 't': sb.Append('\t'); break;
                case 'u':
                    if (pos + 4 > text.Length) throw new FormatException("Invalid unicode escape.");
                    sb.Append((Char)Int32.Parse(text.Substring(pos, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    pos += 4;
                    break;
                default: throw new FormatException($"Invalid escape '\\{esc}'.");
            }
        }
        throw new FormatException("Unterminated string.");
    }

    private static Object? ParseLiteral(String text, ref Int32 pos, String literal, Object? value)
    {
        if (pos + literal.Length > text.Length || String.CompareOrdinal(text, pos, literal, 0, literal.Length) != 0)
            throw new FormatException($"Invalid literal at {pos}.");

        pos += literal.Length;
        return value;
    }

    private static Double ParseNumber(String text, ref Int32 pos)
    {
        var start = pos;
        if (pos < text.Length && (text[pos] == '-' || text[pos] == '+')) pos++;
        while (pos < text.Length && (Char.IsDigit(text[pos]) || text[pos] == '.' || text[pos] == 'e' || text[pos] == 'E' || text[pos] == '-' || text[pos] == '+')) pos++;
        if (pos == start) throw new FormatException($"Invalid number at {start}.");

        return Double.Parse(text.Substring(start, pos - start), CultureInfo.InvariantCulture);
    }

    private static void SkipWhite(String text, ref Int32 pos)
    {
        while (pos < text.Length && (text[pos] == ' ' || text[pos] == '\t' || text[pos] == '\r' || text[pos] == '\n')) pos++;
    }
    #endregion
}
