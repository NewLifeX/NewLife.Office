namespace NewLife.Office.Pdf;

/// <summary>PDF 内容流文本提取器。按 <c>Tf</c> 跟踪当前字体，经 <see cref="PdfFontDecoder"/> 还原字形编码为 Unicode</summary>
/// <remarks>
/// <para>解析内容流中的文本操作符：<c>Tj</c> / <c>TJ</c> / <c>'</c> / <c>"</c>，以及换行信号
/// <c>T*</c> / <c>Td</c> / <c>TD</c> / <c>ET</c>。</para>
/// <para>不依赖字体时（<paramref name="fonts"/> 为空）退化为原先的 UTF-16BE / Latin-1 解读，
/// 保持对旧生成器输出的兼容。</para>
/// </remarks>
internal static class PdfTextExtractor
{
    /// <summary>读取括号字面量字符串为字节（供按位置提取复用）</summary>
    internal static Byte[] ReadLiteralBytes(String content, ref Int32 i) => ReadLiteralString(content, ref i);

    /// <summary>解析十六进制字符串（不含尖括号）为字节</summary>
    internal static Byte[] ParseHexBytes(String hex)
    {
        var digits = new List<Byte>(hex.Length);
        foreach (var ch in hex)
        {
            if (ch is >= '0' and <= '9') digits.Add((Byte)(ch - '0'));
            else if (ch is >= 'a' and <= 'f') digits.Add((Byte)(ch - 'a' + 10));
            else if (ch is >= 'A' and <= 'F') digits.Add((Byte)(ch - 'A' + 10));
        }
        if (digits.Count % 2 == 1) digits.Add(0);

        var bytes = new Byte[digits.Count / 2];
        for (var k = 0; k < bytes.Length; k++)
            bytes[k] = (Byte)((digits[k * 2] << 4) | digits[k * 2 + 1]);

        return bytes;
    }

    /// <summary>按当前字体解码字符串字节；无字体信息时退回 UTF-16BE / Latin-1 兼容解读</summary>
    internal static String DecodeBytes(Byte[] bytes, Dictionary<String, PdfFontDecoder>? fonts, String? fontName)
    {
        if (bytes == null || bytes.Length == 0) return String.Empty;

        if (fonts != null && fontName != null && fonts.TryGetValue(fontName, out var font))
            return font.Decode(bytes);

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return System.Text.Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0)
            return System.Text.Encoding.BigEndianUnicode.GetString(bytes);

        return System.Text.Encoding.GetEncoding(28591).GetString(bytes);
    }

    /// <summary>提取内容流文本并追加到缓冲</summary>
    /// <param name="content">内容流文本（Latin-1 解读）</param>
    /// <param name="fonts">当前页面的字体解码表</param>
    /// <param name="sb">输出缓冲</param>
    public static void ExtractContent(String content, Dictionary<String, PdfFontDecoder>? fonts, System.Text.StringBuilder sb)
    {
        if (content.IsNullOrEmpty()) return;

        var state = new State(fonts);
        var i = 0;
        while (i < content.Length)
        {
            var c = content[i];
            switch (c)
            {
                case '%':
                    // 注释：跳到行尾
                    var nl = content.IndexOf('\n', i);
                    i = nl < 0 ? content.Length : nl + 1;
                    break;

                case '(':
                    state.Operands.Add(ReadLiteralString(content, ref i));
                    break;

                case '<':
                    if (i + 1 < content.Length && content[i + 1] == '<')
                        i = SkipDictionary(content, i);
                    else
                        state.Operands.Add(ReadHexString(content, ref i));
                    break;

                case '[':
                    state.Operands.Add(ReadArray(content, ref i));
                    break;

                case '/':
                    {
                        // 名字操作数（字体名等），保留不含 '/' 的值
                        var nameStart = ++i;
                        while (i < content.Length && !Char.IsWhiteSpace(content[i]) && !IsDelimiter(content[i])) i++;
                        state.Operands.Add(content[nameStart..i]);
                        break;
                    }

                case ']':
                case '>':
                case '{':
                case '}':
                    i++;
                    break;

                default:
                    if (Char.IsWhiteSpace(c))
                    {
                        i++;
                        break;
                    }

                    var start = i;
                    while (i < content.Length && !Char.IsWhiteSpace(content[i]) && !IsDelimiter(content[i])) i++;
                    var token = content[start..i];
                    if (i == start)
                    {
                        i++;
                        break;
                    }

                    ProcessToken(token, state, sb);
                    break;
            }
        }
    }

    private static Boolean IsDelimiter(Char c) => c is '(' or ')' or '<' or '>' or '[' or ']' or '{' or '}' or '/' or '%';

    /// <summary>处理一个操作符或操作数</summary>
    private static void ProcessToken(String token, State state, System.Text.StringBuilder sb)
    {
        // 数字操作数
        if (token.Length > 0 && (Char.IsDigit(token[0]) || token[0] == '-' || token[0] == '+' || token[0] == '.'))
        {
            if (Double.TryParse(token, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var num))
                state.Operands.Add(num);
            return;
        }

        // 名字操作数（字体名等）
        if (token[0] == '/')
        {
            state.Operands.Add(token[1..]);
            return;
        }

        switch (token)
        {
            case "Tf":
                // /FontName size Tf
                if (state.Operands.Count >= 2 && state.Operands[^2] is String fontName)
                    state.CurrentFont = fontName;
                break;

            case "Tj":
                state.ShowLastString(sb);
                break;

            case "TJ":
                if (state.Operands.Count > 0 && state.Operands[^1] is List<Object> items)
                {
                    foreach (var item in items)
                    {
                        if (item is Byte[] bytes)
                            sb.Append(state.Decode(bytes));
                        else if (item is Double adj && adj < -200)
                            sb.Append(' ');
                    }
                }
                break;

            case "'":
                AppendLine(sb);
                state.ShowLastString(sb);
                break;

            case "\"":
                AppendLine(sb);
                state.ShowLastString(sb);
                break;

            case "T*":
                AppendLine(sb);
                break;

            case "Td":
            case "TD":
                // 第二个操作数为纵向位移，非零视为换行
                if (state.Operands.Count >= 2 && state.Operands[^1] is Double dy && Math.Abs(dy) > 0.01)
                    AppendLine(sb);
                break;

            case "ET":
                AppendLine(sb);
                break;
        }

        // 操作符消费掉操作数
        if (token is "Tf" or "Tj" or "TJ" or "'" or "\"" or "T*" or "Td" or "TD" or "ET" or "BT")
            state.Operands.Clear();
    }

    /// <summary>追加换行，避免连续空行</summary>
    private static void AppendLine(System.Text.StringBuilder sb)
    {
        if (sb.Length == 0) return;
        if (sb[^1] == '\n') return;

        sb.Append('\n');
    }

    /// <summary>读取括号字面量字符串为字节</summary>
    private static Byte[] ReadLiteralString(String content, ref Int32 i)
    {
        var bytes = new List<Byte>(32);
        var depth = 1;
        i++; // 跳过 '('

        while (i < content.Length)
        {
            var c = content[i++];
            if (c == '\\' && i < content.Length)
            {
                var e = content[i++];
                switch (e)
                {
                    case 'n': bytes.Add((Byte)'\n'); break;
                    case 'r': bytes.Add((Byte)'\r'); break;
                    case 't': bytes.Add((Byte)'\t'); break;
                    case 'b': bytes.Add((Byte)'\b'); break;
                    case 'f': bytes.Add((Byte)'\f'); break;
                    case '(': bytes.Add((Byte)'('); break;
                    case ')': bytes.Add((Byte)')'); break;
                    case '\\': bytes.Add((Byte)'\\'); break;
                    case >= '0' and <= '7':
                        {
                            var value = e - '0';
                            for (var k = 0; k < 2 && i < content.Length && content[i] >= '0' && content[i] <= '7'; k++)
                                value = value * 8 + (content[i++] - '0');
                            bytes.Add((Byte)value);
                            break;
                        }
                    default:
                        bytes.Add((Byte)e);
                        break;
                }
                continue;
            }

            if (c == '(')
            {
                depth++;
                bytes.Add((Byte)c);
                continue;
            }
            if (c == ')')
            {
                depth--;
                if (depth == 0) break;
                bytes.Add((Byte)c);
                continue;
            }

            bytes.Add((Byte)c);
        }

        return [.. bytes];
    }

    /// <summary>读取十六进制字符串 &lt;hex&gt; 为字节</summary>
    private static Byte[] ReadHexString(String content, ref Int32 i)
    {
        var end = content.IndexOf('>', i + 1);
        if (end < 0)
        {
            i = content.Length;
            return [];
        }

        var hex = content[(i + 1)..end];
        i = end + 1;

        var digits = new List<Byte>(hex.Length);
        foreach (var ch in hex)
        {
            if (ch is >= '0' and <= '9') digits.Add((Byte)(ch - '0'));
            else if (ch is >= 'a' and <= 'f') digits.Add((Byte)(ch - 'a' + 10));
            else if (ch is >= 'A' and <= 'F') digits.Add((Byte)(ch - 'A' + 10));
        }
        if (digits.Count % 2 == 1) digits.Add(0);

        var bytes = new Byte[digits.Count / 2];
        for (var k = 0; k < bytes.Length; k++)
            bytes[k] = (Byte)((digits[k * 2] << 4) | digits[k * 2 + 1]);

        return bytes;
    }

    /// <summary>读取数组操作数（TJ 数组），元素为字符串字节或数字</summary>
    private static List<Object> ReadArray(String content, ref Int32 i)
    {
        var items = new List<Object>();
        i++; // 跳过 '['

        while (i < content.Length)
        {
            var c = content[i];
            if (Char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }
            if (c == ']')
            {
                i++;
                break;
            }

            if (c == '(')
            {
                items.Add(ReadLiteralString(content, ref i));
                continue;
            }
            if (c == '<')
            {
                if (i + 1 < content.Length && content[i + 1] == '<')
                {
                    i = SkipDictionary(content, i);
                    continue;
                }
                items.Add(ReadHexString(content, ref i));
                continue;
            }

            var start = i;
            while (i < content.Length && !Char.IsWhiteSpace(content[i]) && !IsDelimiter(content[i])) i++;
            if (i == start)
            {
                i++;
                continue;
            }

            var token = content[start..i];
            if (Double.TryParse(token, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var num))
                items.Add(num);
        }

        return items;
    }

    /// <summary>跳过 &lt;&lt; ... &gt;&gt; 字典</summary>
    private static Int32 SkipDictionary(String content, Int32 i)
    {
        var depth = 0;
        while (i < content.Length)
        {
            if (content[i] == '<' && i + 1 < content.Length && content[i + 1] == '<')
            {
                depth++;
                i += 2;
                continue;
            }
            if (content[i] == '>' && i + 1 < content.Length && content[i + 1] == '>')
            {
                depth--;
                i += 2;
                if (depth <= 0) return i;
                continue;
            }
            i++;
        }

        return i;
    }

    /// <summary>提取过程中的状态</summary>
    private sealed class State(Dictionary<String, PdfFontDecoder>? fonts)
    {
        private readonly Dictionary<String, PdfFontDecoder>? _fonts = fonts;

        /// <summary>当前字体名</summary>
        public String? CurrentFont { get; set; }

        /// <summary>待处理操作数</summary>
        public List<Object> Operands { get; } = [];

        /// <summary>用当前字体解码字符串</summary>
        public String Decode(Byte[] bytes)
        {
            if (_fonts != null && CurrentFont != null && _fonts.TryGetValue(CurrentFont, out var font))
                return font.Decode(bytes);

            // 无字体信息：沿用历史行为，2 字节按 UTF-16BE，否则 Latin-1
            if (bytes.Length >= 2 && (bytes.Length % 2 == 0) && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return System.Text.Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0 && bytes[1] != 0)
                return System.Text.Encoding.BigEndianUnicode.GetString(bytes);

            return System.Text.Encoding.GetEncoding(28591).GetString(bytes);
        }

        /// <summary>显示最后一个字符串操作数</summary>
        public void ShowLastString(System.Text.StringBuilder sb)
        {
            for (var i = Operands.Count - 1; i >= 0; i--)
            {
                if (Operands[i] is Byte[] bytes)
                {
                    sb.Append(Decode(bytes));
                    return;
                }
            }
        }
    }
}
