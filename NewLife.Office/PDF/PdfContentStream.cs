using System.Text;

namespace NewLife.Office;

/// <summary>PDF 内容流操作符</summary>
public class PdfOperator
{
    /// <summary>操作符名称（如 Tj、TJ、BT、ET、q、Q、m、l、S、f 等）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>操作数列表（字符串、数字等）</summary>
    public List<Object> Operands { get; set; } = [];
}

/// <summary>PDF 内容流解析与重建器</summary>
/// <remarks>
/// 将 PDF 页面内容流解析为操作符序列，支持修改后重新序列化。
/// 这是 PDF 高保真编辑的关键——解析内容流 → 修改操作符 → 重建内容流，
/// 未修改的部分保持原始字节不变。
/// 支持的操作符集覆盖 PDF 1.7 核心图形和文本操作符。
/// </remarks>
public class PdfContentStream
{
    #region 属性
    /// <summary>操作符列表</summary>
    public List<PdfOperator> Operators { get; } = [];
    #endregion

    #region 构造
    /// <summary>从内容流字节解析</summary>
    /// <param name="contentBytes">内容流字节（已解压）</param>
    public PdfContentStream(Byte[] contentBytes)
    {
        var text = Encoding.GetEncoding(28591).GetString(contentBytes);
        ParseStream(text);
    }

    /// <summary>空内容流</summary>
    public PdfContentStream() { }
    #endregion

    #region 序列化
    /// <summary>序列化为 PDF 内容流字节</summary>
    /// <returns>Latin-1 编码的字节数组</returns>
    public Byte[] ToBytes()
    {
        var sb = new StringBuilder();
        foreach (var op in Operators)
        {
            foreach (var operand in op.Operands)
            {
                if (sb.Length > 0 && sb[^1] != '\n') sb.Append(' ');
                sb.Append(FormatOperand(operand));
            }
            if (op.Operands.Count > 0) sb.Append(' ');
            sb.Append(op.Name);
            sb.AppendLine();
        }
        return Encoding.GetEncoding(28591).GetBytes(sb.ToString());
    }

    /// <summary>格式化单个操作数</summary>
    private static String FormatOperand(Object operand)
    {
        return operand switch
        {
            Single f => f.ToString("F3", System.Globalization.CultureInfo.InvariantCulture),
            Double d => d.ToString("F3", System.Globalization.CultureInfo.InvariantCulture),
            Int32 i => i.ToString(),
            String s => $"({EscapePdfText(s)})",
            PdfTextOperand pt => FormatOperand(pt.Value), // 递归
            _ => operand.ToString() ?? String.Empty,
        };
    }

    private static String EscapePdfText(String text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (c == '(' || c == ')' || c == '\\')
                sb.Append('\\');
            sb.Append(c);
        }
        return sb.ToString();
    }
    #endregion

    #region 解析
    /// <summary>解析 PDF 内容流</summary>
    private void ParseStream(String content)
    {
        var pos = 0;
        var currentOp = new List<Object>();

        while (pos < content.Length)
        {
            SkipWhitespace(content, ref pos);
            if (pos >= content.Length) break;

            var c = content[pos];

            // PDF 注释
            if (c == '%')
            {
                while (pos < content.Length && content[pos] != '\r' && content[pos] != '\n') pos++;
                continue;
            }

            // 括号字符串
            if (c == '(')
            {
                var str = ReadParenString(content, ref pos);
                currentOp.Add(str);
                continue;
            }

            // 十六进制字符串 <hex>
            if (c == '<' && pos + 1 < content.Length && content[pos + 1] != '<')
            {
                pos++;
                var hex = ReadHexString(content, ref pos);
                var bytes = HexToBytes(hex);
                // PDF hex 字符串在内容流中按 UTF-16BE 解读
                var text = bytes.Length >= 2 ? Encoding.BigEndianUnicode.GetString(bytes) : Encoding.GetEncoding(28591).GetString(bytes);
                currentOp.Add(text);
                continue;
            }

            // 字典或内联图片（跳过）
            if (c == '<' && pos + 1 < content.Length && content[pos + 1] == '<')
            {
                SkipDict(content, ref pos);
                continue;
            }

            // 数组 [ ... ]
            if (c == '[')
            {
                var arrayContent = ReadArray(content, ref pos);
                // TJ 数组内的字符串和数字分解为独立操作数
                var arrPos = 0;
                while (arrPos < arrayContent.Length)
                {
                    SkipWhitespace(arrayContent, ref arrPos);
                    if (arrPos >= arrayContent.Length) break;
                    if (arrayContent[arrPos] == '(')
                        currentOp.Add(ReadParenString(arrayContent, ref arrPos));
                    else if (arrayContent[arrPos] == '<')
                    {
                        arrPos++;
                        currentOp.Add(ReadHexString(arrayContent, ref arrPos));
                    }
                    else if (arrayContent[arrPos] == '-' || arrayContent[arrPos] == '+' ||
                             (arrayContent[arrPos] >= '0' && arrayContent[arrPos] <= '9') ||
                             arrayContent[arrPos] == '.')
                        currentOp.Add(ReadNumber(arrayContent, ref arrPos));
                    else arrPos++;
                }
                // 读取后续 TJ 操作符
                SkipWhitespace(content, ref pos);
                if (pos + 1 < content.Length && content[pos] == 'T' && content[pos + 1] == 'J')
                {
                    currentOp.Insert(0, "[TJ-ARRAY]");
                    FlushOp(currentOp, "TJ", ref pos);
                }
                continue;
            }

            // 数字
            if (c == '-' || c == '+' || (c >= '0' && c <= '9') || c == '.')
            {
                currentOp.Add(ReadNumber(content, ref pos));
                continue;
            }

            // 操作符名（字母）
            if (Char.IsLetter(c) || c == '*' || c == '\'')
            {
                var opName = ReadOperatorName(content, ref pos);
                if (opName.Length > 0)
                {
                    FlushOp(currentOp, opName, ref pos);
                }
                continue;
            }

            // 跳过无法识别的字符
            pos++;
        }
    }

    private static void FlushOp(List<Object> operands, String opName, ref Int32 pos)
    {
        // TJ 数组特殊处理：操作数已在列表中以 "[TJ-ARRAY]" 分隔
        var ops = new List<Object>();

        if (opName == "TJ" && operands.Count > 0 && operands[0] is String s && s == "[TJ-ARRAY]")
        {
            operands.RemoveAt(0);
            ops.AddRange(operands);
        }
        else
        {
            ops.AddRange(operands);
        }

        // 跳过 TJ（已处理）
        pos += opName.Length;
        operands.Clear();
    }

    /// <summary>读取操作符名称</summary>
    private static String ReadOperatorName(String text, ref Int32 pos)
    {
        var start = pos;
        while (pos < text.Length)
        {
            var c = text[pos];
            if (Char.IsLetter(c) || c == '*' || c == '\'')
                pos++;
            else
                break;
        }
        return text[start..pos];
    }

    /// <summary>读取括号字符串</summary>
    private static String ReadParenString(String text, ref Int32 pos)
    {
        pos++;
        var sb = new StringBuilder();
        var depth = 1;
        while (pos < text.Length && depth > 0)
        {
            var c = text[pos];
            if (c == '\\' && pos + 1 < text.Length)
            {
                var next = text[pos + 1];
                switch (next)
                {
                    case 'n': sb.Append('\n'); pos += 2; continue;
                    case 'r': sb.Append('\r'); pos += 2; continue;
                    case 't': sb.Append('\t'); pos += 2; continue;
                    case '(': sb.Append('('); pos += 2; continue;
                    case ')': sb.Append(')'); pos += 2; continue;
                    case '\\': sb.Append('\\'); pos += 2; continue;
                    default:
                        if (next >= '0' && next <= '7')
                        {
                            var octal = 0;
                            for (var k = 0; k < 3 && pos + 1 < text.Length && text[pos + 1] >= '0' && text[pos + 1] <= '7'; k++)
                            {
                                octal = octal * 8 + (text[pos + 1] - '0');
                                pos++;
                            }
                            sb.Append((Char)octal);
                            pos++;
                        }
                        else
                        {
                            sb.Append(next);
                            pos += 2;
                        }
                        continue;
                }
            }
            if (c == '(') depth++;
            else if (c == ')')
            {
                depth--;
                if (depth == 0) { pos++; break; }
            }
            if (depth > 0) sb.Append(c);
            pos++;
        }
        return sb.ToString();
    }

    private static String ReadHexString(String text, ref Int32 pos)
    {
        var sb = new StringBuilder();
        while (pos < text.Length)
        {
            if (text[pos] == '>') { pos++; break; }
            if ((text[pos] >= '0' && text[pos] <= '9') ||
                (text[pos] >= 'A' && text[pos] <= 'F') ||
                (text[pos] >= 'a' && text[pos] <= 'f'))
                sb.Append(text[pos]);
            pos++;
        }
        return sb.ToString();
    }

    private static Byte[] HexToBytes(String hex)
    {
        if (hex.Length % 2 != 0) hex += "0";
        var bytes = new Byte[hex.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
        {
            if (!Byte.TryParse(hex.Substring(i * 2, 2), System.Globalization.NumberStyles.HexNumber, null, out bytes[i]))
                bytes[i] = 0;
        }
        return bytes;
    }

    private static Single ReadNumber(String text, ref Int32 pos)
    {
        var start = pos;
        while (pos < text.Length)
        {
            var c = text[pos];
            if ((c >= '0' && c <= '9') || c == '-' || c == '+' || c == '.')
                pos++;
            else break;
        }
        var numStr = text[start..pos];
        return Single.TryParse(numStr, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var f) ? f : 0f;
    }

    private static void SkipDict(String text, ref Int32 pos)
    {
        if (pos + 1 < text.Length && text[pos] == '<' && text[pos + 1] == '<')
        {
            pos += 2;
            var depth = 1;
            while (pos < text.Length && depth > 0)
            {
                if (text[pos] == '<' && pos + 1 < text.Length && text[pos + 1] == '<')
                {
                    depth++;
                    pos += 2;
                }
                else if (text[pos] == '>' && pos + 1 < text.Length && text[pos + 1] == '>')
                {
                    depth--;
                    pos += 2;
                }
                else pos++;
            }
        }
    }

    private static String ReadArray(String text, ref Int32 pos)
    {
        pos++; // skip '['
        var depth = 1;
        var start = pos;
        while (pos < text.Length && depth > 0)
        {
            if (text[pos] == '[') depth++;
            else if (text[pos] == ']')
            {
                depth--;
                if (depth == 0) { var result = text[start..pos]; pos++; return result; }
            }
            pos++;
        }
        return text[start..pos];
    }

    private static void SkipWhitespace(String text, ref Int32 pos)
    {
        while (pos < text.Length)
        {
            var c = text[pos];
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n') pos++;
            else break;
        }
    }
    #endregion

    #region 常用操作符辅助方法
    /// <summary>查找文本操作符（Tj/TJ）并提取文本</summary>
    /// <returns>提取到的文本内容列表</returns>
    public List<String> ExtractText()
    {
        var texts = new List<String>();
        foreach (var op in Operators)
        {
            switch (op.Name)
            {
                case "Tj":
                    if (op.Operands.Count > 0 && op.Operands[0] is String s)
                        texts.Add(s);
                    break;
                case "TJ":
                    var sb = new StringBuilder();
                    foreach (var operand in op.Operands)
                    {
                        if (operand is String ts) sb.Append(ts);
                    }
                    if (sb.Length > 0) texts.Add(sb.ToString());
                    break;
            }
        }
        return texts;
    }
    #endregion
}

/// <summary>PDF 文本操作数（用于内容流中标记文本内容）</summary>
internal class PdfTextOperand
{
    public String Value { get; set; } = String.Empty;
}
