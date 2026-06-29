using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 对象解析器，将原始 PDF 语法解析为强类型中间表示</summary>
/// <remarks>
/// 基于 xref 表提供的精确字节偏移，解析间接对象、字典、数组等 PDF 基本数据类型。
/// 支持嵌套结构和间接引用（`N G R`）。
/// 返回类型安全的 <see cref="PdfObject"/> / <see cref="PdfDict"/> / <see cref="PdfArray"/> 等。
/// </remarks>
public static class PdfObjectParser
{
    #region 公共方法
    /// <summary>解析指定对象号的间接对象（从字节数组 + xref 表）</summary>
    /// <param name="data">PDF 完整字节</param>
    /// <param name="xref">交叉引用表</param>
    /// <param name="objNum">对象号</param>
    /// <returns>解析后的对象，失败返回 null</returns>
    public static PdfObject? ReadObject(Byte[] data, PdfXRefTable xref, Int32 objNum)
    {
        // 检查是否为压缩对象（在对象流中）
        var streamObjNum = xref.GetStreamObjNum(objNum);
        if (streamObjNum > 0)
            return ReadObjectFromStream(data, xref, objNum, streamObjNum);

        // 直接偏移定位
        var offset = xref.GetOffset(objNum);
        if (offset < 0 || offset >= data.Length) return null;

        var latin1 = Encoding.GetEncoding(28591);
        var text = latin1.GetString(data);

        var pos = (Int32)offset;

        // 跳过 "N G obj\n" 头部
        var headerEnd = SkipObjHeader(text, ref pos);
        if (headerEnd < 0) return null;

        return ParseValue(data, latin1, text, ref pos, xref);
    }

    /// <summary>从对象流中读取压缩对象</summary>
    /// <param name="data">PDF 完整字节</param>
    /// <param name="xref">交叉引用表</param>
    /// <param name="objNum">目标对象号</param>
    /// <param name="streamObjNum">所在对象流号</param>
    /// <returns>解析后的对象</returns>
    public static PdfObject? ReadObjectFromStream(Byte[] data, PdfXRefTable xref, Int32 objNum, Int32 streamObjNum)
    {
        // 读取对象流本身
        var streamObj = ReadObject(data, xref, streamObjNum);
        if (streamObj is not PdfStream pdfStream) return null;

        // 对象流的第一部分：NN pairs of (objNum, offset)
        // 格式：N 个整数对，每对为 (object number, byte offset in the decompressed stream)

        // 读取 /N 和 /First 参数
        var dict = pdfStream.Dict;
        if (!dict.TryGetValue("N", out var nVal) || nVal is not PdfNumber nNum) return null;
        if (!dict.TryGetValue("First", out var firstVal) || firstVal is not PdfNumber firstNum) return null;

        var n = (Int32)nNum.Value;
        var first = (Int32)firstNum.Value;
        var streamData = pdfStream.Data;

        // 读取对象索引（每个条目 = 对象号 + 流内偏移，各占可数字段）
        var indexData = Encoding.ASCII.GetString(streamData, 0, Math.Min(first, streamData.Length));
        var tokens = Tokenize(indexData);
        var tokenIdx = 0;

        var targetOffset = -1;
        for (var i = 0; i < n && tokenIdx + 1 < tokens.Count; i++)
        {
            if (!Int32.TryParse(tokens[tokenIdx++], out var curObjNum)) break;
            if (!Int32.TryParse(tokens[tokenIdx++], out var curOffset)) break;
            if (curObjNum == objNum)
            {
                targetOffset = curOffset;
                break;
            }
        }

        if (targetOffset < 0 || first + targetOffset >= streamData.Length) return null;

        // 在偏移处解析对象值
        var valueText = Encoding.ASCII.GetString(streamData, first + targetOffset, streamData.Length - first - targetOffset);
        var valTokens = Tokenize(valueText);
        if (valTokens.Count == 0) return null;

        var latin1 = Encoding.GetEncoding(28591);
        var parseText = latin1.GetString(data);
        // 使用简化的解析（对象流内的值不含间接对象头部）
        var valText2 = Encoding.ASCII.GetString(streamData, first + targetOffset, streamData.Length - first - targetOffset);
        return ParseTopLevelValue(valText2, 0, out _);
    }

    /// <summary>从 PDF 字典字符串解析为 PdfDict</summary>
    /// <param name="dictStr">字典字符串（不含外层 &lt;&lt; &gt;&gt;）</param>
    /// <returns>解析后的字典</returns>
    public static PdfDict ParseDictString(String dictStr)
    {
        var dict = new PdfDict();
        var pos = 0;

        // 跳过开头的 <<
        if (pos + 1 < dictStr.Length && dictStr[pos] == '<' && dictStr[pos + 1] == '<')
            pos += 2;

        while (pos < dictStr.Length)
        {
            SkipWhitespace(dictStr, ref pos);
            if (pos >= dictStr.Length || (dictStr[pos] == '>' && pos + 1 < dictStr.Length && dictStr[pos + 1] == '>'))
                break;

            // 读取键（Name: /KeyName）
            if (dictStr[pos] != '/') { pos++; continue; }
            pos++; // 跳过 /
            var key = ReadName(dictStr, ref pos);
            if (key.Length == 0) continue;

            SkipWhitespace(dictStr, ref pos);

            // 读取值
            var value = ParseDictValue(dictStr, ref pos);
            if (value != null)
                dict[key] = value;
        }

        return dict;
    }
    #endregion

    #region 内部解析方法
    /// <summary>解析顶层值（整数/实数/名称/字符串/数组/字典/布尔/null/间接引用）</summary>
    internal static PdfObject? ParseTopLevelValue(String text, Int32 startPos, out Int32 endPos)
    {
        var pos = startPos;
        SkipWhitespace(text, ref pos);

        if (pos >= text.Length) { endPos = pos; return null; }

        var c = text[pos];

        // 数字（整数或实数）
        if (c == '-' || c == '+' || (c >= '0' && c <= '9') || c == '.')
        {
            endPos = pos;
            return ReadNumber(text, ref pos);
        }

        // 名称 /Name
        if (c == '/')
        {
            pos++;
            var name = ReadName(text, ref pos);
            endPos = pos;
            return new PdfName(name);
        }

        // 字典 << ... >>
        if (c == '<' && pos + 1 < text.Length && text[pos + 1] == '<')
        {
            pos += 2;
            var dict = new PdfDict();
            while (pos < text.Length)
            {
                SkipWhitespace(text, ref pos);
                if (pos + 1 < text.Length && text[pos] == '>' && text[pos + 1] == '>')
                {
                    pos += 2;
                    break;
                }
                if (pos >= text.Length || text[pos] != '/') { pos++; continue; }
                pos++;
                var key = ReadName(text, ref pos);
                SkipWhitespace(text, ref pos);
                var val = ParseDictValue(text, ref pos);
                if (val != null) dict[key] = val;
            }
            endPos = pos;
            return new PdfDictObj(dict);
        }

        // 十六进制字符串 <hex>
        if (c == '<' && (pos + 1 >= text.Length || text[pos + 1] != '<'))
        {
            pos++;
            var hex = ReadHexString(text, ref pos);
            endPos = pos;
            return new PdfHexString(hex);
        }

        // 数组 [ ... ]
        if (c == '[')
        {
            pos++;
            var arr = new List<PdfObject>();
            while (pos < text.Length)
            {
                SkipWhitespace(text, ref pos);
                if (pos < text.Length && text[pos] == ']') { pos++; break; }
                var item = ParseDictValue(text, ref pos);
                if (item != null) arr.Add(item);
                else if (pos >= text.Length) break;
            }
            endPos = pos;
            return new PdfArray { Items = arr };
        }

        // 括号字符串 (text)
        if (c == '(')
        {
            var str = ReadParenString(text, ref pos);
            endPos = pos;
            return new PdfString(str);
        }

        // 布尔或 null 关键字
        if (Char.IsLetter(c))
        {
            var word = ReadWord(text, ref pos);
            endPos = pos;
            return word.ToLowerInvariant() switch
            {
                "true" => new PdfBoolean(true),
                "false" => new PdfBoolean(false),
                "null" => new PdfNull(),
                _ => null,
            };
        }

        endPos = pos + 1;
        return null;
    }

    /// <summary>解析字典值（支持数字、名称、字符串、数组、字典、布尔、null、间接引用）</summary>
    internal static PdfObject? ParseDictValue(String text, ref Int32 pos)
    {
        SkipWhitespace(text, ref pos);
        if (pos >= text.Length) return null;

        var c = text[pos];

        // 数字（需要检查是否为间接引用 N G R）
        if (c == '-' || c == '+' || (c >= '0' && c <= '9') || c == '.')
        {
            var num = ReadNumber(text, ref pos);
            if (num is not PdfNumber numObj) return num;

            // 检查后续是否为 "G R"（间接引用）
            var savedPos = pos;
            SkipWhitespace(text, ref pos);
            if (pos < text.Length && (text[pos] >= '0' && text[pos] <= '9'))
            {
                var genNum = ReadNumber(text, ref pos);
                SkipWhitespace(text, ref pos);
                if (pos < text.Length && text[pos] == 'R')
                {
                    pos++; // skip R
                    if (genNum is PdfNumber gn)
                        return new PdfRef((Int32)numObj.Value, (Int32)gn.Value);
                }
            }
            pos = savedPos; // 回退，作为普通数字
            return numObj;
        }

        // 名称
        if (c == '/')
        {
            pos++;
            return new PdfName(ReadName(text, ref pos));
        }

        // 字典
        if (c == '<' && pos + 1 < text.Length && text[pos + 1] == '<')
        {
            pos += 2;
            var dict = new PdfDict();
            while (pos < text.Length)
            {
                SkipWhitespace(text, ref pos);
                if (pos + 1 < text.Length && text[pos] == '>' && text[pos + 1] == '>')
                {
                    pos += 2;
                    break;
                }
                if (pos >= text.Length || text[pos] != '/') { pos++; continue; }
                pos++;
                var key = ReadName(text, ref pos);
                SkipWhitespace(text, ref pos);
                var val = ParseDictValue(text, ref pos);
                if (val != null) dict[key] = val;
            }
            return new PdfDictObj(dict);
        }

        // 十六进制字符串
        if (c == '<' && (pos + 1 >= text.Length || text[pos + 1] != '<'))
        {
            pos++;
            return new PdfHexString(ReadHexString(text, ref pos));
        }

        // 数组
        if (c == '[')
        {
            pos++;
            var arr = new List<PdfObject>();
            while (pos < text.Length)
            {
                SkipWhitespace(text, ref pos);
                if (pos < text.Length && text[pos] == ']') { pos++; break; }
                var item = ParseDictValue(text, ref pos);
                if (item != null) arr.Add(item);
                else if (pos >= text.Length) break;
            }
            return new PdfArray { Items = arr };
        }

        // 括号字符串
        if (c == '(')
            return new PdfString(ReadParenString(text, ref pos));

        // 关键字
        if (Char.IsLetter(c))
        {
            var word = ReadWord(text, ref pos);
            return word.ToLowerInvariant() switch
            {
                "true" => new PdfBoolean(true),
                "false" => new PdfBoolean(false),
                "null" => new PdfNull(),
                _ => null,
            };
        }

        pos++;
        return null;
    }

    /// <summary>解析值（含间接引用处理）</summary>
    internal static PdfObject? ParseValue(Byte[] data, Encoding latin1, String text, ref Int32 pos, PdfXRefTable? xref)
    {
        SkipWhitespace(text, ref pos);
        if (pos >= text.Length) return null;

        var baseValue = ParseDictValue(text, ref pos);
        if (baseValue == null) return null;

        // 检查是否为间接引用（N G R）
        if (baseValue is PdfNumber numObj)
        {
            var savedPos = pos;
            SkipWhitespace(text, ref pos);
            if (pos < text.Length && (text[pos] >= '0' && text[pos] <= '9'))
            {
                var genNum = ReadNumber(text, ref pos);
                SkipWhitespace(text, ref pos);
                if (pos < text.Length && text[pos] == 'R')
                {
                    pos++;
                    if (genNum is PdfNumber gn && xref != null)
                    {
                        var refObjNum = (Int32)numObj.Value;
                        // 解析间接引用指向的对象
                        var resolved = ReadObject(data, xref, refObjNum);
                        if (resolved != null) return resolved;
                        // 如果对象是字典/流等，可能需要返回引用包装
                        // 这里简单返回解析结果
                        return resolved;
                    }
                    return new PdfRef((Int32)numObj.Value, genNum is PdfNumber gn2 ? (Int32)gn2.Value : 0);
                }
            }
            pos = savedPos;
        }

        return baseValue;
    }
    #endregion

    #region 辅助解析方法
    private static Int32 SkipObjHeader(String text, ref Int32 pos)
    {
        SkipWhitespace(text, ref pos);
        // 确认是 "NNN GGG obj"
        var objIdx = text.IndexOf("obj", pos, StringComparison.Ordinal);
        if (objIdx < 0 || objIdx - pos > 30) return -1;
        pos = objIdx + 3;
        SkipWhitespace(text, ref pos);

        // 检查是否后面跟着 stream 字典
        var nextSection = text.IndexOf("<<", pos, Math.Min(100, text.Length - pos), StringComparison.Ordinal);

        return pos; // 返回内容开始位置
    }

    internal static void SkipWhitespace(String text, ref Int32 pos)
    {
        while (pos < text.Length)
        {
            var c = text[pos];
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
                pos++;
            else if (c == '%')
            {
                // PDF 注释
                while (pos < text.Length && text[pos] != '\r' && text[pos] != '\n') pos++;
            }
            else break;
        }
    }

    private static String ReadName(String text, ref Int32 pos)
    {
        var sb = new StringBuilder();
        while (pos < text.Length)
        {
            var c = text[pos];
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n' ||
                c == '/' || c == '<' || c == '>' || c == '[' || c == ']' ||
                c == '(' || c == ')' || c == '{' || c == '}')
                break;
            if (c == '#')
            {
                // PDF 名称转义 #XX
                if (pos + 2 < text.Length &&
                    Byte.TryParse(text.Substring(pos + 1, 2), NumberStyles.HexNumber, null, out var b))
                {
                    sb.Append((Char)b);
                    pos += 3;
                    continue;
                }
            }
            sb.Append(c);
            pos++;
        }
        return sb.ToString();
    }

    private static String ReadWord(String text, ref Int32 pos)
    {
        var sb = new StringBuilder();
        while (pos < text.Length && Char.IsLetterOrDigit(text[pos]))
        {
            sb.Append(text[pos]);
            pos++;
        }
        return sb.ToString();
    }

    private static PdfObject? ReadNumber(String text, ref Int32 pos)
    {
        var sb = new StringBuilder();
        var hasDot = false;
        while (pos < text.Length)
        {
            var c = text[pos];
            if ((c >= '0' && c <= '9') || c == '-' || c == '+' || c == '.')
            {
                if (c == '.') hasDot = true;
                sb.Append(c);
                pos++;
            }
            else break;
        }

        if (sb.Length == 0 || sb.ToString() == "-") return null;

        if (hasDot)
        {
            return Single.TryParse(sb.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var f)
                ? new PdfNumber(f) : null;
        }
        return Int32.TryParse(sb.ToString(), out var i)
            ? new PdfNumber(i)
            : Single.TryParse(sb.ToString(), out var f2) ? new PdfNumber(f2) : null;
    }

    internal static String ReadParenString(String text, ref Int32 pos)
    {
        pos++; // skip '('
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
                    case 'b': sb.Append('\b'); pos += 2; continue;
                    case 'f': sb.Append('\f'); pos += 2; continue;
                    case '(': sb.Append('('); pos += 2; continue;
                    case ')': sb.Append(')'); pos += 2; continue;
                    case '\\': sb.Append('\\'); pos += 2; continue;
                    default:
                        if (next >= '0' && next <= '7')
                        {
                            // 八进制转义 \ddd
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
            else if (c == ')') { depth--; if (depth == 0) { pos++; break; } }
            if (depth > 0) sb.Append(c);
            pos++;
        }
        return sb.ToString();
    }

    private static String ReadHexString(String text, ref Int32 pos)
    {
        var hex = new StringBuilder();
        while (pos < text.Length)
        {
            var c = text[pos];
            if (c == '>') { pos++; break; }
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n') { pos++; continue; }
            if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))
                hex.Append(c);
            pos++;
        }
        return hex.ToString();
    }

    internal static List<String> Tokenize(String text)
    {
        var tokens = new List<String>();
        var pos = 0;
        while (pos < text.Length)
        {
            SkipWhitespace(text, ref pos);
            if (pos >= text.Length) break;

            var c = text[pos];
            if (c == '(')
            {
                // 跳过括号字符串
                var depth = 1;
                var start = pos;
                pos++;
                while (pos < text.Length && depth > 0)
                {
                    if (text[pos] == '\\') pos++;
                    else if (text[pos] == '(') depth++;
                    else if (text[pos] == ')') depth--;
                    pos++;
                }
                tokens.Add(text[start..pos]);
            }
            else if (c == '<')
            {
                tokens.Add("<");
                pos++;
            }
            else if (c == '>')
            {
                tokens.Add(">");
                pos++;
            }
            else if (c == '[' || c == ']')
            {
                tokens.Add(c.ToString());
                pos++;
            }
            else if (c == '/')
            {
                pos++;
                tokens.Add("/" + ReadName(text, ref pos));
            }
            else
            {
                var start = pos;
                while (pos < text.Length && text[pos] != ' ' && text[pos] != '\t' &&
                       text[pos] != '\r' && text[pos] != '\n' &&
                       text[pos] != '/' && text[pos] != '<' && text[pos] != '>' &&
                       text[pos] != '[' && text[pos] != ']' && text[pos] != '(' && text[pos] != ')')
                    pos++;
                if (pos > start)
                    tokens.Add(text[start..pos]);
            }
        }
        return tokens;
    }
    #endregion
}

#region PDF 对象中间表示类型
/// <summary>PDF 对象基类型</summary>
public abstract class PdfObject { }

/// <summary>PDF 数值（整数或实数）</summary>
public class PdfNumber : PdfObject
{
    /// <summary>数值</summary>
    public Single Value { get; }
    /// <summary>实例化数值对象</summary>
    public PdfNumber(Single value) => Value = value;
    /// <summary>实例化数值对象</summary>
    public PdfNumber(Int32 value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => Value.ToString(CultureInfo.InvariantCulture);
}

/// <summary>PDF 名称对象（以 / 开头的标识符）</summary>
public class PdfName : PdfObject
{
    /// <summary>名称值（不含前导 /）</summary>
    public String Value { get; }
    /// <summary>实例化名称对象</summary>
    public PdfName(String value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => $"/{Value}";
}

/// <summary>PDF 字符串对象（括号字符串）</summary>
public class PdfString : PdfObject
{
    /// <summary>字符串值</summary>
    public String Value { get; }
    /// <summary>实例化字符串对象</summary>
    public PdfString(String value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => $"({Value})";
}

/// <summary>PDF 十六进制字符串对象</summary>
public class PdfHexString : PdfObject
{
    /// <summary>十六进制字符串值</summary>
    public String Value { get; }
    /// <summary>实例化十六进制字符串对象</summary>
    public PdfHexString(String value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => $"<{Value}>";
}

/// <summary>PDF 布尔对象</summary>
public class PdfBoolean : PdfObject
{
    /// <summary>布尔值</summary>
    public Boolean Value { get; }
    /// <summary>实例化布尔对象</summary>
    public PdfBoolean(Boolean value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => Value ? "true" : "false";
}

/// <summary>PDF null 对象</summary>
public class PdfNull : PdfObject
{
    /// <inheritdoc/>
    public override String ToString() => "null";
}

/// <summary>PDF 字典对象（包装 PdfDict）</summary>
public class PdfDictObj : PdfObject
{
    /// <summary>字典值</summary>
    public PdfDict Value { get; }
    /// <summary>实例化字典对象</summary>
    public PdfDictObj(PdfDict value) => Value = value;
    /// <inheritdoc/>
    public override String ToString() => Value.ToString() ?? "<<>>";
}

/// <summary>PDF 数组对象</summary>
public class PdfArray : PdfObject
{
    /// <summary>数组项</summary>
    public List<PdfObject> Items { get; set; } = [];
    /// <inheritdoc/>
    public override String ToString() => $"[{String.Join(" ", Items)}]";
}

/// <summary>PDF 间接引用对象（N G R）</summary>
public class PdfRef : PdfObject
{
    /// <summary>引用的对象号</summary>
    public Int32 ObjNum { get; }
    /// <summary>代数号</summary>
    public Int32 GenNum { get; }
    /// <summary>实例化间接引用</summary>
    public PdfRef(Int32 objNum, Int32 genNum) { ObjNum = objNum; GenNum = genNum; }
    /// <inheritdoc/>
    public override String ToString() => $"{ObjNum} {GenNum} R";
}

/// <summary>PDF 流对象（字典 + 二进制数据）</summary>
public class PdfStream : PdfObject
{
    /// <summary>流字典</summary>
    public PdfDict Dict { get; set; } = new();
    /// <summary>流数据字节</summary>
    public Byte[] Data { get; set; } = [];
    /// <inheritdoc/>
    public override String ToString() => $"stream({Data.Length} bytes)";
}

/// <summary>PDF 字典（键值对集合）</summary>
public class PdfDict : Dictionary<String, PdfObject>
{
    /// <inheritdoc/>
    public override String ToString()
    {
        var items = this.Select(kv => $"/{kv.Key} {kv.Value}").ToList();
        return $"<< {String.Join(" ", items)} >>";
    }
}
#endregion
