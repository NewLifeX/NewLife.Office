using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 读取器（基础实现）</summary>
/// <remarks>
/// 直接解析 PDF 字节流，提取文本内容和元数据。
/// 支持 PDF 1.0-1.7，基于对象流扫描方式提取文本（不依赖外部库）。
/// 对加密 PDF 或内嵌 CJK 字体 PDF 的文本提取效果有限。
/// </remarks>
public class PdfReader : IDisposable
{
    #region 属性
    /// <summary>源文件路径</summary>
    public String? FilePath { get; private set; }
    #endregion

    #region 私有字段
    private readonly Byte[] _data;
    #endregion

    #region 构造
    /// <summary>从文件路径打开</summary>
    /// <param name="path">PDF 文件路径</param>
    public PdfReader(String path)
    {
        FilePath = path.GetFullPath();
        _data = File.ReadAllBytes(FilePath);
    }

    /// <summary>从流打开</summary>
    /// <param name="stream">包含 PDF 内容的流</param>
    public PdfReader(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        _data = ms.ToArray();
    }

    /// <summary>释放资源</summary>
    public void Dispose() => GC.SuppressFinalize(this);
    #endregion

    #region 读取方法
    /// <summary>获取总页数（通过 /Count 字段）</summary>
    /// <returns>页数</returns>
    public Int32 GetPageCount()
    {
        var latin1 = Encoding.GetEncoding(1252);
        var pdf = latin1.GetString(_data);
        // 在 Pages 字典中查找 /Count 值
        var countIdx = FindToken(pdf, "/Count");
        if (countIdx < 0) return 0;
        var numStr = ExtractNextToken(pdf, countIdx + 6);
        return Int32.TryParse(numStr.Trim(), out var count) ? count : 0;
    }

    /// <summary>提取全部文本（从所有内容流中）</summary>
    /// <returns>合并后的文本</returns>
    public String ExtractText()
    {
        var sb = new StringBuilder();
        ExtractFromStreams(_data, sb);
        return sb.ToString();
    }

    /// <summary>读取文档元数据</summary>
    /// <returns>元数据对象</returns>
    public PdfMetadata ReadMetadata()
    {
        var meta = new PdfMetadata { PageCount = GetPageCount() };
        var latin1 = Encoding.GetEncoding(1252);
        var pdf = latin1.GetString(_data);

        // 读取 %PDF-x.x 版本
        if (pdf.StartsWith("%PDF-"))
            meta.PdfVersion = pdf.Substring(5, Math.Min(3, pdf.Length - 5));

        // 读取 Info 字典
        var infoStart = FindToken(pdf, "/Info");
        if (infoStart >= 0)
        {
            var dictText = ExtractDict(pdf, infoStart);
            meta.Title = GetDictValue(dictText, "Title");
            meta.Author = GetDictValue(dictText, "Author");
            meta.Subject = GetDictValue(dictText, "Subject");
            meta.CreationDate = GetDictValue(dictText, "CreationDate");
        }

        return meta;
    }

    /// <summary>提取带坐标位置的文本（P06-05）</summary>
    /// <remarks>
    /// 解析内容流中的文本定位/显示操作符，返回各文本段落及其近似坐标。
    /// 坐标系以页面左下角为原点，单位为 PDF 用户空间单位（通常约等于磅/pt）。
    /// 注意：对加密或使用自定义字体映射的 PDF，文本内容可能不准确。
    /// </remarks>
    /// <returns>文本项序列，每项含文本内容和近似 (X, Y) 坐标</returns>
    public IEnumerable<PdfTextItem> ExtractTextWithPositions()
    {
        var latin1 = Encoding.GetEncoding(1252);
        var pdf = latin1.GetString(_data);
        var results = new List<PdfTextItem>();
        var pos = 0;
        while (pos < pdf.Length)
        {
            var streamStart = pdf.IndexOf("stream", pos, StringComparison.Ordinal);
            if (streamStart < 0) break;
            var contentStart = streamStart + 6;
            if (contentStart < pdf.Length && pdf[contentStart] == '\r') contentStart++;
            if (contentStart < pdf.Length && pdf[contentStart] == '\n') contentStart++;
            var streamEnd = pdf.IndexOf("endstream", contentStart, StringComparison.Ordinal);
            if (streamEnd < 0) break;
            var content = pdf.Substring(contentStart, streamEnd - contentStart);
            ExtractPositionedText(content, results);
            pos = streamEnd + 9;
        }
        return results;
    }

    /// <summary>从 PDF 中提取嵌入图片（P06-04）</summary>
    /// <remarks>
    /// 扫描所有流字典，找到类型为 /Subtype /Image 的 XObject 流并提取原始字节。
    /// 对 /Filter /DCTDecode（JPEG）图片返回直接可用的 JPEG 字节。
    /// 其他编码格式返回原始压缩字节，可结合 <see cref="PdfImageStream.Filter"/> 判断。
    /// </remarks>
    /// <returns>图片流对象序列</returns>
    public IEnumerable<PdfImageStream> ExtractImageStreams()
    {
        var latin1 = Encoding.GetEncoding(1252);
        var text = latin1.GetString(_data);
        var pos = 0;
        var imgIdx = 0;
        while (pos < text.Length)
        {
            // 寻找包含 /Subtype /Image 的字典
            var dictStart = text.IndexOf("<<", pos, StringComparison.Ordinal);
            if (dictStart < 0) break;
            var dictEnd = text.IndexOf(">>", dictStart + 2, StringComparison.Ordinal);
            if (dictEnd < 0) break;
            var dict = text.Substring(dictStart, dictEnd - dictStart + 2);

            if (dict.IndexOf("/Subtype", StringComparison.Ordinal) >= 0 && dict.IndexOf("/Image", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // 解析宽高
                var widthTok = GetDictIntValue(dict, "Width");
                var heightTok = GetDictIntValue(dict, "Height");
                var filter = GetDictToken(dict, "Filter");
                var lengthTok = GetDictIntValue(dict, "Length");

                // 找到紧随 >> 之后的 stream
                var strmPos = text.IndexOf("stream", dictEnd, StringComparison.Ordinal);
                if (strmPos >= 0 && strmPos < dictEnd + 100)
                {
                    var dataStart = strmPos + 6;
                    if (dataStart < text.Length && text[dataStart] == '\r') dataStart++;
                    if (dataStart < text.Length && text[dataStart] == '\n') dataStart++;
                    var dataEnd = text.IndexOf("endstream", dataStart, StringComparison.Ordinal);
                    if (dataEnd > dataStart && widthTok > 0 && heightTok > 0)
                    {
                        var rawBytes = _data.AsSpan(dataStart, Math.Min(dataEnd - dataStart, _data.Length - dataStart)).ToArray();
                        yield return new PdfImageStream
                        {
                            Index = imgIdx++,
                            Width = widthTok,
                            Height = heightTok,
                            Filter = filter ?? String.Empty,
                            RawData = rawBytes,
                        };
                        pos = dataEnd + 9;
                        continue;
                    }
                }
            }
            pos = dictEnd + 2;
        }
    }
    #endregion

    #region 私有方法
    /// <summary>从 PDF 内容流中提取文本（解析 Tj/TJ 操作符）</summary>
    private static void ExtractFromStreams(Byte[] pdfData, StringBuilder sb)
    {
        // 扫描所有 stream...endstream 块
        var pdf = Encoding.GetEncoding(1252).GetString(pdfData);
        var pos = 0;
        while (pos < pdf.Length)
        {
            var streamStart = pdf.IndexOf("stream", pos, StringComparison.Ordinal);
            if (streamStart < 0) break;

            // 跳过 "stream\r\n" 或 "stream\n"
            var contentStart = streamStart + 6;
            if (contentStart < pdf.Length && pdf[contentStart] == '\r') contentStart++;
            if (contentStart < pdf.Length && pdf[contentStart] == '\n') contentStart++;

            var streamEnd = pdf.IndexOf("endstream", contentStart, StringComparison.Ordinal);
            if (streamEnd < 0) break;

            var streamContent = pdf.Substring(contentStart, streamEnd - contentStart);
            ExtractTextFromContent(streamContent, sb);
            pos = streamEnd + 9;
        }
    }

    /// <summary>从 PDF 内容流字符串中提取文本操作符</summary>
    private static void ExtractTextFromContent(String content, StringBuilder sb)
    {
        // 解析 (text) Tj 和 [(text)] TJ 操作符
        var i = 0;
        while (i < content.Length)
        {
            if (content[i] == '(')
            {
                // 读取括号字符串
                var str = ReadParenString(content, ref i);
                // 查找后续操作符
                var opPos = i;
                SkipWhitespace(content, ref opPos);
                if (opPos < content.Length - 1)
                {
                    var op = content.Substring(opPos, 2);
                    if (op.StartsWith("Tj") || op.StartsWith("TJ") || op.StartsWith("'") || op.StartsWith("\""))
                    {
                        sb.Append(DecodePdfString(str));
                        i = opPos + (op.StartsWith("Tj") || op.StartsWith("TJ") ? 2 : 1);
                        continue;
                    }
                }
            }
            else if (content[i] == '[')
            {
                // TJ array
                var arrEnd = content.IndexOf(']', i);
                if (arrEnd > i)
                {
                    var arr = content.Substring(i + 1, arrEnd - i - 1);
                    ExtractTextFromContent(arr, sb);
                    i = arrEnd + 1;
                    // skip TJ
                    SkipWhitespace(content, ref i);
                    if (i < content.Length - 1 && content.Substring(i, 2) == "TJ")
                        i += 2;
                    continue;
                }
            }
            else if (content[i] == 'T' && i + 1 < content.Length && content[i + 1] == '*')
            {
                sb.AppendLine();
                i += 2;
                continue;
            }
            else if (content[i] == 'B' && i + 1 < content.Length && content[i + 1] == 'T')
            {
                i += 2;
                continue;
            }
            else if (content[i] == 'E' && i + 3 < content.Length && content.Substring(i, 2) == "ET")
            {
                sb.AppendLine();
                i += 2;
                continue;
            }
            i++;
        }
    }

    private static String ReadParenString(String s, ref Int32 pos)
    {
        pos++; // skip '('
        var sb = new StringBuilder();
        var depth = 1;
        while (pos < s.Length && depth > 0)
        {
            var c = s[pos];
            if (c == '\\' && pos + 1 < s.Length)
            {
                sb.Append(s[pos + 1]);
                pos += 2;
                continue;
            }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) { pos++; break; } }
            if (depth > 0) sb.Append(c);
            pos++;
        }
        return sb.ToString();
    }

    private static void SkipWhitespace(String s, ref Int32 pos)
    {
        while (pos < s.Length && (s[pos] == ' ' || s[pos] == '\t' || s[pos] == '\r' || s[pos] == '\n'))
            pos++;
    }

    private static String DecodePdfString(String s)
    {
        // Basic: remove non-printable control chars, keep Latin-1 printables
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (c >= 32 && c < 256) sb.Append(c);
            else if (c == '\n' || c == '\r') sb.Append(' ');
        }
        return sb.ToString();
    }

    private static Int32 FindToken(String pdf, String token)
    {
        var idx = pdf.IndexOf(token, StringComparison.Ordinal);
        return idx;
    }

    private static String ExtractNextToken(String pdf, Int32 pos)
    {
        SkipWhitespace(pdf, ref pos);
        var end = pos;
        while (end < pdf.Length && pdf[end] != ' ' && pdf[end] != '\n' && pdf[end] != '\r'
               && pdf[end] != '/' && pdf[end] != '<' && pdf[end] != '>')
            end++;
        return pdf.Substring(pos, end - pos);
    }

    private static String ExtractDict(String pdf, Int32 startOffset)
    {
        // find << ... >>
        var start = pdf.IndexOf("<<", startOffset, StringComparison.Ordinal);
        if (start < 0) return String.Empty;
        var end = pdf.IndexOf(">>", start + 2, StringComparison.Ordinal);
        if (end < 0) return String.Empty;
        return pdf.Substring(start, end - start + 2);
    }

    private static String? GetDictValue(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return null;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        if (valStart >= dict.Length) return null;
        if (dict[valStart] == '(')
        {
            var tmp = valStart;
            return ReadParenString(dict, ref tmp);
        }
        return ExtractNextToken(dict, valStart);
    }

    /// <summary>从 PDF 字典字符串中提取整型值</summary>
    /// <param name="dict">字典字符串</param>
    /// <param name="key">键名（不含前导 /）</param>
    /// <returns>整型值，未找到时返回 0</returns>
    private static Int32 GetDictIntValue(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return 0;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        var str = ExtractNextToken(dict, valStart);
        return Int32.TryParse(str.Trim(), out var v) ? v : 0;
    }

    /// <summary>从 PDF 字典字符串中提取 Name 类型的值（不含前导 /）</summary>
    /// <param name="dict">字典字符串</param>
    /// <param name="key">键名（不含前导 /）</param>
    /// <returns>Name 值字符串，未找到时返回 null</returns>
    private static String? GetDictToken(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return null;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        if (valStart >= dict.Length) return null;
        if (dict[valStart] == '/')
        {
            var nameEnd = valStart + 1;
            while (nameEnd < dict.Length && dict[nameEnd] != ' ' && dict[nameEnd] != '\t'
                   && dict[nameEnd] != '\r' && dict[nameEnd] != '\n'
                   && dict[nameEnd] != '/' && dict[nameEnd] != '<' && dict[nameEnd] != '>')
                nameEnd++;
            return dict.Substring(valStart + 1, nameEnd - valStart - 1);
        }
        return ExtractNextToken(dict, valStart).Trim();
    }

    /// <summary>从 PDF 内容流中提取带位置信息的文本</summary>
    /// <param name="content">内容流字符串</param>
    /// <param name="results">结果列表</param>
    private static void ExtractPositionedText(String content, List<PdfTextItem> results)
    {
        var curX = 0f;
        var curY = 0f;
        var fontSize = 0f;
        var inText = false;
        var i = 0;
        var numStack = new List<Single>();

        while (i < content.Length)
        {
            SkipWhitespace(content, ref i);
            if (i >= content.Length) break;

            var c = content[i];

            // PDF 注释行
            if (c == '%')
            {
                while (i < content.Length && content[i] != '\n') i++;
                continue;
            }

            // 括号字符串 (text)
            if (c == '(')
            {
                var s = ReadParenString(content, ref i);
                if (inText)
                {
                    var peek = i;
                    SkipWhitespace(content, ref peek);
                    if (peek + 1 < content.Length && content[peek] == 'T'
                        && (content[peek + 1] == 'j' || content[peek + 1] == 'J'
                            || content[peek + 1] == '\'' || content[peek + 1] == '"'))
                    {
                        var decoded = DecodePdfString(s);
                        if (decoded.Length > 0)
                            results.Add(new PdfTextItem { Text = decoded, X = curX, Y = curY, FontSize = fontSize });
                        i = peek + 2;
                        numStack.Clear();
                        continue;
                    }
                }
                numStack.Clear();
                continue;
            }

            // 嵌套字典 << >>
            if (c == '<' && i + 1 < content.Length && content[i + 1] == '<')
            {
                var end = content.IndexOf(">>", i + 2, StringComparison.Ordinal);
                i = end >= 0 ? end + 2 : i + 2;
                numStack.Clear();
                continue;
            }

            // 十六进制字符串 <hex>
            if (c == '<')
            {
                var end = content.IndexOf('>', i + 1);
                if (end > i)
                {
                    if (inText)
                    {
                        var hex = content.Substring(i + 1, end - i - 1);
                        var decoded = HexToString(hex);
                        if (decoded.Length > 0)
                        {
                            var peek = end + 1;
                            SkipWhitespace(content, ref peek);
                            if (peek + 1 < content.Length && content[peek] == 'T' && content[peek + 1] == 'j')
                            {
                                results.Add(new PdfTextItem { Text = decoded, X = curX, Y = curY, FontSize = fontSize });
                                i = peek + 2;
                                numStack.Clear();
                                continue;
                            }
                        }
                    }
                    i = end + 1;
                }
                else
                    i++;
                continue;
            }

            // TJ 数组 [...]
            if (c == '[')
            {
                var arrEnd = content.IndexOf(']', i);
                if (arrEnd > i && inText)
                {
                    var arr = content.Substring(i + 1, arrEnd - i - 1);
                    var arrSb = new StringBuilder();
                    var ap = 0;
                    while (ap < arr.Length)
                    {
                        SkipWhitespace(arr, ref ap);
                        if (ap >= arr.Length) break;
                        if (arr[ap] == '(')
                        {
                            var s = ReadParenString(arr, ref ap);
                            arrSb.Append(DecodePdfString(s));
                        }
                        else
                        {
                            // 数字（字间距调整）或其他 — 跳过到下一个空白或 (
                            while (ap < arr.Length && arr[ap] != '(' && arr[ap] != ' '
                                   && arr[ap] != '\t' && arr[ap] != '\r' && arr[ap] != '\n')
                                ap++;
                        }
                    }
                    var txt = arrSb.ToString();
                    if (txt.Length > 0)
                        results.Add(new PdfTextItem { Text = txt, X = curX, Y = curY, FontSize = fontSize });
                    i = arrEnd + 1;
                    SkipWhitespace(content, ref i);
                    if (i + 1 < content.Length && content[i] == 'T' && content[i + 1] == 'J')
                        i += 2;
                }
                else
                    i = arrEnd >= 0 ? arrEnd + 1 : i + 1;
                numStack.Clear();
                continue;
            }

            // PDF Name（/name），用于 Tf 的字体名等
            if (c == '/')
            {
                var nameEnd = i + 1;
                while (nameEnd < content.Length && content[nameEnd] != ' ' && content[nameEnd] != '\t'
                       && content[nameEnd] != '\r' && content[nameEnd] != '\n'
                       && content[nameEnd] != '/' && content[nameEnd] != '<' && content[nameEnd] != '>')
                    nameEnd++;
                i = nameEnd;
                continue;
            }

            // 数字（操作符参数）
            if (Char.IsDigit(c) || c == '-'
                || (c == '.' && i + 1 < content.Length && Char.IsDigit(content[i + 1])))
            {
                var numEnd = i + 1;
                while (numEnd < content.Length && (Char.IsDigit(content[numEnd]) || content[numEnd] == '.'))
                    numEnd++;
                if (Single.TryParse(content.Substring(i, numEnd - i),
                    NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                    numStack.Add(num);
                i = numEnd;
                continue;
            }

            // 操作符
            if (Char.IsLetter(c) || c == '\'' || c == '"' || c == '*')
            {
                var opEnd = i + 1;
                while (opEnd < content.Length
                       && (Char.IsLetterOrDigit(content[opEnd]) || content[opEnd] == '*'))
                    opEnd++;
                var op = content.Substring(i, opEnd - i);
                i = opEnd;

                switch (op)
                {
                    case "BT":
                        inText = true;
                        curX = 0;
                        curY = 0;
                        break;
                    case "ET":
                        inText = false;
                        break;
                    case "Td":
                    case "TD":
                        if (numStack.Count >= 2)
                        {
                            curX += numStack[numStack.Count - 2];
                            curY += numStack[numStack.Count - 1];
                        }
                        break;
                    case "Tm":
                        if (numStack.Count >= 6)
                        {
                            curX = numStack[numStack.Count - 2];
                            curY = numStack[numStack.Count - 1];
                        }
                        break;
                    case "Tf":
                        if (numStack.Count >= 1)
                            fontSize = numStack[numStack.Count - 1];
                        break;
                    case "T*":
                        curY -= fontSize > 0 ? fontSize * 1.2f : 12f;
                        break;
                }
                numStack.Clear();
                continue;
            }

            i++;
        }
    }

    /// <summary>将 PDF 十六进制字符串转换为可读文本</summary>
    /// <param name="hex">十六进制字符串（可含空白符）</param>
    /// <returns>可打印字符序列</returns>
    private static String HexToString(String hex)
    {
        var sb = new StringBuilder();
        var clean = hex.ToCharArray();
        var ci = 0;
        var cleanBuf = new StringBuilder(hex.Length);
        while (ci < clean.Length)
        {
            var ch = clean[ci++];
            if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n') continue;
            cleanBuf.Append(ch);
        }
        var s = cleanBuf.ToString();
        for (var k = 0; k + 1 < s.Length; k += 2)
        {
            if (Byte.TryParse(s.Substring(k, 2), NumberStyles.HexNumber, null, out var b) && b >= 32)
                sb.Append((Char)b);
        }
        return sb.ToString();
    }
    #endregion
}