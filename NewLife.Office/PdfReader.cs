#nullable enable
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 元数据</summary>
public class PdfMetadata
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>创建时间字符串（PDF 格式 D:YYYYMMDDHHmmss）</summary>
    public String? CreationDate { get; set; }

    /// <summary>PDF 版本（如 1.4）</summary>
    public String? PdfVersion { get; set; }

    /// <summary>总页数</summary>
    public Int32 PageCount { get; set; }
    #endregion
}

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
    private Boolean _disposed;
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
    public void Dispose()
    {
        _disposed = true;
        GC.SuppressFinalize(this);
    }
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
    #endregion
}
