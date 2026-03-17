using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>EML 邮件文件读取器（RFC 5322 + MIME）</summary>
/// <remarks>
/// 支持读取 .eml 文件或 RFC 5322 格式的字节流/字符串。
/// 支持 text/plain、text/html 正文、multipart/mixed、multipart/alternative、multipart/related。
/// 支持 Q 编码（RFC 2047）和 Base64/Quoted-Printable 内容传输编码。
/// </remarks>
public class EmlReader
{
    #region 读取方法

    /// <summary>从文件路径读取 EML</summary>
    /// <param name="path">EML 文件路径</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage Read(String path)
    {
        var bytes = File.ReadAllBytes(path);
        return Parse(bytes);
    }

    /// <summary>从流读取 EML</summary>
    /// <param name="stream">包含 EML 内容的可读流</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage Read(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return Parse(ms.ToArray());
    }

    /// <summary>从字节数组解析 EML</summary>
    /// <param name="data">EML 原始字节</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage Parse(Byte[] data)
    {
        // 用 Latin-1 保白字节（Unicode 0-255 与 ISO-8859-1 一一对应）
        var text = BytesToLatin1(data);
        return ParseText(text);
    }

    /// <summary>从文本字符串解析 EML</summary>
    /// <param name="text">EML 文本内容（Latin-1 编码保持字节完整）</param>
    /// <returns>解析后的邮件消息</returns>
    public EmlMessage ParseText(String text)
    {
        var msg = new EmlMessage();
        var lines = SplitLines(text);
        var pos = 0;

        // ─── 1. 解析头部 ──────────────────────────────────────────────────
        var headers = ParseHeaders(lines, ref pos);

        // 填充消息字段
        foreach (var kvp in headers)
        {
            msg.Headers[kvp.Key] = kvp.Value;
            var lowKey = kvp.Key.ToLowerInvariant();
            var val = kvp.Value;
            switch (lowKey)
            {
                case "from":     msg.From = DecodeHeaderValue(val); break;
                case "reply-to": msg.ReplyTo = DecodeHeaderValue(val); break;
                case "subject":  msg.Subject = DecodeHeaderValue(val); break;
                case "message-id": msg.MessageId = val.Trim(); break;
                case "date":
                    if (DateTimeOffset.TryParse(val.Trim(), out var dt)) msg.Date = dt;
                    break;
                case "to":
                    foreach (var addr in SplitAddresses(val))
                    {
                        msg.To.Add(DecodeHeaderValue(addr));
                    }
                    break;
                case "cc":
                    foreach (var addr in SplitAddresses(val))
                    {
                        msg.Cc.Add(DecodeHeaderValue(addr));
                    }
                    break;
                case "bcc":
                    foreach (var addr in SplitAddresses(val))
                    {
                        msg.Bcc.Add(DecodeHeaderValue(addr));
                    }
                    break;
            }
        }

        // ─── 2. 解析正文 ──────────────────────────────────────────────────
        String ctHeader;
        var contentType = headers.TryGetValue("content-type", out ctHeader) && !String.IsNullOrEmpty(ctHeader)
            ? ctHeader : "text/plain";
        var body = String.Join("\r\n", lines.Skip(pos));
        ParseBody(msg, contentType, body, headers);

        return msg;
    }

    #endregion

    #region 私有方法

    private static String[] SplitLines(String text)
    {
        return text.Split('\n');
    }

    private static Dictionary<String, String> ParseHeaders(String[] lines, ref Int32 pos)
    {
        var headers = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase);
        String? currentKey = null;
        var currentValue = new StringBuilder();

        while (pos < lines.Length)
        {
            var line = lines[pos].TrimEnd('\r');
            if (line.Length == 0)
            {
                pos++;
                break;  // 空行 = 头部结束
            }

            if ((line[0] == ' ' || line[0] == '\t') && currentKey != null)
            {
                // 续行（folded header）
                currentValue.Append(' ').Append(line.TrimStart());
            }
            else
            {
                // 保存前一个头部
                if (currentKey != null)
                    headers[currentKey] = currentValue.ToString().Trim();

                var colonIdx = line.IndexOf(':');
                if (colonIdx > 0)
                {
                    currentKey = line[..colonIdx].Trim().ToLowerInvariant();
                    currentValue.Clear();
                    currentValue.Append(line[(colonIdx + 1)..]);
                }
            }
            pos++;
        }

        // 保存最后一个头部
        if (currentKey != null)
            headers[currentKey] = currentValue.ToString().Trim();

        return headers;
    }

    private static void ParseBody(EmlMessage msg, String contentType, String body,
        Dictionary<String, String> headers)
    {
        var ctLower = contentType.ToLowerInvariant();

        if (ctLower.StartsWith("multipart/", StringComparison.Ordinal))
        {
            var boundary = GetBoundaryFromContentType(contentType);
            if (boundary != null)
                ParseMultipart(msg, body, boundary, ctLower.IndexOf("alternative", StringComparison.Ordinal) >= 0);
        }
        else
        {
            String ct;
            String enc;
            String cs;
            headers.TryGetValue("content-transfer-encoding", out ct);
            enc = ct ?? "7bit";
            headers.TryGetValue("charset", out cs);
            var charset = GetCharset(contentType) ?? "utf-8";
            var decoded = DecodeContent(body, enc, charset);

            if (ctLower.StartsWith("text/html", StringComparison.Ordinal))
                msg.HtmlBody = decoded;
            else
                msg.TextBody = decoded;
        }
    }

    private static void ParseMultipart(EmlMessage msg, String body, String boundary, Boolean alternative)
    {
        var delimiter = "--" + boundary;
        var endDelimiter = "--" + boundary + "--";

        var lines = body.Split('\n');
        var partLines = new List<String>();
        var inPart = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd('\r');
            if (line.StartsWith(endDelimiter, StringComparison.Ordinal))
            {
                if (inPart && partLines.Count > 0)
                    ProcessPart(msg, String.Join("\r\n", partLines), alternative);
                break;
            }
            if (line.StartsWith(delimiter, StringComparison.Ordinal))
            {
                if (inPart && partLines.Count > 0)
                    ProcessPart(msg, String.Join("\r\n", partLines), alternative);
                partLines.Clear();
                inPart = true;
            }
            else if (inPart)
            {
                partLines.Add(line);
            }
        }
    }

    private static void ProcessPart(EmlMessage msg, String partText, Boolean alternative)
    {
        var lines = SplitLines(partText);
        var pos = 0;
        var partHeaders = ParseHeaders(lines, ref pos);
        var partBody = String.Join("\r\n", lines.Skip(pos));

        var contentType = "text/plain";
        var ctmp = String.Empty;
        if (partHeaders.TryGetValue("content-type", out ctmp) && !String.IsNullOrEmpty(ctmp))
            contentType = ctmp;
        var ctLower = contentType.ToLowerInvariant();

        var encoding = "7bit";
        var etmp = String.Empty;
        if (partHeaders.TryGetValue("content-transfer-encoding", out etmp) && !String.IsNullOrEmpty(etmp))
            encoding = etmp;

        var charset = GetCharset(contentType) ?? "utf-8";

        var contentDisp = String.Empty;
        var dtmp = String.Empty;
        if (partHeaders.TryGetValue("content-disposition", out dtmp) && !String.IsNullOrEmpty(dtmp))
            contentDisp = dtmp;

        String? contentId = null;
        var cidtmp = String.Empty;
        if (partHeaders.TryGetValue("content-id", out cidtmp) && !String.IsNullOrEmpty(cidtmp))
            contentId = cidtmp;

        if (ctLower.StartsWith("multipart/", StringComparison.Ordinal))
        {
            var boundary = GetBoundaryFromContentType(contentType);
            if (boundary != null)
                ParseMultipart(msg, partBody, boundary, ctLower.IndexOf("alternative", StringComparison.Ordinal) >= 0);
            return;
        }

        var isAttachment = contentDisp.ToLowerInvariant().IndexOf("attachment", StringComparison.Ordinal) >= 0;

        if (isAttachment || (!ctLower.StartsWith("text/", StringComparison.Ordinal) && !ctLower.StartsWith("multipart/", StringComparison.Ordinal)))
        {
            // 附件
            var att = new EmlAttachment
            {
                ContentType = contentType.Split(';')[0].Trim(),
                ContentId = contentId?.Trim(),
                FileName = GetFilenameFromDisposition(contentDisp, contentType),
                Data = DecodeBytes(partBody.Trim(), encoding),
            };
            if (contentId != null)
                msg.InlineImages[att.ContentId!] = att;
            else
                msg.Attachments.Add(att);
        }
        else if (ctLower.StartsWith("text/html", StringComparison.Ordinal))
        {
            if (msg.HtmlBody == null || !alternative)
                msg.HtmlBody = DecodeContent(partBody, encoding, charset);
        }
        else
        {
            if (msg.TextBody == null || !alternative)
                msg.TextBody = DecodeContent(partBody, encoding, charset);
        }
    }

    private static String? GetBoundaryFromContentType(String contentType)
    {
        var m = Regex.Match(contentType, @"boundary=""?([^"";\s]+)""?", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : null;
    }

    private static String? GetCharset(String contentType)
    {
        var m = Regex.Match(contentType, @"charset=""?([^"";\s]+)""?", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : null;
    }

    private static String? GetFilenameFromDisposition(String disposition, String contentType)
    {
        var m = Regex.Match(disposition, @"filename\*?=""?([^"";\r\n]+)""?", RegexOptions.IgnoreCase);
        if (m.Success) return DecodeHeaderValue(m.Groups[1].Value.Trim());
        m = Regex.Match(contentType, @"name=""?([^"";\r\n]+)""?", RegexOptions.IgnoreCase);
        return m.Success ? DecodeHeaderValue(m.Groups[1].Value.Trim()) : null;
    }

    private static String DecodeContent(String body, String encoding, String charset)
    {
        var enc = encoding.Trim().ToLowerInvariant();
        if (enc == "base64")
        {
            var clean = Regex.Replace(body, @"\s+", "");
            try
            {
                var bytes = Convert.FromBase64String(clean);
                return GetEncoding(charset).GetString(bytes);
            }
            catch { return body; }
        }
        if (enc == "quoted-printable")
        {
            return DecodeQuotedPrintable(body, charset);
        }
        // 7bit / 8bit / binary：字节重新解析
        var latin = Latin1ToBytes(body);
        return GetEncoding(charset).GetString(latin);
    }

    private static Byte[] DecodeBytes(String body, String encoding)
    {
        var enc = encoding.Trim().ToLowerInvariant();
        if (enc == "base64")
        {
            var clean = Regex.Replace(body, @"\s+", "");
            try { return Convert.FromBase64String(clean); } catch { return new Byte[0]; }
        }
        return Latin1ToBytes(body);
    }

    private static String DecodeQuotedPrintable(String input, String charset)
    {
        var sb = new StringBuilder();
        var ms = new MemoryStream();
        var lines = input.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].TrimEnd('\r');
            if (line.Length > 0 && line[line.Length - 1] == '=')
            {
                line = line[..^1];  // soft line break
                var lineBytes = ParseQpLine(line);
                ms.Write(lineBytes, 0, lineBytes.Length);
            }
            else
            {
                var lineBytes = ParseQpLine(line);
                ms.Write(lineBytes, 0, lineBytes.Length);
                if (i < lines.Length - 1) ms.WriteByte((Byte)'\n');
            }
        }
        return GetEncoding(charset).GetString(ms.ToArray());
    }

    private static Byte[] ParseQpLine(String line)
    {
        var ms = new MemoryStream();
        var i = 0;
        while (i < line.Length)
        {
            if (line[i] == '=' && i + 2 < line.Length)
            {
                if (Byte.TryParse(line.Substring(i + 1, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                {
                    ms.WriteByte(b);
                    i += 3;
                }
                else { ms.WriteByte((Byte)line[i++]); }
            }
            else
            {
                ms.WriteByte((Byte)line[i++]);
            }
        }
        return ms.ToArray();
    }

    /// <summary>解码 RFC 2047 编码词（=?charset?Q/B?text?=）</summary>
    private static String DecodeHeaderValue(String value)
    {
        if (!value.Contains("=?", StringComparison.Ordinal)) return value;
        return Regex.Replace(value, @"=\?([^?]+)\?([QB])\?([^?]*)\?=",
            m =>
            {
                var charset = m.Groups[1].Value;
                var method = m.Groups[2].Value.ToUpperInvariant();
                var encoded = m.Groups[3].Value;
                try
                {
                    var enc = GetEncoding(charset);
                    if (method == "B")
                        return enc.GetString(Convert.FromBase64String(encoded));
                    // Q encoding: _ → space, =HH → byte
                    encoded = encoded.Replace('_', ' ');
                    var bytes = ParseQpLine(encoded);
                    return enc.GetString(bytes);
                }
                catch { return m.Value; }
            }, RegexOptions.IgnoreCase);
    }

    private static IEnumerable<String> SplitAddresses(String value)
    {
        // 简单按逗号分隔（不考虑带逗号的 quoted name）
        return value.Split(',');
    }

    private static Encoding GetEncoding(String charset)
    {
        try { return Encoding.GetEncoding(charset); }
        catch { return Encoding.UTF8; }
    }

    /// <summary>Latin-1 字节保白：将字节数组映射到 Unicode 0-255，避免依赖 Encoding.Latin1</summary>
    private static String BytesToLatin1(Byte[] bytes)
    {
        var chars = new Char[bytes.Length];
        for (var i = 0; i < bytes.Length; i++)
        {
            chars[i] = (Char)bytes[i];
        }
        return new String(chars);
    }

    /// <summary>Latin-1 字节转换：将 Unicode 0-255 字符串转回字节</summary>
    private static Byte[] Latin1ToBytes(String str)
    {
        var bytes = new Byte[str.Length];
        for (var i = 0; i < str.Length; i++)
        {
            bytes[i] = (Byte)str[i];
        }
        return bytes;
    }

    #endregion
}
