using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office.Mail;

/// <summary>EML 邮件文件写入器（RFC 5322 + MIME）</summary>
/// <remarks>
/// 生成符合 RFC 5322 和 MIME 标准的 .eml 文件。
/// 纯文本/HTML 正文自动选择 multipart/alternative 结构；
/// 有附件时使用 multipart/mixed 包装。
/// 默认使用 UTF-8 编码，头部使用 Base64 Q 编码（=?utf-8?B?...?=）。
/// </remarks>
public class EmlWriter
{
    #region 写入方法

    /// <summary>将 EML 消息写入文件</summary>
    /// <param name="message">邮件消息</param>
    /// <param name="path">输出文件路径</param>
    public void Write(Message message, String path)
    {
        var content = Build(message);
        // EML 是纯文本协议文件，不能带 UTF-8 BOM，否则第三方解析器会把首行头部读坏
        File.WriteAllText(path, content, new UTF8Encoding(false));
    }

    /// <summary>将 EML 消息写入流</summary>
    /// <param name="message">邮件消息</param>
    /// <param name="stream">可写输出流</param>
    public void Write(Message message, Stream stream)
    {
        var content = Build(message);
        var bytes = new UTF8Encoding(false).GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>将 EML 消息序列化为字符串</summary>
    /// <param name="message">邮件消息</param>
    /// <returns>EML 格式的字符串</returns>
    public String Build(Message message)
    {
        var sb = new StringBuilder();
        var boundary1 = GenerateBoundary("alt");
        var boundary2 = GenerateBoundary("mix");

        var hasText = message.TextBody != null;
        var hasHtml = message.HtmlBody != null;
        var hasAttachments = message.Attachments.Count > 0 || message.InlineImages.Count > 0;
        var multiBody = hasText && hasHtml;

        // ─── 头部 ─────────────────────────────────────────────────────────
        AppendHeader(sb, "MIME-Version", "1.0");
        AppendHeader(sb, "Date", (message.Date ?? DateTimeOffset.UtcNow).ToString("r"));
        if (message.MessageId != null)
            AppendHeader(sb, "Message-ID", message.MessageId);
        else
            AppendHeader(sb, "Message-ID", $"<{Guid.NewGuid():N}@newlife.office>");
        if (message.From != null)
            AppendHeader(sb, "From", EncodeHeaderValue(message.From));
        foreach (var to in message.To)
        {
            AppendHeader(sb, "To", EncodeHeaderValue(to));
        }
        foreach (var cc in message.Cc)
        {
            AppendHeader(sb, "Cc", EncodeHeaderValue(cc));
        }
        if (message.ReplyTo != null)
            AppendHeader(sb, "Reply-To", EncodeHeaderValue(message.ReplyTo));
        if (message.Subject != null)
            AppendHeader(sb, "Subject", EncodeHeaderValue(message.Subject));

        // ─── Content-Type ─────────────────────────────────────────────────
        if (hasAttachments)
        {
            sb.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary2}\"");
            sb.AppendLine();
            // mixed 外层
            sb.AppendLine($"--{boundary2}");
            // 内层 alternative 或单正文
            AppendBodyPart(sb, boundary1, message, hasText, hasHtml, multiBody);
            // 附件
            foreach (var att in message.Attachments)
            {
                AppendAttachment(sb, boundary2, att);
            }
            foreach (var att in message.InlineImages.Values)
            {
                AppendAttachment(sb, boundary2, att, inline: true);
            }
            sb.AppendLine($"--{boundary2}--");
        }
        else if (multiBody)
        {
            sb.AppendLine($"Content-Type: multipart/alternative; boundary=\"{boundary1}\"");
            sb.AppendLine();
            AppendBodyPart(sb, boundary1, message, hasText, hasHtml, multiBody);
        }
        else if (hasHtml)
        {
            sb.AppendLine("Content-Type: text/html; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: base64");
            sb.AppendLine();
            sb.AppendLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(message.HtmlBody!)));
        }
        else
        {
            sb.AppendLine("Content-Type: text/plain; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
            sb.AppendLine();
            sb.AppendLine(EncodeQuotedPrintable(message.TextBody ?? ""));
        }

        // 协议要求行结束符为 CRLF，不随运行平台变化
        var result = sb.ToString().Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
        return result;
    }

    #endregion

    #region 私有方法

    private static void AppendBodyPart(StringBuilder sb, String boundary, Message message,
        Boolean hasText, Boolean hasHtml, Boolean multiBody)
    {
        if (multiBody)
        {
            sb.AppendLine($"Content-Type: multipart/alternative; boundary=\"{boundary}\"");
            sb.AppendLine();
            if (hasText)
            {
                sb.AppendLine($"--{boundary}");
                sb.AppendLine("Content-Type: text/plain; charset=utf-8");
                sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
                sb.AppendLine();
                sb.AppendLine(EncodeQuotedPrintable(message.TextBody!));
            }
            if (hasHtml)
            {
                sb.AppendLine($"--{boundary}");
                sb.AppendLine("Content-Type: text/html; charset=utf-8");
                sb.AppendLine("Content-Transfer-Encoding: base64");
                sb.AppendLine();
                sb.AppendLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(message.HtmlBody!)));
            }
            sb.AppendLine($"--{boundary}--");
        }
        else if (hasHtml)
        {
            sb.AppendLine("Content-Type: text/html; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: base64");
            sb.AppendLine();
            sb.AppendLine(Convert.ToBase64String(Encoding.UTF8.GetBytes(message.HtmlBody!)));
        }
        else
        {
            sb.AppendLine("Content-Type: text/plain; charset=utf-8");
            sb.AppendLine("Content-Transfer-Encoding: quoted-printable");
            sb.AppendLine();
            sb.AppendLine(EncodeQuotedPrintable(message.TextBody ?? ""));
        }
    }

    private static void AppendAttachment(StringBuilder sb, String boundary, Attachment att, Boolean inline = false)
    {
        sb.AppendLine($"--{boundary}");
        sb.AppendLine($"Content-Type: {StripBreaks(att.ContentType)}");
        if (att.ContentId != null)
            sb.AppendLine($"Content-ID: {StripBreaks(att.ContentId)}");
        var disp = inline ? "inline" : "attachment";
        if (att.FileName != null)
            sb.AppendLine($"Content-Disposition: {disp}; filename=\"{EscapeQuoted(att.FileName)}\"");
        else
            sb.AppendLine($"Content-Disposition: {disp}");
        sb.AppendLine("Content-Transfer-Encoding: base64");
        sb.AppendLine();
        sb.AppendLine(Convert.ToBase64String(att.Data));
    }

    private static void AppendHeader(StringBuilder sb, String name, String value)
    {
        sb.AppendLine($"{name}: {value}");
    }

    /// <summary>转义引号参数中的特殊字符，防止头部注入</summary>
    private static String EscapeQuoted(String value) => StripBreaks(value)
        .Replace("\\", "\\\\")
        .Replace("\"", "\\\"");

    /// <summary>移除 CR/LF，防止头部注入</summary>
    private static String StripBreaks(String value) => value.Replace("\r", String.Empty).Replace("\n", String.Empty);

    /// <summary>对需要编码的头部值使用 RFC 2047 Base64 编码</summary>
    private static String EncodeHeaderValue(String value)
    {
        // 含换行的值一律编码，防止头部注入（CRLF 注入伪造头部）
        var hasBreak = value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0;
        // 如果全为 ASCII 且不含控制字符，无需编码
        if (!hasBreak && value.All(c => c < 128 && c >= 32)) return value;
        return "=?utf-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value)) + "?=";
    }

    /// <summary>Quoted-Printable 编码（UTF-8 正文）</summary>
    private static String EncodeQuotedPrintable(String text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        var sb = new StringBuilder();
        var line = new StringBuilder();
        var i = 0;
        while (i < bytes.Length)
        {
            var b = bytes[i];
            if (b == '\r' || b == '\n')
            {
                // CR/LF/CRLF 一律归一为硬换行，避免产生 "\r\r\n"
                if (b == '\r' && i + 1 < bytes.Length && bytes[i + 1] == '\n') i++;
                AppendQpLine(sb, line, false);
                i++;
                continue;
            }

            var token = b >= 33 && b <= 126 && b != 61 ? ((Char)b).ToString() : $"={b:X2}";
            if (line.Length + token.Length > 75) AppendQpLine(sb, line, true);

            line.Append(token);
            i++;
        }
        AppendQpLine(sb, line, false);

        return sb.ToString();
    }

    /// <summary>输出一行 QP 文本；软换行追加 "="，行尾空白编码为 =XX</summary>
    private static void AppendQpLine(StringBuilder sb, StringBuilder line, Boolean softWrap)
    {
        // 行尾空格/制表符在传输过程中可能被裁剪，必须编码
        while (line.Length > 0 && (line[line.Length - 1] == ' ' || line[line.Length - 1] == '\t'))
        {
            var ch = line[line.Length - 1];
            line.Length--;
            line.Append('=').Append(((Int32)ch).ToString("X2"));
        }

        sb.Append(line);
        sb.Append(softWrap ? "=\r\n" : "\r\n");
        line.Clear();
    }

    private static String GenerateBoundary(String tag)
    {
        return $"----=_Part_{tag}_{Guid.NewGuid():N}";
    }

    #endregion
}
