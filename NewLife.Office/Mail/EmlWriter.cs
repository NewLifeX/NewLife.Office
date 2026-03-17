using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

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
    public void Write(EmlMessage message, String path)
    {
        var content = Build(message);
        File.WriteAllText(path, content, Encoding.UTF8);
    }

    /// <summary>将 EML 消息写入流</summary>
    /// <param name="message">邮件消息</param>
    /// <param name="stream">可写输出流</param>
    public void Write(EmlMessage message, Stream stream)
    {
        var content = Build(message);
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>将 EML 消息序列化为字符串</summary>
    /// <param name="message">邮件消息</param>
    /// <returns>EML 格式的字符串</returns>
    public String Build(EmlMessage message)
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

        return sb.ToString();
    }

    #endregion

    #region 私有方法

    private static void AppendBodyPart(StringBuilder sb, String boundary, EmlMessage message,
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

    private static void AppendAttachment(StringBuilder sb, String boundary, EmlAttachment att, Boolean inline = false)
    {
        sb.AppendLine($"--{boundary}");
        sb.AppendLine($"Content-Type: {att.ContentType}");
        if (att.ContentId != null)
            sb.AppendLine($"Content-ID: {att.ContentId}");
        var disp = inline ? "inline" : "attachment";
        if (att.FileName != null)
            sb.AppendLine($"Content-Disposition: {disp}; filename=\"{att.FileName}\"");
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

    /// <summary>对需要编码的头部值使用 RFC 2047 Base64 编码</summary>
    private static String EncodeHeaderValue(String value)
    {
        // 如果全为 ASCII，无需编码
        if (value.All(c => c < 128 && c >= 32)) return value;
        return "=?utf-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value)) + "?=";
    }

    /// <summary>Quoted-Printable 编码（UTF-8 正文）</summary>
    private static String EncodeQuotedPrintable(String text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        var sb = new StringBuilder();
        var lineLen = 0;
        foreach (var b in bytes)
        {
            String token;
            if (b == '\r') { token = "\r"; }
            else if (b == '\n') { sb.Append("\r\n"); lineLen = 0; continue; }
            else if (b >= 33 && b <= 126 && b != 61) { token = ((Char)b).ToString(); }
            else { token = $"={b:X2}"; }

            if (lineLen + token.Length > 75)
            {
                sb.Append("=\r\n");
                lineLen = 0;
            }
            sb.Append(token);
            lineLen += token.Length;
        }
        return sb.ToString();
    }

    private static String GenerateBoundary(String tag)
    {
        return $"----=_Part_{tag}_{Guid.NewGuid():N}";
    }

    #endregion
}
