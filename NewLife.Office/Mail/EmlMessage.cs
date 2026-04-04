using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office;

/// <summary>EML 邮件消息（RFC 5322 + MIME）</summary>
/// <remarks>
/// 表示一封完整的 EML 格式邮件，包含头部字段、正文和附件。
/// 支持 text/plain、text/html 正文和 multipart/* 结构。
/// </remarks>
public class EmlMessage : ITextExtractable, IMarkdownExtractable
{
    #region 属性

    /// <summary>发件人地址</summary>
    public String? From { get; set; }

    /// <summary>收件人列表（逗号分隔或数组）</summary>
    public List<String> To { get; } = [];

    /// <summary>抄送列表</summary>
    public List<String> Cc { get; } = [];

    /// <summary>密送列表</summary>
    public List<String> Bcc { get; } = [];

    /// <summary>邮件主题</summary>
    public String? Subject { get; set; }

    /// <summary>发送日期</summary>
    public DateTimeOffset? Date { get; set; }

    /// <summary>Message-ID</summary>
    public String? MessageId { get; set; }

    /// <summary>Reply-To 地址</summary>
    public String? ReplyTo { get; set; }

    /// <summary>纯文本正文（text/plain）</summary>
    public String? TextBody { get; set; }

    /// <summary>HTML 正文（text/html）</summary>
    public String? HtmlBody { get; set; }

    /// <summary>附件列表</summary>
    public List<EmlAttachment> Attachments { get; } = [];

    /// <summary>内嵌图片（Content-ID → 附件）</summary>
    public Dictionary<String, EmlAttachment> InlineImages { get; } = [];

    /// <summary>原始头部字段（保留扩展字段）</summary>
    public Dictionary<String, String> Headers { get; } = new(StringComparer.OrdinalIgnoreCase);

    #endregion

    #region 文本提取
    /// <summary>提取纯文本（头部字段+纯文本正文）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText()
    {
        var sb = new StringBuilder();
        if (!String.IsNullOrEmpty(From)) sb.AppendLine($"发件人: {From}");
        if (To.Count > 0) sb.AppendLine($"收件人: {String.Join(", ", To)}");
        if (Cc.Count > 0) sb.AppendLine($"抄送: {String.Join(", ", Cc)}");
        if (!String.IsNullOrEmpty(Subject)) sb.AppendLine($"主题: {Subject}");
        if (Date != null) sb.AppendLine($"日期: {Date:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        if (!String.IsNullOrEmpty(TextBody))
            sb.Append(TextBody);
        else if (!String.IsNullOrEmpty(HtmlBody))
            sb.Append(StripHtml(HtmlBody));

        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（头部加粗+正文）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        var sb = new StringBuilder();
        if (!String.IsNullOrEmpty(From)) sb.AppendLine($"- **发件人**: {From}");
        if (To.Count > 0) sb.AppendLine($"- **收件人**: {String.Join(", ", To)}");
        if (Cc.Count > 0) sb.AppendLine($"- **抄送**: {String.Join(", ", Cc)}");
        if (!String.IsNullOrEmpty(Subject)) sb.AppendLine($"- **主题**: {Subject}");
        if (Date != null) sb.AppendLine($"- **日期**: {Date:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        if (!String.IsNullOrEmpty(TextBody))
            sb.Append(TextBody);
        else if (!String.IsNullOrEmpty(HtmlBody))
            sb.Append(StripHtml(HtmlBody));

        return sb.ToString();
    }

    private static String StripHtml(String html)
    {
        if (String.IsNullOrEmpty(html)) return "";
        var text = Regex.Replace(html, "<[^>]+>", " ");
        text = text.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                   .Replace("&quot;", "\"").Replace("&#39;", "'").Replace("&nbsp;", " ");
        text = Regex.Replace(text, @"\s+", " ").Trim();
        return text;
    }
    #endregion
}

/// <summary>EML 附件</summary>
public class EmlAttachment
{
    #region 属性

    /// <summary>文件名</summary>
    public String? FileName { get; set; }

    /// <summary>Content-Type（如 application/octet-stream、image/png）</summary>
    public String ContentType { get; set; } = "application/octet-stream";

    /// <summary>Content-ID（内嵌图片引用标识，带尖括号）</summary>
    public String? ContentId { get; set; }

    /// <summary>二进制内容</summary>
    public Byte[] Data { get; set; } = [];

    #endregion
}
