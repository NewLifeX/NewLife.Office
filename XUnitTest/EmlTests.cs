using System.ComponentModel;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>EML 邮件格式读写测试</summary>
public class EmlTests
{
    #region 辅助

    private static String BuildSimpleEml(String from, String to, String subject, String body)
    {
        return $"From: {from}\r\nTo: {to}\r\nSubject: {subject}\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Transfer-Encoding: 7bit\r\n\r\n{body}";
    }

    #endregion

    #region 读取测试

    [Fact]
    [DisplayName("解析纯文本邮件")]
    public void Parse_PlainText_ReadsHeaders()
    {
        var eml = BuildSimpleEml("alice@example.com", "bob@example.com", "Hello", "This is the body.");
        var reader = new EmlReader();
        var msg = reader.ParseText(eml);

        Assert.Equal("alice@example.com", msg.From);
        Assert.Contains("bob@example.com", msg.To);
        Assert.Equal("Hello", msg.Subject);
        Assert.NotNull(msg.TextBody);
        Assert.Contains("This is the body", msg.TextBody);
    }

    [Fact]
    [DisplayName("解析 multipart/alternative 邮件")]
    public void Parse_MultipartAlternative_BothBodies()
    {
        var boundary = "boundary_test_001";
        var sb = new StringBuilder();
        sb.Append("From: sender@example.com\r\n");
        sb.Append("To: recv@example.com\r\n");
        sb.Append("Subject: Test Alt\r\n");
        sb.Append("MIME-Version: 1.0\r\n");
        sb.Append($"Content-Type: multipart/alternative; boundary=\"{boundary}\"\r\n");
        sb.Append("\r\n");
        sb.Append($"--{boundary}\r\n");
        sb.Append("Content-Type: text/plain; charset=utf-8\r\n");
        sb.Append("Content-Transfer-Encoding: 7bit\r\n");
        sb.Append("\r\n");
        sb.Append("Plain text content\r\n");
        sb.Append($"--{boundary}\r\n");
        sb.Append("Content-Type: text/html; charset=utf-8\r\n");
        sb.Append("Content-Transfer-Encoding: base64\r\n");
        sb.Append("\r\n");
        sb.Append(Convert.ToBase64String(Encoding.UTF8.GetBytes("<p>HTML content</p>")));
        sb.Append("\r\n");
        sb.Append($"--{boundary}--\r\n");

        var reader = new EmlReader();
        var msg = reader.ParseText(sb.ToString());

        Assert.NotNull(msg.TextBody);
        Assert.Contains("Plain text", msg.TextBody);
        Assert.NotNull(msg.HtmlBody);
        Assert.Contains("HTML content", msg.HtmlBody);
    }

    [Fact]
    [DisplayName("解析带附件的邮件")]
    public void Parse_WithAttachment_AttachmentParsed()
    {
        var boundary = "boundary_attach_001";
        var attData = Encoding.UTF8.GetBytes("attachment content here");
        var sb = new StringBuilder();
        sb.Append("From: a@b.com\r\n");
        sb.Append("To: c@d.com\r\n");
        sb.Append("Subject: File\r\n");
        sb.Append("MIME-Version: 1.0\r\n");
        sb.Append($"Content-Type: multipart/mixed; boundary=\"{boundary}\"\r\n");
        sb.Append("\r\n");
        sb.Append($"--{boundary}\r\n");
        sb.Append("Content-Type: text/plain; charset=utf-8\r\n");
        sb.Append("Content-Transfer-Encoding: 7bit\r\n");
        sb.Append("\r\n");
        sb.Append("See attached file\r\n");
        sb.Append($"--{boundary}\r\n");
        sb.Append("Content-Type: application/octet-stream\r\n");
        sb.Append("Content-Disposition: attachment; filename=\"test.bin\"\r\n");
        sb.Append("Content-Transfer-Encoding: base64\r\n");
        sb.Append("\r\n");
        sb.Append(Convert.ToBase64String(attData));
        sb.Append("\r\n");
        sb.Append($"--{boundary}--\r\n");

        var reader = new EmlReader();
        var msg = reader.ParseText(sb.ToString());

        Assert.NotNull(msg.TextBody);
        Assert.Single(msg.Attachments);
        Assert.Equal("test.bin", msg.Attachments[0].FileName);
        Assert.Equal(attData, msg.Attachments[0].Data);
    }

    [Fact]
    [DisplayName("解析 RFC 2047 Q 编码主题")]
    public void Parse_QEncodedSubject_DecodesCorrectly()
    {
        // =?utf-8?B?5rWL6K+V5LiD6K+V?= => base64 of "测试三号"
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes("测试主题"));
        var eml = $"From: x@y.com\r\nSubject: =?utf-8?B?{encoded}?=\r\n\r\n";

        var reader = new EmlReader();
        var msg = reader.ParseText(eml);

        Assert.Equal("测试主题", msg.Subject);
    }

    [Fact]
    [DisplayName("解析多收件人")]
    public void Parse_MultipleRecipients_AllParsed()
    {
        var eml = "From: a@b.com\r\nTo: x@y.com, z@w.com, q@r.com\r\nSubject: test\r\n\r\n";
        var reader = new EmlReader();
        var msg = reader.ParseText(eml);

        Assert.Equal(3, msg.To.Count);
    }

    #endregion

    #region 写入测试

    [Fact]
    [DisplayName("写入纯文本邮件")]
    public void Write_PlainText_ContainsHeaders()
    {
        var msg = new EmlMessage
        {
            From = "alice@example.com",
            Subject = "Test Subject",
            TextBody = "Hello World",
        };
        msg.To.Add("bob@example.com");

        var writer = new EmlWriter();
        var eml = writer.Build(msg);

        Assert.Contains("From: alice@example.com", eml);
        Assert.Contains("To: bob@example.com", eml);
        Assert.Contains("Subject: Test Subject", eml);
        Assert.Contains("Content-Type: text/plain", eml);
    }

    [Fact]
    [DisplayName("写入包含 HTML 和文本的邮件（multipart/alternative）")]
    public void Write_HtmlAndText_MultipartAlternative()
    {
        var msg = new EmlMessage
        {
            From = "a@b.com",
            Subject = "Alt Test",
            TextBody = "plain",
            HtmlBody = "<p>html</p>",
        };
        msg.To.Add("c@d.com");

        var writer = new EmlWriter();
        var eml = writer.Build(msg);

        Assert.Contains("multipart/alternative", eml);
        Assert.Contains("text/plain", eml);
        Assert.Contains("text/html", eml);
    }

    [Fact]
    [DisplayName("写入带附件的邮件（multipart/mixed）")]
    public void Write_WithAttachment_MultipartMixed()
    {
        var msg = new EmlMessage
        {
            From = "a@b.com",
            TextBody = "body",
        };
        msg.To.Add("b@c.com");
        msg.Attachments.Add(new EmlAttachment
        {
            FileName = "doc.txt",
            ContentType = "text/plain",
            Data = Encoding.UTF8.GetBytes("attachment body"),
        });

        var writer = new EmlWriter();
        var eml = writer.Build(msg);

        Assert.Contains("multipart/mixed", eml);
        Assert.Contains("doc.txt", eml);
        Assert.Contains("base64", eml);
    }

    [Fact]
    [DisplayName("往返测试：写入后读取还原邮件内容")]
    public void RoundTrip_WriteAndRead_PreservesContent()
    {
        var original = new EmlMessage
        {
            From = "sender@test.com",
            Subject = "Round-trip Test",
            TextBody = "Round trip body text.",
        };
        original.To.Add("receiver@test.com");

        var writer = new EmlWriter();
        var eml = writer.Build(original);

        var reader = new EmlReader();
        var parsed = reader.ParseText(eml);

        Assert.Equal(original.From, parsed.From);
        Assert.Equal("Round-trip Test", parsed.Subject);
        Assert.NotNull(parsed.TextBody);
        Assert.Contains("Round trip body text", parsed.TextBody);
    }

    [Fact]
    [DisplayName("写入中文主题使用 Base64 编码")]
    public void Write_ChineseSubject_EncodedAsBase64()
    {
        var msg = new EmlMessage { From = "a@b.com", Subject = "中文主题" };
        msg.To.Add("x@y.com");

        var writer = new EmlWriter();
        var eml = writer.Build(msg);

        Assert.Contains("=?utf-8?B?", eml);
        Assert.DoesNotContain("Subject: 中文主题", eml);  // 不该出现未编码中文
    }

    #endregion

    #region 文件写入集成测试

    [Fact]
    [DisplayName("集成：写入 EML 文件并从文件读取")]
    public void Integration_WriteFile_ThenReadFile()
    {
        var dir = Path.Combine("Bin", "UnitTest", "Artifacts");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "test_output.eml");

        var original = new EmlMessage
        {
            From = "test@newlife.org",
            Subject = "集成测试邮件",
            TextBody = "这是一封集成测试生成的邮件。",
            HtmlBody = "<p>这是一封<b>集成测试</b>生成的邮件。</p>",
        };
        original.To.Add("user@example.com");
        original.Cc.Add("admin@newlife.org");
        original.Attachments.Add(new EmlAttachment
        {
            FileName = "readme.txt",
            ContentType = "text/plain",
            Data = Encoding.UTF8.GetBytes("NewLife.Office EML 集成测试"),
        });

        var writer = new EmlWriter();
        writer.Write(original, path);

        Assert.True(File.Exists(path));

        var reader = new EmlReader();
        var parsed = reader.Read(path);

        Assert.Contains("test@newlife.org", parsed.From ?? "");
        Assert.NotEmpty(parsed.To);
        Assert.NotEmpty(parsed.Cc);
        Assert.Single(parsed.Attachments);
        Assert.Equal("readme.txt", parsed.Attachments[0].FileName);
    }

    #endregion
}
