using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>MsgReader 单元测试（EM02-01/02/03）</summary>
public class MsgReaderTests
{
    // ─── 辅助：构建最小 MSG 内存文档 ─────────────────────────────────────────

    /// <summary>向 CFB 存储写入 MAPI Unicode 字符串属性流</summary>
    private static void AddUnicodeStream(CfbStorage store, String propId, String value)
    {
        var data = Encoding.Unicode.GetBytes(value + "\0");   // UTF-16LE + null
        store.AddStream($"__substg1.0_{propId}001F", data);
    }

    /// <summary>向 CFB 存储写入 MAPI ANSI 字符串属性流</summary>
    private static void AddAnsiStream(CfbStorage store, String propId, String value)
    {
        var data = Encoding.ASCII.GetBytes(value + "\0");
        store.AddStream($"__substg1.0_{propId}001E", data);
    }

    /// <summary>构建含指定字段的 MSG 字节数组</summary>
    private static Byte[] BuildMsgBytes(
        String subject, String body,
        String senderEmail, String senderName,
        String displayTo,
        String? htmlBody = null)
    {
        var doc = new CfbDocument();
        var root = doc.Root;

        AddUnicodeStream(root, "0037", subject);
        AddUnicodeStream(root, "1000", body);
        AddUnicodeStream(root, "0C1A", senderName);
        AddUnicodeStream(root, "0C1F", senderEmail);
        AddUnicodeStream(root, "0E04", displayTo);

        if (htmlBody != null)
            AddUnicodeStream(root, "1013", htmlBody);

        return doc.ToBytes();
    }

    // ─── EM02-01 基础读取 ───────────────────────────────────────────────────

    [Fact]
    [DisplayName("EM02-01 读取 MSG 主题")]
    public void Read_Subject_IsExtracted()
    {
        var bytes = BuildMsgBytes(
            subject: "Hello World",
            body: "Body text",
            senderEmail: "alice@example.com",
            senderName: "Alice",
            displayTo: "Bob");

        var reader = new MsgReader();
        using var ms = new MemoryStream(bytes);
        var msg = reader.Read(ms);

        Assert.Equal("Hello World", msg.Subject);
    }

    [Fact]
    [DisplayName("EM02-01 读取 MSG 纯文本正文")]
    public void Read_PlainBody_IsExtracted()
    {
        var bytes = BuildMsgBytes(
            subject: "Test",
            body: "Plain text body content",
            senderEmail: "a@b.com",
            senderName: "A",
            displayTo: "B");

        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.Equal("Plain text body content", msg.TextBody);
    }

    [Fact]
    [DisplayName("EM02-01 读取 MSG 发件人地址")]
    public void Read_SenderEmail_IsExtracted()
    {
        var bytes = BuildMsgBytes(
            subject: "Sub",
            body: "Body",
            senderEmail: "sender@test.com",
            senderName: "Sender",
            displayTo: "recv@test.com");

        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.NotNull(msg.From);
        Assert.Contains("sender@test.com", msg.From);
        Assert.Contains("Sender", msg.From);
    }

    [Fact]
    [DisplayName("EM02-01 读取 MSG DisplayTo 解析为收件人列表")]
    public void Read_DisplayTo_ParsedToToList()
    {
        var bytes = BuildMsgBytes(
            subject: "Test",
            body: "Body",
            senderEmail: "from@x.com",
            senderName: "From",
            displayTo: "alice@x.com; bob@x.com");

        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        // DisplayTo 提供了 2 个收件人
        Assert.True(msg.To.Count >= 1);
    }

    [Fact]
    [DisplayName("EM02-01 ANSI 编码属性正确读取")]
    public void Read_AnsiProperty_FallbackWorks()
    {
        var doc = new CfbDocument();
        AddAnsiStream(doc.Root, "0037", "ANSI Subject");
        AddAnsiStream(doc.Root, "1000", "ANSI Body");

        using var ms = new MemoryStream(doc.ToBytes());
        var msg = new MsgReader().Read(ms);

        Assert.Equal("ANSI Subject", msg.Subject);
        Assert.Equal("ANSI Body", msg.TextBody);
    }

    // ─── EM02-02 附件提取 ───────────────────────────────────────────────────

    [Fact]
    [DisplayName("EM02-02 提取附件文件名和数据")]
    public void Read_Attachment_NameAndDataExtracted()
    {
        var doc = new CfbDocument();
        AddUnicodeStream(doc.Root, "0037", "Attach Test");

        var attachStorage = doc.Root.AddStorage("__attach_version1.0_#00000000");
        AddUnicodeStream(attachStorage, "3707", "report.pdf");
        AddUnicodeStream(attachStorage, "370E", "application/pdf");
        var pdfData = new Byte[] { 0x25, 0x50, 0x44, 0x46 };   // %PDF
        attachStorage.AddStream("__substg1.0_37010102", pdfData);

        using var ms = new MemoryStream(doc.ToBytes());
        var msg = new MsgReader().Read(ms);

        Assert.Single(msg.Attachments);
        Assert.Equal("report.pdf", msg.Attachments[0].FileName);
        Assert.Equal(4, msg.Attachments[0].Data.Length);
        Assert.Equal(0x25, msg.Attachments[0].Data[0]);   // %
        Assert.Equal("application/pdf", msg.Attachments[0].ContentType);
    }

    [Fact]
    [DisplayName("EM02-02 无附件时 Attachments 为空")]
    public void Read_NoAttachments_EmptyList()
    {
        var bytes = BuildMsgBytes("S", "B", "a@b.com", "A", "B");
        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.Empty(msg.Attachments);
    }

    [Fact]
    [DisplayName("EM02-02 附件无数据流则忽略")]
    public void Read_AttachmentWithoutDataStream_IsSkipped()
    {
        var doc = new CfbDocument();
        AddUnicodeStream(doc.Root, "0037", "Test");
        var attachStorage = doc.Root.AddStorage("__attach_version1.0_#00000000");
        AddUnicodeStream(attachStorage, "3707", "empty.txt");
        // 不添加 __substg1.0_37010102 的数据流

        using var ms = new MemoryStream(doc.ToBytes());
        var msg = new MsgReader().Read(ms);

        Assert.Empty(msg.Attachments);
    }

    [Fact]
    [DisplayName("EM02-02 多个附件均被提取")]
    public void Read_MultipleAttachments_AllExtracted()
    {
        var doc = new CfbDocument();
        AddUnicodeStream(doc.Root, "0037", "Multi attach");

        for (var i = 0; i < 3; i++)
        {
            var st = doc.Root.AddStorage($"__attach_version1.0_#0000000{i}");
            AddUnicodeStream(st, "3707", $"file{i}.bin");
            st.AddStream("__substg1.0_37010102", new Byte[] { (Byte)i });
        }

        using var ms = new MemoryStream(doc.ToBytes());
        var msg = new MsgReader().Read(ms);

        Assert.Equal(3, msg.Attachments.Count);
        Assert.Equal("file0.bin", msg.Attachments[0].FileName);
        Assert.Equal("file2.bin", msg.Attachments[2].FileName);
    }

    // ─── EM02-03 HTML 正文提取 ──────────────────────────────────────────────

    [Fact]
    [DisplayName("EM02-03 读取 HTML 正文")]
    public void Read_HtmlBody_IsExtracted()
    {
        var bytes = BuildMsgBytes(
            subject: "HTML Test",
            body: "Plain fallback",
            senderEmail: "x@y.com",
            senderName: "X",
            displayTo: "Y",
            htmlBody: "<html><body><b>Bold</b></body></html>");

        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.NotNull(msg.HtmlBody);
        Assert.Contains("<b>Bold</b>", msg.HtmlBody);
    }

    [Fact]
    [DisplayName("EM02-03 无 HTML 正文时 HtmlBody 为 null")]
    public void Read_NoHtmlBody_IsNull()
    {
        var bytes = BuildMsgBytes("Sub", "Plain", "a@b.com", "A", "B", htmlBody: null);
        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.Null(msg.HtmlBody);
    }

    [Fact]
    [DisplayName("EM02-03 同时有纯文本和 HTML 正文时均可读取")]
    public void Read_BothPlainAndHtml_BothExtracted()
    {
        var bytes = BuildMsgBytes(
            subject: "Both",
            body: "This is plain text.",
            senderEmail: "a@b.com",
            senderName: "A",
            displayTo: "B",
            htmlBody: "<p>This is HTML.</p>");

        using var ms = new MemoryStream(bytes);
        var msg = new MsgReader().Read(ms);

        Assert.Equal("This is plain text.", msg.TextBody);
        Assert.Contains("This is HTML.", msg.HtmlBody);
    }
}
