using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using NewLife.Office.Calendar;
using NewLife.Office.Epub;
using NewLife.Office.Excel;
using NewLife.Office.Mail;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Ole2;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rtf;
using NewLife.Office.VCard;
using NewLife.Office.Word;
using NewLife.Office.Xps;
using Xunit;

namespace XUnitTest.Common;

/// <summary>跨格式集成测试</summary>
public class CrossFormatIntegrationTests
{
    // ─── RTF 集成：图片嵌入往返 ──────────────────────────────────────────

    [Fact, DisplayName("RTF 图片写入再读取往返测试")]
    public void Rtf_ImageRoundTrip()
    {
        var pngData = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

        var writer = new RtfWriter();
        writer.AddParagraph("Hello RTF");
        writer.AddImage(pngData, "png", 3000, 2000);
        writer.AddParagraph("After image");
        var rtf = writer.ToString();

        var doc = RtfDocument.Parse(rtf);

        Assert.Single(doc.Images);
        Assert.Equal("png", doc.Images[0].Format);
        Assert.Equal(pngData, doc.Images[0].Data);
    }

    // ─── ODS 集成：泛型导出导入往返 ──────────────────────────────────────

    private class ProductInfo
    {
        [DisplayName("商品名称")]
        public String Name { get; set; } = "";

        [DisplayName("单价")]
        public Decimal Price { get; set; }

        [DisplayName("库存")]
        public Int32 Stock { get; set; }
    }

    [Fact, DisplayName("ODS 泛型对象导出再读取往返测试")]
    public void Ods_GenericObjectRoundTrip()
    {
        var items = new[]
        {
            new ProductInfo { Name = "Alpha", Price = 100m, Stock = 10 },
            new ProductInfo { Name = "Beta",  Price = 200m, Stock = 20 },
        };

        using var ms = new MemoryStream();
        var writer = new OdsWriter();
        writer.AddSheet("Products", items);
        writer.Save(ms);

        ms.Position = 0;
        var result = OdsReader.ReadObjects<ProductInfo>(ms).ToList();

        Assert.Equal(2, result.Count);
        Assert.Equal("Alpha", result[0].Name);
        Assert.Equal(100m, result[0].Price);
    }

    // ─── Markdown→Word 集成 ─────────────────────────────────────────────

    [Fact, DisplayName("Markdown→Word 写入并读取验证内容")]
    public void Markdown_ToWord_ContentVerified()
    {
        const String md = "# Integration\n\nMarkdown to Word integration test.\n\n- Item 1\n- Item 2\n";
        var doc     = MarkdownDocument.Parse(md);
        var docxBytes = doc.ToWord();

        // docx 应为合法 ZIP，包含 word/document.xml
        using var ms = new MemoryStream(docxBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/document.xml");
        Assert.NotNull(entry);

        using var sr = new StreamReader(entry!.Open());
        var xml = sr.ReadToEnd();
        Assert.Contains("Integration", xml);
    }

    // ─── XPS 集成：多页往返 ──────────────────────────────────────────────

    [Fact, DisplayName("XPS 多页写入再读取文本一致")]
    public void Xps_MultiPageRoundTrip()
    {
        var writer = new XpsWriter();
        writer.SetProperties(new XpsProperties { Title = "Integration XPS", Creator = "Tests" });
        writer.AddPage(816, 1056, new[] { ("Page 1 content", 96.0, 96.0, 12.0) });
        writer.AddPage(816, 1056, new[] { ("Page 2 content", 96.0, 96.0, 12.0) });

        var bytes = writer.ToBytes();
        using var ms = new MemoryStream(bytes);
        var pages = new XpsReader().Read(ms);

        Assert.Equal(2, pages.Count);
        Assert.Equal("Page 1 content", pages[0].Text);
        Assert.Equal("Page 2 content", pages[1].Text);
    }

    // ─── MSG 集成：构建读取验证 ─────────────────────────────────────────

    [Fact, DisplayName("MSG 构建再读取主题/正文一致")]
    public void Msg_BuildAndReadRoundTrip()
    {
        var msgDoc = new CfbDocument();
        var root   = msgDoc.Root;
        // 写 Unicode 流（MAPI PT_UNICODE = 001F）
        root.AddStream("__substg1.0_0037001F", Encoding.Unicode.GetBytes("Integration Subject\0"));
        root.AddStream("__substg1.0_1000001F", Encoding.Unicode.GetBytes("Integration body text.\0"));
        root.AddStream("__substg1.0_0C1F001F", Encoding.Unicode.GetBytes("from@test.com\0"));
        root.AddStream("__substg1.0_0C1A001F", Encoding.Unicode.GetBytes("FromName\0"));

        using var ms = new MemoryStream(msgDoc.ToBytes());
        var msg = new MsgReader().Read(ms);

        Assert.Equal("Integration Subject", msg.Subject);
        Assert.Equal("Integration body text.", msg.TextBody);
        Assert.Contains("from@test.com", msg.From);
    }
}
