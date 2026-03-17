using System;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>XPS 读取/写入单元测试（XP01/XP02）</summary>
public class XpsTests
{
    // ─── 辅助：构建最小 XPS 字节数组 ─────────────────────────────────────

    private static Byte[] BuildMinimalXps(String pageText = "Hello XPS",
        Double pageW = 816, Double pageH = 1056,
        String? title = null)
    {
        var writer = new XpsWriter();
        if (title != null)
            writer.SetProperties(new XpsProperties { Title = title, Creator = "TestSuite" });
        writer.AddPage(pageW, pageH, new[]
        {
            (pageText, 96.0, 96.0, 16.0)
        });
        return writer.ToBytes();
    }

    // ─── XP02-01 生成 XPS ────────────────────────────────────────────────

    [Fact]
    [DisplayName("XP02-01 ToBytes 返回非空字节数组")]
    public void ToBytes_ReturnsNonEmpty()
    {
        var bytes = BuildMinimalXps();
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 200);
    }

    [Fact]
    [DisplayName("XP02-01 输出是合法 ZIP（PK 头）")]
    public void ToBytes_IsZip()
    {
        var bytes = BuildMinimalXps();
        Assert.Equal(0x50, bytes[0]);   // 'P'
        Assert.Equal(0x4B, bytes[1]);   // 'K'
    }

    [Fact]
    [DisplayName("XP02-01 ZIP 包含 FixedDocumentSequence.fdseq")]
    public void ToBytes_ContainsFdseq()
    {
        var bytes = BuildMinimalXps();
        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("FixedDocumentSequence.fdseq"));
    }

    [Fact]
    [DisplayName("XP02-01 ZIP 包含 FixedDocument 和页面")]
    public void ToBytes_ContainsFixedDocAndPage()
    {
        var bytes = BuildMinimalXps();
        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("Documents/1/1.fds"));
        Assert.NotNull(zip.GetEntry("Documents/1/Pages/1.fpage"));
    }

    [Fact]
    [DisplayName("XP02-01 Save 写入文件")]
    public void Save_WritesFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "test_output.xps");
        try
        {
            var writer = new XpsWriter();
            writer.AddPage(816, 1056, new[] { ("Test", 96.0, 96.0, 14.0) });
            writer.Save(path);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 100);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ─── XP02-02 写入文本 ────────────────────────────────────────────────

    [Fact]
    [DisplayName("XP02-02 页面 XML 包含 Glyphs UnicodeString")]
    public void AddPage_TextAppearsInGlyphs()
    {
        var bytes = BuildMinimalXps("My Custom Text");
        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var pageEntry = zip.GetEntry("Documents/1/Pages/1.fpage")!;
        using var sr = new StreamReader(pageEntry.Open());
        var xml = sr.ReadToEnd();
        Assert.Contains("My Custom Text", xml);
        Assert.Contains("UnicodeString", xml);
    }

    [Fact]
    [DisplayName("XP02-02 多页文档生成正确数量的 fpage 文件")]
    public void AddMultiplePages_AllPagesPresent()
    {
        var writer = new XpsWriter();
        writer.AddPage(816, 1056, new[] { ("Page1", 96.0, 96.0, 12.0) });
        writer.AddPage(816, 1056, new[] { ("Page2", 96.0, 96.0, 12.0) });
        writer.AddPage(816, 1056, new[] { ("Page3", 96.0, 96.0, 12.0) });
        var bytes = writer.ToBytes();

        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("Documents/1/Pages/1.fpage"));
        Assert.NotNull(zip.GetEntry("Documents/1/Pages/2.fpage"));
        Assert.NotNull(zip.GetEntry("Documents/1/Pages/3.fpage"));
    }

    [Fact]
    [DisplayName("XP02-02 特殊字符 XML 转义正确")]
    public void AddPage_SpecialCharsAreEscaped()
    {
        var bytes = BuildMinimalXps("a < b & c > d");
        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var pageEntry = zip.GetEntry("Documents/1/Pages/1.fpage")!;
        using var sr = new StreamReader(pageEntry.Open());
        var xml = sr.ReadToEnd();
        // < 应被转义为 &lt;
        Assert.Contains("&lt;", xml);
        Assert.Contains("&amp;", xml);
        Assert.DoesNotContain("a < b", xml); // 原始 < 不应出现
    }

    // ─── XP02-04 元数据 ─────────────────────────────────────────────────

    [Fact]
    [DisplayName("XP02-04 SetProperties 写入 docProps/core.xml")]
    public void SetProperties_TitleInCoreXml()
    {
        var bytes = BuildMinimalXps(title: "My XPS Title");
        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.GetEntry("docProps/core.xml");
        Assert.NotNull(entry);
        using var sr = new StreamReader(entry!.Open());
        var xml = sr.ReadToEnd();
        Assert.Contains("My XPS Title", xml);
    }

    // ─── XP01-01/02 读取 ────────────────────────────────────────────────

    [Fact]
    [DisplayName("XP01-01 Read 返回正确页数")]
    public void Read_ReturnsCorrectPageCount()
    {
        var writer = new XpsWriter();
        writer.AddPage(816, 1056, new[] { ("p1", 96.0, 96.0, 12.0) });
        writer.AddPage(816, 1056, new[] { ("p2", 96.0, 96.0, 12.0) });
        var bytes = writer.ToBytes();

        using var ms = new MemoryStream(bytes);
        var reader = new XpsReader();
        var pages = reader.Read(ms);

        Assert.Equal(2, pages.Count);
    }

    [Fact]
    [DisplayName("XP01-02 Read 提取 Glyphs 文本")]
    public void Read_ExtractsGlyphsText()
    {
        var bytes = BuildMinimalXps("Hello XPS World");
        using var ms = new MemoryStream(bytes);
        var reader = new XpsReader();
        var pages = reader.Read(ms);

        Assert.Single(pages);
        Assert.Contains("Hello XPS World", pages[0].Text);
    }

    [Fact]
    [DisplayName("XP01-02 多 Glyphs 文本全部提取")]
    public void Read_MultipleGlyphs_AllExtracted()
    {
        var writer = new XpsWriter();
        writer.AddPage(816, 1056, new[]
        {
            ("Line One", 96.0, 96.0, 12.0),
            ("Line Two", 96.0, 120.0, 12.0),
            ("Line Three", 96.0, 144.0, 12.0)
        });
        var bytes = writer.ToBytes();

        using var ms = new MemoryStream(bytes);
        var pages = new XpsReader().Read(ms);

        Assert.Single(pages);
        Assert.Contains("Line One", pages[0].Text);
        Assert.Contains("Line Two", pages[0].Text);
        Assert.Equal(3, pages[0].Glyphs.Count);
    }

    [Fact]
    [DisplayName("XP01-04 Read 提取页面尺寸")]
    public void Read_ExtractsPageDimensions()
    {
        var bytes = BuildMinimalXps(pageW: 595, pageH: 842); // A4
        using var ms = new MemoryStream(bytes);
        var pages = new XpsReader().Read(ms);

        Assert.Single(pages);
        Assert.Equal(595, pages[0].Width, 1);
        Assert.Equal(842, pages[0].Height, 1);
    }

    [Fact]
    [DisplayName("XP01-01 Read 元数据可正确读取")]
    public void ReadProperties_ReturnsTitle()
    {
        var bytes = BuildMinimalXps(title: "My Readable Title");
        using var ms = new MemoryStream(bytes);
        var props = new XpsReader().ReadProperties(ms);
        Assert.Equal("My Readable Title", props.Title);
        Assert.Equal("TestSuite", props.Creator);
    }

    [Fact]
    [DisplayName("XP02-03 AddImage 在 ZIP 中嵌入图片")]
    public void AddImage_EmbeddedInZip()
    {
        var writer = new XpsWriter();
        writer.AddPage(816, 1056, new[] { ("text", 96.0, 96.0, 12.0) });
        writer.AddImage("logo.png", new Byte[] { 0x89, 0x50, 0x4E, 0x47 });
        var bytes = writer.ToBytes();

        using var ms = new MemoryStream(bytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        Assert.NotNull(zip.GetEntry("Resources/Images/logo.png"));
    }

    [Fact]
    [DisplayName("XP01-03 ExtractImages 提取嵌入图片")]
    public void ExtractImages_ReturnsEmbeddedImages()
    {
        var imgData = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // PNG hdr
        var writer = new XpsWriter();
        writer.AddPage(816, 1056, new[] { ("text", 96.0, 96.0, 12.0) });
        writer.AddImage("test.png", imgData);
        var bytes = writer.ToBytes();

        using var ms = new MemoryStream(bytes);
        var images = new XpsReader().ExtractImages(ms).ToList();

        Assert.Single(images);
        Assert.Equal("Resources/Images/test.png", images[0].Path);
        Assert.Equal(imgData, images[0].Data);
    }

    [Fact]
    [DisplayName("XP01 往返测试：写入再读取文本一致")]
    public void RoundTrip_TextIsPreserved()
    {
        const String text = "Round-trip XPS content test";
        var bytes = BuildMinimalXps(text);
        using var ms = new MemoryStream(bytes);
        var pages = new XpsReader().Read(ms);

        Assert.Single(pages);
        Assert.Equal(text, pages[0].Text);
    }
}
