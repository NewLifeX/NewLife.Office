using System.ComponentModel;
using System.IO.Compression;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>EPUB 电子书格式读写测试</summary>
public class EpubTests
{
    #region 辅助

    private static Byte[] BuildEpub(EpubDocument doc)
    {
        var ms = new MemoryStream();
        new EpubWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static EpubDocument ReadEpub(Byte[] data)
    {
        using var ms = new MemoryStream(data);
        return new EpubReader().Read(ms);
    }

    private static EpubDocument MakeDoc()
    {
        var doc = new EpubDocument
        {
            Title = "测试书籍",
            Author = "张三",
            Language = "zh-CN",
            Publisher = "新生命出版社",
            Description = "这是一本测试书籍",
            Identifier = "test-epub-001",
        };
        doc.Chapters.Add(new EpubChapter { Title = "第一章", Content = "<p>第一章内容。</p>", FileName = "chapter01.xhtml" });
        doc.Chapters.Add(new EpubChapter { Title = "第二章", Content = "<p>第二章内容。</p>", FileName = "chapter02.xhtml" });
        return doc;
    }

    #endregion

    #region 写入测试

    [Fact]
    [DisplayName("写入的 EPUB 是有效的 ZIP 文件")]
    public void Write_ValidZip()
    {
        var data = BuildEpub(MakeDoc());
        Assert.True(data.Length > 0);

        // ZIP 魔数 PK
        Assert.Equal(0x50, data[0]);
        Assert.Equal(0x4B, data[1]);
    }

    [Fact]
    [DisplayName("写入的 EPUB 包含必要文件")]
    public void Write_ContainsRequiredFiles()
    {
        var data = BuildEpub(MakeDoc());
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var names = new HashSet<String>(zip.Entries.Select(e => e.FullName));
        Assert.Contains("mimetype", names);
        Assert.Contains("META-INF/container.xml", names);
        Assert.Contains("OEBPS/content.opf", names);
        Assert.Contains("OEBPS/nav.xhtml", names);
        Assert.Contains("OEBPS/chapter01.xhtml", names);
        Assert.Contains("OEBPS/chapter02.xhtml", names);
    }

    [Fact]
    [DisplayName("mimetype 必须不压缩")]
    public void Write_MimetypeNotCompressed()
    {
        var data = BuildEpub(MakeDoc());
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var mt = zip.GetEntry("mimetype");
        Assert.NotNull(mt);
        // 未压缩时 CompressedLength == Length
        Assert.Equal(mt!.Length, mt.CompressedLength);

        using var s = mt.Open();
        using var sr = new StreamReader(s);
        Assert.Equal("application/epub+zip", sr.ReadToEnd());
    }

    [Fact]
    [DisplayName("OPF 包含书名和作者")]
    public void Write_OpfContainsMetadata()
    {
        var data = BuildEpub(MakeDoc());
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var opf = zip.GetEntry("OEBPS/content.opf");
        Assert.NotNull(opf);
        using var s = opf!.Open();
        using var sr = new StreamReader(s);
        var content = sr.ReadToEnd();

        Assert.Contains("测试书籍", content);
        Assert.Contains("张三", content);
        Assert.Contains("zh-CN", content);
        Assert.Contains("新生命出版社", content);
    }

    [Fact]
    [DisplayName("nav.xhtml 包含章节目录")]
    public void Write_NavContainsChapters()
    {
        var data = BuildEpub(MakeDoc());
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var nav = zip.GetEntry("OEBPS/nav.xhtml");
        Assert.NotNull(nav);
        using var s = nav!.Open();
        using var sr = new StreamReader(s);
        var content = sr.ReadToEnd();

        Assert.Contains("第一章", content);
        Assert.Contains("第二章", content);
    }

    [Fact]
    [DisplayName("写入封面图片")]
    public void Write_WithCover_CoverEntryExists()
    {
        var doc = MakeDoc();
        doc.Cover = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // PNG 魔数
        doc.CoverMediaType = "image/png";

        var data = BuildEpub(doc);
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var names = new HashSet<String>(zip.Entries.Select(e => e.FullName));
        Assert.Contains("OEBPS/cover.png", names);
        Assert.Contains("OEBPS/cover.xhtml", names);
    }

    [Fact]
    [DisplayName("章节内容包含标题和正文")]
    public void Write_ChapterContent_WrappedInXhtml()
    {
        var data = BuildEpub(MakeDoc());
        using var ms = new MemoryStream(data);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);

        var ch = zip.GetEntry("OEBPS/chapter01.xhtml");
        Assert.NotNull(ch);
        using var s = ch!.Open();
        using var sr = new StreamReader(s);
        var content = sr.ReadToEnd();

        Assert.Contains("<h1>", content);
        Assert.Contains("第一章", content);
        Assert.Contains("第一章内容", content);
    }

    #endregion

    #region 读取测试

    [Fact]
    [DisplayName("读取 EPUB 元数据")]
    public void Read_ParsesMetadata()
    {
        var data = BuildEpub(MakeDoc());
        var parsed = ReadEpub(data);

        Assert.Equal("测试书籍", parsed.Title);
        Assert.Equal("张三", parsed.Author);
        Assert.Equal("zh-CN", parsed.Language);
        Assert.Equal("新生命出版社", parsed.Publisher);
    }

    [Fact]
    [DisplayName("读取所有章节")]
    public void Read_ParsesChapters()
    {
        var data = BuildEpub(MakeDoc());
        var parsed = ReadEpub(data);

        // nav.xhtml 也在 spine 里
        Assert.True(parsed.Chapters.Count >= 2);
    }

    #endregion

    #region 往返和集成测试

    [Fact]
    [DisplayName("往返测试：写入后可被读取")]
    public void RoundTrip_WriteAndRead()
    {
        var original = MakeDoc();
        var data = BuildEpub(original);
        var parsed = ReadEpub(data);

        Assert.Equal(original.Title, parsed.Title);
        Assert.Equal(original.Author, parsed.Author);
    }

    [Fact]
    [DisplayName("集成：写入 epub 文件并读取")]
    public void Integration_WriteFile_ThenReadFile()
    {
        var dir = Path.Combine("Bin", "UnitTest", "Artifacts");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "test_output.epub");

        var doc = new EpubDocument
        {
            Title = "新生命框架使用指南",
            Author = "NewLife 团队",
            Language = "zh-CN",
            Publisher = "新生命出版社",
            Description = "NewLife.Office 集成测试生成",
            Identifier = "newlife-epub-integration-001",
        };
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第一章：快速入门",
            Content = "<p>本章介绍如何快速上手 NewLife 框架。</p>",
            FileName = "chapter01.xhtml",
        });
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第二章：深入配置",
            Content = "<p>本章介绍各种配置选项。</p>",
            FileName = "chapter02.xhtml",
        });
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第三章：高级功能",
            Content = "<p>本章介绍高级功能，包括扩展插件体系。</p>",
            FileName = "chapter03.xhtml",
        });

        new EpubWriter().Write(doc, path);
        Assert.True(File.Exists(path));
        Assert.True(new FileInfo(path).Length > 500);

        var parsed = new EpubReader().Read(path);
        Assert.Equal("新生命框架使用指南", parsed.Title);
        Assert.Equal("NewLife 团队", parsed.Author);
        Assert.True(parsed.Chapters.Count >= 3);
    }

    #endregion
}
