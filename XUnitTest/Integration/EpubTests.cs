using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>EPUB 格式集成测试</summary>
public class EpubTests : IntegrationTestBase
{
    [Fact, DisplayName("EPUB_复杂写入再读取往返")]
    public void Epub_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.epub");

        var doc = new EpubDocument
        {
            Title = "EPUB集成测试",
            Author = "NewLife Office",
            Language = "zh-CN",
            Publisher = "新生命出版社",
            Description = "这是一本用于集成测试的电子书",
            Identifier = "test-epub-integration-001",
        };
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第一章 引言",
            Content = "<h1>引言</h1><p>欢迎阅读本书。</p><p>本书由 NewLife.Office 自动生成。</p>",
            FileName = "chapter01.xhtml",
        });
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第二章 核心概念",
            Content = "<h1>核心概念</h1><p>本章介绍核心概念。</p><ul><li>概念一</li><li>概念二</li><li>概念三</li></ul>",
            FileName = "chapter02.xhtml",
        });
        doc.Chapters.Add(new EpubChapter
        {
            Title = "第三章 总结",
            Content = "<h1>总结</h1><p>全书完结。感谢阅读。</p>",
            FileName = "chapter03.xhtml",
        });

        new EpubWriter().Write(doc, path);

        Assert.True(File.Exists(path));

        // 读取验证：EpubWriter 可能自动添加目录/封面章节，因此仅验证 >= 3
        var readDoc = new EpubReader().Read(path);
        Assert.Equal("EPUB集成测试", readDoc.Title);
        Assert.Equal("NewLife Office", readDoc.Author);
        Assert.Equal("zh-CN", readDoc.Language);
        Assert.True(readDoc.Chapters.Count >= 3);

        // 按内容验证原始章节存在
        var chapterTitles = readDoc.Chapters.Select(c => c.Title).ToList();
        Assert.Contains("第一章 引言", chapterTitles);
        Assert.True(readDoc.Chapters.Any(c => c.Content != null && c.Content.Contains("欢迎阅读本书")));

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<EpubDocument>(factoryReader);
    }
}
