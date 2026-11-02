using NewLife.Office.Markdown;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>Markdown 导出选项测试：页眉页脚输出与图片落地（W29-04）</summary>
public class MarkdownExportOptionsTests
{
    #region 辅助
    /// <summary>1x1 透明 PNG</summary>
    private static Byte[] PngBytes() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
            writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static MarkdownDocument? ToMarkdown(Byte[] bytes, MarkdownConverterOptions? options)
    {
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        return reader.ToMarkdownDocument(options);
    }

    private static Document BuildDocWithHeaderFooter()
    {
        var doc = new Document();
        doc.AddParagraph("正文第一段");
        doc.Headers.Add(new Header
        {
            Type = "default",
            Elements = [new Element { Type = ElementType.Paragraph, Paragraph = new Paragraph { Runs = { new Run { Text = "机密页眉" } } } }],
        });
        doc.Footers.Add(new Footer
        {
            Type = "default",
            Elements = [new Element { Type = ElementType.Paragraph, Paragraph = new Paragraph { Runs = { new Run { Text = "第 1 页" } } } }],
        });
        return doc;
    }
    #endregion

    [Fact(DisplayName = "Markdown-开启选项后输出页眉页脚引用块")]
    public void Markdown_IncludeHeaderFooter()
    {
        var md = ToMarkdown(Save(BuildDocWithHeaderFooter()), new MarkdownConverterOptions { IncludeHeadersFooters = true });

        Assert.NotNull(md);
        var text = md!.ToMarkdown();
        Assert.Contains("机密页眉", text);
        Assert.Contains("第 1 页", text);

        // 页眉在最前、页脚在最后
        Assert.Contains("Header:", text);
        Assert.Contains("Footer:", text);
        Assert.True(text.IndexOf("Header:", StringComparison.Ordinal) < text.IndexOf("正文第一段", StringComparison.Ordinal), text);
        Assert.True(text.IndexOf("Footer:", StringComparison.Ordinal) > text.IndexOf("正文第一段", StringComparison.Ordinal), text);
    }

    [Fact(DisplayName = "Markdown-默认不输出页眉页脚")]
    public void Markdown_HeaderFooterExcludedByDefault()
    {
        var md = ToMarkdown(Save(BuildDocWithHeaderFooter()), null);

        Assert.NotNull(md);
        var text = md!.ToMarkdown();
        Assert.DoesNotContain("机密页眉", text);
        Assert.DoesNotContain("Header:", text);
    }

    [Fact(DisplayName = "Markdown-图片落盘后按文件名引用")]
    public void Markdown_ImagesWrittenToDirectory()
    {
        var dir = Path.Combine(Path.GetTempPath(), "nlo-md-" + Guid.NewGuid().ToString("N"));
        try
        {
            var doc = new Document();
            doc.AddParagraph("图片前");
            doc.AddPicture(PngBytes(), "png");
            var bytes = Save(doc);

            var md = ToMarkdown(bytes, new MarkdownConverterOptions
            {
                ImageHandling = MarkdownImageHandling.Reference,
                ImageDirectory = dir,
            });

            Assert.NotNull(md);
            var text = md!.ToMarkdown();
            Assert.Contains("image1.png", text);

            var file = Path.Combine(dir, "image1.png");
            Assert.True(File.Exists(file));
            Assert.Equal(PngBytes().Length, new FileInfo(file).Length);
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
    }

    [Fact(DisplayName = "Markdown-图片策略 Skip 时不输出图片")]
    public void Markdown_ImageSkip()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.AddPicture(PngBytes(), "png");

        var md = ToMarkdown(Save(doc), new MarkdownConverterOptions { ImageHandling = MarkdownImageHandling.Skip });

        Assert.NotNull(md);
        Assert.DoesNotContain("image1.png", md!.ToMarkdown());
    }

    [Fact(DisplayName = "Markdown-未指定选项时保持引用占位（历史行为）")]
    public void Markdown_DefaultKeepsReference()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.AddPicture(PngBytes(), "png");

        var md = ToMarkdown(Save(doc), null);

        Assert.NotNull(md);
        Assert.Contains("image1.png", md!.ToMarkdown());
    }

    [Fact(DisplayName = "Markdown-Embed 且无输出目录时内嵌 data URI")]
    public void Markdown_ImageEmbedAsDataUri()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.AddPicture(PngBytes(), "png");

        var md = ToMarkdown(Save(doc), new MarkdownConverterOptions { ImageHandling = MarkdownImageHandling.Embed });

        Assert.NotNull(md);
        Assert.Contains("data:image/png;base64,", md!.ToMarkdown());
    }
}
