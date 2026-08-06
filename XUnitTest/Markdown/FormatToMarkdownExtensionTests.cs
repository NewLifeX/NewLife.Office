using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Epub;
using NewLife.Office.Markdown;
using NewLife.Office.Ods;
using NewLife.Office.Rtf;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>转换中枢扩展测试（MD09，对标 anydoc）</summary>
/// <remarks>
/// 覆盖 ODS/RTF/EPUB → Markdown 的转换（From* 专用入口 + Convert 统一入口 + ToDocument 结构化 AST）。
/// </remarks>
public class FormatToMarkdownExtensionTests
{
    #region 辅助
    private static Byte[] BuildOds()
    {
        var writer = new OdsWriter();
        writer.AddSheet("销售表", new[] { new[] { "产品", "销量" }, new[] { "苹果", "10" }, new[] { "香蕉", "20" } });
        var ms = new MemoryStream();
        writer.Save(ms);
        return ms.ToArray();
    }

    private static Byte[] BuildRtf()
    {
        // RTF 字符串 → 字节（RtfDocument.Parse 支持 UTF-8 字节流）
        var rtf = @"{\rtf1\ansi 第一段文本\par 第二段内容\par}";
        return Encoding.UTF8.GetBytes(rtf);
    }

    private static Byte[] BuildEpub()
    {
        var doc = new EpubDocument
        {
            Title = "测试书籍",
            Author = "张三",
            Language = "zh-CN",
        };
        doc.Chapters.Add(new EpubChapter { Title = "第一章", Content = "<p>第一章内容。</p>", FileName = "chapter01.xhtml" });
        doc.Chapters.Add(new EpubChapter { Title = "第二章", Content = "<p>第二章内容。</p>", FileName = "chapter02.xhtml" });
        var ms = new MemoryStream();
        new EpubWriter().Write(doc, ms);
        return ms.ToArray();
    }
    #endregion

    [Fact, DisplayName("ODS → Markdown：表格转换")]
    public void Ods_ToMarkdown()
    {
        var md = FormatToMarkdown.FromOds(new MemoryStream(BuildOds()));
        Assert.NotNull(md);
        Assert.Contains("产品", md);
        Assert.Contains("苹果", md);
        Assert.Contains("香蕉", md);
    }

    [Fact, DisplayName("RTF → Markdown：文本转换")]
    public void Rtf_ToMarkdown()
    {
        var md = FormatToMarkdown.FromRtf(new MemoryStream(BuildRtf()));
        Assert.NotNull(md);
        Assert.Contains("第一段文本", md);
        Assert.Contains("第二段内容", md);
    }

    [Fact, DisplayName("EPUB → Markdown：章节标题保留")]
    public void Epub_ToMarkdown()
    {
        var md = FormatToMarkdown.FromEpub(new MemoryStream(BuildEpub()));
        Assert.NotNull(md);
        Assert.Contains("第一章", md);
        Assert.Contains("第二章", md);
        Assert.Contains("第一章内容", md);
    }

    [Fact, DisplayName("统一入口 Convert：三种扩展名")]
    public void Convert_ThreeExtensions()
    {
        // 用临时文件（Convert 文件重载需要真实路径）
        var dir = Path.Combine(Path.GetTempPath(), "md09_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var odsPath = Path.Combine(dir, "t.ods");
            File.WriteAllBytes(odsPath, BuildOds());
            var odsMd = FormatToMarkdown.Convert(odsPath);
            Assert.Contains("苹果", odsMd);

            var rtfPath = Path.Combine(dir, "t.rtf");
            File.WriteAllBytes(rtfPath, BuildRtf());
            var rtfMd = FormatToMarkdown.Convert(rtfPath);
            Assert.Contains("第一段文本", rtfMd);

            var epubPath = Path.Combine(dir, "t.epub");
            File.WriteAllBytes(epubPath, BuildEpub());
            var epubMd = FormatToMarkdown.Convert(epubPath);
            Assert.Contains("第一章", epubMd);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("ToDocument：RTF 结构化 AST")]
    public void ToDocument_Rtf()
    {
        var dir = Path.Combine(Path.GetTempPath(), "md09b_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var rtfPath = Path.Combine(dir, "t.rtf");
            File.WriteAllBytes(rtfPath, BuildRtf());
            var doc = FormatToMarkdown.ToDocument(rtfPath);
            Assert.NotNull(doc);
            Assert.True(doc!.Blocks.Count > 0);

            // ODS 结构化 AST
            var odsPath = Path.Combine(dir, "t.ods");
            File.WriteAllBytes(odsPath, BuildOds());
            var odsDoc = FormatToMarkdown.ToDocument(odsPath);
            Assert.NotNull(odsDoc);
            Assert.True(odsDoc!.Blocks.Count > 0);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("异常路径：不支持的格式抛类型化错误")]
    public void Convert_Unsupported_Throws()
    {
        var ex = Assert.Throws<ConvertErrorException>(() => FormatToMarkdown.Convert("data.xyz"));
        Assert.Equal(ConvertErrorType.Unsupported, ex.Type);
    }
}
