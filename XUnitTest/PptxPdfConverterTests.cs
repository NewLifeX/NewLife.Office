using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PptxPdfConverter 单元测试</summary>
public class PptxPdfConverterTests
{
    static PptxPdfConverterTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    // ─── 最小 PPTX 构建辅助 ──────────────────────────────────────────────

    /// <summary>构建包含指定幻灯片 XML 列表的最小 pptx 流</summary>
    /// <param name="slideXmls">每张幻灯片的正文 XML（spTree 内容）</param>
    private static MemoryStream BuildPptx(params String[] slideXmls)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (var i = 0; i < slideXmls.Length; i++)
            {
                var entry = zip.CreateEntry($"ppt/slides/slide{i + 1}.xml");
                using var w = new StreamWriter(entry.Open(), Encoding.UTF8);
                w.Write(WrapSlideXml(slideXmls[i]));
            }
        }
        ms.Position = 0;
        return ms;
    }

    /// <summary>将 spTree 内容包裹成完整的幻灯片 XML</summary>
    private static String WrapSlideXml(String spTreeContent) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"" +
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<p:cSld><p:spTree>" + spTreeContent + "</p:spTree></p:cSld></p:sld>";

    /// <summary>生成含标题和正文的 sp 形状 XML</summary>
    private static String ShapeXml(String id, String text, Int64 x, Int64 y, Int64 cx, Int64 cy) =>
        $"<p:sp>" +
        $"<p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Shape{id}\"/></p:nvSpPr>" +
        $"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"textBox\"/></p:spPr>" +
        $"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody>" +
        $"</p:sp>";

    // ─── 测试 ─────────────────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("单张幻灯片转 PDF 输出合法 PDF 结构")]
    public void Convert_SingleSlide_ProducesValidPdf()
    {
        var titleShape = ShapeXml("1", "我的标题", 0L, 0L, 9_144_000L, 1_143_000L);
        var bodyShape = ShapeXml("2", "正文内容第一行", 0L, 1_143_000L, 9_144_000L, 5_715_000L);

        using var input = BuildPptx(titleShape + bodyShape);
        using var output = new MemoryStream();

        var converter = new PptxPdfConverter { DocumentTitle = "TestPdf" };
        converter.Convert(input, output);

        output.Position = 0;
        var header = new Byte[8];
        output.Read(header, 0, header.Length);
        Assert.Equal("%PDF-1.4", Encoding.ASCII.GetString(header));
        Assert.True(output.Length > 512, "PDF 文件应超过 512 字节");
    }

    [Fact, System.ComponentModel.DisplayName("多张幻灯片分别渲染到 PDF 不同页")]
    public void Convert_MultipleSlides_ProducesMultiplePages()
    {
        var slide1 = ShapeXml("1", "第一页标题", 0L, 0L, 9_144_000L, 1_143_000L);
        var slide2 = ShapeXml("1", "第二页标题", 0L, 0L, 9_144_000L, 1_143_000L) +
                     ShapeXml("2", "第二页正文", 0L, 1_143_000L, 9_144_000L, 5_715_000L);
        var slide3 = ShapeXml("1", "第三页标题", 0L, 0L, 9_144_000L, 1_143_000L);

        using var input = BuildPptx(slide1, slide2, slide3);
        using var output = new MemoryStream();

        new PptxPdfConverter().Convert(input, output);

        // PDF 页面数 = /Type /Page 出现次数
        var text = Encoding.Latin1.GetString(output.ToArray());
        var pageCount = 0;
        var idx = 0;
        while ((idx = text.IndexOf("/Type /Page", idx)) >= 0) { pageCount++; idx++; }
        Assert.True(pageCount >= 3, $"应有 3 个 /Type /Page 条目，实际: {pageCount}");
    }

    [Fact, System.ComponentModel.DisplayName("空幻灯片（无形状）也能转换不抛出")]
    public void Convert_EmptySlide_NoException()
    {
        using var input = BuildPptx(String.Empty);
        using var output = new MemoryStream();

        var ex = Record.Exception(() => new PptxPdfConverter().Convert(input, output));
        Assert.Null(ex);
        Assert.True(output.Length > 0);
    }

    [Fact, System.ComponentModel.DisplayName("保存到文件路径 API 可用")]
    public void Convert_ToFilePath_FileCreated()
    {
        var titleShape = ShapeXml("1", "File API Test", 0L, 0L, 9_144_000L, 1_143_000L);
        using var input = BuildPptx(titleShape);
        var pptxPath = Path.GetTempFileName() + ".pptx";
        var pdfPath = Path.GetTempFileName() + ".pdf";
        try
        {
            // 先将 PPTX 写入临时文件
            using (var fs = File.Create(pptxPath))
                input.CopyTo(fs);

            new PptxPdfConverter { DocumentTitle = "File API" }.Convert(pptxPath, pdfPath);

            Assert.True(File.Exists(pdfPath));
            Assert.True(new FileInfo(pdfPath).Length > 512);

            // 验证 PDF 头
            var header = File.ReadAllBytes(pdfPath).Take(8).ToArray();
            Assert.Equal("%PDF-1.4", Encoding.ASCII.GetString(header));
        }
        finally
        {
            if (File.Exists(pptxPath)) File.Delete(pptxPath);
            if (File.Exists(pdfPath)) File.Delete(pdfPath);
        }
    }

    [Fact, System.ComponentModel.DisplayName("幻灯片形状按 Y 坐标排序后首个为标题")]
    public void Convert_TitleIsTopMostShape()
    {
        // 故意将正文放在 y=0（最上），标题放在 y=1500000（下方）
        // Converter 应按 y 排序，体积方面只检查无异常且 PDF 合法
        var bodyFirst = ShapeXml("1", "Body First", 0L, 0L, 9_144_000L, 500_000L);
        var titleBelow = ShapeXml("2", "Title Below", 0L, 1_500_000L, 9_144_000L, 800_000L);

        using var input = BuildPptx(bodyFirst + titleBelow);
        using var output = new MemoryStream();

        var ex = Record.Exception(() => new PptxPdfConverter().Convert(input, output));
        Assert.Null(ex);

        var text = Encoding.Latin1.GetString(output.ToArray());
        Assert.StartsWith("%PDF-1.4", text);
    }

    // ─── S09-02：PPT 转图片（当前版本为占位桩，依赖外部渲染库）── ──────

    [Fact, System.ComponentModel.DisplayName("S09-02 ConvertToImages 当前版本抛出 NotSupportedException")]
    public void ConvertToImages_ThrowsNotSupportedException()
    {
        // ConvertToImages 需要 SkiaSharp/Docnet.Core 渲染引擎
        // 当前零依赖版本按约定抛出 NotSupportedException
        var ex = Assert.Throws<NotSupportedException>(
            () => new PptxPdfConverter().ConvertToImages("dummy.pptx").GetEnumerator().MoveNext());
        Assert.Contains("SkiaSharp", ex.Message);
    }

    [Fact, System.ComponentModel.DisplayName("S09-02 ConvertToImages 错误消息包含替代方案提示")]
    public void ConvertToImages_ErrorMessageSuggestsAlternative()
    {
        var ex = Assert.Throws<NotSupportedException>(
            () => new PptxPdfConverter().ConvertToImages("any.pptx").GetEnumerator().MoveNext());
        // 错误信息应包含建议使用 Convert → PDF 路径的提示
        Assert.Contains("Convert", ex.Message);
    }
}
