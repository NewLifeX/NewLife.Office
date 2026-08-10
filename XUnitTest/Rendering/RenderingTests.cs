using System.ComponentModel;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Rendering;
using NewLife.Office.Rendering.Imaging;
using NewLife.Office.Rendering.Markdown;
using NewLife.Office.Rendering.Pdf;
using NewLife.Office.Rendering.Ppt;
using NewLife.Office.Rendering.Word;
using NewLife.Office.Word;
using SkiaSharp;
using Xunit;

using XUnitTest.Common;

namespace XUnitTest.Rendering;

/// <summary>Rendering 渲染扩展包测试（REN 模块）</summary>
/// <remarks>
/// 覆盖 REN01-REN13 全部功能点：PDF/DOCX/PPTX/Markdown → 图片渲染、图片处理工具、文字/二维码 → 图片、缩略图网格、综合文档预览。
/// 测试文档复用主库 PdfWriter/WordWriter/PptxWriter 程序化生成，不依赖外部文件。
/// </remarks>
public class RenderingTests : IntegrationTestBase
{
    #region 辅助
    /// <summary>生成指定页数的测试 PDF 字节</summary>
    private static Byte[] CreateTestPdf(Int32 pages)
    {
        using var ms = new MemoryStream();
        using (var writer = new PdfWriter())
        {
            for (var i = 0; i < pages; i++)
            {
                writer.BeginPage();
                writer.DrawText($"Page {i + 1}", 56, 780, 16);
                writer.DrawText("NewLife.Office Rendering Test", 56, 750, 12);
                writer.EndPage();
            }
            writer.Save(ms);
        }
        return ms.ToArray();
    }

    /// <summary>生成测试 PDF 文件路径</summary>
    private static String CreateTestPdfFile(Int32 pages)
    {
        var path = Path.Combine(OutputDir, $"ren_pdf_{Guid.NewGuid():N}.pdf");
        File.WriteAllBytes(path, CreateTestPdf(pages));
        return path;
    }

    /// <summary>生成测试 DOCX 字节</summary>
    private static Byte[] CreateTestDocx()
    {
        using var ms = new MemoryStream();
        using (var w = new WordWriter())
        {
            w.DocumentProperties.Title = "渲染测试";
            w.AppendHeading("渲染测试", 1);
            w.AppendParagraph("Word 渲染扩展包测试段落。");
            w.AppendParagraph("第二段：表格与列表。", ParagraphStyle.Normal, new RunProperties { Bold = true, FontSize = 14f });
            w.Save(ms);
        }
        return ms.ToArray();
    }

    /// <summary>生成测试 DOCX 文件路径</summary>
    private static String CreateTestDocxFile()
    {
        var path = Path.Combine(OutputDir, $"ren_docx_{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, CreateTestDocx());
        return path;
    }

    /// <summary>生成测试 PPTX 字节</summary>
    private static Byte[] CreateTestPptx()
    {
        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.AddSlide();
            writer.AddTextBox(0, "渲染测试", 2, 2, 8, 3);
            writer.Save(ms);
        }
        return ms.ToArray();
    }

    /// <summary>生成测试 PPTX 文件路径</summary>
    private static String CreateTestPptxFile()
    {
        var path = Path.Combine(OutputDir, $"ren_pptx_{Guid.NewGuid():N}.pptx");
        File.WriteAllBytes(path, CreateTestPptx());
        return path;
    }

    /// <summary>生成测试 Markdown 文件路径</summary>
    private static String CreateTestMdFile()
    {
        var path = Path.Combine(OutputDir, $"ren_md_{Guid.NewGuid():N}.md");
        File.WriteAllText(path, "# 渲染测试\n\nMarkdown 渲染扩展包测试段落。\n\n- 列表项一\n- 列表项二\n");
        return path;
    }

    /// <summary>生成指定尺寸的测试 PNG 图片字节</summary>
    private static Byte[] CreateTestImage(Int32 width, Int32 height)
    {
        using var bmp = new SKBitmap(width, height);
        using (var canvas = new SKCanvas(bmp))
        {
            canvas.Clear(SKColors.LightGray);
            using var paint = new SKPaint { Color = SKColors.Blue };
            canvas.DrawRect(0, 0, width / 2f, height / 2f, paint);
            using var paint2 = new SKPaint { Color = SKColors.Red };
            canvas.DrawCircle(width * 0.75f, height * 0.75f, Math.Min(width, height) * 0.2f, paint2);
        }
        using var img = SKImage.FromBitmap(bmp);
        return img.Encode(SKEncodedImageFormat.Png, 100).ToArray();
    }

    /// <summary>验证 PNG magic bytes</summary>
    private static void AssertPng(Byte[] data)
    {
        Assert.NotNull(data);
        Assert.True(data.Length > 8, "PNG 数据过短");
        Assert.Equal(0x89, data[0]);
        Assert.Equal(0x50, data[1]); // P
        Assert.Equal(0x4E, data[2]); // N
        Assert.Equal(0x47, data[3]); // G
    }

    /// <summary>验证 JPEG magic bytes</summary>
    private static void AssertJpeg(Byte[] data)
    {
        Assert.NotNull(data);
        Assert.True(data.Length > 3, "JPEG 数据过短");
        Assert.Equal(0xFF, data[0]);
        Assert.Equal(0xD8, data[1]);
        Assert.Equal(0xFF, data[2]);
    }

    /// <summary>验证 WebP magic bytes（RIFF....WEBP）</summary>
    private static void AssertWebp(Byte[] data)
    {
        Assert.NotNull(data);
        Assert.True(data.Length > 12, "WebP 数据过短");
        Assert.Equal(0x52, data[0]); // R
        Assert.Equal(0x49, data[1]); // I
        Assert.Equal(0x46, data[2]); // F
        Assert.Equal(0x46, data[3]); // F
        Assert.Equal(0x57, data[8]); // W
        Assert.Equal(0x45, data[9]); // E
        Assert.Equal(0x42, data[10]); // B
        Assert.Equal(0x50, data[11]); // P
    }

    /// <summary>验证 BMP magic bytes（BM）</summary>
    private static void AssertBmp(Byte[] data)
    {
        Assert.NotNull(data);
        Assert.True(data.Length > 2, "BMP 数据过短");
        Assert.Equal(0x42, data[0]); // B
        Assert.Equal(0x4D, data[1]); // M
    }
    #endregion

    #region REN01 PDF → 图片
    [Fact, DisplayName("REN01-01 PDF 单页渲染为 PNG（路径/流/字节三入口）")]
    public void Pdf_RenderSinglePage()
    {
        var pdf = CreateTestPdf(2);

        // 字节入口
        AssertPng(PdfRenderer.RenderPage(pdf, 0, 96));
        AssertPng(PdfRenderer.RenderPage(pdf, 1, 96));

        // 流入口
        using var ms = new MemoryStream(pdf);
        AssertPng(PdfRenderer.RenderPage(ms, 0, 96));

        // 路径入口
        var path = CreateTestPdfFile(1);
        AssertPng(PdfRenderer.RenderPage(path, 0, 96));
    }

    [Fact, DisplayName("REN01-02 PDF 全页渲染为图片序列")]
    public void Pdf_RenderAllPages()
    {
        var pdf = CreateTestPdf(3);
        var pages = PdfRenderer.RenderAllPages(pdf, 96).ToList();
        Assert.Equal(3, pages.Count);
        Assert.All(pages, AssertPng);
    }

    [Fact, DisplayName("REN01-03 PDF 多格式输出（PNG/JPEG/WebP/BMP）")]
    public void Pdf_RenderMultiFormat()
    {
        var pdf = CreateTestPdf(1);
        AssertPng(PdfRenderer.RenderPage(pdf, 0, 96, ImageFormat.Png));
        AssertJpeg(PdfRenderer.RenderPage(pdf, 0, 96, ImageFormat.Jpeg, 80));
        AssertWebp(PdfRenderer.RenderPage(pdf, 0, 96, ImageFormat.Webp, 80));
        AssertBmp(PdfRenderer.RenderPage(pdf, 0, 96, ImageFormat.Bmp));
    }
    #endregion

    #region REN02 DOCX → 图片
    [Fact, DisplayName("REN02-01 DOCX 单页渲染为 PNG（路径/流双入口）")]
    public void Word_RenderSinglePage()
    {
        var docx = CreateTestDocx();

        // 流入口
        using (var ms = new MemoryStream(docx))
        {
            AssertPng(WordRenderer.RenderPage(ms, 0, 96));
        }

        // 路径入口
        var path = CreateTestDocxFile();
        AssertPng(WordRenderer.RenderPage(path, 0, 96));
    }

    [Fact, DisplayName("REN02-02 DOCX 全页渲染为 PNG 序列")]
    public void Word_RenderAllPages()
    {
        var path = CreateTestDocxFile();
        var pages = WordRenderer.RenderAllPages(path, 96).ToList();
        Assert.NotEmpty(pages);
        Assert.All(pages, AssertPng);
    }
    #endregion

    #region REN03 PPTX → 图片
    [Fact, DisplayName("REN03-01 PPTX 幻灯片渲染为 PNG（路径/流双入口）")]
    public void Ppt_RenderSingleSlide()
    {
        var pptx = CreateTestPptx();

        // 流入口
        using (var ms = new MemoryStream(pptx))
        {
            AssertPng(PptRenderer.RenderSlide(ms, 0, 96));
        }

        // 路径入口
        var path = CreateTestPptxFile();
        AssertPng(PptRenderer.RenderSlide(path, 0, 96));
    }

    [Fact, DisplayName("REN03-02 PPTX 全幻灯片渲染为 PNG 序列")]
    public void Ppt_RenderAllSlides()
    {
        var path = CreateTestPptxFile();
        var pages = PptRenderer.RenderAllSlides(path, 96).ToList();
        Assert.NotEmpty(pages);
        Assert.All(pages, AssertPng);
    }
    #endregion

    #region REN04 Markdown → 图片
    [Fact, DisplayName("REN04-01 Markdown 文本渲染为 PNG")]
    public void Md_RenderText()
    {
        var png = MdRenderer.RenderToImage("# Hello\n\nThis is a paragraph.\n\n- item 1\n- item 2\n", 96);
        AssertPng(png);
    }

    [Fact, DisplayName("REN04-02 Markdown 文件渲染为 PNG")]
    public void Md_RenderFile()
    {
        var path = CreateTestMdFile();
        var png = MdRenderer.RenderFileToImage(path, 96);
        AssertPng(png);
    }
    #endregion

    #region REN05-REN09 图片处理工具
    [Fact, DisplayName("REN05-01 图片缩放（保持宽高比）")]
    public void Image_Resize()
    {
        var img = CreateTestImage(100, 50);

        // 保持宽高比：100x50 → 50x25
        var resized = ImageHelper.Resize(img, 50, 25);
        using (var bmp = SKBitmap.Decode(resized))
        {
            Assert.NotNull(bmp);
            Assert.Equal(50, bmp.Width);
            Assert.Equal(25, bmp.Height);
        }

        // 不保持宽高比：100x50 → 40x30（严格拉伸）
        var stretched = ImageHelper.Resize(img, 40, 30, keepAspectRatio: false);
        using (var bmp = SKBitmap.Decode(stretched))
        {
            Assert.NotNull(bmp);
            Assert.Equal(40, bmp.Width);
            Assert.Equal(30, bmp.Height);
        }
    }

    [Fact, DisplayName("REN06-01 图片格式转换（PNG→JPEG/WebP/BMP）")]
    public void Image_ConvertFormat()
    {
        var img = CreateTestImage(20, 20);
        AssertJpeg(ImageHelper.Convert(img, ImageFormat.Jpeg));
        AssertWebp(ImageHelper.Convert(img, ImageFormat.Webp));
        AssertBmp(ImageHelper.Convert(img, ImageFormat.Bmp));
        AssertPng(ImageHelper.Convert(img, ImageFormat.Png));
    }

    [Fact, DisplayName("REN07-01 图片裁剪与旋转")]
    public void Image_CropAndRotate()
    {
        var img = CreateTestImage(100, 50);

        // 裁剪：100x50 中裁 40x20
        var cropped = ImageHelper.Crop(img, 10, 10, 40, 20);
        using (var bmp = SKBitmap.Decode(cropped))
        {
            Assert.NotNull(bmp);
            Assert.Equal(40, bmp.Width);
            Assert.Equal(20, bmp.Height);
        }

        // 旋转 90°：宽高交换
        var rotated = ImageHelper.Rotate(img, 90);
        using (var bmp = SKBitmap.Decode(rotated))
        {
            Assert.NotNull(bmp);
            Assert.Equal(50, bmp.Width);
            Assert.Equal(100, bmp.Height);
        }
    }

    [Fact, DisplayName("REN08-01 文字水印（透明度+居中）")]
    public void Image_TextWatermark()
    {
        var img = CreateTestImage(120, 60);
        var watermarked = ImageHelper.AddTextWatermark(img, "内部资料", 18, 0.5f);
        AssertPng(watermarked);

        // 水印后尺寸不变
        using var bmp = SKBitmap.Decode(watermarked);
        Assert.NotNull(bmp);
        Assert.Equal(120, bmp.Width);
        Assert.Equal(60, bmp.Height);
    }

    [Fact, DisplayName("REN09-01 图片叠加（透明度+位置）")]
    public void Image_Overlay()
    {
        var bg = CreateTestImage(100, 100);
        var ov = CreateTestImage(30, 30);
        var result = ImageHelper.OverlayImage(bg, ov, 10, 10, 0.5f);
        AssertPng(result);
    }
    #endregion

    #region REN10-REN11 文字/二维码 → 图片
    [Fact, DisplayName("REN10-01 文字渲染为图片（多行+自动换行）")]
    public void Text_RenderToImage()
    {
        var png = TextRenderer.Render("Hello World\n第二行中文内容", fontSize: 20, width: 400, backColor: SKColors.White);
        AssertPng(png);
        using var bmp = SKBitmap.Decode(png);
        Assert.NotNull(bmp);
        Assert.Equal(400, bmp.Width);
        Assert.True(bmp.Height >= 60, "多行文字高度应足够");
    }

    [Fact, DisplayName("REN11-01 二维码渲染（默认+定制颜色/尺寸/格式）")]
    public void QrCode_Render()
    {
        // 默认渲染
        var png = QrCodeRenderer.Render("https://newlifex.com", 4);
        AssertPng(png);

        // 定制：指定尺寸 + 前景色 + JPEG 格式
        var styled = QrCodeRenderer.RenderStyled("https://newlifex.com", size: 200, format: ImageFormat.Png, foreColor: SKColors.Blue, backColor: SKColors.White);
        AssertPng(styled);
        using var bmp = SKBitmap.Decode(styled);
        Assert.NotNull(bmp);
        Assert.Equal(200, bmp.Width);
        Assert.Equal(200, bmp.Height);
    }
    #endregion

    #region REN12-REN13 缩略图与综合预览
    [Fact, DisplayName("REN12-01 PDF 缩略图网格（多页拼图+页码）")]
    public void Pdf_RenderThumbnails()
    {
        var path = CreateTestPdfFile(4);
        var thumb = PdfRenderer.RenderThumbnails(path, cols: 2, thumbWidth: 100, dpi: 72);
        AssertPng(thumb);

        // 网格尺寸：2 列 × 2 行（4 页）
        using var bmp = SKBitmap.Decode(thumb);
        Assert.NotNull(bmp);
        Assert.True(bmp.Width > 200, "网格宽度应容纳两列");
        Assert.True(bmp.Height > 150, "网格高度应容纳两行");
    }

    [Fact, DisplayName("REN13-01 综合文档预览（按扩展名路由）")]
    public void Preview_RenderFirstPage()
    {
        // PDF
        var pdfPath = CreateTestPdfFile(1);
        AssertPng(DocumentPreview.RenderFirstPage(pdfPath, 96));

        // DOCX
        var docxPath = CreateTestDocxFile();
        AssertPng(DocumentPreview.RenderFirstPage(docxPath, 96));

        // PPTX
        var pptxPath = CreateTestPptxFile();
        AssertPng(DocumentPreview.RenderFirstPage(pptxPath, 96));

        // Markdown
        var mdPath = CreateTestMdFile();
        AssertPng(DocumentPreview.RenderFirstPage(mdPath, 96));

        // 不支持格式
        Assert.Throws<NotSupportedException>(() => DocumentPreview.RenderFirstPage(Path.Combine(OutputDir, "unknown.txt")));
    }
    #endregion

    #region 边界与异常
    [Fact, DisplayName("REN-边界：null 参数抛出 ArgumentNullException")]
    public void Renderer_NullArguments()
    {
        Assert.Throws<ArgumentNullException>(() => PdfRenderer.RenderPage((String)null!));
        Assert.Throws<ArgumentNullException>(() => PdfRenderer.RenderPage((Byte[])null!));
        // yield 迭代器惰性执行，需枚举才抛出参数校验异常
        Assert.Throws<ArgumentNullException>(() => PdfRenderer.RenderAllPages((Byte[])null!).ToList());
        Assert.Throws<ArgumentNullException>(() => WordRenderer.RenderPage((String)null!));
        Assert.Throws<ArgumentNullException>(() => PptRenderer.RenderSlide((String)null!));
        Assert.Throws<ArgumentNullException>(() => MdRenderer.RenderToImage(null!));
        Assert.Throws<ArgumentNullException>(() => MdRenderer.RenderFileToImage(null!));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.Resize(null!, 10, 10));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.Convert(null!, ImageFormat.Png));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.Crop(null!, 0, 0, 10, 10));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.Rotate(null!, 90));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.AddTextWatermark(null!, "水印"));
        Assert.Throws<ArgumentNullException>(() => ImageHelper.OverlayImage(null!, new Byte[] { 1 }, 0, 0));
        Assert.Throws<ArgumentNullException>(() => TextRenderer.Render(null!));
        Assert.Throws<ArgumentNullException>(() => DocumentPreview.RenderFirstPage(null!));
    }

    [Fact, DisplayName("REN-边界：PDF 页面索引越界抛出 ArgumentOutOfRangeException")]
    public void Pdf_RenderInvalidPageIndex()
    {
        var pdf = CreateTestPdf(2);
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfRenderer.RenderPage(pdf, 5, 96));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfRenderer.RenderPage(pdf, -1, 96));
    }

    [Fact, DisplayName("REN-异常：无效图片数据抛出 InvalidOperationException")]
    public void Image_InvalidData()
    {
        var invalid = new Byte[] { 0x01, 0x02, 0x03, 0x04 };
        Assert.Throws<InvalidOperationException>(() => ImageHelper.Resize(invalid, 10, 10));
        Assert.Throws<InvalidOperationException>(() => ImageHelper.Convert(invalid, ImageFormat.Png));
    }
    #endregion
}
