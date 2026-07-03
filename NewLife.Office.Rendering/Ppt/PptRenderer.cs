using NewLife.Office.Ppt;
using NewLife.Office.Rendering.Pdf;

namespace NewLife.Office.Rendering.Ppt;

/// <summary>PowerPoint 幻灯片渲染器</summary>
/// <remarks>
/// 将 PPTX 幻灯片渲染为图片。内部先将 PPTX 转换为 PDF，再渲染 PDF 页面。
/// 渲染质量受限于 PptxPdfConverter 的文本映射精度（非像素级，不含图形/图片的原始渲染）。
/// </remarks>
/// <example>
/// <code>
/// var png = PptRenderer.RenderSlide("presentation.pptx", 0, 150);
/// File.WriteAllBytes("slide0.png", png);
/// </code>
/// </example>
public static class PptRenderer
{
    /// <summary>将 PPTX 指定幻灯片渲染为 PNG 图片</summary>
    /// <param name="pptxPath">pptx 文件路径</param>
    /// <param name="slideIndex">幻灯片索引（0-based，默认第一张）</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderSlide(String pptxPath, Int32 slideIndex = 0, Int32 dpi = 150)
    {
        if (pptxPath == null) throw new ArgumentNullException(nameof(pptxPath));

        var pdfBytes = ConvertToPdf(pptxPath);
        return PdfRenderer.RenderPage(pdfBytes, slideIndex, dpi);
    }

    /// <summary>将 PPTX 指定幻灯片渲染为图片（从流）</summary>
    /// <param name="pptxStream">pptx 数据流</param>
    /// <param name="slideIndex">幻灯片索引（0-based）</param>
    /// <param name="dpi">输出分辨率（DPI）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderSlide(Stream pptxStream, Int32 slideIndex = 0, Int32 dpi = 150)
    {
        if (pptxStream == null) throw new ArgumentNullException(nameof(pptxStream));

        var pdfBytes = ConvertToPdf(pptxStream);
        return PdfRenderer.RenderPage(pdfBytes, slideIndex, dpi);
    }

    /// <summary>将 PPTX 所有幻灯片渲染为 PNG 序列</summary>
    /// <param name="pptxPath">pptx 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>各幻灯片 PNG 图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllSlides(String pptxPath, Int32 dpi = 150)
    {
        if (pptxPath == null) throw new ArgumentNullException(nameof(pptxPath));

        var pdfBytes = ConvertToPdf(pptxPath);
        return PdfRenderer.RenderAllPages(pdfBytes, dpi);
    }

    #region 辅助
    private static Byte[] ConvertToPdf(String path)
    {
        using var fs = File.OpenRead(path);
        return ConvertToPdf(fs);
    }

    private static Byte[] ConvertToPdf(Stream stream)
    {
        using var ms = new MemoryStream();
        var converter = new PptxPdfConverter();
        converter.Convert(stream, ms);
        return ms.ToArray();
    }
    #endregion
}
