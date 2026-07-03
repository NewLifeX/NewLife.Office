using NewLife.Office.Word;
using NewLife.Office.Rendering.Pdf;

namespace NewLife.Office.Rendering.Word;

/// <summary>Word 文档渲染器</summary>
/// <remarks>
/// 将 DOCX 页面渲染为图片。内部先将 DOCX 转换为 PDF，再渲染 PDF 页面。
/// 渲染质量受限于 WordPdfConverter 的文本映射精度（非像素级）。
/// </remarks>
/// <example>
/// <code>
/// var png = WordRenderer.RenderPage("document.docx", 0, 150);
/// File.WriteAllBytes("page0.png", png);
/// </code>
/// </example>
public static class WordRenderer
{
    /// <summary>将 DOCX 指定页面渲染为 PNG 图片</summary>
    /// <param name="docxPath">docx 文件路径</param>
    /// <param name="pageIndex">页面索引（0-based，默认第一页）</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderPage(String docxPath, Int32 pageIndex = 0, Int32 dpi = 150)
    {
        if (docxPath == null) throw new ArgumentNullException(nameof(docxPath));

        // DOCX → PDF
        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docxPath);
        if (pdfBytes == null || pdfBytes.Length == 0)
            throw new InvalidOperationException("Word 转 PDF 失败");

        // PDF → Image
        return PdfRenderer.RenderPage(pdfBytes, pageIndex, dpi);
    }

    /// <summary>将 DOCX 指定页面渲染为图片（从流）</summary>
    /// <param name="docxStream">docx 数据流</param>
    /// <param name="pageIndex">页面索引（0-based）</param>
    /// <param name="dpi">输出分辨率（DPI）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderPage(Stream docxStream, Int32 pageIndex = 0, Int32 dpi = 150)
    {
        if (docxStream == null) throw new ArgumentNullException(nameof(docxStream));

        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docxStream);
        if (pdfBytes == null || pdfBytes.Length == 0)
            throw new InvalidOperationException("Word 转 PDF 失败");

        return PdfRenderer.RenderPage(pdfBytes, pageIndex, dpi);
    }

    /// <summary>将 DOCX 所有页面渲染为 PNG 序列</summary>
    /// <param name="docxPath">docx 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>各页 PNG 图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllPages(String docxPath, Int32 dpi = 150)
    {
        if (docxPath == null) throw new ArgumentNullException(nameof(docxPath));

        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docxPath);
        if (pdfBytes == null || pdfBytes.Length == 0)
            throw new InvalidOperationException("Word 转 PDF 失败");

        return PdfRenderer.RenderAllPages(pdfBytes, dpi);
    }
}
