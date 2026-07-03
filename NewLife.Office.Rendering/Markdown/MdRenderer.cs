using NewLife.Office.Markdown;
using NewLife.Office.Rendering.Pdf;

namespace NewLife.Office.Rendering.Markdown;

/// <summary>Markdown 文档渲染器</summary>
/// <remarks>
/// 将 Markdown 文本渲染为图片。内部先将 Markdown 转换为 PDF，再渲染 PDF 页面。
/// </remarks>
/// <example>
/// <code>
/// var md = "# Hello\n\nThis is a paragraph.";
/// var png = MdRenderer.RenderToImage(md, dpi: 150);
/// File.WriteAllBytes("output.png", png);
/// </code>
/// </example>
public static class MdRenderer
{
    /// <summary>将 Markdown 文本渲染为 PNG 图片</summary>
    /// <param name="markdown">Markdown 文本</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderToImage(String markdown, Int32 dpi = 150)
    {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));

        var pdfBytes = ConvertToPdf(markdown);
        return PdfRenderer.RenderPage(pdfBytes, 0, dpi);
    }

    /// <summary>将 Markdown 文件渲染为 PNG 图片</summary>
    /// <param name="mdPath">Markdown 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderFileToImage(String mdPath, Int32 dpi = 150)
    {
        if (mdPath == null) throw new ArgumentNullException(nameof(mdPath));

        var markdown = File.ReadAllText(mdPath);
        return RenderToImage(markdown, dpi);
    }

    /// <summary>将 Markdown 文本渲染为多页 PNG 序列</summary>
    /// <param name="markdown">Markdown 文本</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>各页 PNG 图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllPages(String markdown, Int32 dpi = 150)
    {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));

        var pdfBytes = ConvertToPdf(markdown);
        return PdfRenderer.RenderAllPages(pdfBytes, dpi);
    }

    #region 辅助
    private static Byte[] ConvertToPdf(String markdown)
    {
        var doc = MarkdownDocument.Parse(markdown);
        var converter = new MarkdownPdfConverter();
        return converter.ToBytes(doc);
    }
    #endregion
}
