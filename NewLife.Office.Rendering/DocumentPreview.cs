using NewLife.Office.Rendering.Pdf;
using NewLife.Office.Rendering.Word;
using NewLife.Office.Rendering.Ppt;
using NewLife.Office.Rendering.Markdown;

namespace NewLife.Office.Rendering;

/// <summary>综合文档预览渲染器</summary>
/// <remarks>
/// 自动检测文件格式，路由到对应的渲染器。
/// 支持 PDF、DOCX、PPTX、Markdown 格式。
/// </remarks>
/// <example>
/// <code>
/// var preview = DocumentPreview.RenderFirstPage("document.pdf", 150);
/// File.WriteAllBytes("preview.png", preview);
/// </code>
/// </example>
public static class DocumentPreview
{
    /// <summary>自动检测格式并渲染第一页为 PNG</summary>
    /// <param name="filePath">文档文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>PNG 图片字节</returns>
    public static Byte[] RenderFirstPage(String filePath, Int32 dpi = 150)
    {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".pdf" => PdfRenderer.RenderPage(filePath, 0, dpi),
            ".docx" => WordRenderer.RenderPage(filePath, 0, dpi),
            ".pptx" => PptRenderer.RenderSlide(filePath, 0, dpi),
            ".md" => MdRenderer.RenderFileToImage(filePath, dpi),
            ".markdown" => MdRenderer.RenderFileToImage(filePath, dpi),
            _ => throw new NotSupportedException($"不支持的文件格式: {ext}"),
        };
    }

    /// <summary>自动检测格式并渲染所有页为 PNG 序列</summary>
    /// <param name="filePath">文档文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <returns>各页 PNG 图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllPages(String filePath, Int32 dpi = 150)
    {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        switch (ext)
        {
            case ".pdf": return PdfRenderer.RenderAllPages(filePath, dpi);
            case ".docx": return WordRenderer.RenderAllPages(filePath, dpi);
            case ".pptx": return PptRenderer.RenderAllSlides(filePath, dpi);
            default: throw new NotSupportedException($"不支持的文件格式: {ext}");
        }
    }
}
