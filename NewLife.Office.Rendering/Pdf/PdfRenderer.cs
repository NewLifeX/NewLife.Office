using NewLife.Office.Pdf;
using NewLife.Office.Rendering.Imaging;
using SkiaSharp;

namespace NewLife.Office.Rendering.Pdf;

/// <summary>PDF 页面渲染器（公共 API）</summary>
/// <remarks>
/// 将 PDF 页面渲染为 PNG/JPEG/WebP/BMP 等格式的图片。
/// 基于 SkiaSharp 实现内容流解析和 2D 图形渲染，无需 PDFium 等外部原生库。
/// <para>目前支持的 PDF 元素：纯文字、线段/矩形/简单路径、嵌入图片、纯色填充。</para>
/// <para>尚未支持：渐变填充、透明度组、复杂裁剪路径。</para>
/// </remarks>
/// <example>
/// <code>
/// // 渲染单页
/// var png = PdfRenderer.RenderPage("document.pdf", 0, 150);
/// File.WriteAllBytes("page0.png", png);
///
/// // 渲染所有页
/// var pages = PdfRenderer.RenderAllPages("document.pdf", 150);
/// var i = 0;
/// foreach (var page in pages)
///     File.WriteAllBytes($"page_{i++}.png", page);
/// </code>
/// </example>
public static class PdfRenderer
{
    #region 单页渲染
    /// <summary>将 PDF 指定页面渲染为 PNG 图片</summary>
    /// <param name="pdfPath">PDF 文件路径</param>
    /// <param name="pageIndex">页面索引（0-based，默认第一页）</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <param name="quality">JPEG/WebP 质量（0-100，默认 85）</param>
    /// <returns>图片字节</returns>
    public static Byte[] RenderPage(String pdfPath, Int32 pageIndex = 0, Int32 dpi = 150,
        ImageFormat format = ImageFormat.Png, Int32 quality = 85)
    {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));
        using var fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return RenderPage(fs, pageIndex, dpi, format, quality);
    }

    /// <summary>将 PDF 指定页面渲染为图片</summary>
    /// <param name="pdfStream">PDF 数据流</param>
    /// <param name="pageIndex">页面索引（0-based，默认第一页）</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <param name="quality">JPEG/WebP 质量（0-100，默认 85）</param>
    /// <returns>图片字节</returns>
    public static Byte[] RenderPage(Stream pdfStream, Int32 pageIndex = 0, Int32 dpi = 150,
        ImageFormat format = ImageFormat.Png, Int32 quality = 85)
    {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        Byte[] pdfBytes;
        if (pdfStream is MemoryStream ms)
            pdfBytes = ms.ToArray();
        else
        {
            using var tmpMs = new MemoryStream();
            pdfStream.CopyTo(tmpMs);
            pdfBytes = tmpMs.ToArray();
        }
        return RenderPageInternal(pdfBytes, pageIndex, dpi, format, quality);
    }

    /// <summary>从 PDF 字节数组渲染指定页面</summary>
    public static Byte[] RenderPage(Byte[] pdfBytes, Int32 pageIndex = 0, Int32 dpi = 150,
        ImageFormat format = ImageFormat.Png, Int32 quality = 85)
    {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
        return RenderPageInternal(pdfBytes, pageIndex, dpi, format, quality);
    }

    private static Byte[] RenderPageInternal(Byte[] pdfBytes, Int32 pageIndex, Int32 dpi,
        ImageFormat format, Int32 quality)
    {
        var xref = new PdfXRefTable(pdfBytes);
        // 通过 PdfReader 简单获取页数
        var pageCount = GetPdfPageCount(pdfBytes);
        if (pageIndex < 0 || pageIndex >= pageCount)
            throw new ArgumentOutOfRangeException(nameof(pageIndex), $"页面索引 {pageIndex} 超出范围 [0, {pageCount - 1}]");

        // PDF 页面尺寸（A4 默认：595pt × 842pt）
        // 转换为像素：pt * dpi / 72
        var width = (Int32)(595 * dpi / 72.0);
        var height = (Int32)(842 * dpi / 72.0);

        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;

        var pageRenderer = new PdfPageRenderer();
        pageRenderer.Render(pdfBytes, xref, pageIndex, canvas, width, height, dpi);

        // 编码输出（BMP 手写编码，Skia 默认不含 BMP 编码器）
        using var img = surface.Snapshot();
        if (format == ImageFormat.Bmp)
        {
            using var bmp = SKBitmap.FromImage(img);
            return ImageHelper.EncodeBmp(bmp);
        }
        var encodeFormat = format switch
        {
            ImageFormat.Jpeg => SKEncodedImageFormat.Jpeg,
            ImageFormat.Webp => SKEncodedImageFormat.Webp,
            _ => SKEncodedImageFormat.Png,
        };
        var data = img.Encode(encodeFormat, quality);
        return data.ToArray();
    }
    #endregion

    #region 全部页面渲染
    /// <summary>将 PDF 所有页面渲染为 PNG 图片</summary>
    /// <param name="pdfPath">PDF 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>各页图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllPages(String pdfPath, Int32 dpi = 150,
        ImageFormat format = ImageFormat.Png)
    {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));

        var pdfBytes = File.ReadAllBytes(pdfPath);
        return RenderAllPages(pdfBytes, dpi, format);
    }

    /// <summary>将 PDF 所有页面渲染为 PNG 图片（从字节数组）</summary>
    /// <param name="pdfBytes">PDF 字节数组</param>
    /// <param name="dpi">输出分辨率（DPI，默认 150）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>各页图片字节的序列</returns>
    public static IEnumerable<Byte[]> RenderAllPages(Byte[] pdfBytes, Int32 dpi = 150,
        ImageFormat format = ImageFormat.Png)
    {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
        var pageCount = GetPdfPageCount(pdfBytes);
        for (var i = 0; i < pageCount; i++)
            yield return RenderPageInternal(pdfBytes, i, dpi, format, 85);
    }
    #endregion

    #region 缩略图
    /// <summary>生成 PDF 缩略图网格</summary>
    /// <param name="pdfPath">PDF 文件路径</param>
    /// <param name="cols">每行列数（默认 4）</param>
    /// <param name="thumbWidth">缩略图宽度（像素，默认 200）</param>
    /// <param name="dpi">渲染分辨率（DPI，默认 72）</param>
    /// <returns>缩略图网格 PNG 字节</returns>
    public static Byte[] RenderThumbnails(String pdfPath, Int32 cols = 4, Int32 thumbWidth = 200, Int32 dpi = 72)
    {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));

        var pdfBytes = File.ReadAllBytes(pdfPath);
        var pageCount = GetPdfPageCount(pdfBytes);
        if (pageCount == 0) return Array.Empty<Byte>();

        // 计算缩略图高度（A4 比例：842/595 ≈ 0.707）
        var thumbHeight = (Int32)(thumbWidth * 842.0 / 595);

        var rows = (Int32)Math.Ceiling((Single)pageCount / cols);
        var totalWidth = cols * thumbWidth + (cols + 1) * 4;
        var totalHeight = rows * thumbHeight + (rows + 1) * 4;

        using var surface = SKSurface.Create(new SKImageInfo(totalWidth, totalHeight));
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.LightGray);

        var pageRenderer = new PdfPageRenderer();
        for (var i = 0; i < pageCount; i++)
        {
            var row = i / cols;
            var col = i % cols;
            var x = col * thumbWidth + (col + 1) * 4;
            var y = row * thumbHeight + (row + 1) * 4;

            // 保存画布状态，裁剪到缩略图区域
            canvas.Save();
            canvas.ClipRect(new SKRect(x, y, x + thumbWidth, y + thumbHeight));
            canvas.Translate(x, y);
            canvas.Scale(thumbWidth / (595f * dpi / 72f));

            try
            {
                var xref = new PdfXRefTable(pdfBytes);
                pageRenderer.Render(pdfBytes, xref, i, canvas, thumbWidth, thumbHeight, dpi);
            }
            catch
            {
                // 渲染失败，画灰色占位
                using var failPaint = new SKPaint { Color = SKColors.DarkGray };
                canvas.DrawRect(0, 0, thumbWidth, thumbHeight, failPaint);
            }

            canvas.Restore();

            // 画页面编号
            using var textFont = new SKFont(SKTypeface.Default, 11, 1, 0);
            using var textPaint = new SKPaint
            {
                Color = SKColors.White,
                IsAntialias = true,
            };
            canvas.DrawText($"{i + 1}", x + 4, y + 14, textFont, textPaint);
        }

        using var img = surface.Snapshot();
        return img.Encode(SKEncodedImageFormat.Png, 100).ToArray();
    }
    #endregion

    #region 辅助
    /// <summary>获取 PDF 页数</summary>
    private static Int32 GetPdfPageCount(Byte[] data)
    {
        var latin1 = System.Text.Encoding.GetEncoding(28591);
        var text = latin1.GetString(data);
        var countIdx = -1;
        var pos = 0;
        while (true)
        {
            var idx = text.IndexOf("/Count", pos, StringComparison.Ordinal);
            if (idx < 0) break;
            // 确保 /Count 后面是数字
            var after = idx + 6;
            while (after < text.Length && (text[after] == ' ' || text[after] == '\r' || text[after] == '\n')) after++;
            if (after < text.Length && Char.IsDigit(text[after]))
            {
                countIdx = after;
                break;
            }
            pos = idx + 6;
        }
        if (countIdx < 0) return 0;
        var end = countIdx;
        while (end < text.Length && Char.IsDigit(text[end])) end++;
        return Int32.TryParse(text.Substring(countIdx, end - countIdx), out var count) ? count : 0;
    }
    #endregion
}
