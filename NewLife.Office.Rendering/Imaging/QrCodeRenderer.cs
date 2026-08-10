using SkiaSharp;
using NewLife.Office.Pdf;

namespace NewLife.Office.Rendering.Imaging;

/// <summary>二维码渲染器</summary>
/// <remarks>
/// 基于 NewLife.Office 内置的 QR 码生成引擎（纯 C#），
/// 使用 SkiaSharp 提供更灵活的尺寸控制和格式输出。
/// </remarks>
public static class QrCodeRenderer
{
    /// <summary>生成 QR 码 PNG 图片</summary>
    /// <param name="content">要编码的内容（URL 或短文本）</param>
    /// <param name="moduleSize">每模块像素数（默认 4，推荐 2-8）</param>
    /// <returns>PNG 格式图片字节</returns>
    public static Byte[] Render(String content, Int32 moduleSize = 4)
    {
        // 直接复用内置 QR 码生成器（已输出 PNG）
        return PdfQRCode.Generate(content, moduleSize);
    }

    /// <summary>生成 QR 码图片（指定格式）</summary>
    /// <param name="content">要编码的内容</param>
    /// <param name="size">输出图片边长（像素），null 则自动计算</param>
    /// <param name="format">输出格式</param>
    /// <param name="foreColor">前景色，默认黑色</param>
    /// <param name="backColor">背景色，默认白色</param>
    /// <returns>图片字节</returns>
    public static Byte[] RenderStyled(String content, Int32? size = null,
        ImageFormat format = ImageFormat.Png, SKColor? foreColor = null, SKColor? backColor = null)
    {
        // 先用原生引擎生成标准 PNG，再解码后用 SkiaSharp 重绘
        var pngBytes = PdfQRCode.Generate(content, 1);

        using var src = SKBitmap.Decode(pngBytes);
        if (src == null) throw new InvalidOperationException("QR 码生成失败");

        var moduleCount = src.Width; // 1:1 模块对应像素
        var dstSize = size ?? moduleCount * 4;

        using var surface = SKSurface.Create(new SKImageInfo(dstSize, dstSize));
        var canvas = surface.Canvas;

        // 背景
        canvas.Clear(backColor ?? SKColors.White);

        // 按模块绘制
        var fg = foreColor ?? SKColors.Black;
        var modulePixelSize = (Single)dstSize / moduleCount;
        for (var y = 0; y < moduleCount; y++)
        {
            for (var x = 0; x < moduleCount; x++)
            {
                var pixel = src.GetPixel(x, y);
                if (pixel.Red < 128) // 暗色 = 前景
                {
                    using var paint = new SKPaint { Color = fg, IsAntialias = false };
                    canvas.DrawRect(x * modulePixelSize, y * modulePixelSize,
                        modulePixelSize + 1, modulePixelSize + 1, paint);
                }
            }
        }

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
        return img.Encode(encodeFormat, 100).ToArray();
    }
}
