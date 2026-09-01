using SkiaSharp;

namespace NewLife.Office.Rendering.Imaging;

/// <summary>图片格式枚举</summary>
public enum ImageFormat
{
    /// <summary>PNG 格式</summary>
    Png,

    /// <summary>JPEG 格式</summary>
    Jpeg,

    /// <summary>WebP 格式</summary>
    Webp,

    /// <summary>BMP 格式</summary>
    Bmp
}

/// <summary>图片处理辅助类</summary>
/// <remarks>
/// 基于 SkiaSharp 提供图片缩放、裁剪、旋转、格式转换、水印叠加等功能。
/// </remarks>
public static class ImageHelper
{
    #region 缩放
    /// <summary>缩放图片到指定尺寸</summary>
    /// <param name="imageData">源图片字节</param>
    /// <param name="width">目标宽度（像素）</param>
    /// <param name="height">目标高度（像素）</param>
    /// <param name="keepAspectRatio">是否保持宽高比（默认 true）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <param name="quality">JPEG/WebP 质量（0-100，默认 85）</param>
    /// <returns>缩放后的图片字节</returns>
    /// <exception cref="ArgumentNullException">imageData 为 null</exception>
    public static Byte[] Resize(Byte[] imageData, Int32 width, Int32 height,
        Boolean keepAspectRatio = true, ImageFormat format = ImageFormat.Png, Int32 quality = 85)
    {
        if (imageData == null) throw new ArgumentNullException(nameof(imageData));

        using var src = DecodeOrThrow(imageData);

        Int32 dstW, dstH;
        if (keepAspectRatio)
        {
            var scale = Math.Min((Single)width / src.Width, (Single)height / src.Height);
            dstW = (Int32)(src.Width * scale);
            dstH = (Int32)(src.Height * scale);
        }
        else
        {
            dstW = width;
            dstH = height;
        }

        using var dst = src.Resize(new SKSizeI(dstW, dstH), new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
        if (dst == null) throw new InvalidOperationException("缩放失败");

        return EncodeToFormat(dst, format, quality);
    }
    #endregion

    #region 格式转换
    /// <summary>转换图片格式</summary>
    /// <param name="imageData">源图片字节</param>
    /// <param name="format">目标格式</param>
    /// <param name="quality">JPEG/WebP 质量（0-100，默认 85）</param>
    /// <returns>转换后的图片字节</returns>
    public static Byte[] Convert(Byte[] imageData, ImageFormat format, Int32 quality = 85)
    {
        if (imageData == null) throw new ArgumentNullException(nameof(imageData));

        using var src = DecodeOrThrow(imageData);

        return EncodeToFormat(src, format, quality);
    }
    #endregion

    #region 裁剪
    /// <summary>裁剪图片</summary>
    /// <param name="imageData">源图片字节</param>
    /// <param name="x">裁剪区域左上角 X</param>
    /// <param name="y">裁剪区域左上角 Y</param>
    /// <param name="width">裁剪区域宽度</param>
    /// <param name="height">裁剪区域高度</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>裁剪后的图片字节</returns>
    public static Byte[] Crop(Byte[] imageData, Int32 x, Int32 y, Int32 width, Int32 height,
        ImageFormat format = ImageFormat.Png)
    {
        if (imageData == null) throw new ArgumentNullException(nameof(imageData));

        using var src = DecodeOrThrow(imageData);

        var rect = new SKRectI(x, y, x + width, y + height);
        using var dst = new SKBitmap(width, height);
        src.ExtractSubset(dst, rect);

        return EncodeToFormat(dst, format);
    }
    #endregion

    #region 旋转
    /// <summary>旋转图片</summary>
    /// <param name="imageData">源图片字节</param>
    /// <param name="degrees">旋转角度（顺时针）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>旋转后的图片字节</returns>
    public static Byte[] Rotate(Byte[] imageData, Single degrees, ImageFormat format = ImageFormat.Png)
    {
        if (imageData == null) throw new ArgumentNullException(nameof(imageData));

        using var src = DecodeOrThrow(imageData);

        var radians = degrees * (Single)Math.PI / 180f;
        var cos = Math.Abs((Single)Math.Cos(radians));
        var sin = Math.Abs((Single)Math.Sin(radians));
        var dstW = (Int32)(src.Width * cos + src.Height * sin);
        var dstH = (Int32)(src.Width * sin + src.Height * cos);

        using var surface = SKSurface.Create(new SKImageInfo(dstW, dstH));
        var canvas = surface.Canvas;
        canvas.Translate(dstW / 2f, dstH / 2f);
        canvas.RotateDegrees(degrees);
        canvas.Translate(-src.Width / 2f, -src.Height / 2f);
        canvas.DrawBitmap(src, 0, 0);

        using var img = surface.Snapshot();
        using var dst = SKBitmap.FromImage(img);
        return EncodeToFormat(dst!, format);
    }
    #endregion

    #region 水印
    /// <summary>添加文字水印</summary>
    /// <param name="imageData">源图片字节</param>
    /// <param name="text">水印文字</param>
    /// <param name="fontSize">字号（默认 24）</param>
    /// <param name="opacity">透明度（0-1，默认 0.3）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>带水印的图片字节</returns>
    public static Byte[] AddTextWatermark(Byte[] imageData, String text,
        Single fontSize = 24, Single opacity = 0.3f, ImageFormat format = ImageFormat.Png)
    {
        if (imageData == null) throw new ArgumentNullException(nameof(imageData));

        using var src = DecodeOrThrow(imageData);

        using var surface = SKSurface.Create(new SKImageInfo(src.Width, src.Height));
        var canvas = surface.Canvas;
        canvas.DrawBitmap(src, 0, 0);

        using var font = new SKFont(SKTypeface.Default, fontSize, 1, 0);
        using var paint = new SKPaint
        {
            Color = SKColors.White.WithAlpha((Byte)(opacity * 255)),
            IsAntialias = true,
            Style = SKPaintStyle.Fill,
        };

        // 文字居中
        var textWidth = font.MeasureText(text, paint);
        var x = (src.Width - textWidth) / 2f;
        var y = src.Height / 2f + fontSize / 3f;

        canvas.DrawText(text, x, y, font, paint);

        using var img = surface.Snapshot();
        using var dst = SKBitmap.FromImage(img);
        return EncodeToFormat(dst!, format);
    }

    /// <summary>叠加图片水印</summary>
    /// <param name="background">背景图片字节</param>
    /// <param name="overlay">叠加图片字节</param>
    /// <param name="x">叠加位置 X</param>
    /// <param name="y">叠加位置 Y</param>
    /// <param name="opacity">叠加图片透明度（0-1，默认 0.5）</param>
    /// <param name="format">输出格式（默认 PNG）</param>
    /// <returns>合成后的图片字节</returns>
    public static Byte[] OverlayImage(Byte[] background, Byte[] overlay,
        Int32 x, Int32 y, Single opacity = 0.5f, ImageFormat format = ImageFormat.Png)
    {
        if (background == null) throw new ArgumentNullException(nameof(background));
        if (overlay == null) throw new ArgumentNullException(nameof(overlay));

        using var bg = DecodeOrThrow(background);
        using var ov = DecodeOrThrow(overlay);

        using var surface = SKSurface.Create(new SKImageInfo(bg.Width, bg.Height));
        var canvas = surface.Canvas;
        canvas.DrawBitmap(bg, 0, 0);

        using var paint = new SKPaint { Color = SKColors.White.WithAlpha((Byte)(opacity * 255)) };
        canvas.DrawBitmap(ov, x, y, paint);

        using var img = surface.Snapshot();
        using var dst = SKBitmap.FromImage(img);
        return EncodeToFormat(dst!, format);
    }
    #endregion

    #region 辅助
    /// <summary>解码图片字节，失败抛 InvalidOperationException（SkiaSharp 3.x 对无效数据抛异常而非返回 null）</summary>
    private static SKBitmap DecodeOrThrow(Byte[] imageData)
    {
        try
        {
            var bmp = SKBitmap.Decode(imageData);
            return bmp ?? throw new InvalidOperationException("无法解码图片数据");
        }
        catch (InvalidOperationException) { throw; }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法解码图片数据", ex);
        }
    }

    /// <summary>手写 BMP 编码（Skia 默认不含 BMP 编码器，Encode(SKEncodedImageFormat.Bmp) 返回 null）</summary>
    /// <param name="bitmap">源位图</param>
    /// <returns>BMP 格式图片字节</returns>
    internal static Byte[] EncodeBmp(SKBitmap bitmap)
    {
        var width = bitmap.Width;
        var height = bitmap.Height;
        if (width <= 0 || height <= 0) throw new InvalidOperationException("无效的图片尺寸");

        // 每行 4 字节对齐（BGR 24bit）
        var rowSize = (width * 3 + 3) / 4 * 4;
        var pixelSize = rowSize * height;
        var fileSize = 14 + 40 + pixelSize;

        var pixels = bitmap.Pixels;
        using var ms = new MemoryStream(fileSize);
        using var bw = new BinaryWriter(ms);

        // BITMAPFILEHEADER（14 字节）
        bw.Write((Byte)'B');
        bw.Write((Byte)'M');
        bw.Write(fileSize);
        bw.Write((Int16)0);
        bw.Write((Int16)0);
        bw.Write(14 + 40);

        // BITMAPINFOHEADER（40 字节）
        bw.Write(40);
        bw.Write(width);
        bw.Write(height);
        bw.Write((Int16)1);   // 平面数
        bw.Write((Int16)24);  // 每像素位数
        bw.Write(0);          // BI_RGB 不压缩
        bw.Write(pixelSize);
        bw.Write(2835);       // 水平分辨率 96DPI
        bw.Write(2835);       // 垂直分辨率 96DPI
        bw.Write(0);          // 颜色表数
        bw.Write(0);          // 重要颜色数

        // 像素数据（BMP 为 bottom-up，BGR 顺序）
        var row = new Byte[rowSize];
        for (var y = height - 1; y >= 0; y--)
        {
            var baseIdx = y * width;
            for (var x = 0; x < width; x++)
            {
                var c = pixels[baseIdx + x];
                row[x * 3] = c.Blue;
                row[x * 3 + 1] = c.Green;
                row[x * 3 + 2] = c.Red;
            }
            bw.Write(row, 0, rowSize);
        }

        bw.Flush();
        return ms.ToArray();
    }

    private static Byte[] EncodeToFormat(SKBitmap bitmap, ImageFormat format, Int32 quality = 85)
    {
        if (format == ImageFormat.Bmp)
            return EncodeBmp(bitmap);

        using var img = SKImage.FromBitmap(bitmap);
        var encodeFormat = format switch
        {
            ImageFormat.Jpeg => SKEncodedImageFormat.Jpeg,
            ImageFormat.Webp => SKEncodedImageFormat.Webp,
            _ => SKEncodedImageFormat.Png,
        };
        var data = img.Encode(encodeFormat, format == ImageFormat.Jpeg || format == ImageFormat.Webp ? quality : 100);

        return data.ToArray();
    }
    #endregion
}
