using SkiaSharp;

namespace NewLife.Office.Rendering.Imaging;

/// <summary>文字渲染器</summary>
/// <remarks>将文字渲染为 PNG 图片，支持多行、自动换行、字体指定。</remarks>
public static class TextRenderer
{
    /// <summary>将文字渲染为 PNG 图片</summary>
    /// <param name="text">要渲染的文字</param>
    /// <param name="fontFamily">字体名称，null 使用系统默认</param>
    /// <param name="fontSize">字号（默认 16）</param>
    /// <param name="width">图片宽度（像素，默认 800）</param>
    /// <param name="foreColor">文字颜色，默认黑色</param>
    /// <param name="backColor">背景颜色，null 为透明</param>
    /// <param name="padding">内边距（像素，默认 20）</param>
    /// <returns>PNG 格式图片字节</returns>
    public static Byte[] Render(String text, String? fontFamily = null, Single fontSize = 16,
        Int32 width = 800, SKColor? foreColor = null, SKColor? backColor = null, Int32 padding = 20)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        using var typeface = GetTypeface(fontFamily);
        using var font = new SKFont(typeface, fontSize, 1, 0);
        using var paint = new SKPaint
        {
            IsAntialias = true,
            Color = foreColor ?? SKColors.Black,
        };

        // 计算行高和文字高度
        var fontMetrics = font.Metrics;
        var lineHeight = fontMetrics.Descent - fontMetrics.Ascent + fontMetrics.Leading;
        var contentWidth = width - padding * 2;

        // 自动换行：逐字符累加宽度，超出则换行
        var lines = WrapText(text, font, paint, contentWidth);

        var totalHeight = (Int32)(lines.Count * lineHeight + padding * 2 + 10);

        // 创建画布
        using var surface = SKSurface.Create(new SKImageInfo(width, totalHeight));
        var canvas = surface.Canvas;

        // 背景
        if (backColor != null)
            canvas.Clear(backColor.Value);

        // 绘制文字
        var y = padding + fontMetrics.Ascent;
        foreach (var line in lines)
        {
            canvas.DrawText(line, padding, -y, font, paint);
            y += lineHeight;
        }

        using var img = surface.Snapshot();
        using var bitmap = SKBitmap.FromImage(img);
        var result = img.Encode(SKEncodedImageFormat.Png, 100);
        return result.ToArray();
    }

    #region 辅助
    private static SKTypeface GetTypeface(String? fontFamily)
    {
        if (String.IsNullOrEmpty(fontFamily))
            return SKTypeface.Default;

        var tf = SKTypeface.FromFamilyName(fontFamily);
        if (tf != null)
            return tf;

        // 回退：尝试常见中文字体
        var fallbackFonts = new[] { "Microsoft YaHei", "SimSun", "WenQuanYi Micro Hei", "Noto Sans CJK SC", "sans-serif" };
        foreach (var name in fallbackFonts)
        {
            tf = SKTypeface.FromFamilyName(name);
            if (tf != null)
                return tf;
        }

        return SKTypeface.Default;
    }

    private static List<String> WrapText(String text, SKFont font, SKPaint paint, Single maxWidth)
    {
        var lines = new List<String>();
        var currentLine = String.Empty;

        foreach (var ch in text)
        {
            if (ch == '\n')
            {
                lines.Add(currentLine);
                currentLine = String.Empty;
                continue;
            }

            var testLine = currentLine + ch;
            if (font.MeasureText(testLine, paint) > maxWidth && currentLine.Length > 0)
            {
                lines.Add(currentLine);
                currentLine = ch.ToString();
            }
            else
            {
                currentLine = testLine;
            }
        }

        if (currentLine.Length > 0)
            lines.Add(currentLine);

        if (lines.Count == 0)
            lines.Add(String.Empty);

        return lines;
    }
    #endregion
}
