using System.Text;

namespace NewLife.Office.Word;

/// <summary>图片元信息探测：格式识别、像素尺寸与 DPI（用于按自然尺寸插入图片）</summary>
/// <remarks>
/// 仅支持常见位图容器（PNG/JPEG/GIF/BMP）。未识别时调用方回退默认尺寸。
/// DPI 缺失时按 96 处理（Windows/Word 默认），与 python-docx 的 72 不同：96 更贴近实际所见大小。
/// </remarks>
public static class ImageInfo
{
    /// <summary>默认 DPI（图片未声明时使用）</summary>
    public const Double DefaultDpi = 96;

    #region 格式识别
    /// <summary>按文件头识别图片格式</summary>
    /// <param name="data">图片字节</param>
    /// <returns>扩展名（png/jpeg/gif/bmp/svg），无法识别返回 null</returns>
    public static String? DetectExtension(Byte[] data)
    {
        if (data == null || data.Length < 4) return null;

        if (data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return "png";
        if (data[0] == 0xFF && data[1] == 0xD8) return "jpeg";
        if (data[0] == 'G' && data[1] == 'I' && data[2] == 'F') return "gif";
        if (data[0] == 'B' && data[1] == 'M') return "bmp";

        // SVG 是文本，检查前缀里是否出现 <svg
        var head = Encoding.UTF8.GetString(data, 0, Math.Min(data.Length, 256));
        if (head.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0) return "svg";

        return null;
    }
    #endregion

    #region 尺寸探测
    /// <summary>探测图片像素尺寸与 DPI</summary>
    /// <param name="data">图片字节</param>
    /// <param name="width">像素宽</param>
    /// <param name="height">像素高</param>
    /// <param name="dpi">水平 DPI（缺失按 <see cref="DefaultDpi"/>）</param>
    /// <returns>识别成功返回 true</returns>
    public static Boolean TryGetPixelSize(Byte[] data, out Int32 width, out Int32 height, out Double dpi)
    {
        width = 0;
        height = 0;
        dpi = DefaultDpi;
        // 最短的位图容器（GIF 头部）也需要 10 字节，各格式解码器再做自身长度校验
        if (data == null || data.Length < 10) return false;

        try
        {
            return DetectExtension(data) switch
            {
                "png" => TryPng(data, ref width, ref height, ref dpi),
                "jpeg" => TryJpeg(data, ref width, ref height, ref dpi),
                "gif" => TryGif(data, ref width, ref height),
                "bmp" => TryBmp(data, ref width, ref height),
                _ => false,
            };
        }
        catch
        {
            // 图片损坏或结构异常：交由调用方回退默认尺寸
            return false;
        }
    }

    private static Boolean TryPng(Byte[] data, ref Int32 width, ref Int32 height, ref Double dpi)
    {
        // 8 字节签名 + 4 字节长度 + 4 字节 "IHDR" → 宽高各 4 字节大端
        if (data.Length < 24) return false;
        width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
        height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];

        // 可选 pHYs 块：x 像素/米（unit=1 表示米制）
        for (var i = 8; i + 12 < data.Length; i++)
        {
            if (data[i] != (Byte)'p' || data[i + 1] != (Byte)'H' || data[i + 2] != (Byte)'Y' || data[i + 3] != (Byte)'s') continue;
            var unit = data[i + 12];
            if (unit != 1) break;
            var ppm = (data[i + 4] << 24) | (data[i + 5] << 16) | (data[i + 6] << 8) | data[i + 7];
            if (ppm > 0) dpi = Math.Round(ppm * 0.0254, 1);
            break;
        }
        return width > 0 && height > 0;
    }

    private static Boolean TryJpeg(Byte[] data, ref Int32 width, ref Int32 height, ref Double dpi)
    {
        var pos = 2;
        while (pos + 4 < data.Length)
        {
            if (data[pos] != 0xFF) { pos++; continue; }

            var marker = data[pos + 1];
            if (marker == 0xD8 || marker == 0x01 || marker is >= 0xD0 and <= 0xD7) { pos += 2; continue; }
            if (marker == 0xD9) break;

            var segLen = (data[pos + 2] << 8) | data[pos + 3];
            if (segLen < 2) break;

            // SOF0/1/2/3/5/6/7/9/A/B/D/E/F：段内 1 字节精度 + 2 字节高 + 2 字节宽
            if (marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF)
            {
                if (pos + 9 < data.Length)
                {
                    height = (data[pos + 5] << 8) | data[pos + 6];
                    width = (data[pos + 7] << 8) | data[pos + 8];
                }
                return width > 0 && height > 0;
            }

            // APP0 JFIF：xdensity/ydensity 在段内偏移 9/11（1 = 每英寸）
            if (marker == 0xE0 && segLen >= 12 && pos + 13 < data.Length
                && data[pos + 4] == (Byte)'J' && data[pos + 5] == (Byte)'F' && data[pos + 6] == (Byte)'I' && data[pos + 7] == (Byte)'F')
            {
                var units = data[pos + 11];
                var xd = (data[pos + 12] << 8) | data[pos + 13];
                if (units == 1 && xd > 0) dpi = xd;
            }
            pos += 2 + segLen;
        }
        return width > 0 && height > 0;
    }

    private static Boolean TryGif(Byte[] data, ref Int32 width, ref Int32 height)
    {
        if (data.Length < 10) return false;
        width = data[6] | (data[7] << 8);
        height = data[8] | (data[9] << 8);
        return width > 0 && height > 0;
    }

    private static Boolean TryBmp(Byte[] data, ref Int32 width, ref Int32 height)
    {
        if (data.Length < 26) return false;
        width = data[18] | (data[19] << 8) | (data[20] << 16) | (data[21] << 24);
        var h = data[22] | (data[23] << 8) | (data[24] << 16) | (data[25] << 24);
        height = Math.Abs(h);
        return width > 0 && height > 0;
    }
    #endregion
}
