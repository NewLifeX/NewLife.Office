using System.Text;

namespace NewLife.Office;

/// <summary>纯文本字节解码：BOM → 严格 UTF-8 → GBK 兜底</summary>
/// <remarks>
/// 供 Markdown、CSV 等文本格式共用，保证同一份字节在不同入口得到一致的解码结果，
/// 避免中文 Windows 导出的 GBK 文本出现乱码。
/// </remarks>
internal static class OfficeTextDecoder
{
    #region 探测
    /// <summary>探测字节样本的编码</summary>
    /// <param name="sample">字节样本，可截断在字符中间</param>
    /// <param name="bomLength">样本开头的 BOM 字节数；无 BOM 时为 0</param>
    /// <returns>编码；样本为空时返回 UTF-8</returns>
    public static Encoding DetectEncoding(Byte[] sample, out Int32 bomLength)
    {
        bomLength = 0;
        if (sample == null || sample.Length == 0) return Encoding.UTF8;

        if (sample.Length >= 3 && sample[0] == 0xEF && sample[1] == 0xBB && sample[2] == 0xBF)
        {
            bomLength = 3;
            return Encoding.UTF8;
        }
        if (sample.Length >= 2 && sample[0] == 0xFF && sample[1] == 0xFE)
        {
            bomLength = 2;
            return Encoding.Unicode;
        }
        if (sample.Length >= 2 && sample[0] == 0xFE && sample[1] == 0xFF)
        {
            bomLength = 2;
            return Encoding.BigEndianUnicode;
        }

        return IsValidUtf8(sample) ? Encoding.UTF8 : GetGbk() ?? Encoding.UTF8;
    }
    #endregion

    #region 解码
    /// <summary>解码字节数组</summary>
    /// <param name="bytes">字节数组，空数组返回空字符串</param>
    /// <returns>解码后的文本</returns>
    public static String Decode(Byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0) return "";

        var encoding = DetectEncoding(bytes, out var bomLength);
        return bomLength > 0
            ? encoding.GetString(bytes, bomLength, bytes.Length - bomLength)
            : encoding.GetString(bytes);
    }
    #endregion

    #region 辅助
    private static Boolean IsValidUtf8(Byte[] bytes)
    {
        try
        {
            // flush:false 允许样本尾部被截断的多字节字符，只把真正的非法序列判为非 UTF-8
            new UTF8Encoding(false, true).GetDecoder().GetCharCount(bytes, 0, bytes.Length, false);
            return true;
        }
        catch (DecoderFallbackException)
        {
            return false;
        }
    }

    private static Encoding? GetGbk()
    {
        try
        {
            return Encoding.GetEncoding(936);
        }
        catch
        {
            // netstandard 未注册 CodePages 提供程序时回退 UTF-8
            return null;
        }
    }
    #endregion
}
