using System;
using System.Net;

namespace NewLife.Office.Markdown;

/// <summary>HTML 实体（字符引用）解码器</summary>
/// <remarks>
/// 支持 CommonMark 规范的全部字符引用：
/// <list type="bullet">
/// <item>具名实体：HTML5 完整实体表（经 <see cref="WebUtility.HtmlDecode(String)"/>，约 2000+ 项）</item>
/// <item>数字字符引用：十进制 <c>&amp;#65;</c> 与十六进制 <c>&amp;#x41;</c>，含补充平面代理对（<c>&amp;#128512;</c> → 😀）</item>
/// </list>
/// 未识别的实体返回 <c>null</c>，由调用方保留原文；无效数字码点（0 / 代理区 / 超出 U+10FFFF）
/// 按 CommonMark 规范解码为 U+FFFD 替换符。
/// <para>供 <see cref="MarkdownParser"/> 与 <see cref="HtmlToMarkdownConverter"/> 共享使用，保证两处解码行为一致。</para>
/// </remarks>
internal static class EntityDecoder
{
    /// <summary>解码单个实体（不含 &amp; 与 ;）</summary>
    /// <param name="entity">实体名，如 "amp"、"#65"、"#x1F600"</param>
    /// <returns>解码后字符串；未识别返回 <c>null</c>（调用方保留原文）</returns>
    public static String? Decode(String entity)
    {
        if (String.IsNullOrEmpty(entity)) return null;

        // 数字字符引用：&#123; / &#x1F;
        if (entity[0] == '#')
            return DecodeNumeric(entity);

        // 具名实体：WebUtility.HtmlDecode 内置完整 HTML5 实体表；
        // 未识别实体解码后原样返回（含 &...;），据此判定
        var input = "&" + entity + ";";
        var decoded = WebUtility.HtmlDecode(input);
        return decoded == input ? null : decoded;
    }

    /// <summary>解码数字字符引用（十进制 / 十六进制），含补充平面代理对</summary>
    /// <param name="entity">实体名（以 # 开头）</param>
    /// <returns>解码后字符串；无效码点返回 U+FFFD（CommonMark 规范）</returns>
    private static String? DecodeNumeric(String entity)
    {
        try
        {
            Int32 code;
            if (entity.Length > 2 && (entity[1] == 'x' || entity[1] == 'X'))
                code = Convert.ToInt32(entity.Substring(2), 16);
            else
                code = Convert.ToInt32(entity.Substring(1), 10);

            // 无效码点：0 / 代理区 0xD800-0xDFFF / 超出 0x10FFFF → U+FFFD
            if (code <= 0 || code > 0x10FFFF || code >= 0xD800 && code <= 0xDFFF)
                return "\uFFFD";
            return Char.ConvertFromUtf32(code);
        }
        catch
        {
            return "\uFFFD";
        }
    }
}
