using SkiaSharp;

namespace NewLife.Office.Rendering.Pdf;

/// <summary>PDF 字体到 SkiaSharp 字体的映射器</summary>
/// <remarks>
/// 维护 PDF 字体名 → SKTypeface 的映射表。
/// 映射策略：
///   1. 标准 14 字体 → 系统对应字体（如 Helvetica → sans-serif）
///   2. 系统已安装字体（按名称匹配）
///   3. 通用回退字体
///   4. CJK 中文回退字体链
/// </remarks>
internal sealed class PdfFontMapper
{
    private readonly Dictionary<String, SKTypeface> _cache = new(StringComparer.OrdinalIgnoreCase);

    // 标准 PDF 14 字体 → 系统字体映射
    private static readonly Dictionary<String, String> Standard14Map = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Helvetica"] = "sans-serif",
        ["Helvetica-Bold"] = "sans-serif",
        ["Helvetica-Oblique"] = "sans-serif",
        ["Helvetica-BoldOblique"] = "sans-serif",
        ["Times-Roman"] = "serif",
        ["Times-Bold"] = "serif",
        ["Times-Italic"] = "serif",
        ["Times-BoldItalic"] = "serif",
        ["Courier"] = "monospace",
        ["Courier-Bold"] = "monospace",
        ["Courier-Oblique"] = "monospace",
        ["Courier-BoldOblique"] = "monospace",
        ["Symbol"] = "serif",
        ["ZapfDingbats"] = "serif",
    };

    // CJK 中文字体回退链（按优先级排列）
    private static readonly String[] CjkFallbackFonts =
    {
        "Microsoft YaHei",
        "SimSun",
        "SimHei",
        "Noto Sans CJK SC",
        "WenQuanYi Micro Hei",
        "WenQuanYi Zen Hei",
        "Source Han Sans SC",
        "STSong",
        "FangSong",
        "KaiTi",
        "sans-serif",
    };

    // 缓存 CJK typeface（懒加载）
    private SKTypeface? _cjkTypeface;
    private Boolean _cjkSearched;

    /// <summary>获取或创建指定 PDF 字体对应的 SKTypeface</summary>
    /// <param name="pdfFontName">PDF 字体名（如 "F1" 或 "Helvetica"）</param>
    /// <param name="baseFontName">BaseFont 名称（来自 PDF 字体字典）</param>
    /// <param name="fontData">嵌入字体字节（可为 null）</param>
    /// <returns>SKTypeface，不会返回 null</returns>
    public SKTypeface GetTypeface(String pdfFontName, String? baseFontName, Byte[]? fontData)
    {
        // 优先走缓存
        if (_cache.TryGetValue(pdfFontName, out var cached))
            return cached;

        SKTypeface? tf = null;

        // 1. 嵌入字体优先
        if (fontData != null && fontData.Length > 0)
        {
            tf = SKTypeface.FromData(SKData.CreateCopy(fontData));
            if (tf != null)
            {
                _cache[pdfFontName] = tf;
                return tf;
            }
        }

        // 2. 标准 14 字体映射
        var mappedName = baseFontName ?? pdfFontName;
        if (Standard14Map.TryGetValue(mappedName, out var systemName))
        {
            tf = SKTypeface.FromFamilyName(systemName);
        }

        // 3. 直接用 PDF 字体名查找系统字体
        if (tf == null)
            tf = SKTypeface.FromFamilyName(mappedName);

        // 4. 去掉连字符后缀重试（如 "Helvetica-Bold" → "Helvetica"）
        if (tf == null && mappedName.Contains('-'))
        {
            var shortName = mappedName[..mappedName.LastIndexOf('-')];
            tf = SKTypeface.FromFamilyName(shortName);
        }

        // 5. CJK 回退
        if (tf == null)
            tf = GetCjkTypeface();

        // 6. 终极回退
        if (tf == null)
            tf = SKTypeface.Default;

        _cache[pdfFontName] = tf;
        return tf;
    }

    /// <summary>获取 CJK 中文字体</summary>
    private SKTypeface GetCjkTypeface()
    {
        if (_cjkSearched) return _cjkTypeface ?? SKTypeface.Default;
        _cjkSearched = true;

        foreach (var fontName in CjkFallbackFonts)
        {
            var tf = SKTypeface.FromFamilyName(fontName);
            if (tf != null)
            {
                _cjkTypeface = tf;
                return tf;
            }
        }

        return SKTypeface.Default;
    }

    /// <summary>判断是否为 CJK 字体名</summary>
    public static Boolean IsCjkFont(String? baseFontName)
    {
        if (String.IsNullOrEmpty(baseFontName)) return false;
        var upper = baseFontName.ToUpperInvariant();
        return upper.Contains("CJK") || upper.Contains("SONG") ||
               upper.Contains("HEI") || upper.Contains("MING") ||
               upper.Contains("GOTHIC") || upper.Contains("MINCHO") ||
               upper.Contains("STSONG") || upper.Contains("KAITI") ||
               upper.Contains("FANGSONG") || upper.Contains("MS ") ||
               upper.Contains("KOZUKA") || upper.Contains("ADOBE") && upper.Contains("STD");
    }
}
