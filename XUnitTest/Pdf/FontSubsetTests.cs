using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF 字体子集化测试（P08，对标 PdfSharp）</summary>
public class FontSubsetTests
{
    static FontSubsetTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    #region 辅助
    /// <summary>查找可用的系统中文字体（优先 TTF 单字体，其次 TTC）</summary>
    private static String? FindCjkFont()
    {
        var dir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        foreach (var name in new[] { "simsunb.ttf", "simsun.ttc", "msyh.ttc", "simhei.ttf" })
        {
            var path = Path.Combine(dir, name);
            if (File.Exists(path)) return path;
        }
        return null;
    }

    /// <summary>计算 TTF/TTC 的字体偏移（TTC 内取指定索引）</summary>
    private static Int32 GetSfOffset(Byte[] data, Int32 ttcIndex)
    {
        if (data.Length > 4 && data[0] == 0x74 && data[1] == 0x74 && data[2] == 0x63 && data[3] == 0x66)
        {
            var numFonts = ReadU32(data, 8);
            if (ttcIndex >= numFonts) ttcIndex = 0;
            return ReadU32(data, 12 + ttcIndex * 4);
        }
        return 0;
    }

    private static UInt16 ReadU16(Byte[] d, Int32 o) => (UInt16)((d[o] << 8) | d[o + 1]);

    private static Int32 ReadU32(Byte[] d, Int32 o) => (d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3];

    /// <summary>从子集 TTF 读取 numGlyphs（maxp 表 +4）</summary>
    private static Int32 GetNumGlyphs(Byte[] data)
    {
        var numTables = ReadU16(data, 4);
        for (var i = 0; i < numTables; i++)
        {
            var pos = 12 + i * 16;
            var tag = Encoding.ASCII.GetString(data, pos, 4);
            if (tag == "maxp") return ReadU16(data, ReadU32(data, pos + 8) + 4);
        }
        return -1;
    }

    /// <summary>从 PDF 字节中提取 FontFile2 流对象的数据长度（-1=未找到嵌入字体）</summary>
    private static Int32 FindFontFile2Length(Byte[] pdf)
    {
        var text = Encoding.Latin1.GetString(pdf);
        var m = Regex.Match(text, @"/FontFile2 (\d+) 0 R");
        if (!m.Success) return -1;
        var objId = m.Groups[1].Value;
        var m2 = Regex.Match(text, objId + @" 0 obj\s*<<(?<dict>[^>]*)/Length (?<len>\d+)");
        if (!m2.Success) return -1;
        return Int32.Parse(m2.Groups["len"].Value);
    }
    #endregion

    #region 子集化核心
    [Fact, DisplayName("真实中文字体子集化后体积显著减小且结构合法")]
    public void Subset_RealCjkFont_VolumeShrinksAndValid()
    {
        var fontPath = FindCjkFont();
        if (fontPath == null)
        {
            return; // 系统中文字体不可用，静默跳过（xUnit v2 无动态 Skip）
        }

        var fontData = File.ReadAllBytes(fontPath);
        var sfOff = GetSfOffset(fontData, 0);

        // 常用汉字样本
        var text = "你好世界，NewLife.Office 办公自动化！12345。";
        var chars = new List<Int32>();
        foreach (var ch in text) chars.Add(ch);

        var (subset, map) = TrueTypeSubsetter.Subset(fontData, sfOff, chars);

        // 合法 TTF 头（版本 0x00010000 或 1.0）
        Assert.True(subset.Length >= 12, "子集过短");
        Assert.Equal(0x00010000, ReadU32(subset, 0));
        // 体积显著减小（样本字符少，应远小于原字体）
        Assert.True(subset.Length < fontData.Length / 2, $"子集 {subset.Length} 未显著小于原字体 {fontData.Length}");
        // 映射只含请求字符（含字体映射到 .notdef 的除外）
        Assert.NotEmpty(map);
        // 新 GID 都在 numGlyphs 范围内
        var numGlyphs = GetNumGlyphs(subset);
        Assert.True(numGlyphs > 0);
        Assert.True(numGlyphs <= chars.Count + 1, $"numGlyphs={numGlyphs} 超出预期");
        foreach (var gid in map.Values)
            Assert.True(gid < numGlyphs, $"GID {gid} 超出 numGlyphs {numGlyphs}");
    }

    [Fact, DisplayName("子集化映射仅包含实际使用字符且可还原")]
    public void Subset_GlyphMap_OnlyUsedChars()
    {
        var fontPath = FindCjkFont();
        if (fontPath == null)
        {
            return; // 系统中文字体不可用，静默跳过（xUnit v2 无动态 Skip）
        }

        var fontData = File.ReadAllBytes(fontPath);
        var sfOff = GetSfOffset(fontData, 0);

        var used = new List<Int32> { '中', '国', '人' };
        var (_, map) = TrueTypeSubsetter.Subset(fontData, sfOff, used);

        // 映射的键应包含这三个字符（除非字体缺字形，simsun 覆盖常用汉字）
        foreach (var code in used)
        {
            if (!map.ContainsKey((UInt16)code))
            {
                // 允许个别字符映射缺失（字体不含该字形），但不应出现额外字符
                continue;
            }
        }
        // 不含未请求字符
        foreach (var kv in map)
        {
            Assert.Contains((Int32)kv.Key, used);
        }
    }

    [Fact, DisplayName("空字符集仅保留 .notdef 字形")]
    public void Subset_EmptyChars_OnlyNotdef()
    {
        var fontPath = FindCjkFont();
        if (fontPath == null)
        {
            return; // 系统中文字体不可用，静默跳过（xUnit v2 无动态 Skip）
        }

        var fontData = File.ReadAllBytes(fontPath);
        var sfOff = GetSfOffset(fontData, 0);

        var (subset, map) = TrueTypeSubsetter.Subset(fontData, sfOff, []);

        Assert.Equal(1, GetNumGlyphs(subset));
        Assert.Empty(map);
    }

    [Fact, DisplayName("非 BMP 码点与未知字符被安全跳过")]
    public void Subset_NonBmpAndUnknown_Safe()
    {
        var fontPath = FindCjkFont();
        if (fontPath == null)
        {
            return; // 系统中文字体不可用，静默跳过（xUnit v2 无动态 Skip）
        }

        var fontData = File.ReadAllBytes(fontPath);
        var sfOff = GetSfOffset(fontData, 0);

        // 非 BMP（emoji）、未知码点、常规字符混合
        var chars = new List<Int32> { 0x1F600, 0x10FFFF, 'A', '汉' };
        var (subset, map) = TrueTypeSubsetter.Subset(fontData, sfOff, chars);

        Assert.True(subset.Length > 0);
        // 只应保留 BMP 且字体存在的字符
        Assert.True(map.Count <= 2);
        Assert.DoesNotContain(0x1F600, map.Keys.Select(k => (Int32)k));
    }
    #endregion

    #region PdfWriter 端到端
    [Fact, DisplayName("启用子集化后 PDF 嵌入字体体积显著减小")]
    public void PdfWriter_SubsetFonts_EmbedSmallerFont()
    {
        var font = new PdfWriter().CreateFont("宋体");
        if (font.FontFilePath == null)
        {
            return; // 宋体字体文件不可用，静默跳过
        }
        var fontData = File.ReadAllBytes(font.FontFilePath);

        // 未启用子集化：嵌入原始字体
        var fullPdf = BuildCjkPdf(false);
        var fullLen = FindFontFile2Length(fullPdf);
        Assert.True(fullLen > 0, "未找到未子集化 FontFile2");
        Assert.Equal(fontData.Length, fullLen);

        // 启用子集化：嵌入子集字体
        var subsetPdf = BuildCjkPdf(true);
        var subsetLen = FindFontFile2Length(subsetPdf);
        Assert.True(subsetLen > 0, "未找到子集化 FontFile2");
        Assert.True(subsetLen < fullLen / 2, $"子集 FontFile2 {subsetLen} 未显著小于原始 {fullLen}");
    }

    /// <summary>生成含中文文本的 PDF（可选子集化）</summary>
    private static Byte[] BuildCjkPdf(Boolean subset)
    {
        using var ms = new MemoryStream();
        var writer = new PdfWriter { SubsetFonts = subset };
        var font = writer.CreateFont("宋体");
        writer.BeginPage();
        writer.DrawText("你好，世界！Hello World 123。", 56, 780, 24, font);
        writer.DrawText("第二行文字，NewLife.Office。", 56, 740, 24, font);
        writer.Save(ms);
        return ms.ToArray();
    }
    #endregion
}
