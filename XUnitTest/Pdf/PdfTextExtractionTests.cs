using System;
using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF 文本提取测试。覆盖 ToUnicode CMap 映射（CID 字体）与 bfrange 区间映射</summary>
/// <remarks>回归背景：内容流里的文本是字形编码而非 Unicode，Type0/Identity-H 字体必须经 /ToUnicode
/// 才能还原。早期实现直接把 2 字节当成 UTF-16BE，导致带子集字体的中文 PDF 提取出乱码。</remarks>
public class PdfTextExtractionTests
{
    [Fact]
    [DisplayName("PDF_bfchar：CID 字体经 ToUnicode 还原为中文")]
    public void ExtractText_Type0WithBfChar()
    {
        var cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
            + "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n"
            + "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n"
            + "3 beginbfchar\n<0001> <4E2D>\n<0002> <6587>\n<0003> <6D4B>\nendbfchar\n"
            + "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend";

        var pdf = BuildCidPdf("<000100020003> Tj", cmap);

        using var reader = new PdfReader(new MemoryStream(pdf));
        var text = reader.ExtractText();

        Assert.Contains("中文测", text);
    }

    [Fact]
    [DisplayName("PDF_bfrange：区间映射按增量还原")]
    public void ExtractText_Type0WithBfRange()
    {
        var cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
            + "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n"
            + "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n"
            + "1 beginbfrange\n<0001> <0003> <4E2D>\nendbfrange\n"
            + "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend";

        // 4E2D/4E2E/4E2F 为连续码点
        var pdf = BuildCidPdf("<000100020003> Tj", cmap);

        using var reader = new PdfReader(new MemoryStream(pdf));
        var text = reader.ExtractText();

        Assert.Contains("中", text);
        Assert.Equal(3, text.Trim().Length);
    }

    [Fact]
    [DisplayName("PDF_无 ToUnicode：Identity 编码按码点兜底")]
    public void ExtractText_Type0WithoutToUnicode()
    {
        var pdf = BuildCidPdf("<4E2D65876D4B> Tj", cmap: null);

        using var reader = new PdfReader(new MemoryStream(pdf));
        var text = reader.ExtractText();

        Assert.Contains("中文测", text);
    }

    /// <summary>构造只含一页 Type0 字体的最小 PDF。cmap 为空时不写 /ToUnicode</summary>
    /// <param name="showText">内容流里的文本显示操作</param>
    /// <param name="cmap">ToUnicode CMap 内容</param>
    /// <returns>PDF 字节</returns>
    private static Byte[] BuildCidPdf(String showText, String? cmap)
    {
        var content = $"BT /F1 12 Tf 72 720 Td {showText} ET";

        var objects = new System.Collections.Generic.List<String>
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            cmap == null
                ? "<< /Type /Font /Subtype /Type0 /BaseFont /SimSun /Encoding /Identity-H /DescendantFonts [6 0 R] >>"
                : "<< /Type /Font /Subtype /Type0 /BaseFont /SimSun /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 7 0 R >>",
            $"<< /Length {content.Length} >>\nstream\n{content}\nendstream",
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /SimSun /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>",
        };
        if (cmap != null) objects.Add($"<< /Length {cmap.Length} >>\nstream\n{cmap}\nendstream");

        var sb = new StringBuilder();
        sb.Append("%PDF-1.7\n");

        var offsets = new Int32[objects.Count];
        for (var i = 0; i < objects.Count; i++)
        {
            offsets[i] = sb.Length;
            sb.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }

        var xrefPos = sb.Length;
        sb.Append($"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets)
            sb.Append($"{offset:D10} 00000 n \n");
        sb.Append($"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\nstartxref\n{xrefPos}\n%%EOF\n");

        return Encoding.Latin1.GetBytes(sb.ToString());
    }
}
