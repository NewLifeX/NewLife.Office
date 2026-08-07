using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office;
using NewLife.Office.Ole2;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>doc → docx 转换测试（W20）</summary>
/// <remarks>
/// 基于手写最小 .doc（OLE2 + FIB + CLX）验证 DocReader.ToDocument/SaveAsDocx：
/// 段落文本、带格式（粗体/居中）、表格交错顺序均保留。
/// </remarks>
public class DocxConversionTests
{
    #region .doc 构造辅助（与 DocReaderTests 同模式）

    private static Byte[] LE2(UInt16 v) => new Byte[] { (Byte)(v & 0xFF), (Byte)(v >> 8) };
    private static Byte[] LE4(UInt32 v) =>
        new Byte[] { (Byte)(v & 0xFF), (Byte)((v >> 8) & 0xFF), (Byte)((v >> 16) & 0xFF), (Byte)(v >> 24) };

    private static Byte[] Concat(params Byte[][] parts)
    {
        var total = 0;
        foreach (var p in parts) total += p.Length;
        var buf = new Byte[total];
        var pos = 0;
        foreach (var p in parts) { Array.Copy(p, 0, buf, pos, p.Length); pos += p.Length; }
        return buf;
    }

    /// <summary>构建最小 WordDocument 流（文本按字节写入，\r 转为段落符，保留 \x07 表格标记）</summary>
    private static Byte[] BuildWordDocStream(String text, Int32 fibSize = 400)
    {
        var chars = new List<Byte>();
        foreach (var ch in text)
        {
            if (ch == '\r' || ch == '\n') chars.Add(0x0D);
            else chars.Add((Byte)ch);
        }
        chars.Add(0x0D); // 末段落符

        var textBytes = chars.ToArray();
        var textLen = textBytes.Length;

        const Int32 FibSizeBase = 400;
        var ClxStart = fibSize;
        const Int32 PlcPcdSize = 2 * 4 + 1 * 8;
        const Int32 ClxBlockSize = 1 + 4 + PlcPcdSize;
        var textOffset = ClxStart + ClxBlockSize;
        var fcValue = textOffset * 2;

        var fib = new Byte[fibSize];
        fib[0] = 0xEC; fib[1] = 0xA5;
        fib[2] = 0xC1; fib[3] = 0x00;
        fib[32] = 14; fib[33] = 0;
        fib[62] = 22; fib[63] = 0;
        fib[152] = 74; fib[153] = 0;
        var fcClxBytes = LE4((UInt32)ClxStart);
        var lcbClxBytes = LE4((UInt32)ClxBlockSize);
        Array.Copy(fcClxBytes, 0, fib, 258, 4);
        Array.Copy(lcbClxBytes, 0, fib, 262, 4);

        var plcPcd = Concat(LE4(0), LE4((UInt32)textLen), Concat(LE2(0), LE4((UInt32)(fcValue | (1 << 30))), LE2(0)));
        var clxBlock = Concat(new Byte[] { 0x02 }, LE4((UInt32)plcPcd.Length), plcPcd);

        return Concat(fib, clxBlock, textBytes);
    }

    /// <summary>打包为 .doc 字节</summary>
    private static Byte[] BuildDoc(String text)
    {
        var wordDoc = BuildWordDocStream(text);
        var cfb = new CfbDocument();
        cfb.Root.AddStream("WordDocument", wordDoc);
        return cfb.ToBytes();
    }

    /// <summary>构建带 CHPX/PAPX 格式的 WordDocument 流（"BoldText\rPlainText\r"，首段粗体+居中）</summary>
    private static Byte[] BuildFormattedWordDocStream()
    {
        const String text = "BoldText\rPlainText\r";
        var textBytes = Encoding.ASCII.GetBytes(text);
        var textLen = textBytes.Length;

        const Int32 FibSize = 400;
        const Int32 ClxStart = FibSize;
        const Int32 PlcPcdSize = 2 * 4 + 1 * 8;
        const Int32 ClxBlockSize = 1 + 4 + PlcPcdSize;
        var textOffset = ClxStart + ClxBlockSize;
        var fcValue = textOffset * 2;

        // CHPX：首段粗体
        var chpx0 = new Byte[] { 3, 0x88, 0x35, 0x01 };
        var chpx1 = new Byte[] { 0 };
        var chpxData0 = textOffset + textLen;
        var chpxData1 = chpxData0 + chpx0.Length;
        var chpxPlcBteStart = chpxData1 + chpx1.Length;
        var chpxPlcBte = Concat(LE4(0), LE4(8), LE4(19), LE4((UInt32)chpxData0), LE4((UInt32)chpxData1));
        var lcbChpx = chpxPlcBte.Length;

        // PAPX：首段居中（sprmPJc=0x2403 → [0xA4,0x03] + 0x01）+ PHE(12)
        var papx0 = new Byte[17];
        papx0[0] = 3; papx0[1] = 0xA4; papx0[2] = 0x03; papx0[3] = 0x01;
        var papx1 = new Byte[13];
        var papxData0 = chpxPlcBteStart + lcbChpx;
        var papxData1 = papxData0 + papx0.Length;
        var papxPlcBteStart = papxData1 + papx1.Length;
        var papxPlcBte = Concat(LE4(0), LE4(9), LE4(19), LE4((UInt32)papxData0), LE4((UInt32)papxData1));
        var lcbPapx = papxPlcBte.Length;

        var fib = new Byte[FibSize];
        fib[0] = 0xEC; fib[1] = 0xA5;
        fib[2] = 0xC1; fib[3] = 0x00;
        fib[32] = 14; fib[33] = 0;
        fib[62] = 22; fib[63] = 0;
        fib[152] = 74; fib[153] = 0;
        Array.Copy(LE4((UInt32)chpxPlcBteStart), 0, fib, 250, 4);
        Array.Copy(LE4((UInt32)lcbChpx), 0, fib, 254, 4);
        Array.Copy(LE4((UInt32)ClxStart), 0, fib, 258, 4);
        Array.Copy(LE4((UInt32)ClxBlockSize), 0, fib, 262, 4);
        Array.Copy(LE4((UInt32)papxPlcBteStart), 0, fib, 266, 4);
        Array.Copy(LE4((UInt32)lcbPapx), 0, fib, 270, 4);

        var plcPcd = Concat(LE4(0), LE4((UInt32)textLen), Concat(LE2(0), LE4((UInt32)(fcValue | (1 << 30))), LE2(0)));
        var clxBlock = Concat(new Byte[] { 0x02 }, LE4((UInt32)plcPcd.Length), plcPcd);

        return Concat(fib, clxBlock, textBytes, chpx0, chpx1, chpxPlcBte, papx0, papx1, papxPlcBte);
    }

    private static Byte[] BuildFormattedDoc()
    {
        var cfb = new CfbDocument();
        cfb.Root.AddStream("WordDocument", BuildFormattedWordDocStream());
        return cfb.ToBytes();
    }

    #endregion

    #region 测试

    [Fact, DisplayName("W20_doc转docx_纯文本段落转换")]
    public void DocToDocx_BasicText()
    {
        var bytes = BuildDoc("First paragraph\rSecond paragraph\r");
        var outPath = Path.Combine(Path.GetTempPath(), $"conv_{Guid.NewGuid():N}.docx");
        try
        {
            using var ms = new MemoryStream(bytes);
            using (var reader = new DocReader(ms))
                reader.SaveAsDocx(outPath);

            Assert.True(File.Exists(outPath));
            using var reader2 = new WordReader(outPath);
            var paras = reader2.ReadParagraphs().ToList();
            Assert.Contains("First paragraph", paras);
            Assert.Contains("Second paragraph", paras);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W20_doc转docx_带格式段落转换")]
    public void DocToDocx_Formatted()
    {
        var bytes = BuildFormattedDoc();
        var outPath = Path.Combine(Path.GetTempPath(), $"conv_{Guid.NewGuid():N}.docx");
        try
        {
            using var ms = new MemoryStream(bytes);
            using (var reader = new DocReader(ms))
                reader.SaveAsDocx(outPath);

            using var reader2 = new WordReader(outPath);
            var doc = reader2.ReadDocument();
            var paras = doc.Elements.Where(e => e.Paragraph != null).Select(e => e.Paragraph!).ToList();
            Assert.True(paras.Count >= 2);
            // 首段粗体
            var first = paras[0];
            Assert.Contains("BoldText", String.Concat(first.Runs.Select(r => r.Text)));
            Assert.True(first.Runs.Any(r => r.Properties?.Bold == true), "粗体格式应保留");
            Assert.Equal("center", first.Alignment);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W20_doc转docx_表格转换")]
    public void DocToDocx_Table()
    {
        // 构造：段落 + 表格行（0x07 单元格标记）+ 段落
        // 注意：\u0007 固定4位转义，避免 \x07B 被解析成 0x7B='{'
        const String Sep = "\u0007";
        var bytes = BuildDoc("Pre\rA" + Sep + "B" + Sep + "C" + Sep + "\rPost\r");
        var outPath = Path.Combine(Path.GetTempPath(), $"conv_{Guid.NewGuid():N}.docx");
        try
        {
            using var ms = new MemoryStream(bytes);
            using (var reader = new DocReader(ms))
                reader.SaveAsDocx(outPath);

            using var reader2 = new WordReader(outPath);
            var doc = reader2.ReadDocument();
            var elements = doc.Elements.ToList();

            // 顺序：段落 → 表格 → 段落
            Assert.Equal(3, elements.Count);
            Assert.Equal(ElementType.Paragraph, elements[0].Type);
            Assert.Equal(ElementType.Table, elements[1].Type);
            Assert.Equal(ElementType.Paragraph, elements[2].Type);

            // 表格单元格内容
            var rows = elements[1].TableRows!;
            Assert.Single(rows);
            Assert.Equal(3, rows[0].Count);
            Assert.Equal("A", String.Concat(rows[0][0].Paragraphs[0].Runs.Select(r => r.Text)));
            Assert.Equal("B", String.Concat(rows[0][1].Paragraphs[0].Runs.Select(r => r.Text)));
            Assert.Equal("C", String.Concat(rows[0][2].Paragraphs[0].Runs.Select(r => r.Text)));

            // 后置段落
            var last = String.Concat(elements[2].Paragraph!.Runs.Select(r => r.Text));
            Assert.Contains("Post", last);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W20_doc转docx_模型ToDocument直接验证")]
    public void DocToDocx_ToDocument()
    {
        var bytes = BuildDoc("Alpha\rBeta\r");
        using var ms = new MemoryStream(bytes);
        using var reader = new DocReader(ms);
        var doc = reader.ToDocument();
        Assert.Equal(2, doc.Elements.Count);
        Assert.Equal("Alpha", String.Concat(doc.Elements[0].Paragraph!.Runs.Select(r => r.Text)));
        Assert.Equal("Beta", String.Concat(doc.Elements[1].Paragraph!.Runs.Select(r => r.Text)));
    }

    [Fact, DisplayName("W20_doc转docx_OfficeFactory便捷转换")]
    public void DocToDocx_OfficeFactory()
    {
        var bytes = BuildDoc("FactoryDoc\r");
        var inPath = Path.Combine(Path.GetTempPath(), $"in_{Guid.NewGuid():N}.doc");
        var outPath = Path.Combine(Path.GetTempPath(), $"out_{Guid.NewGuid():N}.docx");
        try
        {
            File.WriteAllBytes(inPath, bytes);
            OfficeFactory.ConvertDocToDocx(inPath, outPath);
            Assert.True(File.Exists(outPath));

            using var reader = new WordReader(outPath);
            Assert.Contains("FactoryDoc", reader.ReadFullText());
        }
        finally
        {
            if (File.Exists(inPath)) File.Delete(inPath);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    #endregion
}
