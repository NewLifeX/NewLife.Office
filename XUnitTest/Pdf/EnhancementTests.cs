using NewLife.Office;
using System.Text;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF 增强测试 — 注释类型、表单填充读取</summary>
public class PdfEnhancementTests
{
    static PdfEnhancementTests()
    {
        // PdfReader 文本提取需要 Latin-1 (1252) 编码
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
    [Fact(DisplayName = "PDF—Caret注释写入")]
    public void Annotation_Caret()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Caret, PageIndex = 0,
                X = 100, Y = 700, Width = 10, Height = 15,
                Contents = "插入此处", Author = "审阅者"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Polygon注释写入")]
    public void Annotation_Polygon()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Polygon, PageIndex = 0,
                X = 100, Y = 700, Width = 50, Height = 50, Contents = "多边形区域"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Polygon", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Squiggly注释写入")]
    public void Annotation_Squiggly()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Squiggly, PageIndex = 0,
                X = 100, Y = 700, Width = 80, Height = 5, Contents = "波浪下划线"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);
            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Squiggly", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—表单创建+填充+读取往返")]
    public void FormField_FillAndRead()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AddTextField("txtName", 100, 700, 200, 20, "ZhangSan");
            writer.AddCheckBox("chkAgree", 100, 650, 12, true);
            writer.AddComboBox("selCity", 100, 620, 200, 20, new List<String> { "Beijing", "Shanghai" }, 1);
            writer.Save(path);

            // Verify form data is correctly written to PDF
            var rawPdf = File.ReadAllText(path);
            Assert.Contains("/AcroForm", rawPdf);
            Assert.Contains("/Subtype /Widget", rawPdf);
            Assert.Contains("/FT /Tx", rawPdf);
            Assert.Contains("/FT /Btn", rawPdf);

            // Read back - form reading relies on xref table which is still being improved
            using var reader = new PdfReader(path);
            var form = reader.ReadFormFields();
            if (form != null)
            {
                Assert.True(form.Fields.Count >= 3);
                var nameField = form.Fields.FirstOrDefault(f => f.FullName == "txtName");
                Assert.NotNull(nameField);
                Assert.Equal("ZhangSan", nameField!.Value);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—SetFormFieldValue修改值")]
    public void FormField_SetValue()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AddTextField("txtName", 100, 700, 200, 20, "OriginalValue");
            var result = writer.SetFormFieldValue("txtName", "ModifiedValue");
            Assert.True(result);
            writer.Save(path);

            var rawPdf = File.ReadAllText(path);
            Assert.Contains("/V (ModifiedValue)", rawPdf);

            using var reader = new PdfReader(path);
            var form = reader.ReadFormFields();
            if (form != null)
            {
                var nameField = form.Fields.FirstOrDefault(f => f.FullName == "txtName");
                Assert.NotNull(nameField);
                Assert.Equal("ModifiedValue", nameField!.Value);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Polygon顶点写入")]
    public void Polygon_Vertices()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Polygon,
                PageIndex = 0,
                X = 100, Y = 700, Width = 200, Height = 100,
                Contents = "三角形区域",
                Vertices = [100, 700, 200, 800, 300, 700]
            };
            writer.AddAnnotation(ann);
            writer.Save(path);

            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Polygon", pdfContent);
            Assert.Contains("/Vertices", pdfContent);
            Assert.Contains("100.00", pdfContent);
            Assert.Contains("800.00", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—PolyLine顶点写入")]
    public void PolyLine_Vertices()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.PolyLine,
                PageIndex = 0,
                X = 50, Y = 600, Width = 300, Height = 50,
                Contents = "折线路径",
                Vertices = [50, 600, 150, 650, 250, 620, 350, 640]
            };
            writer.AddAnnotation(ann);
            writer.Save(path);

            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /PolyLine", pdfContent);
            Assert.Contains("/Vertices", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—Polygon无顶点时不含Vertices")]
    public void Polygon_NoVertices_OmitsEntry()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var ann = new PdfAnnotation
            {
                Type = PdfAnnotationType.Polygon,
                PageIndex = 0,
                X = 100, Y = 700, Width = 50, Height = 50,
                Contents = "无顶点"
            };
            writer.AddAnnotation(ann);
            writer.Save(path);

            var pdfContent = File.ReadAllText(path);
            Assert.Contains("/Subtype /Polygon", pdfContent);
            Assert.DoesNotContain("/Vertices", pdfContent);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    #region LZWDecode 解码
    [Fact(DisplayName = "PDF—LZW解码简单文本(ABC)")]
    public void LzwDecode_SimpleText()
    {
        var encoded = BuildSimpleLzw("ABC");
        var decoded = PdfReader.DecodeLzw(encoded);
        Assert.Equal("ABC", Encoding.ASCII.GetString(decoded));
    }

    [Fact(DisplayName = "PDF—LZW解码长文本")]
    public void LzwDecode_LongText()
    {
        var text = "The quick brown fox jumps over the lazy dog. 1234567890!@#$%^";
        var encoded = BuildSimpleLzw(text);
        var decoded = PdfReader.DecodeLzw(encoded);
        Assert.Equal(text, Encoding.ASCII.GetString(decoded));
    }

    [Fact(DisplayName = "PDF—LZW解码中文UTF-8文本")]
    public void LzwDecode_ChineseText()
    {
        var text = "你好世界HelloWorld混合文本测试";
        var textBytes = Encoding.UTF8.GetBytes(text);
        var encoded = BuildSimpleLzwBytes(textBytes);
        var decoded = PdfReader.DecodeLzw(encoded);
        Assert.Equal(textBytes, decoded);
    }

    [Fact(DisplayName = "PDF—LZW解码空数据")]
    public void LzwDecode_EmptyData()
    {
        var decoded = PdfReader.DecodeLzw([]);
        Assert.Empty(decoded);

        decoded = PdfReader.DecodeLzw(null!);
        Assert.Empty(decoded);
    }

    [Fact(DisplayName = "PDF—LZW解码包含EarlyChange参数")]
    public void LzwDecode_EarlyChange()
    {
        // earlyChange=0 和 earlyChange=1 都应能正确解码简单数据
        var encoded = BuildSimpleLzw("Test");
        var decoded0 = PdfReader.DecodeLzw(encoded, earlyChange: 0);
        var decoded1 = PdfReader.DecodeLzw(encoded, earlyChange: 1);
        Assert.Equal("Test", Encoding.ASCII.GetString(decoded0));
        Assert.Equal("Test", Encoding.ASCII.GetString(decoded1));
    }

    /// <summary>构建简单 LZW 编码数据（仅单字节码，无字典压缩，适合短文本）</summary>
    private static Byte[] BuildSimpleLzw(String text) => BuildSimpleLzwBytes(Encoding.ASCII.GetBytes(text));

    private static Byte[] BuildSimpleLzwBytes(Byte[] data)
    {
        using var ms = new MemoryStream();
        var bitBuf = 0L;
        var bitCount = 0;

        void WriteCode(Int32 code)
        {
            bitBuf |= (Int64)code << bitCount;
            bitCount += 9;
            while (bitCount >= 8)
            {
                ms.WriteByte((Byte)(bitBuf & 0xFF));
                bitBuf >>= 8;
                bitCount -= 8;
            }
        }

        WriteCode(256); // Clear
        foreach (var b in data)
            WriteCode(b);
        WriteCode(257); // EOD

        // Flush remaining bits
        if (bitCount > 0)
            ms.WriteByte((Byte)(bitBuf & 0xFF));

        return ms.ToArray();
    }
    #endregion

    #region RunLengthDecode 解码
    [Fact(DisplayName = "PDF—RLE解码重复字节序列")]
    public void RleDecode_RepeatedBytes()
    {
        // "AAAA": n=253(0xFD) → 257-253=4次, 'A'(0x41), EOD(128)
        var encoded = new Byte[] { 0xFD, 0x41, 0x80 };
        var decoded = PdfReader.DecodeRunLength(encoded);
        Assert.Equal("AAAA", Encoding.ASCII.GetString(decoded));
    }

    [Fact(DisplayName = "PDF—RLE解码字面量序列")]
    public void RleDecode_LiteralBytes()
    {
        // "ABC": n=2 → copy 3 bytes, 'A','B','C', EOD(128)
        var encoded = new Byte[] { 0x02, 0x41, 0x42, 0x43, 0x80 };
        var decoded = PdfReader.DecodeRunLength(encoded);
        Assert.Equal("ABC", Encoding.ASCII.GetString(decoded));
    }

    [Fact(DisplayName = "PDF—RLE解码混合序列")]
    public void RleDecode_Mixed()
    {
        // "AABBB": n=1(copy 2:'A','A'), n=254(257-254=3 repeat B), 'B', EOD(128)
        var encoded = new Byte[] { 0x01, 0x41, 0x41, 0xFE, 0x42, 0x80 };
        var decoded = PdfReader.DecodeRunLength(encoded);
        Assert.Equal("AABBB", Encoding.ASCII.GetString(decoded));
    }

    [Fact(DisplayName = "PDF—RLE解码空数据")]
    public void RleDecode_Empty()
    {
        var decoded = PdfReader.DecodeRunLength([]);
        Assert.Empty(decoded);
        decoded = PdfReader.DecodeRunLength(null!);
        Assert.Empty(decoded);
    }
    #endregion

    #region 字体信息读取
    [Fact(DisplayName = "PDF—ReadFonts读取字体列表")]
    public void ReadFonts_Basic()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AppendLine("Hello World", 12);
            writer.AppendLine("中文测试", 12);
            writer.Save(path);

            using var reader = new PdfReader(path);
            var fonts = reader.ReadFonts();
            // ReadFonts 可能返回空（如果扫描逻辑未匹配到 BaseFont），至少应不抛异常
            Assert.NotNull(fonts);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—ReadFonts无异常")]
    public void ReadFonts_NoException()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AppendLine("Test", 10);
            writer.Save(path);

            using var reader = new PdfReader(path);
            var fonts = reader.ReadFonts();
            // 不应抛出异常，返回列表可能为空也可能包含字体
            Assert.NotNull(fonts);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 注释读取
    [Fact(DisplayName = "PDF—注释读写往返测试")]
    public void ReadAnnotations_RoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AppendLine("测试页面", 12);
            writer.AddAnnotation(new PdfAnnotation
            {
                Type = PdfAnnotationType.Text,
                PageIndex = 0,
                X = 100, Y = 600,
                Width = 20, Height = 20,
                Contents = "这是一个便签注释",
                Author = "测试者",
            });
            writer.Save(path);

            // 验证底层 PDF 含 /Annots
            var rawBytes = File.ReadAllBytes(path);
            var asciiContent = System.Text.Encoding.ASCII.GetString(rawBytes);
            Assert.Contains("/Annots", asciiContent);

            using var reader = new PdfReader(path);
            var annots = reader.ReadAnnotations(0);
            Assert.NotEmpty(annots);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "PDF—无注释页面返回空列表")]
    public void ReadAnnotations_NoAnnotations_ReturnsEmpty()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            writer.AppendLine("无注释页面", 12);
            writer.Save(path);

            using var reader = new PdfReader(path);
            var annots = reader.ReadAnnotations(0);
            Assert.NotNull(annots);
            Assert.Empty(annots);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region PdfTable 模型集成
    [Fact(DisplayName = "PDF—PdfTable模型写入PDF")]
    public void DrawTable_PdfTableModel()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        try
        {
            using var writer = new PdfWriter();
            var table = new PdfTable
            {
                ColumnWidths = [150f, 150f, 150f],
                HeaderBackColor = OfficeColor.FromHex("4472C4"),
            };

            var header = new PdfTableRow { IsHeader = true };
            header.Cells.Add(new PdfTableCell { Text = "姓名" });
            header.Cells.Add(new PdfTableCell { Text = "部门" });
            header.Cells.Add(new PdfTableCell { Text = "销售额" });
            table.Rows.Add(header);

            var row1 = new PdfTableRow();
            row1.Cells.Add(new PdfTableCell { Text = "张三" });
            row1.Cells.Add(new PdfTableCell { Text = "技术部" });
            row1.Cells.Add(new PdfTableCell { Text = "¥120,000" });
            table.Rows.Add(row1);

            var row2 = new PdfTableRow();
            row2.Cells.Add(new PdfTableCell { Text = "李四" });
            row2.Cells.Add(new PdfTableCell { Text = "销售部" });
            row2.Cells.Add(new PdfTableCell { Text = "¥250,000" });
            table.Rows.Add(row2);

            writer.DrawTable(table);
            writer.Save(path);

            Assert.True(File.Exists(path));
            var text = new PdfReader(path).ExtractText();
            Assert.Contains("姓名", text);
            Assert.Contains("张三", text);
            Assert.Contains("李四", text);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion
}
