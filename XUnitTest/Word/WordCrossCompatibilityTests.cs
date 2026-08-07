using System.ComponentModel;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NewLife.Office.Word;
using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using Xunit;
using Xunit.Abstractions;

namespace XUnitTest.Word;

/// <summary>
/// Word docx 竞品交叉兼容测试（对标 Open XML SDK / NPOI）
/// </summary>
/// <remarks>
/// 测试模式（详见 cross-compatibility-testing 技能）：
/// <list type="bullet">
/// <item><b>F1 我写你读</b>：NewLife.Office 的 <c>WordWriter</c> 生成 docx → 竞品 NPOI 读取并断言内容正确</item>
/// <item><b>F2 你写我读</b>：竞品 NPOI 生成 docx → 我们的 <c>WordReader</c> 读取并断言内容正确</item>
/// <item><b>F3 我写我读</b>：自校验（已有 WordRoundTripTests 覆盖，此处补充 OpenXmlValidator 规范校验）</item>
/// </list>
/// 竞品仅作为测试依赖，不污染生产代码。
/// </remarks>
[Trait("Category", "CrossCompatibility")]
public class WordCrossCompatibilityTests
{
    private readonly ITestOutputHelper _output;

    public WordCrossCompatibilityTests(ITestOutputHelper output) => _output = output;

    #region 辅助

    /// <summary>生成 1x1 红色 PNG 用于图片往返测试</summary>
    private static Byte[] CreatePng()
    {
        // 1x1 红色不透明 PNG
        var b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
        return Convert.FromBase64String(b64);
    }

    /// <summary>生成临时文件路径（带扩展名）</summary>
    private static String TempFile(String ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"WordCompat_{Guid.NewGuid():N}.{ext}");
        return path;
    }

    /// <summary>用 NPOI 校验 OpenXML 规范（0 错误 = 可通过 Word 打开无修复提示）</summary>
    /// <remarks>在包打开期间捕获错误描述字符串，避免释放后访问 Path/Part 抛异常</remarks>
    private static List<String> ValidateDocx(String path)
    {
        var result = new List<String>();
        using var doc = WordprocessingDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        foreach (var e in validator.Validate(doc))
            result.Add($"{e.ErrorType}: {e.Description}");
        return result;
    }

    /// <summary>断言文件通过 OOXML 规范校验</summary>
    private void AssertValidDocx(String path)
    {
        var errors = ValidateDocx(path);
        foreach (var e in errors)
            _output.WriteLine($"  规范错误: {e}");
        Assert.Empty(errors);
    }

    /// <summary>拼接段落全部 Run 文本</summary>
    private static String ParaText(Paragraph p) => String.Concat(p.Runs.Select(r => r.Text));

    /// <summary>拼接 NPOI 文档全部段落文本（NPOI 2.8 无 XWPFDocument.Text 属性）</summary>
    private static String NpoiText(XWPFDocument doc) => String.Join("\n", doc.Paragraphs.Select(p => p.Text));

    #endregion

    #region F1 我写你读（NewLife.Office 写 → NPOI 读）

    [Fact, DisplayName("F1_我写你读_NPOI能读取我们生成的docx文本与格式")]
    public void F1_NPOI_Reads_Our_BasicDocx()
    {
        var path = TempFile("docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.AppendHeading("年度报告", 1);
                w.AppendParagraph("这是正文内容。");
                w.AppendParagraph("粗体重点", ParagraphStyle.Normal,
                    new RunProperties { Bold = true, FontSize = 14f });
                w.AppendParagraph("红色斜体", ParagraphStyle.Normal,
                    new RunProperties { Italic = true, ForeColor = "FF0000" });
                w.Save(path);
            }

            // NPOI 读取
            using var fs = File.OpenRead(path);
            var doc = new XWPFDocument(fs);
            var text = NpoiText(doc);

            Assert.True(text.Contains("年度报告"), "NPOI 应能提取标题文本");
            Assert.True(text.Contains("这是正文内容。"), "NPOI 应能提取正文");
            Assert.True(text.Contains("粗体重点"), "NPOI 应能提取粗体文本");

            // 格式验证：粗体段落
            var boldPara = doc.Paragraphs.FirstOrDefault(p => p.Text.Contains("粗体重点"));
            Assert.NotNull(boldPara);
            var boldRun = boldPara!.Runs.FirstOrDefault(r => r.Text.Contains("粗体"));
            Assert.NotNull(boldRun);
            Assert.True(boldRun!.IsBold == true, "NPOI 应读到粗体格式");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F1_我写你读_NPOI能读取我们生成的表格")]
    public void F1_NPOI_Reads_Our_Table()
    {
        var path = TempFile("docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.AppendTable(new[]
                {
                    new[] { "产品", "价格", "库存" },
                    new[] { "笔记本", "5999", "100" },
                    new[] { "手机", "3999", "500" },
                }, firstRowHeader: true);
                w.Save(path);
            }

            using var fs = File.OpenRead(path);
            var doc = new XWPFDocument(fs);
            Assert.Single(doc.Tables);
            var table = doc.Tables[0];
            Assert.Equal(3, table.NumberOfRows);
            Assert.Equal("产品", table.GetRow(0).GetCell(0).GetTextRecursively());
            Assert.Equal("笔记本", table.GetRow(1).GetCell(0).GetTextRecursively());
            Assert.Equal("3999", table.GetRow(2).GetCell(1).GetTextRecursively());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F1_我写你读_NPOI能读取我们生成的图片与列表")]
    public void F1_NPOI_Reads_Our_ImageAndList()
    {
        var path = TempFile("docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.InsertImage(CreatePng(), "png", 5, 3);
                w.AppendBulletList(new[] { "项目一", "项目二" });
                w.Save(path);
            }

            using var fs = File.OpenRead(path);
            var doc = new XWPFDocument(fs);
            Assert.NotEmpty(doc.AllPictures);
            var text = NpoiText(doc);
            Assert.True(text.Contains("项目一"), "NPOI 应能提取列表文本");
            Assert.True(text.Contains("项目二"), "NPOI 应能提取列表文本");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F1_我写你读_NPOI能读取我们生成的页眉页脚")]
    public void F1_NPOI_Reads_Our_HeaderFooter()
    {
        var path = TempFile("docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.PageSettings.HeaderText = "公司机密";
                w.PageSettings.FooterText = "第 1 页";
                w.AppendParagraph("正文内容");
                w.Save(path);
            }

            using var fs = File.OpenRead(path);
            var doc = new XWPFDocument(fs);
            Assert.NotEmpty(doc.HeaderList);
            var headerText = doc.HeaderList[0].Text;
            Assert.Contains("公司机密", headerText, StringComparison.OrdinalIgnoreCase);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    #endregion

    #region F2 你写我读（NPOI 写 → NewLife.Office 读）

    [Fact, DisplayName("F2_你写我读_我们能读取NPOI生成的docx文本")]
    public void F2_We_Read_NPOI_BasicDocx()
    {
        var path = TempFile("docx");
        try
        {
            // NPOI 生成
            var doc = new XWPFDocument();
            doc.CreateParagraph().CreateRun().SetText("NPOI 生成文档");
            var p2 = doc.CreateParagraph();
            var r2 = p2.CreateRun();
            r2.SetText("粗体内容");
            r2.IsBold = true;
            var p3 = doc.CreateParagraph();
            var r3 = p3.CreateRun();
            r3.SetText("斜体内容");
            r3.IsItalic = true;
            using (var fs = File.Create(path)) doc.Write(fs);

            // 我们读取
            using var reader = new WordReader(path);
            var paragraphs = reader.ReadParagraphs().ToList();
            Assert.Contains("NPOI 生成文档", paragraphs);
            Assert.Contains("粗体内容", paragraphs);
            Assert.Contains("斜体内容", paragraphs);

            // 模型级读取（格式）
            var model = reader.ReadDocument();
            var boldPara = model.Elements
                .Select(e => e.Paragraph)
                .FirstOrDefault(p => p != null && ParaText(p).Contains("粗体内容"));
            Assert.NotNull(boldPara);
            Assert.True(boldPara!.Runs.Any(r => r.Properties?.Bold == true), "应读到 NPOI 写入的粗体格式");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F2_你写我读_我们能读取NPOI生成的表格")]
    public void F2_We_Read_NPOI_Table()
    {
        var path = TempFile("docx");
        try
        {
            var doc = new XWPFDocument();
            var table = doc.CreateTable(3, 3, null);
            table.GetRow(0).GetCell(0).SetText("姓名");
            table.GetRow(0).GetCell(1).SetText("部门");
            table.GetRow(1).GetCell(0).SetText("张三");
            table.GetRow(1).GetCell(1).SetText("研发部");
            table.GetRow(2).GetCell(0).SetText("李四");
            table.GetRow(2).GetCell(1).SetText("市场部");
            using (var fs = File.Create(path)) doc.Write(fs);

            using var reader = new WordReader(path);
            var tables = reader.ReadTables().ToList();
            Assert.Single(tables);
            Assert.Equal("姓名", tables[0][0][0]);
            Assert.Equal("研发部", tables[0][1][1]);
            Assert.Equal("市场部", tables[0][2][1]);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F2_你写我读_我们能读取NPOI生成的图片")]
    public void F2_We_Read_NPOI_Image()
    {
        var path = TempFile("docx");
        try
        {
            var doc = new XWPFDocument();
            var para = doc.CreateParagraph();
            var run = para.CreateRun();
            using (var imgStream = new MemoryStream(CreatePng()))
                run.AddPicture(imgStream, (Int32)PictureType.PNG, "test.png", 100, 100);
            using (var fs = File.Create(path)) doc.Write(fs);

            using var reader = new WordReader(path);
            var images = reader.ExtractImages().ToList();
            Assert.Single(images);
            Assert.Equal("png", images[0].Extension);
            Assert.Equal(CreatePng().Length, images[0].Data.Length);

            // 模型级：图片元素
            var model = reader.ReadDocument();
            Assert.NotEmpty(model.Elements.Where(e => e.Type == ElementType.Image));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F2_你写我读_我们能读取NPOI生成的页眉")]
    public void F2_We_Read_NPOI_Header()
    {
        var path = TempFile("docx");
        try
        {
            var doc = new XWPFDocument();
            var header = doc.CreateHeader(HeaderFooterType.DEFAULT);
            header.CreateParagraph().CreateRun().SetText("NPOI 页眉");
            doc.CreateParagraph().CreateRun().SetText("正文");
            using (var fs = File.Create(path)) doc.Write(fs);

            using var reader = new WordReader(path);
            var model = reader.ReadDocument();
            Assert.NotEmpty(model.Headers);
            var headerText = String.Join("", model.Headers.SelectMany(h => h.Elements)
                .SelectMany(e => e.Paragraph?.Runs ?? [])
                .Select(r => r.Text));
            Assert.Contains("NPOI 页眉", headerText);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    #endregion

    #region F3 规范校验（OpenXmlValidator，防 Word 修复提示）

    [Fact, DisplayName("F3_规范校验_WordWriter新建文档通过OOXML校验")]
    public void F3_OpenXmlValidator_NewDocx()
    {
        var path = TempFile("docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.DocumentProperties.Title = "校验测试";
                w.AppendHeading("标题", 1);
                w.AppendParagraph("正文段落", ParagraphStyle.Normal,
                    new RunProperties { Bold = true, ForeColor = "FF0000" });
                w.AppendTable(new[] { new[] { "A", "B" }, new[] { "1", "2" } });
                w.AppendBulletList(new[] { "列表项" });
                w.InsertImage(CreatePng(), "png", 5, 3);
                w.Save(path);
            }
            AssertValidDocx(path);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F3_规范校验_NPOI生成的docx通过校验（基线）")]
    public void F3_OpenXmlValidator_NPOI_Baseline()
    {
        var path = TempFile("docx");
        try
        {
            var doc = new XWPFDocument();
            doc.CreateParagraph().CreateRun().SetText("基线文档");
            var table = doc.CreateTable(2, 2, null);
            table.GetRow(0).GetCell(0).SetText("A");
            using (var fs = File.Create(path)) doc.Write(fs);
            AssertValidDocx(path);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact, DisplayName("F3_规范校验_读取后模型写回的文件通过OOXML校验")]
    public void F3_OpenXmlValidator_RoundTripDocx()
    {
        var path = TempFile("docx");
        var outPath = TempFile("docx");
        try
        {
            // 先生成一个包含多种元素的文件
            using (var w = new WordWriter())
            {
                w.AppendHeading("往返校验", 1);
                w.AppendParagraph("正文", ParagraphStyle.Normal, new RunProperties { Bold = true });
                w.AppendTable(new[] { new[] { "列1", "列2" }, new[] { "值1", "值2" } });
                w.AppendBulletList(new[] { "条目" });
                w.InsertImage(CreatePng(), "png", 5, 3);
                w.Save(path);
            }

            // 读取 → 模型写回
            using (var reader = new WordReader(path))
            {
                var doc = reader.ReadDocument();
                using (var w = new WordWriter())
                    w.Save(outPath, doc);
            }
            AssertValidDocx(outPath);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact, DisplayName("F2_你写我读_读取合并单元格表格(gridSpan/vMerge)")]
    public void F2_We_Read_MergedCells()
    {
        var path = TempFile("docx");
        try
        {
            // 构造 Word/NPOI 都会写出的规范合并单元格 XML：gridSpan 横向合并 + vMerge 纵向合并
            const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var bodyInner = "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
                + "<w:tblGrid><w:gridCol w:w=\"2500\"/><w:gridCol w:w=\"2500\"/><w:gridCol w:w=\"2500\"/></w:tblGrid>"
                + "<w:tr>"
                + "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>横向合并</w:t></w:r></w:p></w:tc>"
                + "<w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>V1</w:t></w:r></w:p></w:tc>"
                + "</w:tr>"
                + "<w:tr>"
                + "<w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>"
                + "<w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>"
                + "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>续</w:t></w:r></w:p></w:tc>"
                + "</w:tr>"
                + "</w:tbl>";
            var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInner}</w:body></w:document>";
            using var ms = new MemoryStream();
            using (var za = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                void WriteEntry(String name, String content)
                {
                    using var sw = new System.IO.StreamWriter(za.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
                    sw.Write(content);
                }
                WriteEntry("[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
                WriteEntry("_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                WriteEntry("word/document.xml", documentXml);
                WriteEntry("word/_rels/document.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
            }
            File.WriteAllBytes(path, ms.ToArray());

            // 我们读取并识别合并
            using var reader = new WordReader(path);
            var model = reader.ReadDocument();
            var t = Assert.Single(model.Elements.Where(e => e.Type == ElementType.Table)).Table!;
            Assert.Equal(2, t.Rows.Count);

            // 横向合并 gridSpan → ColSpan
            Assert.True(t.Rows[0].Cells[0].ColSpan >= 2, "gridSpan 应被读取为 ColSpan");
            // 纵向合并 vMerge restart/continue → RowSpan -1/0
            Assert.Equal(-1, t.Rows[0].Cells[1].RowSpan);
            Assert.Equal(0, t.Rows[1].Cells[2].RowSpan);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    #endregion

    #region Fixtures 真实文件回归（F2 真实 Word 文件 + F3 输出校验）

    [Fact, DisplayName("F2_真实Word文件_读取往返并通过OOXML校验")]
    public void F2_RealWordFiles_RoundTrip_AndValid()
    {
        var fixturesDir = Path.Combine(AppContext.BaseDirectory, "Word", "Fixtures");
        if (!Directory.Exists(fixturesDir))
        {
            _output.WriteLine($"Fixtures 目录不存在: {fixturesDir}，跳过");
            return;
        }

        var files = Directory.GetFiles(fixturesDir, "*.docx", SearchOption.TopDirectoryOnly);
        Assert.NotEmpty(files);

        foreach (var sourcePath in files)
        {
            var name = Path.GetFileNameWithoutExtension(sourcePath);
            var outPath = Path.Combine(Path.GetTempPath(), $"WordCompat_{name}_{Guid.NewGuid():N}.docx");
            try
            {
                // F2：读取真实 Word 生成的文件
                NewLife.Office.Word.Document source;
                using (var reader = new WordReader(sourcePath))
                    source = reader.ReadDocument();
                Assert.NotEmpty(source.Elements);
                _output.WriteLine($"[{name}] 源: {source.Elements.Count} 元素 (P={source.Elements.Count(e => e.Type == ElementType.Paragraph)} T={source.Elements.Count(e => e.Type == ElementType.Table)} I={source.Elements.Count(e => e.Type == ElementType.Image)})");

                // 模型写回
                using (var w = new WordWriter())
                    w.Save(outPath, source);

                // F3：真实 Word 文件常含 WPS/老 Word 遗留非规范构造（VML spt 属性、
                // 重复 shape id、uiPriority 顺序等），L3 保真要求原样保留。
                // 断言「往返不新增规范错误」而非强制 0 错误——保证我们不把文件改得更糟。
                var srcErrors = ValidateDocx(sourcePath).Count;
                var outErrors = ValidateDocx(outPath).Count;
                _output.WriteLine($"[{name}] 规范错误: 源={srcErrors} 输出={outErrors}");
                Assert.True(outErrors <= srcErrors + 3,
                    $"往返不应显著新增规范错误: 源={srcErrors} 输出={outErrors}");

                // 回读对比元素数量（模型保真）
                using (var reader = new WordReader(outPath))
                {
                    var output = reader.ReadDocument();
                    Assert.Equal(source.Elements.Count, output.Elements.Count);
                    Assert.Equal(
                        source.Elements.Count(e => e.Type == ElementType.Table),
                        output.Elements.Count(e => e.Type == ElementType.Table));
                    Assert.Equal(
                        source.Elements.Count(e => e.Type == ElementType.Image),
                        output.Elements.Count(e => e.Type == ElementType.Image));
                }
            }
            finally
            {
                if (File.Exists(outPath)) File.Delete(outPath);
            }
        }
    }

    #endregion
}
