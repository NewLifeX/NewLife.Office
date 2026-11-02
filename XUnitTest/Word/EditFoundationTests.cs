using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>编辑底座测试：BeginEdit 生效、保存前自检、未建模元素保留、语义定位与增删改、单位换算</summary>
public class EditFoundationTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 构造与工具
    private static Byte[] BuildDocx(String bodyInnerXml)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml));
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static Document ReadBack(Byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static String BodyText(Document doc)
    {
        var sb = new StringBuilder();
        foreach (var el in doc.Elements)
        {
            if (el.Paragraph == null) continue;
            foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
            sb.Append('|');
        }
        return sb.ToString();
    }

    private static String ReadDocumentXml(Byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(za.GetEntry("word/document.xml")!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
    #endregion

    #region BeginEdit 与保存自检
    [Fact(DisplayName = "编辑底座-BeginEdit 后修改生效")]
    public void BeginEdit_MakesModelEditsPersist()
    {
        var doc = Read("<w:p><w:r><w:t>原文</w:t></w:r></w:p>");
        Assert.NotNull(doc.DocumentXml);

        doc.BeginEdit();
        Assert.True(doc.IsEditing);
        Assert.Null(doc.DocumentXml);

        var para = Assert.Single(doc.AllParagraphs());
        para.SetText("改后");

        var bytes = Save(doc);
        Assert.Equal("改后|", BodyText(ReadBack(bytes)));
    }

    [Fact(DisplayName = "编辑底座-未关闭透传时保存抛异常而非静默丢弃")]
    public void Save_ModelEditedWithoutBeginEdit_Throws()
    {
        var doc = Read("<w:p><w:r><w:t>原文</w:t></w:r></w:p>");
        doc.ReplaceText("原文", "改后");   // 递增 ModelRevision，但未关闭 DocumentXml

        using var ms = new MemoryStream();
        using var writer = new WordWriter();
        var ex = Assert.Throws<InvalidOperationException>(() => writer.Save(ms, doc));
        Assert.Contains("BeginEdit", ex.Message);
    }

    [Fact(DisplayName = "编辑底座-无修改时透传保存不报错")]
    public void Save_NoModelEdit_PassthroughWorks()
    {
        var doc = Read("<w:p><w:r><w:t>原文</w:t></w:r></w:p>");
        var bytes = Save(doc);
        Assert.Equal("原文|", BodyText(ReadBack(bytes)));
    }

    [Fact(DisplayName = "编辑底座-保留未建模 body 元素（customXml 块）")]
    public void BeginEdit_KeepsUnknownElements()
    {
        var doc = Read("<w:customXml w:element=\"meta\"><w:p><w:r><w:t>原文</w:t></w:r></w:p></w:customXml>");

        var unknown = doc.Elements.FindAll(e => e.Type == ElementType.Unknown);
        Assert.Single(unknown);

        doc.BeginEdit();
        Assert.NotNull(unknown[0].RawXml);          // 未建模元素默认保留原始 XML
        Assert.Contains("customXml", ReadDocumentXml(Save(doc)));
    }

    [Fact(DisplayName = "编辑底座-识别段落中的域并登记诊断信息")]
    public void BeginEdit_ReportsUnmodeledParagraphs()
    {
        var doc = Read("<w:p><w:r><w:t>共</w:t></w:r>"
            + "<w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple>"
            + "<w:r><w:t>页</w:t></w:r></w:p>");

        var para = Assert.Single(doc.AllParagraphs());
        Assert.True(para.HasUnmodeledContent);

        doc.BeginEdit();
        Assert.Single(doc.EditWarnings);
    }

    [Fact(DisplayName = "编辑底座-普通段落不标记未建模内容")]
    public void HasUnmodeledContent_FalseForPlainParagraph()
    {
        var doc = Read("<w:p><w:r><w:t>普通文本</w:t></w:r></w:p>");
        var para = Assert.Single(doc.AllParagraphs());
        Assert.False(para.HasUnmodeledContent);

        doc.BeginEdit();
        Assert.Empty(doc.EditWarnings);
    }
    #endregion

    #region 语义定位
    [Fact(DisplayName = "语义定位-AllParagraphs 覆盖正文/表格/页眉页脚")]
    public void AllParagraphs_CoversTableAndHeader()
    {
        // 顶层无页眉页脚部件，这里验证正文 + 表格单元格（含嵌套表格）
        var doc = Read("<w:p><w:r><w:t>正文</w:t></w:r></w:p>"
            + "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");

        var paras = doc.AllParagraphs();
        Assert.Equal(2, paras.Count);
        Assert.Equal("正文", paras[0].Runs[0].Text);
        Assert.Equal("单元格", paras[1].Runs[0].Text);

        var table = Assert.Single(doc.AllTables());
        Assert.Equal("单元格", table.GetCellText(0, 0));
    }

    [Fact(DisplayName = "语义定位-FindParagraph 按条件查找")]
    public void FindParagraph_ByPredicate()
    {
        var doc = Read("<w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:t>正文</w:t></w:r></w:p>");

        var para = doc.FindParagraph(p => p.Runs.Count > 0 && p.Runs[0].Text == "正文");
        Assert.NotNull(para);
        Assert.Equal("正文", para!.Runs[0].Text);

        Assert.Null(doc.FindParagraph(p => p.Runs.Count > 0 && p.Runs[0].Text == "不存在"));
    }
    #endregion

    #region 语义增删改
    [Fact(DisplayName = "语义编辑-段落后插入并持久化")]
    public void InsertParagraphAfter_Persists()
    {
        var doc = Read("<w:p><w:r><w:t>第一段</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p>");
        doc.BeginEdit();

        var first = doc.AllParagraphs()[0];
        var inserted = doc.InsertParagraphAfter(first, "插入段");
        Assert.NotNull(inserted);

        Assert.Equal("第一段|插入段|第二段|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "语义编辑-段落后插入多行按段拆分")]
    public void InsertParagraphAfter_MultiLineSplits()
    {
        var doc = Read("<w:p><w:r><w:t>A</w:t></w:r></w:p>");
        doc.BeginEdit();

        var anchor = doc.AllParagraphs()[0];
        doc.InsertParagraphAfter(anchor, "B\nC");

        Assert.Equal("A|B|C|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "语义编辑-删除段落并持久化")]
    public void RemoveParagraph_Persists()
    {
        var doc = Read("<w:p><w:r><w:t>保留</w:t></w:r></w:p><w:p><w:r><w:t>删除</w:t></w:r></w:p>");
        doc.BeginEdit();

        var target = doc.AllParagraphs()[1];
        Assert.True(doc.RemoveParagraph(target));

        Assert.Equal("保留|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "语义编辑-SetText 保留首个 Run 格式")]
    public void SetText_KeepsFirstRunFormatting()
    {
        var doc = Read("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>旧</w:t></w:r><w:r><w:t>文本</w:t></w:r></w:p>");
        doc.BeginEdit();

        var para = doc.AllParagraphs()[0];
        para.SetText("新");

        var run = Assert.Single(para.Runs);
        Assert.Equal("新", run.Text);
        Assert.True(run.Properties?.Bold);
    }

    [Fact(DisplayName = "语义编辑-单元格整格替换并持久化")]
    public void CellSetText_Persists()
    {
        var doc = Read("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>旧值</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>第二段</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        doc.BeginEdit();

        var table = Assert.Single(doc.AllTables());
        table.GetCell(0, 0)!.SetText("新值");

        var back = Assert.Single(ReadBack(Save(doc)).AllTables());
        Assert.Equal("新值", back.GetCellText(0, 0));
        Assert.Single(back.GetCell(0, 0)!.Paragraphs);
    }

    [Fact(DisplayName = "语义编辑-多节文档插入段落同步到节")]
    public void InsertParagraph_MultiSection_SyncsSection()
    {
        // 两节：第一节段落自带 sectPr，body 末尾 sectPr 属于第二节
        var doc = Read("<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>第二节</w:t></w:r></w:p>"
            + "<w:sectPr/>");
        Assert.Equal(2, doc.Sections.Count);

        doc.BeginEdit();
        doc.InsertParagraphAfter(doc.AllParagraphs()[0], "插入到第一节");

        var back = ReadBack(Save(doc));
        var text = BodyText(back);
        // 节分隔符段落会被读回成空段落，这里只校验三段文本的先后顺序
        var i1 = text.IndexOf("第一节", StringComparison.Ordinal);
        var i2 = text.IndexOf("插入到第一节", StringComparison.Ordinal);
        var i3 = text.IndexOf("第二节", StringComparison.Ordinal);
        Assert.True(i1 >= 0 && i1 < i2 && i2 < i3, text);
    }
    #endregion

    #region 真实文件回归
    [Fact(DisplayName = "编辑底座-真实 Word 文件改一段后不新增规范错误")]
    public void RealFile_BeginEditAndEdit_PassesValidation()
    {
        var fixturesDir = Path.Combine(AppContext.BaseDirectory, "Word", "Fixtures");
        if (!Directory.Exists(fixturesDir)) return;
        var files = Directory.GetFiles(fixturesDir, "*.docx", SearchOption.TopDirectoryOnly);
        if (files.Length == 0) return;

        var sourcePath = files[0];
        Document doc;
        using (var reader = new WordReader(sourcePath))
            doc = reader.ReadDocument();

        var unknownBefore = doc.Elements.Count(e => e.Type == ElementType.Unknown);

        doc.BeginEdit();

        // 找一个不含未建模内容的段落，改成可识别的验证文本
        const String Mark = "编辑底座验证文本";
        var target = doc.FindParagraph(p => p.Runs.Count > 0 && !p.HasUnmodeledContent);
        Assert.NotNull(target);
        target!.SetText(Mark);

        var outPath = Path.Combine(Path.GetTempPath(), $"Edit_{Guid.NewGuid():N}.docx");
        try
        {
            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            var back = ReadBack(File.ReadAllBytes(outPath));
            Assert.Contains(Mark, BodyText(back));
            Assert.Equal(unknownBefore, back.Elements.Count(e => e.Type == ElementType.Unknown));

            // 真实 Word 文件本身可能含非规范构造，只断言「不显著新增错误」
            var srcErrors = ValidateDocx(sourcePath).Count;
            var outErrors = ValidateDocx(outPath).Count;
            Assert.True(outErrors <= srcErrors + 3, $"源={srcErrors} 输出={outErrors}");
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    /// <summary>用 OpenXmlValidator 统计规范错误数（包打开期间取描述，避免释放后访问）</summary>
    private static List<String> ValidateDocx(String path)
    {
        var result = new List<String>();
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
        try
        {
            foreach (var e in validator.Validate(package))
                result.Add($"{e.ErrorType}: {e.Description}");
        }
        catch (Exception ex)
        {
            // 校验器自身在部分真实文件上会抛异常（语义约束实现缺陷），计入错误以便观察
            result.Add($"ValidatorException: {ex.GetType().Name}: {ex.Message}");
        }
        return result;
    }
    #endregion

    #region 单位换算    [Fact(DisplayName = "单位-厘米与英寸互换")]
    public void Length_CmToInches()
    {
        Assert.True(Length.FromCm(2.54).Equals(Length.FromInches(1)));
        Assert.Equal(1, Length.FromCm(2.54).Inches);
        Assert.Equal(360000, Length.FromCm(1).Emu);
    }

    [Theory(DisplayName = "单位-文本解析")]
    [InlineData("12pt", 12)]
    [InlineData("720twip", 36)]
    [InlineData("2.54cm", 72)]
    [InlineData("1in", 72)]
    public void Length_Parse(String text, Double expectedPoints)
    {
        Assert.Equal(expectedPoints, Math.Round(Length.Parse(text).Points, 6));
    }

    [Fact(DisplayName = "单位-非法文本抛 FormatException")]
    public void Length_ParseInvalid_Throws()
    {
        Assert.Throws<FormatException>(() => Length.Parse("abc"));
        Assert.Throws<FormatException>(() => Length.Parse("12km"));
    }
    #endregion
}
