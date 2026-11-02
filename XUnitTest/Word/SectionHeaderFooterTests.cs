using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>各节独立页眉页脚（取消“链接到上一节”）测试</summary>
public class SectionHeaderFooterTests
{
    #region 辅助
    private static Byte[] Save(Document doc)
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
            writer.Save(ms, doc);
        return ms.ToArray();
    }

    private static String? ReadEntry(Byte[] bytes, String name)
    {
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = za.GetEntry(name);
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    private static Boolean HasEntry(Byte[] bytes, String name)
    {
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        return za.GetEntry(name) != null;
    }

    private static Document Read(Byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private static Footer BuildFooter(String text) => new()
    {
        Type = "default",
        Elements = [new Element { Type = ElementType.Paragraph, Paragraph = new Paragraph { Runs = { new Run { Text = text } } } }],
    };

    private static Element Paragraph(String text) => new() { Type = ElementType.Paragraph, Paragraph = new Paragraph { Runs = { new Run { Text = text } } } };

    /// <summary>构造两节文档</summary>
    private static Document BuildTwoSections()
    {
        var doc = new Document();
        doc.AddParagraph("第一节内容");
        doc.AddSection("nextPage");
        doc.AddParagraph("第二节内容");
        return doc;
    }
    #endregion

    [Fact(DisplayName = "独立页脚-取消链接后生成独立部件")]
    public void UnlinkedFooter_CreatesOwnPart()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));
        doc.Sections[1].LinkFootersToPrevious = false;
        doc.Sections[1].Footers.Add(BuildFooter("第二节页脚"));

        var bytes = Save(doc);

        // 两个页脚部件：footer1=文档级，footer2=第二节独立
        Assert.True(HasEntry(bytes, "word/footer1.xml"));
        Assert.True(HasEntry(bytes, "word/footer2.xml"));
        Assert.Contains("全文页脚", ReadEntry(bytes, "word/footer1.xml")!);
        Assert.Contains("第二节页脚", ReadEntry(bytes, "word/footer2.xml")!);

        // 第二节的 sectPr 引用 rFtr2
        var documentXml = ReadEntry(bytes, "word/document.xml")!;
        Assert.Contains("rFtr2", documentXml);
        Assert.Contains("rFtr1", documentXml);
    }

    [Fact(DisplayName = "独立页脚-内容为空时生成空部件以清空继承")]
    public void UnlinkedFooter_Empty_ClearsViaEmptyPart()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));
        doc.Sections[1].LinkFootersToPrevious = false;   // 断链但不提供内容

        var bytes = Save(doc);

        Assert.True(HasEntry(bytes, "word/footer2.xml"));
        var xml = ReadEntry(bytes, "word/footer2.xml")!;
        Assert.DoesNotContain("<w:t", xml);      // 空页脚：没有任何文本
        Assert.Contains("rFtr2", ReadEntry(bytes, "word/document.xml")!);
    }

    [Fact(DisplayName = "默认链接-第二节不再生成独立部件")]
    public void LinkedByDefault_NoOwnPart()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));

        var bytes = Save(doc);

        Assert.True(HasEntry(bytes, "word/footer1.xml"));
        Assert.False(HasEntry(bytes, "word/footer2.xml"));
    }

    [Fact(DisplayName = "独立页眉-取消链接后生成独立页眉部件")]
    public void UnlinkedHeader_CreatesOwnPart()
    {
        var doc = BuildTwoSections();
        doc.Headers.Add(new Header { Type = "default", Elements = [Paragraph("全文页眉")] });
        doc.Sections[1].LinkHeadersToPrevious = false;
        doc.Sections[1].Headers.Add(new Header { Type = "default", Elements = [Paragraph("第二节页眉")] });

        var bytes = Save(doc);

        Assert.Contains("全文页眉", ReadEntry(bytes, "word/header1.xml")!);
        Assert.Contains("第二节页眉", ReadEntry(bytes, "word/header2.xml")!);
        Assert.Contains("rHdr2", ReadEntry(bytes, "word/document.xml")!);

        // 内容类型与关系需齐全，否则 Word 打不开
        Assert.Contains("/word/header2.xml", ReadEntry(bytes, "[Content_Types].xml")!);
        Assert.Contains("header2.xml", ReadEntry(bytes, "word/_rels/document.xml.rels")!);
    }

    [Fact(DisplayName = "独立页脚-保存后读回链接标志与内容")]
    public void UnlinkedFooter_RoundTrip()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));
        doc.Sections[1].LinkFootersToPrevious = false;
        doc.Sections[1].Footers.Add(BuildFooter("第二节页脚"));

        var back = Read(Save(doc));

        Assert.Equal(2, back.Sections.Count);
        Assert.False(back.Sections[1].LinkFootersToPrevious);
        Assert.True(back.Sections[0].LinkFootersToPrevious);
        var footer = Assert.Single(back.Sections[1].Footers);
        Assert.Equal("第二节页脚", footer.Elements[0].Paragraph!.Runs[0].Text);
    }

    [Fact(DisplayName = "默认链接-读回后第二节仍为链接状态")]
    public void LinkedByDefault_RoundTrip()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));

        var back = Read(Save(doc));

        Assert.Equal(2, back.Sections.Count);
        Assert.True(back.Sections[1].LinkFootersToPrevious);
        Assert.Empty(back.Sections[1].Footers);
    }

    [Fact(DisplayName = "单节文档-链接标志默认且不生成额外部件")]
    public void SingleSection_NoExtraPart()
    {
        var doc = new Document();
        doc.AddParagraph("只有一节");
        doc.Footers.Add(BuildFooter("页脚"));

        var bytes = Save(doc);

        Assert.True(HasEntry(bytes, "word/footer1.xml"));
        Assert.False(HasEntry(bytes, "word/footer2.xml"));

        var back = Read(bytes);
        Assert.True(back.Sections[0].LinkFootersToPrevious);
    }

    [Fact(DisplayName = "第一节独立页脚-显式设置本节内容后生效")]
    public void FirstSection_OwnFooterUsed()
    {
        var doc = BuildTwoSections();
        doc.Footers.Add(BuildFooter("全文页脚"));
        // 第一节没有“上一节”，但显式给本节页脚时不应被静默忽略
        doc.Sections[0].Footers.Add(BuildFooter("第一节专属页脚"));

        var bytes = Save(doc);

        Assert.True(HasEntry(bytes, "word/footer2.xml"));
        Assert.Contains("第一节专属页脚", ReadEntry(bytes, "word/footer2.xml")!);
        Assert.Contains("全文页脚", ReadEntry(bytes, "word/footer1.xml")!);
        // 第一节的 sectPr 引用本节独立部件
        Assert.Contains("rFtr2", ReadEntry(bytes, "word/document.xml")!);
    }

    [Fact(DisplayName = "多节文档-首节页眉读取后编辑保存仍保留且引用有效")]
    public void MultiSection_FirstSectionHeader_SurvivesEditAndSave()
    {
        // 源文档：首节有页眉，末节（body 级 sectPr）没有任何页眉
        var source = BuildTwoSectionDocxWithFirstHeader();

        Document doc;
        using (var ms = new MemoryStream(source))
        using (var reader = new WordReader(ms))
            doc = reader.ReadDocument();

        // 首节的页眉必须能被读到（不能恒为空）
        Assert.True(doc.Sections.Count > 1);
        var firstHeader = Assert.Single(doc.Sections[0].Headers);
        Assert.Equal("第一节页眉", firstHeader.Elements[0].Paragraph!.Runs[0].Text);

        // 编辑正文后保存
        doc.BeginEdit();
        doc.Elements[0].Paragraph!.SetText("首节正文已改");
        var bytes = Save(doc);

        // 首节页眉仍在
        var back = Read(bytes);
        var backHeader = Assert.Single(back.Sections[0].Headers);
        Assert.Equal("第一节页眉", backHeader.Elements[0].Paragraph!.Runs[0].Text);

        // 引用不悬空：document.xml 里出现的每个 r:id 都要在 rels 中存在
        var documentXml = ReadEntry(bytes, "word/document.xml")!;
        var rels = ReadEntry(bytes, "word/_rels/document.xml.rels")!;
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(documentXml, "r:id=\\\"([^\\\"]+)\\\""))
        {
            var id = m.Groups[1].Value;
            Assert.True(rels.Contains($"Id=\"{id}\""), $"引用 {id} 在 rels 中不存在（悬空引用）");
        }
    }

    /// <summary>构造两节 docx：首节 sectPr 带页眉，末节 sectPr 无页眉</summary>
    private static Byte[] BuildTwoSectionDocxWithFirstHeader()
    {
        const String NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{NS}\" xmlns:r=\"{R}\"><w:body>"
            + "<w:p><w:r><w:t>首节正文</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>"
            + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\"/>"
            + "</w:sectPr></w:pPr></w:p>"
            + "<w:p><w:r><w:t>次节正文</w:t></w:r></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\"/></w:sectPr>"
            + "</w:body></w:document>";
        var headerXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:hdr xmlns:w=\"{NS}\"><w:p><w:r><w:t>第一节页眉</w:t></w:r></w:p></w:hdr>";

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZip(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
                + "</Types>");
            WriteZip(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteZip(za, "word/document.xml", documentXml);
            WriteZip(za, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
                + "</Relationships>");
            WriteZip(za, "word/header1.xml", headerXml);
        }
        return ms.ToArray();
    }

    private static void WriteZip(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }
}
