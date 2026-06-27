using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>WordWriter/WordReader 单元测试 — 分栏/SDT/往返保真</summary>
public class WordWriterTests
{
    #region 分栏（Columns）
    [Fact(DisplayName = "分栏写入—2栏文档")]
    public void WriteWithColumns_TwoColumns()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.PageSettings.ColumnCount = 2;
            writer.PageSettings.ColumnSpacing = 720;
            writer.AppendParagraph("第一栏内容");
            writer.AppendParagraph("第二栏也会看到这些文字");
            writer.Save(tempFile);

            Assert.True(File.Exists(tempFile));
            // 读取回验证
            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(2, doc.PageSettings.ColumnCount);
            Assert.Equal(720, doc.PageSettings.ColumnSpacing);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "分栏写入—3栏文档")]
    public void WriteWithColumns_ThreeColumns()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.PageSettings.ColumnCount = 3;
            writer.PageSettings.ColumnSpacing = 360;
            writer.AppendParagraph("三栏排版");
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(3, doc.PageSettings.ColumnCount);
            Assert.Equal(360, doc.PageSettings.ColumnSpacing);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "分栏写入—默认单栏")]
    public void WriteWithoutColumns_DefaultSingleColumn()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.AppendParagraph("普通文档");
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(1, doc.PageSettings.ColumnCount);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 分栏+其他页面设置组合
    [Fact(DisplayName = "分栏—含分栏的完整页面设置往返")]
    public void Columns_WithFullPageSettings_RoundTrip()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.PageSettings.Landscape = true;
            writer.PageSettings.ColumnCount = 2;
            writer.PageSettings.MarginLeft = 2000;
            writer.PageSettings.MarginRight = 2000;
            writer.AppendHeading("横向两栏文档", 1);
            writer.AppendParagraph("在横向页面上使用两栏布局。");
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(2, doc.PageSettings.ColumnCount);
            Assert.True(doc.PageSettings.Landscape);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region SDT 内容控件
    [Fact(DisplayName = "SDT—构建含纯文本控件的 docx 并读取")]
    public void Sdt_ReadPlainText()
    {
        // 手工构建含 w:sdt 的 document.xml
        var docxContent = BuildDocxWithSdt("<w:sdt>" +
            "<w:sdtPr><w:tag w:val=\"Name\"/><w:alias w:val=\"姓名\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>张三</w:t></w:r></w:p></w:sdtContent>" +
            "</w:sdt>");

        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            File.WriteAllBytes(tempFile, docxContent);
            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.SdtElements);
            Assert.Equal("Name", doc.SdtElements[0].Tag);
            Assert.Equal("姓名", doc.SdtElements[0].Alias);
            Assert.Contains("张三", doc.SdtElements[0].Content);
            Assert.Equal(WordSdtType.PlainText, doc.SdtElements[0].SdtType);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "SDT—读取日期控件类型")]
    public void Sdt_ReadDateType()
    {
        var docxContent = BuildDocxWithSdt("<w:sdt>" +
            "<w:sdtPr><w:tag w:val=\"SignDate\"/><w:date/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>2026-06-27</w:t></w:r></w:p></w:sdtContent>" +
            "</w:sdt>");

        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            File.WriteAllBytes(tempFile, docxContent);
            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.SdtElements);
            Assert.Equal(WordSdtType.Date, doc.SdtElements[0].SdtType);
            Assert.Contains("2026-06-27", doc.SdtElements[0].Content);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "SDT—读取下拉列表控件类型")]
    public void Sdt_ReadDropDownList()
    {
        var docxContent = BuildDocxWithSdt("<w:sdt>" +
            "<w:sdtPr><w:tag w:val=\"Dept\"/><w:dropDownList w:lastValue=\"技术部\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>技术部</w:t></w:r></w:p></w:sdtContent>" +
            "</w:sdt>");

        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            File.WriteAllBytes(tempFile, docxContent);
            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.SdtElements);
            Assert.Equal(WordSdtType.DropDownList, doc.SdtElements[0].SdtType);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "SDT—无 sdtPr 时默认为纯文本")]
    public void Sdt_NoPr_DefaultPlainText()
    {
        var docxContent = BuildDocxWithSdt("<w:sdt>" +
            "<w:sdtContent><w:p><w:r><w:t>默认类型</w:t></w:r></w:p></w:sdtContent>" +
            "</w:sdt>");

        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            File.WriteAllBytes(tempFile, docxContent);
            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.SdtElements);
            Assert.Equal(WordSdtType.PlainText, doc.SdtElements[0].SdtType);
            Assert.Contains("默认类型", doc.SdtElements[0].Content);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 分页符格式保留
    [Fact(DisplayName = "分页符—带有格式的段落加分页符后格式保留")]
    public void PageBreak_FormatPreservation()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.AppendHeading("第一章", 1);
            writer.AppendParagraph("这是第一章的内容。", WordParagraphStyle.Normal,
                new WordRunProperties { Bold = true, FontSize = 12 });
            writer.AppendPageBreak();
            writer.AppendHeading("第二章", 2);
            writer.AppendParagraph("这是第二章的内容。");
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.True(doc.Elements.Count >= 4);
            var heading1 = doc.Elements[0];
            Assert.Equal(WordElementType.Paragraph, heading1.Type);
            Assert.Equal(WordParagraphStyle.Heading1, heading1.Paragraph!.Style);
            var pageBreak = doc.Elements[2];
            Assert.Equal(WordElementType.Paragraph, pageBreak.Type);
            Assert.True(pageBreak.Paragraph!.IsPageBreak);
            var heading2 = doc.Elements[3];
            Assert.Equal(WordElementType.Paragraph, heading2.Type);
            Assert.Equal(WordParagraphStyle.Heading2, heading2.Paragraph!.Style);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 多级嵌套列表
    [Fact(DisplayName = "多级列表—2级嵌套写入+读取")]
    public void MultiLevelBulletList_WriteAndRead()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.AppendMultiLevelBulletList([
                ("一级项目A", 0),
                ("二级项目1", 1),
                ("二级项目2", 1),
                ("一级项目B", 0),
                ("三级项目x", 2),
            ]);
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(5, doc.Elements.Count);
            var level0 = doc.Elements[0].Paragraph!;
            Assert.True(level0.IsBullet);
            Assert.Equal(0, level0.ListLevel);
            var level1 = doc.Elements[1].Paragraph!;
            Assert.True(level1.IsBullet);
            Assert.Equal(1, level1.ListLevel);
            var level2 = doc.Elements[4].Paragraph!;
            Assert.True(level2.IsBullet);
            Assert.Equal(2, level2.ListLevel);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "多级列表—默认级别为0")]
    public void MultiLevelBulletList_DefaultLevel()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.AppendBulletList(["普通项目1", "普通项目2"]);
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.All(doc.Elements, e => Assert.Equal(0, e.Paragraph!.ListLevel));
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "多级列表—与普通段落混合")]
    public void MultiLevelBulletList_MixedWithParagraphs()
    {
        var tempFile = Path.GetTempFileName() + ".docx";
        try
        {
            using var writer = new WordWriter();
            writer.AppendHeading("章节标题", 1);
            writer.AppendParagraph("正文段落");
            writer.AppendMultiLevelBulletList([
                ("要点一", 0),
                ("要点一-1", 1),
                ("要点二", 0),
            ]);
            writer.AppendParagraph("结束段落");
            writer.Save(tempFile);

            using var reader = new WordReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal(6, doc.Elements.Count);
            Assert.Equal(WordParagraphStyle.Heading1, doc.Elements[0].Paragraph!.Style);
            Assert.False(doc.Elements[1].Paragraph!.IsBullet);
            Assert.True(doc.Elements[2].Paragraph!.IsBullet);
            Assert.False(doc.Elements[5].Paragraph!.IsBullet);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 辅助方法
    /// <summary>构建包含指定 SDT 元素的最小合法 docx 文件</summary>
    private static Byte[] BuildDocxWithSdt(String sdtXml)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\">" +
            "<w:body>" +
            sdtXml +
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"720\" w:footer=\"720\"/></w:sectPr>" +
            "</w:body></w:document>";

        var contentTypeXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "</Types>";

        var relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>";

        var docRelsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "</Relationships>";

        using var ms = new MemoryStream();
        using (var za = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
        {
            WriteZipEntry(za, "[Content_Types].xml", contentTypeXml);
            WriteZipEntry(za, "_rels/.rels", relsXml);
            WriteZipEntry(za, "word/document.xml", documentXml);
            WriteZipEntry(za, "word/_rels/document.xml.rels", docRelsXml);
        }
        return ms.ToArray();
    }

    private static void WriteZipEntry(System.IO.Compression.ZipArchive za, String path, String content)
    {
        var entry = za.CreateEntry(path, System.IO.Compression.CompressionLevel.Optimal);
        using var sw = new StreamWriter(entry.Open(), System.Text.Encoding.UTF8);
        sw.Write(content);
    }
    #endregion
}
