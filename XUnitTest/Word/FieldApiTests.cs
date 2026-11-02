using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>域与目录 API 测试：Paragraph.AddField / Document.AddField / AddTableOfContents / updateFields</summary>
public class FieldApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 工具
    private static Byte[] BuildDocx(String bodyInnerXml, String? settingsInnerXml = null)
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
            if (settingsInnerXml != null)
            {
                WriteEntry(za, "word/settings.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + $"<w:settings xmlns:w=\"{W}\">{settingsInnerXml}</w:settings>");
            }
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml, String? settingsInnerXml = null)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml, settingsInnerXml));
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

    private static String ReadEntry(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = za.GetEntry(path);
        if (entry == null) return String.Empty;
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
    #endregion

    #region 域
    [Fact(DisplayName = "域-段落追加 PAGE 域输出复杂域结构")]
    public void AddField_Paragraph_WritesComplexField()
    {
        var doc = new Document();
        var el = doc.AddParagraph("第 ");
        el.Paragraph!.AddField("PAGE", "1");

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.Contains("w:fldCharType=\"begin\"", xml);
        Assert.Contains("PAGE", xml);
        Assert.Contains("w:fldCharType=\"separate\"", xml);
        Assert.Contains("w:fldCharType=\"end\"", xml);
        Assert.Contains("<w:t>1</w:t>", xml);
    }

    [Fact(DisplayName = "域-Document.AddField 生成独立段落域")]
    public void AddField_Document_WritesFieldParagraph()
    {
        var doc = new Document();
        doc.AddField("NUMPAGES", "3");

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.Contains("NUMPAGES", xml);
        Assert.Contains("<w:t>3</w:t>", xml);
    }

    [Fact(DisplayName = "域-空指令抛异常")]
    public void AddField_EmptyCode_Throws()
    {
        var doc = new Document();
        var el = doc.AddParagraph("x");
        Assert.Throws<ArgumentException>(() => el.Paragraph!.AddField(""));
        Assert.Throws<ArgumentException>(() => doc.AddField((String)null!));
    }

    [Fact(DisplayName = "域-带结果的域回读文本包含结果")]
    public void AddField_Result_ReadableAsText()
    {
        var doc = new Document();
        doc.AddField("PAGE", "7");

        using var ms = new MemoryStream(Save(doc));
        using var reader = new WordReader(ms);
        Assert.Contains("7", reader.ReadFullText());
    }
    #endregion

    #region 目录
    [Fact(DisplayName = "目录-AddTableOfContents 写入 TOC 域并开启更新域")]
    public void AddTableOfContents_WritesTocFieldAndUpdateFields()
    {
        var doc = new Document();
        doc.AddHeading("第一章", 1);
        doc.AddTableOfContents(levels: 3, title: "目录");

        Assert.True(doc.UpdateFields);

        var bytes = Save(doc);
        var xml = ReadEntry(bytes, "word/document.xml");
        Assert.Contains("TOC", xml);
        Assert.Contains("fldCharType=\"begin\"", xml);
        Assert.Contains("w:updateFields", ReadEntry(bytes, "word/settings.xml"));
    }

    [Fact(DisplayName = "目录-不输出标题时只生成域段落")]
    public void AddTableOfContents_NoTitle()
    {
        var doc = new Document();
        doc.AddTableOfContents();

        // 1 个目录域段落（标题为 null 时不额外生成段落）
        Assert.Single(doc.Elements);
    }

    [Fact(DisplayName = "目录-levels 超界自动收敛到 1~9")]
    public void AddTableOfContents_LevelsClamped()
    {
        var doc = new Document();
        doc.AddTableOfContents(levels: 0);
        Assert.Contains("1-1", ReadEntry(Save(doc), "word/document.xml"));

        var doc2 = new Document();
        doc2.AddTableOfContents(levels: 99);
        Assert.Contains("1-9", ReadEntry(Save(doc2), "word/document.xml"));
    }
    #endregion

    #region updateFields 透传
    [Fact(DisplayName = "域-透传文档开启 UpdateFields 后改写 settings.xml")]
    public void UpdateFields_PassthroughSettings_Patched()
    {
        var doc = Read("<w:p><w:r><w:t>正文</w:t></w:r></w:p>", "<w:defaultTabStop w:val=\"720\"/>");
        doc.BeginEdit();
        doc.UpdateFields = true;
        doc.AddField("PAGE", "1");

        var settings = ReadEntry(Save(doc), "word/settings.xml");
        Assert.Contains("w:updateFields", settings);
        Assert.Contains("defaultTabStop", settings);   // 原有内容保留
    }

    [Fact(DisplayName = "域-未开启 UpdateFields 时不写 updateFields")]
    public void UpdateFields_NotSet_Absent()
    {
        var doc = Read("<w:p><w:r><w:t>正文</w:t></w:r></w:p>", "<w:defaultTabStop w:val=\"720\"/>");
        doc.BeginEdit();
        doc.AddField("PAGE", "1");

        Assert.DoesNotContain("w:updateFields", ReadEntry(Save(doc), "word/settings.xml"));
    }
    #endregion
}
