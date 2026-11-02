using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>修订写入与页码域测试：w:ins/w:del 输出、回读修订记录、页脚页码域</summary>
public class RevisionApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 工具
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

    #region 修订
    [Fact(DisplayName = "修订-段落改写生成 w:del + w:ins 并回读修订记录")]
    public void SetTextAsRevision_WritesInsAndDel()
    {
        var doc = Read("<w:p><w:r><w:t>原文</w:t></w:r></w:p>");
        doc.BeginEdit();
        doc.StartRevisions("AI 审阅", new DateTime(2026, 11, 2, 9, 0, 0, DateTimeKind.Utc));

        var para = doc.AllParagraphs()[0];
        doc.SetTextAsRevision(para, "新文");

        var bytes = Save(doc);
        var xml = ReadEntry(bytes, "word/document.xml");
        Assert.Contains("<w:del ", xml);
        Assert.Contains("<w:delText>原文</w:delText>", xml);
        Assert.Contains("<w:ins ", xml);
        Assert.Contains("<w:t>新文</w:t>", xml);
        Assert.Contains("w:author=\"AI 审阅\"", xml);

        var back = ReadBack(bytes);
        Assert.Contains(back.TrackChanges, c => c.Type == TrackChangeType.Delete && c.Text == "原文");
        Assert.Contains(back.TrackChanges, c => c.Type == TrackChangeType.Insert && c.Text == "新文");
    }

    [Fact(DisplayName = "修订-Run 直接标记插入生成 w:ins")]
    public void RunRevision_Insert()
    {
        var doc = new Document();
        doc.StartRevisions("作者A");
        var el = doc.AddParagraph("保留");
        el.Paragraph!.Runs.Add(new Run { Text = "新增内容", Revision = TrackChangeType.Insert });

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.Contains("<w:ins ", xml);
        Assert.Contains("w:author=\"作者A\"", xml);
        Assert.Contains("<w:t>新增内容</w:t>", xml);
    }

    [Fact(DisplayName = "修订-删除 Run 使用 delText 而非 t")]
    public void RunRevision_Delete_UsesDelText()
    {
        var doc = new Document();
        var el = doc.AddParagraph("");
        el.Paragraph!.Runs.Add(new Run { Text = "要删掉", Revision = TrackChangeType.Delete });

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.Contains("<w:delText>要删掉</w:delText>", xml);
        Assert.DoesNotContain("<w:t>要删掉</w:t>", xml);
    }

    [Fact(DisplayName = "修订-未标记修订的 Run 不产生 w:ins/w:del")]
    public void NoRevision_NoMarkup()
    {
        var doc = new Document();
        doc.AddParagraph("普通文本");

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.DoesNotContain("<w:ins ", xml);
        Assert.DoesNotContain("<w:del ", xml);
    }

    [Fact(DisplayName = "修订-StartRevisions 参数校验")]
    public void StartRevisions_InvalidArgs_Throws()
    {
        var doc = new Document();
        Assert.Throws<ArgumentException>(() => doc.StartRevisions(""));
        Assert.Throws<ArgumentNullException>(() => ((Document)null!).StartRevisions("A"));
        Assert.Throws<ArgumentNullException>(() => doc.SetTextAsRevision(null!, "x"));
    }
    #endregion

    #region 页码
    [Fact(DisplayName = "页码-页脚写入 PAGE/NUMPAGES 域并开启更新域")]
    public void AddPageNumber_WritesFields()
    {
        var doc = new Document();
        doc.AddParagraph("正文");
        doc.AddPageNumber();

        Assert.True(doc.UpdateFields);
        Assert.Single(doc.Footers);

        var bytes = Save(doc);
        var footer = ReadEntry(bytes, "word/footer1.xml");
        Assert.Contains("PAGE", footer);
        Assert.Contains("NUMPAGES", footer);
        Assert.Contains("fldCharType=\"begin\"", footer);
        Assert.Contains("第 ", footer);
    }

    [Fact(DisplayName = "页码-同一类型页脚重复追加不新建部件")]
    public void AddPageNumber_Twice_SingleFooter()
    {
        var doc = new Document();
        doc.AddPageNumber();
        doc.AddPageNumber("第 ", " / ", " 页");

        Assert.Single(doc.Footers);
        Assert.Equal(2, doc.Footers[0].Elements.Count);
    }

    [Fact(DisplayName = "页码-首页页脚类型独立")]
    public void AddPageNumber_FirstPageType()
    {
        var doc = new Document();
        doc.AddPageNumber(type: "first");
        Assert.Equal("first", doc.Footers[0].Type);
    }
    #endregion
}
