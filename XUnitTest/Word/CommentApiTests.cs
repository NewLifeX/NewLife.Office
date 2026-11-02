using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>批注 API 测试：AddComment 写入 comments.xml 与锚点、读取回模型、不误伤已有批注</summary>
public class CommentApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #region 工具
    private static Byte[] BuildDocx(String bodyInnerXml, String? commentsInnerXml = null)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var types = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>";
            if (commentsInnerXml != null)
                types += "<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/>";
            types += "</Types>";
            WriteEntry(za, "[Content_Types].xml", types);
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
            if (commentsInnerXml != null)
                rels += "<Relationship Id=\"rComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"comments.xml\"/>";
            rels += "</Relationships>";
            WriteEntry(za, "word/_rels/document.xml.rels", rels);
            if (commentsInnerXml != null)
                WriteEntry(za, "word/comments.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + $"<w:comments xmlns:w=\"{W}\">{commentsInnerXml}</w:comments>");
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml, String? commentsInnerXml = null)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml, commentsInnerXml));
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

    private static Boolean HasEntry(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        return za.GetEntry(path) != null;
    }
    #endregion

    #region 写入
    [Fact(DisplayName = "批注-新建文档写入 comments.xml 与锚点")]
    public void AddComment_NewDoc_WritesPartsAndAnchors()
    {
        var doc = new Document();
        var el = doc.AddParagraph("这段结论需要数据支撑。");
        var comment = doc.AddComment(el.Paragraph!, "请补充数据来源。", "AI 审阅", "AI");

        Assert.Equal(1, doc.Comments.Count);
        Assert.False(String.IsNullOrEmpty(comment.Author));

        var bytes = Save(doc);
        Assert.True(HasEntry(bytes, "word/comments.xml"));
        Assert.Contains("comments.xml", ReadEntry(bytes, "[Content_Types].xml"));
        Assert.Contains("rComments", ReadEntry(bytes, "word/_rels/document.xml.rels"));

        var docXml = ReadEntry(bytes, "word/document.xml");
        Assert.Contains($"<w:commentRangeStart w:id=\"{comment.Id}\"/>", docXml);
        Assert.Contains($"<w:commentRangeEnd w:id=\"{comment.Id}\"/>", docXml);
        Assert.Contains($"<w:commentReference w:id=\"{comment.Id}\"/>", docXml);

        var commentsXml = ReadEntry(bytes, "word/comments.xml");
        Assert.Contains("AI 审阅", commentsXml);
        Assert.Contains("请补充数据来源。", commentsXml);
    }

    [Fact(DisplayName = "批注-读写往返可回读到批注内容")]
    public void AddComment_RoundTrip()
    {
        var doc = new Document();
        var el = doc.AddParagraph("正文");
        doc.AddComment(el.Paragraph!, "第一条", "作者A");

        var back = ReadBack(Save(doc));
        var c = Assert.Single(back.Comments);
        Assert.Equal("第一条", c.Text);
        Assert.Equal("作者A", c.Author);
        Assert.Equal(1, c.Paragraphs.Count);
    }

    [Fact(DisplayName = "批注-多条批注 ID 唯一")]
    public void AddComment_Multiple_UniqueIds()
    {
        var doc = new Document();
        var a = doc.AddParagraph("段落一");
        var b = doc.AddParagraph("段落二");
        var c1 = doc.AddComment(a.Paragraph!, "批注一");
        var c2 = doc.AddComment(b.Paragraph!, "批注二");

        Assert.NotEqual(c1.Id, c2.Id);
        Assert.Equal(2, doc.Comments.Count);

        var back = ReadBack(Save(doc));
        Assert.Equal(2, back.Comments.Count);
    }

    [Fact(DisplayName = "批注-空段落也能挂批注")]
    public void AddComment_EmptyParagraph()
    {
        var doc = new Document();
        var el = doc.AddParagraph("");
        doc.AddComment(el.Paragraph!, "空段批注");

        var xml = ReadEntry(Save(doc), "word/document.xml");
        Assert.Contains("commentRangeStart", xml);
        Assert.Contains("commentReference", xml);
    }

    [Fact(DisplayName = "批注-参数为空抛异常")]
    public void AddComment_NullArgs_Throws()
    {
        var doc = new Document();
        var el = doc.AddParagraph("x");
        Assert.Throws<ArgumentNullException>(() => doc.AddComment(null!, "t"));
        Assert.Throws<ArgumentNullException>(() => ((Document)null!).AddComment(el.Paragraph!, "t"));
    }
    #endregion

    #region 规范校验
    [Fact(DisplayName = "批注-输出不新增 OOXML 规范错误")]
    public void AddComment_DoesNotAddValidationErrors()
    {
        var baseline = new Document();
        baseline.AddParagraph("正文");

        var withComment = new Document();
        var el = withComment.AddParagraph("正文");
        withComment.AddComment(el.Paragraph!, "批注内容", "AI");

        var plainErrors = CountErrors(Save(baseline));
        var commentErrors = CountErrors(Save(withComment));
        Assert.True(commentErrors <= plainErrors, $"基线={plainErrors} 带批注={commentErrors}");
    }

    /// <summary>用 OpenXmlValidator 统计规范错误数（校验器自身异常也计入，避免测试直接爆炸）</summary>
    private static Int32 CountErrors(Byte[] data)
    {
        var path = Path.Combine(Path.GetTempPath(), $"cmt_{Guid.NewGuid():N}.docx");
        try
        {
            File.WriteAllBytes(path, data);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            var count = 0;
            try
            {
                foreach (var _ in validator.Validate(package)) count++;
            }
            catch (Exception)
            {
                count++;
            }
            return count;
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 已有批注不被误伤
    [Fact(DisplayName = "批注-读取已有批注后再改正文，原 comments.xml 原样保留")]
    public void ExistingComments_SurviveBodyEdit()
    {
        var comments = "<w:comment w:id=\"1\" w:author=\"原作者\" w:date=\"2026-01-01T00:00:00Z\">"
            + "<w:p><w:r><w:t>原有批注</w:t></w:r></w:p></w:comment>";
        var doc = Read("<w:p><w:r><w:t>原文</w:t></w:r></w:p>", comments);

        var existing = Assert.Single(doc.Comments);
        Assert.Equal("原有批注", existing.Text);

        doc.BeginEdit();
        doc.AllParagraphs()[0].SetText("改后正文");

        var bytes = Save(doc);
        // 未新增锚点 → 不透传重建，原批注部件保持不变
        Assert.True(HasEntry(bytes, "word/comments.xml"));
        Assert.Contains("原有批注", ReadEntry(bytes, "word/comments.xml"));
    }
    #endregion
}
