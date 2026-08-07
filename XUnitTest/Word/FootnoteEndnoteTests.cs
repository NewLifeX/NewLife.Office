using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>脚注/尾注模型读取测试（W22）</summary>
public class FootnoteEndnoteTests
{
    #region 辅助

    private static void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    /// <summary>构建含脚注/尾注的完整 docx</summary>
    private static Byte[] BuildDocxWithNotes()
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>"
            + "<w:p><w:r><w:t>正文带脚注</w:t></w:r>"
            + "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r>"
            + "</w:p>"
            + "<w:p><w:r><w:t>正文带尾注</w:t></w:r>"
            + "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:endnoteReference w:id=\"2\"/></w:r>"
            + "</w:p>"
            + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>"
            + "</w:body></w:document>";

        var footnotesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:footnotes xmlns:w=\"{W}\">"
            + "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
            + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
            + "<w:footnote w:id=\"1\"><w:p><w:r><w:t>这是脚注内容</w:t></w:r></w:p></w:footnote>"
            + "</w:footnotes>";

        var endnotesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:endnotes xmlns:w=\"{W}\">"
            + "<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>"
            + "<w:endnote w:id=\"2\"><w:p><w:r><w:t>这是尾注内容</w:t></w:r></w:p></w:endnote>"
            + "</w:endnotes>";

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>"
                + "<Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/footnotes.xml", footnotesXml);
            WriteEntry(za, "word/endnotes.xml", endnotesXml);
            WriteEntry(za, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>"
                + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/>"
                + "</Relationships>");
        }
        return ms.ToArray();
    }

    #endregion

    #region 测试

    [Fact, DisplayName("W22_脚注_读取模型且跳过内置分隔符")]
    public void Footnotes_Parsed()
    {
        var bytes = BuildDocxWithNotes();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        // 仅普通脚注（id=1），内置分隔符（id=-1/0）被跳过
        var fn = Assert.Single(doc.Footnotes);
        Assert.Equal(1, fn.Id);
        Assert.Equal("这是脚注内容", fn.Text);
        Assert.Single(fn.Paragraphs);
        Assert.Equal("这是脚注内容", String.Concat(fn.Paragraphs[0].Runs.Select(r => r.Text)));
    }

    [Fact, DisplayName("W22_尾注_读取模型")]
    public void Endnotes_Parsed()
    {
        var bytes = BuildDocxWithNotes();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        var en = Assert.Single(doc.Endnotes);
        Assert.Equal(2, en.Id);
        Assert.Equal("这是尾注内容", en.Text);
    }

    [Fact, DisplayName("W22_脚注_无脚注文件时列表为空")]
    public void Footnotes_Empty_WhenNoFile()
    {
        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
        {
            writer.AppendParagraph("无脚注文档");
            writer.Save(ms);
        }
        ms.Position = 0;
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        Assert.Empty(doc.Footnotes);
        Assert.Empty(doc.Endnotes);
    }

    [Fact, DisplayName("W22_脚注_读取往返不丢失脚注部件")]
    public void Footnotes_RoundTrip_Preserved()
    {
        var bytes = BuildDocxWithNotes();
        var outPath = Path.Combine(Path.GetTempPath(), $"notes_{Guid.NewGuid():N}.docx");
        try
        {
            Document doc;
            using (var ms = new MemoryStream(bytes))
            using (var reader = new WordReader(ms))
                doc = reader.ReadDocument();

            Assert.Single(doc.Footnotes);

            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            using (var reader = new WordReader(outPath))
            {
                var re = reader.ReadDocument();
                // 透传保留 footnotes.xml，脚注文本不丢
                Assert.Single(re.Footnotes);
                Assert.Equal("这是脚注内容", re.Footnotes[0].Text);
            }
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W22_脚注_FindText命中脚注文本")]
    public void Footnotes_FindText_Hits()
    {
        var bytes = BuildDocxWithNotes();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();

        var matches = doc.FindText("这是脚注内容");
        Assert.Single(matches);

        var endnoteMatches = doc.FindText("这是尾注内容");
        Assert.Single(endnoteMatches);
    }

    [Fact, DisplayName("W22_脚注_ReplaceText替换并持久化")]
    public void Footnotes_ReplaceText_Persists()
    {
        var bytes = BuildDocxWithNotes();
        var outPath = Path.Combine(Path.GetTempPath(), $"notes_{Guid.NewGuid():N}.docx");
        try
        {
            Document doc;
            using (var ms = new MemoryStream(bytes))
            using (var reader = new WordReader(ms))
                doc = reader.ReadDocument();

            var count = doc.ReplaceText("这是脚注内容", "更新后的脚注");
            Assert.Equal(1, count);

            // footnotes.xml 部件已同步（无需清 DocumentXml，脚注是独立部件）
            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            using (var reader = new WordReader(outPath))
            {
                var re = reader.ReadDocument();
                Assert.Equal("更新后的脚注", re.Footnotes[0].Text);
                // 尾注未改动，保持原文
                Assert.Equal("这是尾注内容", re.Endnotes[0].Text);
            }
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W22_脚注_Markdown转换包含脚注")]
    public void Footnotes_Markdown_Included()
    {
        var bytes = BuildDocxWithNotes();
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        var md = reader.ExtractMarkdown();
        Assert.NotNull(md);
        Assert.Contains("脚注 1: 这是脚注内容", md);
        Assert.Contains("尾注 2: 这是尾注内容", md);
    }

    [Fact, DisplayName("W22_脚注_模板填充覆盖脚注内占位符")]
    public void Template_Fill_CoversFootnotes()
    {
        var templatePath = Path.Combine(Path.GetTempPath(), $"tpl_{Guid.NewGuid():N}.docx");
        var outPath = Path.Combine(Path.GetTempPath(), $"out_{Guid.NewGuid():N}.docx");
        try
        {
            // 生成含脚注占位符的模板（脚注正文在独立 footnotes.xml 部件）
            using (var w = new WordWriter())
            {
                w.AppendFootnote("条款", "当事人 {{Name}} 应于 {{Date}} 前履行");
                w.Save(templatePath);
            }

            var tpl = new WordTemplate(templatePath);
            tpl.Fill(outPath, new Dictionary<String, Object?>
            {
                ["Name"] = "张三",
                ["Date"] = "2026-08-07",
            });

            // 脚注部件中的占位符应被替换（WordTemplate 对所有 .xml 条目做替换）
            using var za = System.IO.Compression.ZipFile.OpenRead(outPath);
            var entry = za.GetEntry("word/footnotes.xml");
            Assert.NotNull(entry);
            using var sr = new System.IO.StreamReader(entry!.Open(), System.Text.Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("张三", xml);
            Assert.DoesNotContain("{{Name}}", xml);
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    #endregion

    #region 写入（WordWriter 程序化创建脚注/尾注）

    [Fact, DisplayName("W22_脚注_WordWriter程序化创建并读回")]
    public void Writer_AppendFootnote_RoundTrip()
    {
        var outPath = Path.Combine(Path.GetTempPath(), $"notes_{Guid.NewGuid():N}.docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.AppendFootnote("研究背景", "这是第一个脚注");
                w.AppendParagraph("正文内容");
                w.AppendEndnote("附录说明", "这是尾注");
                w.Save(outPath);
            }

            using (var reader = new WordReader(outPath))
            {
                var doc = reader.ReadDocument();
                Assert.Single(doc.Footnotes);
                Assert.Equal("这是第一个脚注", doc.Footnotes[0].Text);
                Assert.Equal(1, doc.Footnotes[0].Id);
                Assert.Single(doc.Endnotes);
                Assert.Equal("这是尾注", doc.Endnotes[0].Text);
            }
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W22_脚注_WordWriter创建含脚注文件通过规范校验")]
    public void Writer_AppendFootnote_ValidDocx()
    {
        var outPath = Path.Combine(Path.GetTempPath(), $"notes_{Guid.NewGuid():N}.docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.AppendFootnote("段落带脚注", "脚注内容");
                w.AppendParagraph("第二段");
                w.Save(outPath);
            }

            // OpenXmlValidator 校验（引用 DocumentFormat.OpenXml）
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outPath, false);
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    [Fact, DisplayName("W22_脚注_模型写回保留脚注（读入→写出，无重复 ZIP 条目）")]
    public void Writer_ModelWriteBack_PreservesNotes()
    {
        var bytes = BuildDocxWithNotes();
        var outPath = Path.Combine(Path.GetTempPath(), $"notes_{Guid.NewGuid():N}.docx");
        try
        {
            Document doc;
            using (var ms = new MemoryStream(bytes))
            using (var reader = new WordReader(ms))
                doc = reader.ReadDocument();

            // 模型输出：DocumentXml=null 触发 Writer 生成 footnotes.xml
            doc.DocumentXml = null;
            using (var writer = new WordWriter())
                writer.Save(outPath, doc);

            // 1. ZIP 部件唯一性：脚注/尾注/Content_Types/rels 不得重复条目（回归 #1）
            using (var za = ZipFile.OpenRead(outPath))
            {
                var names = za.Entries.Select(e => e.FullName).ToList();
                Assert.Single(names.Where(n => n == "word/footnotes.xml"));
                Assert.Single(names.Where(n => n == "word/endnotes.xml"));
                Assert.Single(names.Where(n => n == "[Content_Types].xml"));
                Assert.Single(names.Where(n => n == "word/_rels/document.xml.rels"));
            }

            // 2. 重读：脚注/尾注模型完整
            using (var reader = new WordReader(outPath))
            {
                var re = reader.ReadDocument();
                Assert.Single(re.Footnotes);
                Assert.Equal("这是脚注内容", re.Footnotes[0].Text);
                Assert.Single(re.Endnotes);
                Assert.Equal("这是尾注内容", re.Endnotes[0].Text);
            }

            // 3. OpenXmlValidator 校验
            using var doc2 = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(outPath, false);
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            var errors = validator.Validate(doc2).ToList();
            Assert.Empty(errors);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    #endregion
}
