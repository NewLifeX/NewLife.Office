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

    #endregion
}
