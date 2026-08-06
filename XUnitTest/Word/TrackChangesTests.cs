using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>Word 修订追踪（Track Changes）测试（W11，对标 Aspose.Words）</summary>
/// <remarks>
/// 覆盖 w:ins/w:del 的模型级读取：插入/删除类型、作者、日期、修订文本与段落上下文。
/// </remarks>
public class TrackChangesTests
{
    private static Byte[] BuildDocxWithTrackChanges()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            var ct = zip.CreateEntry("[Content_Types].xml");
            using (var s = ct.Open())
            {
                var content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                    "</Types>";
                var bytes = Encoding.UTF8.GetBytes(content);
                s.Write(bytes, 0, bytes.Length);
            }

            var rels = zip.CreateEntry("_rels/.rels");
            using (var s = rels.Open())
            {
                var content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                    "</Relationships>";
                var bytes = Encoding.UTF8.GetBytes(content);
                s.Write(bytes, 0, bytes.Length);
            }

            var doc = zip.CreateEntry("word/document.xml");
            using (var s = doc.Open())
            {
                var content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body>" +
                    "<w:p>" +
                    "<w:r><w:t>原有文本</w:t></w:r>" +
                    "<w:ins w:id=\"1\" w:author=\"张三\" w:date=\"2024-01-01T10:00:00Z\"><w:r><w:t>插入内容</w:t></w:r></w:ins>" +
                    "<w:del w:id=\"2\" w:author=\"李四\" w:date=\"2024-02-02T10:00:00Z\"><w:r><w:delText>删除内容</w:delText></w:r></w:del>" +
                    "</w:p>" +
                    "<w:ins w:id=\"3\" w:author=\"王五\"><w:p><w:r><w:t>整个段落插入</w:t></w:r></w:p></w:ins>" +
                    "</w:body>" +
                    "</w:document>";
                var bytes = Encoding.UTF8.GetBytes(content);
                s.Write(bytes, 0, bytes.Length);
            }
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("修订追踪：读取插入与删除记录")]
    public void TrackChanges_ReadInsertDelete()
    {
        using var reader = new WordReader(new MemoryStream(BuildDocxWithTrackChanges()));
        var doc = reader.ReadDocument();

        Assert.Equal(3, doc.TrackChanges.Count);

        // 行内插入
        var ins = doc.TrackChanges[0];
        Assert.Equal(TrackChangeType.Insert, ins.Type);
        Assert.Equal("张三", ins.Author);
        Assert.Equal("2024-01-01T10:00:00Z", ins.Date);
        Assert.Equal("插入内容", ins.Text);
        Assert.Equal("原有文本插入内容删除内容", ins.ParagraphText);

        // 行内删除
        var del = doc.TrackChanges[1];
        Assert.Equal(TrackChangeType.Delete, del.Type);
        Assert.Equal("李四", del.Author);
        Assert.Equal("删除内容", del.Text);

        // 段落级插入
        var paraIns = doc.TrackChanges[2];
        Assert.Equal(TrackChangeType.Insert, paraIns.Type);
        Assert.Equal("王五", paraIns.Author);
        Assert.Equal("整个段落插入", paraIns.Text);
    }

    [Fact, DisplayName("修订追踪：段落级修订段落被解析为元素")]
    public void TrackChanges_ParagraphLevelUnwrapped()
    {
        using var reader = new WordReader(new MemoryStream(BuildDocxWithTrackChanges()));
        var doc = reader.ReadDocument();

        // 段落级修订包裹的段落也应出现在 Elements 中
        Assert.Contains(doc.Elements, e => e.Type == ElementType.Paragraph &&
            e.Paragraph != null && e.Paragraph.Runs.Any(r => r.Text.Contains("整个段落插入")));
    }

    [Fact, DisplayName("修订追踪：无修订文档返回空列表")]
    public void TrackChanges_Empty()
    {
        using var writer = new WordWriter();
        using var ms = new MemoryStream();
        var d = new Document();
        d.Elements.Add(new Element
        {
            Type = ElementType.Paragraph,
            Paragraph = new Paragraph { Runs = { new Run { Text = "无修订" } } },
        });
        writer.Save(ms, d);
        ms.Position = 0;

        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        Assert.Empty(doc.TrackChanges);
    }
}
