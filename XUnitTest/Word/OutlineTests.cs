using System.Text.Json;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>文档结构快照测试：锚点、块类型、表格矩阵、JSON 输出</summary>
public class OutlineTests
{
    private static Document BuildSample()
    {
        var doc = new Document();
        doc.AddHeading("第一章 概述", 1);
        doc.AddParagraph("这是正文第一段。");
        doc.AddParagraph("这是正文第二段。");
        doc.AddTable([["姓名", "部门"], ["张三", "研发"]]);
        return doc;
    }

    [Fact(DisplayName = "快照-块类型与锚点")]
    public void ToOutline_BlocksAndAnchors()
    {
        var outline = BuildSample().ToOutline();

        Assert.Equal(4, outline.Blocks.Count);
        Assert.Equal(1, outline.HeadingCount);
        Assert.Equal(1, outline.TableCount);

        Assert.Equal("p#0", outline.Blocks[0].Anchor);
        Assert.Equal("heading", outline.Blocks[0].Type);
        Assert.Equal(1, outline.Blocks[0].Level);
        Assert.Equal("第一章 概述", outline.Blocks[0].Text);

        Assert.Equal("paragraph", outline.Blocks[1].Type);
        Assert.Equal("这是正文第一段。", outline.Blocks[1].Text);

        Assert.Equal("tbl#3", outline.Blocks[3].Anchor);
        Assert.Equal("table", outline.Blocks[3].Type);
        Assert.Equal(2, outline.Blocks[3].Rows);
        Assert.Equal(2, outline.Blocks[3].Columns);
        Assert.Equal("张三", outline.Blocks[3].Cells![1][0]);
    }

    [Fact(DisplayName = "快照-JSON 可被解析且含锚点")]
    public void ToJson_Parsable()
    {
        var json = BuildSample().ToOutline().ToJson();

        using var parsed = JsonDocument.Parse(json);
        var root = parsed.RootElement;
        Assert.Equal(1, root.GetProperty("headingCount").GetInt32());
        Assert.Equal(1, root.GetProperty("tableCount").GetInt32());

        var blocks = root.GetProperty("blocks");
        Assert.Equal(4, blocks.GetArrayLength());
        Assert.Equal("p#0", blocks[0].GetProperty("anchor").GetString());
        Assert.Equal(1, blocks[0].GetProperty("level").GetInt32());
        Assert.Equal("tbl#3", blocks[3].GetProperty("anchor").GetString());
        Assert.Equal("研发", blocks[3].GetProperty("cells")[1][1].GetString());
    }

    [Fact(DisplayName = "快照-JSON 紧凑模式无换行")]
    public void ToJson_Compact()
    {
        var json = BuildSample().ToOutline().ToJson(pretty: false);
        Assert.DoesNotContain("\n", json);
        using var parsed = JsonDocument.Parse(json);
        Assert.Equal(4, parsed.RootElement.GetProperty("blocks").GetArrayLength());
    }

    [Fact(DisplayName = "快照-特殊字符转义")]
    public void ToJson_EscapesSpecialChars()
    {
        var doc = new Document();
        var el = doc.AddParagraph("");
        // 直接构造含换行/制表符的 Run（AddParagraph 会按换行拆段）
        el.Paragraph!.Runs.Add(new Run { Text = "引号\"反斜杠\\换行\n制表\t" });

        var json = doc.ToOutline().ToJson();
        using var parsed = JsonDocument.Parse(json);
        var text = parsed.RootElement.GetProperty("blocks")[0].GetProperty("text").GetString();
        Assert.Contains("\"", text);
        Assert.Contains("\\", text);
        Assert.Contains("\n", text);
    }

    [Fact(DisplayName = "快照-未建模内容被标记")]
    public void ToOutline_FlagsUnmodeled()
    {
        // 含 fldSimple 的段落：Reader 会标记 HasUnmodeledContent
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var body = "<w:p><w:r><w:t>共</w:t></w:r>"
            + "<w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>";

        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{body}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
        {
            void Write(String name, String content)
            {
                using var sw = new StreamWriter(za.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
                sw.Write(content);
            }
            Write("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Write("_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Write("word/document.xml", documentXml);
            Write("word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        }

        ms.Position = 0;
        using var reader = new WordReader(ms);
        var outline = reader.ReadOutline();

        var block = Assert.Single(outline.Blocks);
        Assert.True(block.HasUnmodeledContent);
        Assert.Contains("\"hasUnmodeledContent\": true", outline.ToJson());
    }

    [Fact(DisplayName = "快照-空文档返回空块列表")]
    public void ToOutline_Empty()
    {
        var outline = new Document().ToOutline();
        Assert.Empty(outline.Blocks);
        Assert.Contains("\"blocks\": []", outline.ToJson(pretty: false));
    }

    [Fact(DisplayName = "快照-域 Run 取缓存结果文本")]
    public void ToOutline_FieldUsesResult()
    {
        var doc = new Document();
        doc.AddField("PAGE", "7");

        var block = Assert.Single(doc.ToOutline().Blocks);
        Assert.Equal("7", block.Text);
    }
}
