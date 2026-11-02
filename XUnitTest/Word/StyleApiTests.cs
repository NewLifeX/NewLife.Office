using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>样式对象 API 测试：读取/新增/修改/删除 styles.xml 里的样式定义</summary>
public class StyleApiTests
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

    #region 读取
    [Fact(DisplayName = "样式-新文档按默认样式表读到 Heading1/Normal")]
    public void GetStyle_DefaultStyles()
    {
        var doc = new Document();

        var heading1 = doc.GetStyle("Heading1");
        Assert.NotNull(heading1);
        Assert.Equal("paragraph", heading1!.Type);
        Assert.True(heading1.Bold);
        Assert.Equal(20, heading1.FontSize);

        var normal = doc.GetStyle("Normal");
        Assert.NotNull(normal);
        Assert.Null(doc.GetStyle("NotExists"));
    }

    [Fact(DisplayName = "样式-GetStyles 返回全部样式")]
    public void GetStyles_ReturnsAll()
    {
        var doc = new Document();
        var styles = doc.GetStyles();
        Assert.Contains(styles, s => s.StyleId == "Normal");
        Assert.Contains(styles, s => s.StyleId == "Heading1");
        Assert.Contains(styles, s => s.StyleId == "TableGrid");
    }
    #endregion

    #region 修改
    [Fact(DisplayName = "样式-设置字体字号颜色并持久化")]
    public void SetStyleFont_Persists()
    {
        var doc = new Document();
        var info = doc.SetStyleFont("Normal", "微软雅黑", 10.5, bold: true, color: "#333333");

        Assert.NotNull(info);
        Assert.Equal("微软雅黑", info!.FontName);
        Assert.Equal(10.5, info.FontSize);
        Assert.True(info.Bold);
        Assert.Equal("333333", info.Color);
        Assert.True(doc.StylesModified);

        // 回读校验（不依赖 XML 序列化细节）
        var back = ReadBack(Save(doc));
        var saved = back.GetStyle("Normal")!;
        Assert.Equal("微软雅黑", saved.FontName);
        Assert.Equal(10.5, saved.FontSize);
        Assert.True(saved.Bold);
    }

    [Fact(DisplayName = "样式-设置段落格式并持久化")]
    public void SetStyleParagraph_Persists()
    {
        var doc = new Document();
        var info = doc.SetStyleParagraph("Normal", alignment: "both", spaceBefore: 120, spaceAfter: 240, lineSpacingPct: 150, indentLeft: 360, firstLineIndent: 420);

        Assert.NotNull(info);
        Assert.Equal("both", info!.Alignment);
        Assert.Equal(120, info.SpaceBefore);
        Assert.Equal(240, info.SpaceAfter);
        Assert.Equal(150, info.LineSpacingPct);
        Assert.Equal(360, info.IndentLeft);
        Assert.Equal(420, info.FirstLineIndent);

        var stylesXml = ReadEntry(Save(doc), "word/styles.xml");
        Assert.Contains("w:line=\"360\"", stylesXml);       // 150% → 360
        Assert.Contains("w:jc w:val=\"both\"", stylesXml);
    }

    [Fact(DisplayName = "样式-取消加粗时显式写 w:b val=0")]
    public void SetStyleFont_BoldFalse_ExplicitOff()
    {
        var doc = new Document();
        Assert.True(doc.GetStyle("Heading1")!.Bold);

        var info = doc.SetStyleFont("Heading1", bold: false);
        Assert.False(info!.Bold);

        var back = ReadBack(Save(doc));
        Assert.False(back.GetStyle("Heading1")!.Bold);
    }

    [Fact(DisplayName = "样式-不存在的样式返回 null")]
    public void SetStyleFont_Missing_ReturnsNull()
    {
        var doc = new Document();
        Assert.Null(doc.SetStyleFont("NoSuchStyle", "Arial"));
        Assert.Null(doc.SetStyleParagraph("NoSuchStyle", alignment: "center"));
    }
    #endregion

    #region 增删
    [Fact(DisplayName = "样式-新增段落样式并可被段落引用")]
    public void AddStyle_New_UsableByParagraph()
    {
        var doc = new Document();
        var info = doc.AddStyle("QuoteStyle", "引用", "paragraph", "Normal");
        Assert.Equal("QuoteStyle", info.StyleId);
        Assert.Equal("引用", info.Name);

        doc.SetStyleFont("QuoteStyle", italic: true, color: "666666");
        doc.SetStyleParagraph("QuoteStyle", indentLeft: 720);

        var el = doc.AddParagraph("这是一段引用。");
        el.Paragraph!.StyleId = "QuoteStyle";

        var bytes = Save(doc);
        var stylesXml = ReadEntry(bytes, "word/styles.xml");
        Assert.Contains("w:styleId=\"QuoteStyle\"", stylesXml);
        Assert.Contains("666666", stylesXml);
        Assert.Contains("w:pStyle w:val=\"QuoteStyle\"", ReadEntry(bytes, "word/document.xml"));
    }

    [Fact(DisplayName = "样式-重复新增返回现有样式")]
    public void AddStyle_Existing_ReturnsExisting()
    {
        var doc = new Document();
        var first = doc.AddStyle("MyStyle", "我的样式");
        var second = doc.AddStyle("MyStyle", "另一个名字");

        Assert.Equal(first.StyleId, second.StyleId);
        Assert.Equal("我的样式", second.Name);
    }

    [Fact(DisplayName = "样式-删除样式")]
    public void RemoveStyle_Works()
    {
        var doc = new Document();
        doc.AddStyle("TempStyle", "临时");
        Assert.NotNull(doc.GetStyle("TempStyle"));

        Assert.True(doc.RemoveStyle("TempStyle"));
        Assert.Null(doc.GetStyle("TempStyle"));
        Assert.False(doc.RemoveStyle("TempStyle"));
    }

    [Fact(DisplayName = "样式-参数校验")]
    public void AddStyle_InvalidArgs_Throws()
    {
        var doc = new Document();
        Assert.Throws<ArgumentException>(() => doc.AddStyle(""));
        Assert.Throws<ArgumentNullException>(() => ((Document)null!).AddStyle("X"));
    }
    #endregion

    #region 透传覆盖
    [Fact(DisplayName = "样式-透传文档改样式后写入 styles.xml 部件并补引用")]
    public void Styles_ModifiedOnReadDoc_WritesPartAndRefs()
    {
        var doc = Read("<w:p><w:r><w:t>正文</w:t></w:r></w:p>");
        doc.BeginEdit();
        doc.SetStyleFont("Normal", "微软雅黑", 12);

        var bytes = Save(doc);
        var stylesXml = ReadEntry(bytes, "word/styles.xml");
        Assert.Contains("微软雅黑", stylesXml);
        Assert.Contains("styles.xml", ReadEntry(bytes, "[Content_Types].xml"));
        Assert.Contains("styles.xml", ReadEntry(bytes, "word/_rels/document.xml.rels"));
    }
    #endregion
}
