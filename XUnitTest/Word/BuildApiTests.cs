using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>文档构建 API 测试：AddHeading/AddTable/AddPageBreak/AddSection、正则查找替换、页面设置副本</summary>
public class BuildApiTests
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

    private static String BodyText(Document doc)
    {
        var sb = new StringBuilder();
        foreach (var el in doc.Elements)
        {
            if (el.Paragraph == null) continue;
            foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
            sb.Append('|');
        }
        return sb.ToString();
    }
    #endregion

    #region 构建
    [Fact(DisplayName = "构建-AddHeading 生成标题样式并持久化")]
    public void AddHeading_Persists()
    {
        var doc = new Document();
        doc.AddHeading("一级标题", 1);
        doc.AddHeading("三级标题", 3);

        var back = ReadBack(Save(doc));
        var paras = back.AllParagraphs();
        Assert.Equal(ParagraphStyle.Heading1, paras[0].Style);
        Assert.Equal(ParagraphStyle.Heading3, paras[1].Style);
    }

    [Fact(DisplayName = "构建-AddTable(rows,cols) 生成空网格并持久化")]
    public void AddTable_BySize_Persists()
    {
        var doc = new Document();
        doc.AddTable(2, 3);

        var back = ReadBack(Save(doc));
        var table = Assert.Single(back.AllTables());
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count);
        Assert.Equal(3, table.Rows[1].Cells.Count);
    }

    [Fact(DisplayName = "构建-AddTable(数据数组) 首行作表头")]
    public void AddTable_ByData_Persists()
    {
        var doc = new Document();
        doc.AddTable([["姓名", "部门"], ["张三", "研发"]]);

        var back = ReadBack(Save(doc));
        var table = Assert.Single(back.AllTables());
        Assert.Equal("姓名", table.GetCellText(0, 0));
        Assert.Equal("研发", table.GetCellText(1, 1));
        Assert.True(table.Rows[0].IsHeader);
    }

    [Fact(DisplayName = "构建-AddTable 非法尺寸抛异常")]
    public void AddTable_InvalidSize_Throws()
    {
        var doc = new Document();
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddTable(0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddTable(2, 0));
        Assert.Throws<ArgumentNullException>(() => doc.AddTable((IEnumerable<String[]>)null!));
    }

    [Fact(DisplayName = "构建-AddPageBreak 生成分页符并持久化")]
    public void AddPageBreak_Persists()
    {
        var doc = new Document();
        doc.AddParagraph("第一页");
        doc.AddPageBreak();
        doc.AddParagraph("第二页");

        var back = ReadBack(Save(doc));
        Assert.Contains(back.AllParagraphs(), p => p.IsPageBreak);
    }

    [Fact(DisplayName = "构建-AddSection 生成新节且其后内容归属新节")]
    public void AddSection_SplitsSections()
    {
        var doc = new Document();
        doc.AddParagraph("第一节内容");

        var sec = doc.AddSection("continuous");
        Assert.Equal(2, doc.Sections.Count);
        Assert.Equal("continuous", sec.Type);
        Assert.Equal(1, doc.Sections[0].Elements.Count);

        doc.AddParagraph("第二节内容");
        Assert.Equal(1, doc.Sections[1].Elements.Count);
        Assert.NotSame(doc.Sections[0].PageSettings, sec.PageSettings);

        // 回读后仍为两节
        var back = ReadBack(Save(doc));
        Assert.Equal(2, back.Sections.Count);
    }
    #endregion

    #region 正则
    [Fact(DisplayName = "正则-跨 Run 查找")]
    public void FindRegex_CrossRun()
    {
        var doc = Read("<w:p><w:r><w:t>订单编号：A-</w:t></w:r><w:r><w:t>001 号</w:t></w:r></w:p>");

        var matches = doc.FindRegex(@"A-(\d+)");
        var m = Assert.Single(matches);
        Assert.Equal("A-001", m.Text);
        Assert.Equal(5, m.StartOffset);
    }

    [Fact(DisplayName = "正则-跨 Run 替换支持分组引用并持久化")]
    public void ReplaceRegex_CrossRun_Persists()
    {
        var doc = Read("<w:p><w:r><w:t>订单编号：A-</w:t></w:r><w:r><w:t>001 号</w:t></w:r></w:p>");
        doc.BeginEdit();

        var count = doc.ReplaceRegex(@"A-(\d+)", "B-$1");
        Assert.Equal(1, count);

        Assert.Equal("订单编号：B-001 号|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "正则-忽略大小写选项生效")]
    public void ReplaceRegex_IgnoreCase()
    {
        var doc = Read("<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>");
        doc.BeginEdit();

        var count = doc.ReplaceRegex("world", "NewLife", RegexOptions.IgnoreCase);
        Assert.Equal(1, count);
        Assert.Equal("Hello NewLife|", BodyText(ReadBack(Save(doc))));
    }

    [Fact(DisplayName = "正则-空模式不替换")]
    public void ReplaceRegex_EmptyPattern_ReturnsZero()
    {
        var doc = Read("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>");
        Assert.Equal(0, doc.ReplaceRegex("", "X"));
        Assert.Empty(doc.FindRegex(""));
    }
    #endregion

    #region 页面设置
    [Fact(DisplayName = "页面设置-Clone 为独立副本")]
    public void PageSettings_Clone_IsIndependent()
    {
        var ps = new PageSettings { PageWidth = 1000, Landscape = true };
        ps.HeaderRefs["default"] = "rHdr1";

        var copy = ps.Clone();
        copy.PageWidth = 2000;
        copy.HeaderRefs["default"] = "rHdr2";

        Assert.Equal(1000, ps.PageWidth);
        Assert.Equal("rHdr1", ps.HeaderRefs["default"]);
        Assert.True(copy.Landscape);
    }
    #endregion
}
