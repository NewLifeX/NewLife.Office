using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>样式默认/表格补全/嵌套表格测试</summary>
public class StyleAndNestedTableTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static Byte[] BuildDocx(String bodyInnerXml, String? stylesXml = null)
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
            if (stylesXml != null) WriteEntry(za, "word/styles.xml", stylesXml);
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String name, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(name).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static Document Read(String bodyInnerXml, String? stylesXml = null)
    {
        using var ms = new MemoryStream(BuildDocx(bodyInnerXml, stylesXml));
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    private const String StylesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + $"<w:styles xmlns:w=\"{W}\">"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/>"
        + "<w:pPr><w:jc w:val=\"center\"/><w:ind w:left=\"240\" w:right=\"120\"/>"
        + "<w:spacing w:before=\"240\" w:after=\"120\" w:line=\"480\" w:lineRule=\"auto\"/></w:pPr>"
        + "<w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
        + "</w:styles>";

    #region 段落样式默认
    [Fact(DisplayName = "样式—段落样式 pPr 默认应用到段落模型")]
    public void Style_ParaDefaultsApplied()
    {
        var body = "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>";
        var doc = Read(body, StylesXml);
        var para = Assert.Single(doc.Elements).Paragraph!;
        Assert.Equal(ParagraphStyle.Heading1, para.Style);
        Assert.Equal("center", para.Alignment);
        Assert.Equal(240, para.IndentLeft);
        Assert.Equal(120, para.IndentRight);
        Assert.Equal(240, para.SpaceBefore);
        Assert.Equal(120, para.SpaceAfter);
        Assert.Equal(200, para.LineSpacingPct); // 480/240*100
    }

    [Fact(DisplayName = "样式—段落内联格式优先于样式默认")]
    public void Style_InlineOverridesStyleDefault()
    {
        var body = "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/><w:jc w:val=\"right\"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>";
        var doc = Read(body, StylesXml);
        var para = Assert.Single(doc.Elements).Paragraph!;
        Assert.Equal("right", para.Alignment); // 内联覆盖样式默认 center
        Assert.Equal(240, para.IndentLeft);    // 缩进仍继承样式
    }
    #endregion

    #region 表格读取补全
    [Fact(DisplayName = "表格—tblStyle/对齐/宽度读取")]
    public void Table_StyleAlignWidth()
    {
        var body = "<w:tbl><w:tblPr>"
            + "<w:tblStyle w:val=\"TableGrid\"/>"
            + "<w:tblW w:w=\"6000\" w:type=\"dxa\"/>"
            + "<w:jc w:val=\"center\"/>"
            + "</w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var doc = Read(body);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Table);
        var table = el.Table!;
        Assert.Equal("TableGrid", table.StyleId);
        Assert.Equal("center", table.Alignment);
        Assert.Equal(6000, table.Width);
    }

    [Fact(DisplayName = "表格—StyleId 写回 tblStyle")]
    public void Table_StyleIdWriteBack()
    {
        var table = new Table
        {
            StyleId = "TableGrid",
            Width = 6000,
            Alignment = "center",
        };
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "A" } } } } });
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "B" } } } } });

        var doc = new Document { DocumentXml = null };
        doc.Elements.Add(new Element { Type = ElementType.Table, Table = table });

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);
            using var za = ZipFile.OpenRead(path);
            using var sr = new StreamReader(za.GetEntry("word/document.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("<w:tblStyle w:val=\"TableGrid\"/>", xml);
            Assert.Contains("<w:jc w:val=\"center\"/>", xml);
            Assert.Contains("<w:tblW w:w=\"6000\"", xml);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion

    #region 嵌套表格
    [Fact(DisplayName = "嵌套表格—读取到 Cell.NestedTables")]
    public void NestedTable_Read()
    {
        var body = "<w:tbl><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:p><w:r><w:t>嵌套A</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>"
            + "<w:p><w:r><w:t>外层</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Table);
        var outer = el.Table!;
        var cell = outer.Rows[0].Cells[0];
        Assert.Single(cell.NestedTables);
        var nested = cell.NestedTables[0];
        Assert.Equal("嵌套A", nested.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
        // 外层段落仍保留
        Assert.Equal("外层", cell.Paragraphs[0].Runs[0].Text);
    }

    [Fact(DisplayName = "嵌套表格—写回后重新读取结构一致")]
    public void NestedTable_WriteBack()
    {
        var body = "<w:tbl><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:p><w:r><w:t>内层数据</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>"
            + "<w:p><w:r><w:t>外层数据</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc1 = Read(body);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc1);
            using var reader = new WordReader(path);
            var doc2 = reader.ReadDocument();
            var el = Assert.Single(doc2.Elements, e => e.Type == ElementType.Table);
            var cell = el.Table!.Rows[0].Cells[0];
            Assert.Single(cell.NestedTables);
            Assert.Equal("内层数据", cell.NestedTables[0].Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
            Assert.Equal("外层数据", cell.Paragraphs[0].Runs[0].Text);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "嵌套表格—查找替换覆盖嵌套表格内文本")]
    public void NestedTable_ReplaceText()
    {
        var body = "<w:tbl><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:p><w:r><w:t>旧值内层</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>"
            + "<w:p><w:r><w:t>旧值外层</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var count = doc.ReplaceText("旧值", "新值");
        Assert.Equal(2, count);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Table);
        var cell = el.Table!.Rows[0].Cells[0];
        Assert.Equal("新值内层", cell.NestedTables[0].Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
        Assert.Equal("新值外层", cell.Paragraphs[0].Runs[0].Text);
    }
    #endregion
}
