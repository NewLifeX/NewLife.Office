using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>日常编辑 API 测试：查找替换（跨 Run/忽略大小写/整词）、表格行列操作</summary>
public class EditApiTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

    private static String GetBodyText(Document doc)
    {
        var sb = new StringBuilder();
        foreach (var el in doc.Elements)
        {
            if (el.Paragraph != null)
            {
                foreach (var r in el.Paragraph.Runs) sb.Append(r.Text);
                sb.Append('|');
            }
            else if (el.TableRows != null)
            {
                foreach (var row in el.TableRows)
                    foreach (var cell in row)
                        foreach (var p in cell.Paragraphs)
                            foreach (var r in p.Runs) sb.Append(r.Text);
                sb.Append('|');
            }
        }
        return sb.ToString();
    }

    #region 查找
    [Fact(DisplayName = "查找—单 Run 内匹配")]
    public void Find_SingleRun()
    {
        var doc = Read("<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>");
        var matches = doc.FindText("World");
        var m = Assert.Single(matches);
        Assert.Equal("World", m.Text);
        Assert.Equal(6, m.StartOffset);
    }

    [Fact(DisplayName = "查找—跨 Run 匹配")]
    public void Find_CrossRun()
    {
        // "旧公司名" 被拆到两个 Run
        var doc = Read("<w:p><w:r><w:t>旧公</w:t></w:r><w:r><w:t>司名</w:t></w:r></w:p>");
        var matches = doc.FindText("旧公司名");
        var m = Assert.Single(matches);
        Assert.Equal("旧公司名", m.Text);
        Assert.Equal(0, m.StartOffset);
    }
    #endregion

    #region 替换
    [Fact(DisplayName = "替换—单 Run 内保留格式")]
    public void Replace_SingleRunKeepFormat()
    {
        var doc = Read("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello World</w:t></w:r></w:p>");
        var count = doc.ReplaceText("World", "NewLife");
        Assert.Equal(1, count);
        var run = Assert.Single(doc.Elements[0].Paragraph!.Runs);
        Assert.Equal("Hello NewLife", run.Text);
        Assert.True(run.Properties!.Bold); // 格式保留
    }

    [Fact(DisplayName = "替换—跨 Run 匹配保留首 Run 格式")]
    public void Replace_CrossRunKeepFirstFormat()
    {
        var doc = Read("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>旧公</w:t></w:r><w:r><w:t>司名</w:t></w:r></w:p>");
        var count = doc.ReplaceText("旧公司名", "新公司");
        Assert.Equal(1, count);
        var runs = doc.Elements[0].Paragraph!.Runs;
        Assert.Equal("新公司", runs[0].Text);
        Assert.True(runs[0].Properties!.Bold);
        Assert.Equal(String.Empty, runs[1].Text);
    }

    [Fact(DisplayName = "替换—忽略大小写")]
    public void Replace_IgnoreCase()
    {
        var doc = Read("<w:p><w:r><w:t>hello WORLD hello</w:t></w:r></w:p>");
        var count = doc.ReplaceText("hello", "hi", ignoreCase: true);
        Assert.Equal(2, count);
        Assert.Equal("hi WORLD hi", Assert.Single(doc.Elements[0].Paragraph!.Runs).Text);
    }

    [Fact(DisplayName = "替换—整词匹配")]
    public void Replace_WholeWord()
    {
        var doc = Read("<w:p><w:r><w:t>cat category cat</w:t></w:r></w:p>");
        var count = doc.ReplaceText("cat", "dog", wholeWord: true);
        Assert.Equal(2, count);
        Assert.Equal("dog category dog", Assert.Single(doc.Elements[0].Paragraph!.Runs).Text);
    }

    [Fact(DisplayName = "替换—表格单元格内替换")]
    public void Replace_InTableCell()
    {
        var body = "<w:tbl><w:tr>"
            + "<w:tc><w:p><w:r><w:t>旧值</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>保持</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var count = doc.ReplaceText("旧值", "新值");
        Assert.Equal(1, count);
        Assert.Contains("新值", GetBodyText(doc));
        Assert.DoesNotContain("旧值", GetBodyText(doc));
    }

    [Fact(DisplayName = "替换—多处替换并返回次数")]
    public void Replace_Multiple()
    {
        var doc = Read("<w:p><w:r><w:t>a X b X c</w:t></w:r></w:p>");
        var count = doc.ReplaceText("X", "Y");
        Assert.Equal(2, count);
        Assert.Equal("a Y b Y c", Assert.Single(doc.Elements[0].Paragraph!.Runs).Text);
    }
    #endregion

    #region 表格行列操作
    [Fact(DisplayName = "表格操作—AddRow/InsertRow/RemoveRow")]
    public void Table_RowOps()
    {
        var table = new Table();
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "A" } } } } });
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "B" } } } } });
        Assert.Equal(2, table.Rows.Count);

        var inserted = table.InsertRow(1);
        inserted.Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "C" } } } } });
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal("C", table.Rows[1].Cells[0].Paragraphs[0].Runs[0].Text);

        table.RemoveRow(0);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("C", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
    }

    [Fact(DisplayName = "表格操作—AddColumn/RemoveColumn")]
    public void Table_ColumnOps()
    {
        var table = new Table();
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "A1" } } } } });
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "A2" } } } } });
        table.AddColumn();
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal(2, table.Rows[1].Cells.Count);

        table.AddColumn(0); // 插入到最前
        Assert.Equal(3, table.Rows[0].Cells.Count);
        Assert.Equal("A1", table.Rows[0].Cells[1].Paragraphs[0].Runs[0].Text);

        table.RemoveColumn(0); // 移除插入的空列，A1 回到首位
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal("A1", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
    }
    #endregion

    #region 表格单元格便利 API
    [Fact(DisplayName = "单元格API—GetCell/GetCellText")]
    public void TableCell_Get()
    {
        var table = new Table();
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "产品" } } } } });
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "笔记本" } } } } });

        Assert.NotNull(table.GetCell(0, 0));
        Assert.Equal("产品", table.GetCellText(0, 0));
        Assert.Equal("笔记本", table.GetCellText(1, 0));
        Assert.Equal(String.Empty, table.GetCellText(0, 5)); // 越界返回空
        Assert.Null(table.GetCell(9, 0));
    }

    [Fact(DisplayName = "单元格API—SetCellText 修改后写回")]
    public void TableCell_SetCellText()
    {
        var table = new Table();
        table.AddRow().Cells.Add(new Cell { Paragraphs = { new Paragraph { Runs = { new Run { Text = "旧值" } } } } });

        table.SetCellText(0, 0, "新值");
        Assert.Equal("新值", table.GetCellText(0, 0));

        // 越界自动补齐行列
        table.SetCellText(2, 3, "补值");
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal("补值", table.GetCellText(2, 3));
    }

    [Fact(DisplayName = "单元格API—读入表格改值后写回往返")]
    public void TableCell_ReadModifyWriteBack()
    {
        var body = "<w:tbl><w:tr>"
            + "<w:tc><w:p><w:r><w:t>旧产品名</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>旧价格</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var el = Assert.Single(doc.Elements, e => e.Type == ElementType.Table);
        var table = el.Table!;
        Assert.Equal("旧产品名", table.GetCellText(0, 0));

        // 通过模型修改内容需显式重建：关闭 document.xml 透传 + 清空 RawXml 兜底
        doc.DocumentXml = null;
        doc.ClearRawXml();
        table.SetCellText(0, 0, "新产品名");
        table.SetCellText(0, 1, "新价格");

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);
            using var reader = new WordReader(path);
            var doc2 = reader.ReadDocument();
            var table2 = Assert.Single(doc2.Elements, e => e.Type == ElementType.Table).Table!;
            Assert.Equal("新产品名", table2.GetCellText(0, 0));
            Assert.Equal("新价格", table2.GetCellText(0, 1));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "GetText—模型全文提取（段落+表格）")]
    public void Document_GetText()
    {
        var body = "<w:p><w:r><w:t>标题行</w:t></w:r></w:p>"
            + "<w:tbl><w:tr>"
            + "<w:tc><w:p><w:r><w:t>列A</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>列B</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var doc = Read(body);
        var text = doc.GetText();
        Assert.Contains("标题行", text);
        Assert.Contains("列A", text);
        Assert.Contains("列B", text);
    }
    #endregion
}
