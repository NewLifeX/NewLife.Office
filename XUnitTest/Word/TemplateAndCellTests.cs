using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>WordTemplate 拆分占位符合并 + 单元格边框读写测试</summary>
public class TemplateAndCellTests
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

    private static void WriteEntry(ZipArchive za, String name, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(name).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static String ReadDocumentXml(String docxPath)
    {
        using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return String.Empty;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    #region WordTemplate 拆分占位符合并
    [Fact(DisplayName = "模板—跨 Run 拆分占位符合并替换")]
    public void Template_SplitPlaceholder()
    {
        var templatePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        var outputPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            // 用两个 Run 模拟 Word 拆分占位符：{{Na | me}}
            using (var w = new WordWriter())
            {
                w.AppendFormattedParagraph(new[]
                {
                    new Run { Text = "你好 {{Na" },
                    new Run { Text = "me}}，欢迎！" },
                });
                w.Save(templatePath);
            }

            var template = new WordTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object?> { ["Name"] = "张三" });

            using var reader = new WordReader(outputPath);
            var text = reader.ExtractText() ?? "";
            Assert.Contains("你好 张三，欢迎！", text);
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact(DisplayName = "模板—拆分占位符含 run 格式不误替换")]
    public void Template_SplitPlaceholder_NoFalsePositive()
    {
        var templatePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        var outputPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter())
            {
                // 非占位符文本（{{ 中间有真实内容），不应被误压缩
                w.AppendFormattedParagraph(new[]
                {
                    new Run { Text = "价格 {{10" },
                    new Run { Text = "0}} 元" },
                });
                w.Save(templatePath);
            }

            var template = new WordTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object?> { ["Name"] = "张三" });

            using var reader = new WordReader(outputPath);
            var text = reader.ExtractText() ?? "";
            Assert.Contains("价格 {{100}} 元", text); // 非占位符保持原样
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
    #endregion

    #region 单元格边框
    [Fact(DisplayName = "单元格边框—tcBorders 读取")]
    public void CellBorder_Read()
    {
        var body = "<w:tbl><w:tr>"
            + "<w:tc><w:tcPr><w:tcBorders>"
            + "<w:top w:val=\"single\" w:sz=\"8\" w:color=\"FF0000\"/>"
            + "<w:bottom w:val=\"double\" w:sz=\"6\" w:color=\"0000FF\"/>"
            + "</w:tcBorders></w:tcPr>"
            + "<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        using var ms = new MemoryStream(BuildDocx(body));
        using var reader = new WordReader(ms);
        var doc = reader.ReadDocument();
        var cell = doc.Elements[0].Table!.Rows[0].Cells[0];
        Assert.NotNull(cell.Borders);
        Assert.Equal(BorderStyle.Single, cell.Borders!.Top!.Style);
        Assert.Equal("FF0000", cell.Borders.Top.Color);
        Assert.Equal(BorderStyle.Double, cell.Borders.Bottom!.Style);
    }

    [Fact(DisplayName = "单元格边框—模型写回 tcBorders")]
    public void CellBorder_WriteBack()
    {
        var table = new Table();
        var row = table.AddRow();
        var cell = new Cell
        {
            Borders = new TableBorders
            {
                Top = new Border { Style = BorderStyle.Single, Color = "FF0000", Width = 8 },
            },
        };
        cell.Paragraphs.Add(new Paragraph { Runs = { new Run { Text = "A" } } });
        row.Cells.Add(cell);

        var doc = new Document { DocumentXml = null };
        doc.Elements.Add(new Element { Type = ElementType.Table, Table = table });

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter()) w.Save(path, doc);
            var xml = ReadDocumentXml(path);
            Assert.Contains("<w:tcBorders>", xml);
            Assert.Contains("<w:top w:val=\"single\"", xml);
            Assert.Contains("w:color=\"FF0000\"", xml);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
    #endregion
}
