using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

public class ExcelTemplateTests
{
    /// <summary>创建一个包含占位符的模板xlsx文件</summary>
    private static String CreateTemplateTempFile(String placeholder = "{{Name}}", String sheetPlaceholder = "{{Title}}")
    {
        var path = Path.Combine(Path.GetTempPath(), "template_test_" + Guid.NewGuid().ToString("N") + ".xlsx");

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var za = new ZipArchive(fs, ZipArchiveMode.Create, false, Encoding.UTF8);

        // [Content_Types].xml
        using (var sw = new StreamWriter(za.CreateEntry("[Content_Types].xml").Open(), Encoding.UTF8))
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>");

        // _rels/.rels
        using (var sw = new StreamWriter(za.CreateEntry("_rels/.rels").Open(), Encoding.UTF8))
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

        // workbook
        using (var sw = new StreamWriter(za.CreateEntry("xl/workbook.xml").Open(), Encoding.UTF8))
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");

        // workbook rels
        using (var sw = new StreamWriter(za.CreateEntry("xl/_rels/workbook.xml.rels").Open(), Encoding.UTF8))
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");

        // styles
        using (var sw = new StreamWriter(za.CreateEntry("xl/styles.xml").Open(), Encoding.UTF8))
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs></styleSheet>");

        // sharedStrings - 包含占位符
        using (var sw = new StreamWriter(za.CreateEntry("xl/sharedStrings.xml").Open(), Encoding.UTF8))
            sw.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\"><si><t>{placeholder}</t></si><si><t>Static</t></si></sst>");

        // sheet1 - 包含内联占位符
        using (var sw = new StreamWriter(za.CreateEntry("xl/worksheets/sheet1.xml").Open(), Encoding.UTF8))
            sw.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>{sheetPlaceholder}</t></is></c></row></sheetData></worksheet>");

        return path;
    }

    [Fact, DisplayName("Fill替换共享字符串中的占位符")]
    public void Fill_ReplacesPlaceholdersInSharedStrings()
    {
        var templatePath = CreateTemplateTempFile();
        var outputPath = Path.Combine(Path.GetTempPath(), "output_" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            var template = new ExcelTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object>
            {
                ["Name"] = "TestValue",
                ["Title"] = "Report"
            });

            Assert.True(File.Exists(outputPath));

            // 读取输出文件验证替换
            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var za = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8);

            var sharedEntry = za.GetEntry("xl/sharedStrings.xml");
            Assert.NotNull(sharedEntry);
            using var sr = new StreamReader(sharedEntry!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();

            Assert.Contains("TestValue", xml);
            Assert.DoesNotContain("{{Name}}", xml);
            Assert.Contains("Static", xml); // 未被替换
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact, DisplayName("Fill替换工作表内联文本中的占位符")]
    public void Fill_ReplacesPlaceholdersInSheet()
    {
        var templatePath = CreateTemplateTempFile();
        var outputPath = Path.Combine(Path.GetTempPath(), "output_" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            var template = new ExcelTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object>
            {
                ["Name"] = "Replaced",
                ["Title"] = "MyReport"
            });

            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var za = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8);

            using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();

            Assert.Contains("MyReport", xml);
            Assert.DoesNotContain("{{Title}}", xml);
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact, DisplayName("Fill到流")]
    public void Fill_ToStream()
    {
        var templatePath = CreateTemplateTempFile();

        try
        {
            var template = new ExcelTemplate(templatePath);
            using var output = new MemoryStream();
            template.Fill(output, new Dictionary<String, Object>
            {
                ["Name"] = "StreamValue",
                ["Title"] = "StreamTitle"
            });

            output.Position = 0;
            using var za = new ZipArchive(output, ZipArchiveMode.Read, true, Encoding.UTF8);

            using var sr = new StreamReader(za.GetEntry("xl/sharedStrings.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("StreamValue", xml);
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
        }
    }

    [Fact, DisplayName("Fill未匹配的占位符保持原样")]
    public void Fill_UnmatchedPlaceholders_Preserved()
    {
        var templatePath = CreateTemplateTempFile("{{Name}}", "{{Unknown}}");
        var outputPath = Path.Combine(Path.GetTempPath(), "output_" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            var template = new ExcelTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object>
            {
                ["Name"] = "Filled"
            });

            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var za = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8);

            using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("{{Unknown}}", xml); // 未匹配保持原样
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact, DisplayName("Fill特殊字符自动转义")]
    public void Fill_SpecialCharsEscaped()
    {
        var templatePath = CreateTemplateTempFile();
        var outputPath = Path.Combine(Path.GetTempPath(), "output_" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            var template = new ExcelTemplate(templatePath);
            template.Fill(outputPath, new Dictionary<String, Object>
            {
                ["Name"] = "<script>alert('xss')</script>",
                ["Title"] = "A & B"
            });

            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var za = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8);

            using var sr = new StreamReader(za.GetEntry("xl/sharedStrings.xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.DoesNotContain("<script>", xml); // 转义处理
            Assert.Contains("&lt;script&gt;", xml);
        }
        finally
        {
            if (File.Exists(templatePath)) File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact, DisplayName("构造函数空路径抛出异常")]
    public void Constructor_NullPath_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => new ExcelTemplate(null!));
        Assert.Throws<ArgumentNullException>(() => new ExcelTemplate(""));
    }
}
