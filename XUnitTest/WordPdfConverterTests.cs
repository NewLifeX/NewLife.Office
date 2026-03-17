using System.IO.Compression;
using System.Text;
using System.Xml;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>WordPdfConverter docx→PDF 转换器单元测试</summary>
public class WordPdfConverterTests
{
    // ─── docx 构建辅助（与 WordHtmlConverterTests 相同模式）──────────────

    private static Stream BuildDocx(String documentXml)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(zip, "word/document.xml", documentXml);
            AddEntry(zip, "[Content_Types].xml",
                "<?xml version=\"1.0\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>");
        }
        ms.Position = 0;
        return ms;
    }

    private static void AddEntry(ZipArchive zip, String name, String content)
    {
        var entry = zip.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
        writer.Write(content);
    }

    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static String Para(String text, String styleId = null)
    {
        var stylePart = styleId == null ? String.Empty
            : $"<w:pPr><w:pStyle w:val=\"{styleId}\"/></w:pPr>";
        return $"<w:p>{stylePart}<w:r><w:t>{System.Security.SecurityElement.Escape(text)}</w:t></w:r></w:p>";
    }

    private static String Table(params String[][] rows)
    {
        var sb = new StringBuilder("<w:tbl>");
        foreach (var row in rows)
        {
            sb.Append("<w:tr>");
            foreach (var cell in row)
                sb.Append($"<w:tc><w:p><w:r><w:t>{System.Security.SecurityElement.Escape(cell)}</w:t></w:r></w:p></w:tc>");
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
        return sb.ToString();
    }

    private static String WrapDoc(String bodyContent)
        => $"<?xml version=\"1.0\"?><w:document xmlns:w=\"{W}\"><w:body>{bodyContent}</w:body></w:document>";

    // ─── 测试 ─────────────────────────────────────────────────────────────

    [Fact, System.ComponentModel.DisplayName("ConvertToBytes 单段落返回合法 PDF 字节")]
    public void ConvertToBytes_SingleParagraph_ReturnsValidPdf()
    {
        using var docx = BuildDocx(WrapDoc(Para("Hello from Word")));
        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 100);
        // PDF magic bytes: %PDF
        Assert.Equal((Byte)'%', pdfBytes[0]);
        Assert.Equal((Byte)'P', pdfBytes[1]);
        Assert.Equal((Byte)'D', pdfBytes[2]);
        Assert.Equal((Byte)'F', pdfBytes[3]);
    }

    [Fact, System.ComponentModel.DisplayName("ConvertToBytes 多段落不抛异常")]
    public void ConvertToBytes_MultipleParagraphs_NoException()
    {
        var body = Para("First paragraph") + Para("Second paragraph") + Para("Third paragraph");
        using var docx = BuildDocx(WrapDoc(body));
        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotEmpty(pdfBytes);
    }

    [Fact, System.ComponentModel.DisplayName("Heading1 样式段落不抛异常（字号更大）")]
    public void ConvertToBytes_Heading1Paragraph_NoException()
    {
        var body = Para("Chapter 1", "Heading1") + Para("Body text below heading");
        using var docx = BuildDocx(WrapDoc(body));
        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotEmpty(pdfBytes);
    }

    [Fact, System.ComponentModel.DisplayName("含表格的文档正常转换")]
    public void ConvertToBytes_WithTable_NoException()
    {
        var tbl = Table(new[] { "Name", "Age" }, new[] { "Alice", "30" }, new[] { "Bob", "25" });
        var body = Para("Employee List") + tbl;
        using var docx = BuildDocx(WrapDoc(body));
        var converter = new WordPdfConverter();
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotEmpty(pdfBytes);
    }

    [Fact, System.ComponentModel.DisplayName("空文档返回空 PDF 字节（但不抛异常）")]
    public void ConvertToBytes_EmptyDocument_ReturnsBytes()
    {
        // 空 body 不含任何内容
        using var docx = BuildDocx(WrapDoc(String.Empty));
        var converter = new WordPdfConverter();
        // 即使无内容，PdfFluentDocument 也会生成有效的空 PDF
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotNull(pdfBytes);
    }

    [Fact, System.ComponentModel.DisplayName("ConvertToFile 输出文件到流")]
    public void ConvertToFile_WritesToPath()
    {
        using var docx = BuildDocx(WrapDoc(Para("Test output")));
        var tmpPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pdf");
        try
        {
            var converter = new WordPdfConverter();
            // 使用流版本 ConvertToFile
            docx.Position = 0;
            converter.ConvertToFile(docx, tmpPath);
            Assert.True(File.Exists(tmpPath));
            var fileBytes = File.ReadAllBytes(tmpPath);
            Assert.True(fileBytes.Length > 100);
        }
        finally
        {
            if (File.Exists(tmpPath)) File.Delete(tmpPath);
        }
    }

    [Fact, System.ComponentModel.DisplayName("可配置字号属性")]
    public void BodyFontSize_IsConfigurable()
    {
        using var docx = BuildDocx(WrapDoc(Para("Font size test")));
        var converter = new WordPdfConverter { BodyFontSize = 14f, H1FontSize = 28f };
        var pdfBytes = converter.ConvertToBytes(docx);
        Assert.NotEmpty(pdfBytes);
    }
}
