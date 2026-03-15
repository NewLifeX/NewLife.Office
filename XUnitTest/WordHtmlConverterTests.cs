using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>WordHtmlConverter 单元测试</summary>
public class WordHtmlConverterTests
{
    // ─── 构建最小 docx 的辅助 ──────────────────────────────────────────────

    /// <summary>构建一个最小合法的 docx 字节流</summary>
    private static Byte[] BuildDocx(String documentXml, String? relsXml = null, String? mediaName = null, Byte[]? mediaBytes = null)
    {
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: Encoding.UTF8))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>");

            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>");

            WriteEntry(za, "word/document.xml", documentXml);

            WriteEntry(za, "word/_rels/document.xml.rels",
                relsXml ?? "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");

            if (mediaName != null && mediaBytes != null)
            {
                var e = za.CreateEntry("word/media/" + mediaName);
                using var es = e.Open();
                es.Write(mediaBytes, 0, mediaBytes.Length);
            }
        }
        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive za, String name, String content)
    {
        var e = za.CreateEntry(name);
        using var sw = new StreamWriter(e.Open(), Encoding.UTF8);
        sw.Write(content);
    }

    private static String W => "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
    private static String WR => "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    // ─── 基础段落 ──────────────────────────────────────────────────────────

    [Fact, DisplayName("普通段落转为 <p> 元素")]
    public void NormalParagraph_Renders_P_Tag()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>";
        var doc = BuildDocx(xml);
        var conv = new WordHtmlConverter { FullPage = false };
        var html = conv.Convert(new MemoryStream(doc));
        Assert.Contains("<p>", html);
        Assert.Contains("Hello World", html);
        Assert.Contains("</p>", html);
    }

    [Fact, DisplayName("空段落仍然生成 <p> 元素")]
    public void EmptyParagraph_Renders_Empty_P()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body><w:p/></w:body></w:document>";
        var doc = BuildDocx(xml);
        var conv = new WordHtmlConverter { FullPage = false };
        var html = conv.Convert(new MemoryStream(doc));
        Assert.Contains("<p>", html);
    }

    // ─── 标题 ─────────────────────────────────────────────────────────────

    [Fact, DisplayName("Heading1 样式段落转为 <h1>")]
    public void Heading1_Renders_H1()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<h1>", html);
        Assert.Contains("Title", html);
        Assert.Contains("</h1>", html);
    }

    [Fact, DisplayName("w:val=\"2\" 识别为 Heading2，转为 <h2>")]
    public void HeadingByNumber_Renders_H2()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:pPr><w:pStyle w:val=\"2\"/></w:pPr><w:r><w:t>Sec</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<h2>", html);
    }

    [Fact, DisplayName("HEADING3（大写+数字）识别为 <h3>")]
    public void HeadingUppercase_Renders_H3()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:pPr><w:pStyle w:val=\"HEADING3\"/></w:pPr><w:r><w:t>Sub</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<h3>", html);
    }

    // ─── 文字格式 ──────────────────────────────────────────────────────────

    [Fact, DisplayName("粗体 run 包裹 <strong>")]
    public void Bold_Run_Renders_Strong()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<strong>", html);
        Assert.Contains("Bold", html);
    }

    [Fact, DisplayName("斜体 run 包裹 <em>")]
    public void Italic_Run_Renders_Em()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<em>", html);
        Assert.Contains("Italic", html);
    }

    [Fact, DisplayName("下划线 run 包裹 <u>")]
    public void Underline_Run_Renders_U()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t>Underline</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<u>", html);
    }

    [Fact, DisplayName("w:color 生成 color 内联 CSS")]
    public void ColoredRun_Renders_SpanStyle()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:r><w:rPr><w:color w:val=\"FF0000\"/></w:rPr><w:t>Red</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("color:#FF0000", html);
        Assert.Contains("Red", html);
    }

    // ─── 对齐 ─────────────────────────────────────────────────────────────

    [Fact, DisplayName("居中段落带 text-align:center 样式")]
    public void CenteredParagraph_Renders_TextAlignCenter()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("text-align:center", html);
    }

    // ─── 超链接 ───────────────────────────────────────────────────────────

    [Fact, DisplayName("超链接段落生成 <a href=...> 元素")]
    public void Hyperlink_Renders_Anchor()
    {
        var docXml = $"<?xml version=\"1.0\"?><w:document {W} {WR}><w:body>" +
                     "<w:p><w:hyperlink r:id=\"rId1\"><w:r><w:t>Click</w:t></w:r></w:hyperlink></w:p>" +
                     "</w:body></w:document>";
        var relsXml = "<?xml version=\"1.0\"?>" +
                      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                      "</Relationships>";
        var doc = BuildDocx(docXml, relsXml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<a href=", html);
        Assert.Contains("https://example.com", html);
        Assert.Contains("Click", html);
    }

    // ─── 表格 ─────────────────────────────────────────────────────────────

    [Fact, DisplayName("表格生成 <table>/<th>/<td> 结构")]
    public void Table_Renders_HtmlTable()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:tbl>" +
                  "<w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Age</w:t></w:r></w:p></w:tc></w:tr>" +
                  "<w:tr><w:tc><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>30</w:t></w:r></w:p></w:tc></w:tr>" +
                  "</w:tbl>" +
                  "</w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("<table", html);
        Assert.Contains("<th>", html);
        Assert.Contains("<td>", html);
        Assert.Contains("Name", html);
        Assert.Contains("Alice", html);
    }

    // ─── FullPage ─────────────────────────────────────────────────────────

    [Fact, DisplayName("FullPage=true 输出完整 DOCTYPE 页面")]
    public void FullPage_True_Renders_DocType()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = true, PageTitle = "TestDoc" }.Convert(new MemoryStream(doc));
        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("<title>TestDoc</title>", html);
        Assert.Contains("<body>", html);
        Assert.Contains("</html>", html);
    }

    [Fact, DisplayName("FullPage=false 只输出 body 片段，不含 DOCTYPE")]
    public void FullPage_False_Renders_Fragment()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>";
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.DoesNotContain("<!DOCTYPE", html);
        Assert.DoesNotContain("<html>", html);
        Assert.Contains("<p>", html);
    }

    // ─── HTML 特殊字符转义 ────────────────────────────────────────────────

    [Fact, DisplayName("文本中的 < > & 被正确转义")]
    public void HtmlSpecialChars_AreEscaped()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                  "<w:p><w:r><w:t xml:space=\"preserve\">a &lt; b &amp; c</w:t></w:r></w:p>" +
                  "</w:body></w:document>";
        // XML 中已经转义，经过 XmlDocument 解析后 InnerText 会是 "a < b & c"
        // HtmlEncode 应将其转为 "a &lt; b &amp; c"
        var doc = BuildDocx(xml);
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(doc));
        Assert.Contains("&lt;", html);
        Assert.Contains("&amp;", html);
    }

    // ─── ConvertFromFile ──────────────────────────────────────────────────

    [Fact, DisplayName("ConvertFromFile 从临时文件转换正确")]
    public void ConvertFromFile_Works()
    {
        var xml = $"<?xml version=\"1.0\"?><w:document {W}><w:body><w:p><w:r><w:t>FileTest</w:t></w:r></w:p></w:body></w:document>";
        var bytes = BuildDocx(xml);
        var path = Path.Combine(Path.GetTempPath(), $"whc_{Guid.NewGuid():N}.docx");
        try
        {
            File.WriteAllBytes(path, bytes);
            var html = new WordHtmlConverter { FullPage = false }.ConvertFromFile(path);
            Assert.Contains("FileTest", html);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
