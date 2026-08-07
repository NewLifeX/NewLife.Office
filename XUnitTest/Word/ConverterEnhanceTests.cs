using System.IO.Compression;
using System.Text;
using NewLife.Office.Pdf;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>转换器增强测试：HTML 列表/图片/丰富格式/合并单元格，PDF 列表/粗体/分页</summary>
public class ConverterEnhanceTests
{
    static ConverterEnhanceTests() => System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static Byte[] BuildDocx(String documentXml, String? numberingXml = null,
        String? relsXml = null, String mediaName = "", Byte[]? mediaBytes = null)
    {
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels",
                relsXml ?? "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
            if (numberingXml != null) WriteEntry(za, "word/numbering.xml", numberingXml);
            if (mediaName.Length > 0 && mediaBytes != null)
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
        using var sw = new StreamWriter(za.CreateEntry(name).Open(), new UTF8Encoding(false));
        sw.Write(content);
    }

    private static String Doc(String body) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>{body}</w:body></w:document>";

    private const String NumberingXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + $"<w:numbering xmlns:w=\"{W}\">"
        + "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/></w:lvl></w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>"
        + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
        + "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>"
        + "</w:numbering>";

    #region HTML 转换
    [Fact(DisplayName = "HTML—无序列表生成 ul/li")]
    public void Html_BulletList()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>苹果</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>香蕉</w:t></w:r></w:p>";
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body), NumberingXml)));
        Assert.Contains("<ul>", html);
        Assert.Equal(2, CountOccurrences(html, "<li>"));
        Assert.Contains("<li>苹果</li>", html);
        Assert.Contains("<li>香蕉</li>", html);
    }

    [Fact(DisplayName = "HTML—有序列表生成 ol/li")]
    public void Html_OrderedList()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>第一</w:t></w:r></w:p>";
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body), NumberingXml)));
        Assert.Contains("<ol>", html);
        Assert.Contains("<li>第一</li>", html);
    }

    [Fact(DisplayName = "HTML—删除线/上标/下标/高亮")]
    public void Html_RichRunFormats()
    {
        var body = "<w:p>"
            + "<w:r><w:rPr><w:strike/></w:rPr><w:t>划线</w:t></w:r>"
            + "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>2</w:t></w:r>"
            + "<w:r><w:rPr><w:vertAlign w:val=\"subscript\"/></w:rPr><w:t>3</w:t></w:r>"
            + "<w:r><w:rPr><w:highlight w:val=\"yellow\"/></w:rPr><w:t>高亮</w:t></w:r>"
            + "</w:p>";
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body))));
        Assert.Contains("<s>划线</s>", html);
        Assert.Contains("<sup>2</sup>", html);
        Assert.Contains("<sub>3</sub>", html);
        Assert.Contains("<mark>高亮</mark>", html);
    }

    [Fact(DisplayName = "HTML—图片渲染为 img")]
    public void Html_Image()
    {
        var body = "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"3600000\" cy=\"2700000\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "<pic:pic><pic:blipFill><a:blip r:embed=\"rIdImg\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
            + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3600000\" cy=\"2700000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
            + "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdImg\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>"
            + "</Relationships>";
        var png = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body), null, rels, "image1.png", png)));
        Assert.Contains("<img", html);
    }

    [Fact(DisplayName = "HTML—EmbedImages 内嵌图片 data URI（回归#2）")]
    public void Html_EmbedImages()
    {
        var body = "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"3600000\" cy=\"2700000\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            + "<pic:pic><pic:blipFill><a:blip r:embed=\"rIdImg\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
            + "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3600000\" cy=\"2700000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
            + "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdImg\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>"
            + "</Relationships>";
        var png = new Byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        // EmbedImages=true：以 rId 为 key 查找媒体（修复前以文件名查找导致内嵌失效）
        var html = new WordHtmlConverter { FullPage = false, EmbedImages = true }
            .Convert(new MemoryStream(BuildDocx(Doc(body), null, rels, "image1.png", png)));
        Assert.Contains("data:image/png;base64,", html);
        Assert.Contains("<img", html);
    }

    [Fact(DisplayName = "HTML—合并单元格 colspan/rowspan")]
    public void Html_MergedCells()
    {
        var body = "<w:tbl><w:tr>"
            + "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>合并2列</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>"
            + "</w:tr><w:tr>"
            + "<w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>跨行</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>"
            + "</w:tr><w:tr>"
            + "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
            + "<w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>G</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>";
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body))));
        Assert.Contains("colspan=\"2\"", html);
        Assert.Contains("rowspan=\"2\"", html);
    }

    [Fact(DisplayName = "HTML—分页符输出 page-break")]
    public void Html_PageBreak()
    {
        var body = "<w:p><w:r><w:t>第一页</w:t></w:r></w:p>"
            + "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
            + "<w:p><w:r><w:t>第二页</w:t></w:r></w:p>";
        var html = new WordHtmlConverter { FullPage = false }.Convert(new MemoryStream(BuildDocx(Doc(body))));
        Assert.Contains("page-break-after:always", html);
        Assert.Contains("第二页", html);
    }
    #endregion

    #region PDF 转换
    [Fact(DisplayName = "PDF—列表前缀与粗体")]
    public void Pdf_ListAndBold()
    {
        var body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>项目一</w:t></w:r></w:p>"
            + "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗内容</w:t></w:r></w:p>";
        var pdfBytes = new WordPdfConverter().ConvertToBytes(new MemoryStream(BuildDocx(Doc(body), NumberingXml)));
        Assert.NotNull(pdfBytes);
        Assert.True(pdfBytes.Length > 100);

        // 读回 PDF 验证文本内容
        using var pdfReader = new PdfReader(new MemoryStream(pdfBytes));
        var text = pdfReader.ExtractText();
        Assert.Contains("• 项目一", text);
        Assert.Contains("加粗内容", text);
    }

    [Fact(DisplayName = "PDF—分页符生成多页")]
    public void Pdf_PageBreak()
    {
        var body = "<w:p><w:r><w:t>第一页内容</w:t></w:r></w:p>"
            + "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
            + "<w:p><w:r><w:t>第二页内容</w:t></w:r></w:p>";
        var pdfBytes = new WordPdfConverter().ConvertToBytes(new MemoryStream(BuildDocx(Doc(body))));
        using var pdfReader = new PdfReader(new MemoryStream(pdfBytes));
        Assert.True(pdfReader.GetPageCount() >= 2);
        var text = pdfReader.ExtractText();
        Assert.Contains("第一页内容", text);
        Assert.Contains("第二页内容", text);
    }
    #endregion

    private static Int32 CountOccurrences(String haystack, String needle)
    {
        var count = 0;
        var idx = 0;
        while ((idx = haystack.IndexOf(needle, idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += needle.Length;
        }
        return count;
    }
}
