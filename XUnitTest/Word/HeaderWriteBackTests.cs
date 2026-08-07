using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>读入→修改页眉页脚→写回测试（用户显式重建 DocumentXml=null 时生效）</summary>
public class HeaderWriteBackTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static Byte[] BuildDocxWithHeader(String headerText)
    {
        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>"
            + "<w:p><w:r><w:t>正文内容</w:t></w:r></w:p>"
            + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>"
            + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>"
            + "</w:body></w:document>";
        var headerXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:t>{headerText}</w:t></w:r></w:p></w:hdr>";
        var relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + $"<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
            + "</Relationships>";

        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(za, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
                + "</Types>");
            WriteEntry(za, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            WriteEntry(za, "word/document.xml", documentXml);
            WriteEntry(za, "word/_rels/document.xml.rels", relsXml);
            WriteEntry(za, "word/header1.xml", headerXml);
            WriteEntry(za, "word/_rels/header1.xml.rels",
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

    private static Document Read(Byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var reader = new WordReader(ms);
        return reader.ReadDocument();
    }

    [Fact(DisplayName = "页眉写回—替换 Header 后保存生效")]
    public void Header_WriteBack()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            var src = Read(BuildDocxWithHeader("旧页眉"));
            Assert.Single(src.Headers);
            Assert.Equal("旧页眉", src.HeaderText);

            // 用户替换页眉并决定重建正文
            src.DocumentXml = null;
            src.Headers.Clear();
            src.Headers.Add(new Header
            {
                Type = "default",
                Elements =
                [
                    new Element
                    {
                        Type = ElementType.Paragraph,
                        Paragraph = new Paragraph
                        {
                            Runs = { new Run { Text = "新页眉", Properties = new RunProperties { Bold = true } } },
                        },
                    },
                ],
            });

            using (var w = new WordWriter()) w.Save(path, src);

            // 验证生成的文件结构
            using (var za = ZipFile.OpenRead(path))
            {
                using (var sr = new StreamReader(za.GetEntry("word/header1.xml")!.Open(), Encoding.UTF8))
                    Assert.Contains("新页眉", sr.ReadToEnd());
                using (var sr = new StreamReader(za.GetEntry("[Content_Types].xml")!.Open(), Encoding.UTF8))
                    Assert.Contains("header1.xml", sr.ReadToEnd());
                using (var sr = new StreamReader(za.GetEntry("word/_rels/document.xml.rels")!.Open(), Encoding.UTF8))
                    Assert.Contains("header1.xml", sr.ReadToEnd());
                using (var sr = new StreamReader(za.GetEntry("word/document.xml")!.Open(), Encoding.UTF8))
                {
                    var xml = sr.ReadToEnd();
                    Assert.Contains("headerReference", xml);
                    Assert.Contains("rHdr1", xml);
                }
            }

            // 读回验证
            var dst = Read(File.ReadAllBytes(path));
            Assert.Single(dst.Headers);
            Assert.Equal("新页眉", dst.HeaderText);
            var para = Assert.Single(dst.Headers[0].Elements).Paragraph!;
            Assert.Equal("新页眉", Assert.Single(para.Runs).Text);
            Assert.True(Assert.Single(para.Runs).Properties!.Bold);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "页脚写回—替换 Footer 后保存生效")]
    public void Footer_WriteBack()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            // 用带页眉的 docx 构造，再附加页脚场景：读入后新增页脚
            var src = Read(BuildDocxWithHeader("表头"));
            src.DocumentXml = null;
            src.Footers.Clear();
            src.Footers.Add(new Footer
            {
                Type = "default",
                Elements =
                [
                    new Element
                    {
                        Type = ElementType.Paragraph,
                        Paragraph = new Paragraph { Runs = { new Run { Text = "第 1 页" } } },
                    },
                ],
            });

            using (var w = new WordWriter()) w.Save(path, src);

            using (var za = ZipFile.OpenRead(path))
            {
                Assert.NotNull(za.GetEntry("word/footer1.xml"));
                using var sr = new StreamReader(za.GetEntry("word/footer1.xml")!.Open(), Encoding.UTF8);
                Assert.Contains("第 1 页", sr.ReadToEnd());
            }

            var dst = Read(File.ReadAllBytes(path));
            Assert.Single(dst.Footers);
            Assert.Contains("第 1 页", dst.Footers[0].Elements[0].Paragraph!.Runs[0].Text);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
