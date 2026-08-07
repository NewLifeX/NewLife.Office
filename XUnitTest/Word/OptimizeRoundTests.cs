using System.IO.Compression;
using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>优化验证测试：ContentTypes 图片扩展名覆盖、ReadTables 嵌套表格排除</summary>
public class OptimizeRoundTests
{
    private const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact(DisplayName = "ContentTypes—gif/svg 等扩展名图片 Default 覆盖")]
    public void ContentTypes_ImageExtensions()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var w = new WordWriter())
            {
                w.AppendParagraph("图片测试");
                w.InsertImage(new Byte[] { 0x47, 0x49, 0x46, 0x38 }, "gif", 3, 2);
                w.InsertImage(new Byte[] { 0x3C, 0x73, 0x76, 0x67 }, "svg", 3, 2);
                w.Save(path);
            }

            using var za = ZipFile.OpenRead(path);
            using var sr = new StreamReader(za.GetEntry("[Content_Types].xml")!.Open(), Encoding.UTF8);
            var xml = sr.ReadToEnd();
            Assert.Contains("<Default Extension=\"gif\" ContentType=\"image/gif\"/>", xml);
            Assert.Contains("<Default Extension=\"svg\" ContentType=\"image/svg+xml\"/>", xml);
            // 生成的媒体部件存在
            var media = za.Entries.Where(e => e.FullName.StartsWith("word/media/")).Select(e => e.FullName).ToList();
            Assert.Contains(media, m => m.EndsWith(".gif"));
            Assert.Contains(media, m => m.EndsWith(".svg"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact(DisplayName = "ReadTables—嵌套表格不作为独立表格返回")]
    public void ReadTables_NestedExcluded()
    {
        var body = "<w:tbl><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>"
            + "<w:tc><w:p><w:r><w:t>嵌套A</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>"
            + "<w:p><w:r><w:t>外层</w:t></w:r></w:p></w:tc>"
            + "</w:tr>"
            + "<w:tr><w:tc><w:p><w:r><w:t>第二行</w:t></w:r></w:p></w:tc></w:tr>"
            + "</w:tbl>";

        var documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{W}\"><w:body>{body}</w:body></w:document>";
        using var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var sw = new StreamWriter(za.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
            sw.Write(documentXml);
        }
        ms.Position = 0;

        using var reader = new WordReader(ms);
        var tables = reader.ReadTables().ToList();
        // 只返回外层 1 张表（2 行），嵌套表格不独立返回
        var tbl = Assert.Single(tables);
        Assert.Equal(2, tbl.Length);
        // 外层单元格文本含嵌套表格文本
        Assert.Contains("嵌套A", tbl[0][0]);
    }
}
