using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPT 文档属性测试（S14-01/02/03，审计补测）</summary>
/// <remarks>
/// 覆盖 docProps/core.xml 与 app.xml 的写入（PptxWriter.WriteDocProps）
/// 与读取（PptxReader.ParseDocProps）往返，含空属性边界与缺失部件容错。
/// </remarks>
public class DocumentPropertiesTests
{
    #region 辅助
    /// <summary>用指定文档属性构建 pptx 字节（经 Save(stream, document) 走文档模型通道）</summary>
    private static Byte[] BuildPptx(DocumentProperties props)
    {
        var doc = new Presentation { Properties = props };
        var slide = new Slide { Layout = "title_only" };
        slide.TextBoxes.Add(new TextBox { Text = "标题", Role = "title" });
        doc.Slides.Add(slide);

        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.Save(ms, doc);
        }
        return ms.ToArray();
    }

    /// <summary>读取 pptx ZIP 中指定部件的文本，不存在返回空串</summary>
    private static String ReadZipEntry(Byte[] pptx, String path)
    {
        using var ms = new MemoryStream(pptx);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = za.GetEntry(path);
        if (entry == null) return String.Empty;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }
    #endregion

    #region S14-01 core.xml 写入
    [Fact(DisplayName = "S14-01 core.xml 写入：Title/Author/Subject/Description")]
    public void DocProps_WriteCoreXml()
    {
        var xml = ReadZipEntry(BuildPptx(new DocumentProperties
        {
            Title = "季度报告",
            Author = "张三",
            Subject = "销售数据",
            Description = "2026 Q2 报告",
        }), "docProps/core.xml");

        Assert.NotEmpty(xml);
        Assert.Contains("<dc:title>季度报告</dc:title>", xml);
        Assert.Contains("<dc:creator>张三</dc:creator>", xml);
        Assert.Contains("<dc:subject>销售数据</dc:subject>", xml);
        Assert.Contains("<dc:description>2026 Q2 报告</dc:description>", xml);
    }

    [Fact(DisplayName = "S14-01 core.xml 边界：空属性不生成 docProps 部件")]
    public void DocProps_Empty_NoParts()
    {
        var core = ReadZipEntry(BuildPptx(new DocumentProperties()), "docProps/core.xml");
        var app = ReadZipEntry(BuildPptx(new DocumentProperties()), "docProps/app.xml");

        Assert.Equal(String.Empty, core);
        Assert.Equal(String.Empty, app);
    }
    #endregion

    #region S14-02 app.xml 写入
    [Fact(DisplayName = "S14-02 app.xml 写入：Slides 数量与 Company")]
    public void DocProps_WriteAppXml()
    {
        var xml = ReadZipEntry(BuildPptx(new DocumentProperties { Author = "张三" }), "docProps/app.xml");

        Assert.NotEmpty(xml);
        Assert.Contains("<Slides>1</Slides>", xml);
        Assert.Contains("<Company>张三</Company>", xml);
    }
    #endregion

    #region S14-03 docProps 读取
    [Fact(DisplayName = "S14-03 docProps 往返读取：四字段精确相等")]
    public void DocProps_ReadRoundtrip()
    {
        var bytes = BuildPptx(new DocumentProperties
        {
            Title = "季度报告",
            Author = "张三",
            Subject = "销售数据",
            Description = "2026 Q2 报告",
        });

        using var ms = new MemoryStream(bytes);
        using var reader = new PptxReader(ms);
        var doc = reader.ReadDocument();

        Assert.NotNull(doc);
        Assert.Equal("季度报告", doc.Properties.Title);
        Assert.Equal("张三", doc.Properties.Author);
        Assert.Equal("销售数据", doc.Properties.Subject);
        Assert.Equal("2026 Q2 报告", doc.Properties.Description);
    }

    [Fact(DisplayName = "S14-03 docProps 容错：缺失部件时读取安全，属性保持默认")]
    public void DocProps_MissingParts_ReadSafe()
    {
        // 不带文档属性的演示文稿（无 docProps 部件）
        var doc = new Presentation();
        var slide = new Slide { Layout = "blank" };
        doc.Slides.Add(slide);

        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.Save(ms, doc);
        }

        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var read = reader.ReadDocument();

        Assert.NotNull(read);
        Assert.Null(read.Properties.Title);
        Assert.Null(read.Properties.Author);
        Assert.Null(read.Properties.Subject);
        Assert.Null(read.Properties.Description);
    }
    #endregion
}
