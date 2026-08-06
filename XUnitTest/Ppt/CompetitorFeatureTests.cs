using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPT 竞品功能测试（S17-S20，对标 Aspose.Slides/ShapeCrawler）</summary>
/// <remarks>
/// 覆盖打开密码 AES 加密（S17）、形状图片填充（S18）、图表样式配色（S19）、超链接跳转文件（S20）。
/// </remarks>
public class CompetitorFeatureTests
{
    #region 辅助
    private static Byte[] BuildPptx(Action<PptxWriter> build)
    {
        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.AddSlide();
            build(writer);
            writer.Save(ms);
        }
        return ms.ToArray();
    }

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

    #region S17 打开密码 AES
    [Fact, DisplayName("S17 加密：SaveEncrypted 生成 OLE2 加密容器")]
    public void Encrypt_SaveEncrypted()
    {
        var path = Path.Combine(Path.GetTempPath(), "enc_" + Guid.NewGuid().ToString("N") + ".pptx");
        try
        {
            using (var writer = new PptxWriter())
            {
                writer.AddSlide();
                writer.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 8000000, Height = 400000, Runs = { new Run { Text = "加密内容" } } });
                writer.SaveEncrypted(path, "pwd123");
            }

            var bytes = File.ReadAllBytes(path);
            // OLE2 CFB 魔数
            Assert.True(bytes.Length > 8);
            Assert.Equal(0xD0, bytes[0]);
            Assert.Equal(0xCF, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xE0, bytes[3]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact, DisplayName("S17 加密：带密码读取还原内容")]
    public void Encrypt_ReadWithPassword()
    {
        var path = Path.Combine(Path.GetTempPath(), "enc_" + Guid.NewGuid().ToString("N") + ".pptx");
        try
        {
            using (var writer = new PptxWriter())
            {
                writer.AddSlide();
                writer.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 8000000, Height = 400000, Runs = { new Run { Text = "加密内容" } } });
                writer.SaveEncrypted(path, "pwd123");
            }

            using var reader = new PptxReader(path, "pwd123");
            var doc = reader.ReadDocument();
            Assert.NotNull(doc);
            Assert.True(doc.Slides.Count >= 1);
            var slide = doc.Slides[0];
            var text = slide.TextBoxes.FirstOrDefault()?.Text ?? "";
            Assert.Contains("加密内容", text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact, DisplayName("S17 加密：无密码读取抛异常")]
    public void Encrypt_NoPassword_Throws()
    {
        var path = Path.Combine(Path.GetTempPath(), "enc_" + Guid.NewGuid().ToString("N") + ".pptx");
        try
        {
            using (var writer = new PptxWriter())
            {
                writer.AddSlide();
                writer.SaveEncrypted(path, "pwd123");
            }
            Assert.Throws<InvalidOperationException>(() => new PptxReader(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact, DisplayName("S17 加密：错误密码读取抛异常")]
    public void Encrypt_WrongPassword_Throws()
    {
        var path = Path.Combine(Path.GetTempPath(), "enc_" + Guid.NewGuid().ToString("N") + ".pptx");
        try
        {
            using (var writer = new PptxWriter())
            {
                writer.AddSlide();
                writer.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 8000000, Height = 400000, Runs = { new Run { Text = "机密内容" } } });
                writer.SaveEncrypted(path, "pwd123");
            }
            // 错误密码：Agile 解密密钥不符 → 解密/解包失败抛异常
            Assert.ThrowsAny<Exception>(() => new PptxReader(path, "wrongpwd"));
        }
        finally
        {
            File.Delete(path);
        }
    }
    #endregion

    #region S18 形状图片填充
    [Fact, DisplayName("S18 图片填充：写入与读取还原")]
    public void FillImage_WriteRead()
    {
        // 1x1 PNG
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].Shapes.Add(new Shape
            {
                ShapeType = "rect",
                Left = 1000000, Top = 1000000, Width = 3000000, Height = 2000000,
                FillImage = png,
                FillImageExt = "png",
            });
        });

        // slide rels 含图片填充关系
        var rels = ReadZipEntry(bytes, "ppt/slides/_rels/slide1.xml.rels");
        Assert.Contains("rImgFill", rels);

        // 媒体文件存在
        var mediaEntry = GetEntry(bytes, "ppt/media/rImgFill");
        Assert.NotNull(mediaEntry);

        // 读取还原
        using var reader = new PptxReader(new MemoryStream(bytes));
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        var shapes = doc.Slides[0].Shapes;
        var fillShape = shapes.FirstOrDefault(s => s.FillImage is { Length: > 0 });
        Assert.NotNull(fillShape);
        Assert.Equal(png.Length, fillShape!.FillImage!.Length);
    }

    [Fact, DisplayName("S18 图片填充：组形状内写入与读取还原")]
    public void FillImage_GroupShape()
    {
        // 1x1 PNG
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        var bytes = BuildPptx(w =>
        {
            var grp = new GroupShape { Left = 500000, Top = 500000, Width = 4000000, Height = 3000000 };
            grp.Shapes.Add(new Shape
            {
                ShapeType = "ellipse",
                Left = 1000000, Top = 1000000, Width = 2000000, Height = 1500000,
                FillImage = png,
                FillImageExt = "png",
            });
            w.Slides[0].Groups.Add(grp);
        });

        // slide rels 含图片填充关系（组内形状）
        var rels = ReadZipEntry(bytes, "ppt/slides/_rels/slide1.xml.rels");
        Assert.Contains("rImgFill", rels);

        // 媒体文件存在
        var mediaEntry = GetEntry(bytes, "ppt/media/rImgFill");
        Assert.NotNull(mediaEntry);

        // 读取还原：组内形状 FillImage
        using var reader = new PptxReader(new MemoryStream(bytes));
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        Assert.NotEmpty(doc.Slides[0].Groups);
        var grpShapes = doc.Slides[0].Groups[0].Shapes;
        var fillShape = grpShapes.FirstOrDefault(s => s.FillImage is { Length: > 0 });
        Assert.NotNull(fillShape);
        Assert.Equal(png.Length, fillShape!.FillImage!.Length);
    }

    private static ZipArchiveEntry? GetEntry(Byte[] pptx, String prefix)
    {
        using var ms = new MemoryStream(pptx);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read);
        return za.Entries.FirstOrDefault(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
    }
    #endregion

    #region S19 图表样式配色
    [Fact, DisplayName("S19 图表：样式索引与系列配色写入")]
    public void Chart_StyleAndColor()
    {
        var bytes = BuildPptx(w =>
        {
            var chart = new Chart
            {
                ChartType = "bar",
                Title = "销售",
                Categories = ["一月", "二月", "三月"],
                StyleIndex = 10,
                Left = 1000000, Top = 1000000, Width = 8000000, Height = 4000000,
            };
            chart.Series.Add(new ChartSeries { Name = "销售额", Values = [100, 150, 200], Color = "FF8800" });
            w.Slides[0].Charts.Add(chart);
        });

        // ChartNumber 为 internal，程序化 Add 时默认 0 → chart0.xml
        var chartXml = ReadZipEntry(bytes, "ppt/charts/chart0.xml");
        Assert.Contains("<c:style val=\"10\"/>", chartXml);
        Assert.Contains("FF8800", chartXml);
    }
    #endregion

    #region S20 超链接跳转文件
    [Fact, DisplayName("S20 超链接：TextBox 跳转文件写入与读取还原")]
    public void Hyperlink_FileLink()
    {
        var target = @"C:\data\report.pdf";
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].TextBoxes.Add(new TextBox
            {
                Left = 100000, Top = 100000, Width = 5000000, Height = 400000,
                HyperlinkUrl = target,
                Runs = { new Run { Text = "点击打开报告" } },
            });
        });

        // slide rels 含 External 超链接（文件路径）
        var rels = ReadZipEntry(bytes, "ppt/slides/_rels/slide1.xml.rels");
        Assert.Contains("relationships/hyperlink", rels);
        Assert.Contains("report.pdf", rels);
        Assert.Contains("TargetMode=\"External\"", rels);

        // 读取还原
        using var reader = new PptxReader(new MemoryStream(bytes));
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        var tb = doc.Slides[0].TextBoxes.FirstOrDefault();
        Assert.NotNull(tb);
        Assert.Equal(target, tb!.HyperlinkUrl);
    }
    #endregion
}
