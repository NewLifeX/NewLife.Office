using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPT 功能覆盖补测（S10-04/S12/S13，审计缺口）</summary>
/// <remarks>
/// 覆盖跨文件幻灯片复制（媒体重命名/越界）、动画 XML 写入与回读、
/// 批注/连接器/全局页眉页脚 XML 写入。
/// </remarks>
public class FeatureCoverageTests
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

    private static readonly Byte[] Png1x1 = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
    #endregion

    #region S10-04 跨文件幻灯片复制
    [Fact, DisplayName("S10-04 跨文件复制：媒体重命名与内容保留")]
    public void CopySlideFrom_MediaRename()
    {
        // 源文档：第 1 张含图片 + 文本框，第 2 张含文本框
        var src = BuildPptx(w =>
        {
            w.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "源幻灯片文本" } } });
            w.Slides[0].Images.Add(new Picture { Data = Png1x1, Extension = "png", RelId = "rImg1" });
            w.AddSlide();
            w.Slides[1].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "第二张" } } });
        });

        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.AddSlide();
            writer.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "目标原有" } } });
            var idx = writer.CopySlideFrom(src, 0);
            Assert.Equal(1, idx);
            writer.Save(ms);
        }
        var bytes = ms.ToArray();

        // 复制后媒体被重命名为 m1.png（非 rImg1.png）
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.NotNull(za.GetEntry("ppt/media/m1.png"));
        Assert.Null(za.GetEntry("ppt/media/rImg1.png"));

        // 读取验证：复制来的幻灯片文本保留
        using var reader = new PptxReader(new MemoryStream(bytes));
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        Assert.True(doc.Slides.Count >= 2);
        var texts = doc.Slides.SelectMany(s => s.TextBoxes).Select(t => t.Text).ToList();
        Assert.Contains("目标原有", texts);
        Assert.Contains("源幻灯片文本", texts);
    }

    [Fact, DisplayName("S10-04 跨文件复制：越界索引抛异常")]
    public void CopySlideFrom_OutOfRange()
    {
        var src = BuildPptx(w => w.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000 }));
        using var writer = new PptxWriter();
        writer.AddSlide();
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.CopySlideFrom(src, 99));
    }

    [Fact, DisplayName("S10-04 跨文件复制：图表部件复制与读回")]
    public void CopySlideFrom_ChartParts()
    {
        var src = BuildPptx(w =>
        {
            var c = w.AddBarChart(0, ["分类A", "分类B"]);
            c.Series.Add(new ChartSeries { Name = "复制图表", Values = [11, 22] });
        });

        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.AddSlide();
            var idx = writer.CopySlideFrom(src, 0);
            Assert.Equal(1, idx);
            writer.Save(ms);
        }
        var bytes = ms.ToArray();

        // 图表部件已复制
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.NotNull(za.GetEntry("ppt/charts/chart1.xml"));

        // 读回图表数据
        using var reader = new PptxReader(new MemoryStream(bytes));
        var slides = reader.ReadAllSlides().ToList();
        Assert.True(slides.Count >= 2);
        var chart = slides[1].Charts.FirstOrDefault();
        Assert.NotNull(chart);
        Assert.Equal("复制图表", chart!.Series[0].Name);
        Assert.Equal(22, chart.Series[0].Values[1]);
    }

    [Fact, DisplayName("S10-04 跨文件复制：批注部件复制与读回")]
    public void CopySlideFrom_Comments()
    {
        var src = BuildPptx(w =>
        {
            w.Slides[0].Comments.Add(new Comment { Author = "甲", Text = "复制来的批注" });
        });

        using var ms = new MemoryStream();
        using (var writer = new PptxWriter())
        {
            writer.AddSlide();
            var idx = writer.CopySlideFrom(src, 0);
            Assert.Equal(1, idx);
            writer.Save(ms);
        }
        var bytes = ms.ToArray();

        // 批注部件按输出幻灯片编号命名（目标第 2 张 → comment2.xml），含作者部件
        using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.NotNull(za.GetEntry("ppt/comments/comment2.xml"));
        Assert.NotNull(za.GetEntry("ppt/comments/commentAuthors.xml"));

        // 读回批注
        using var reader = new PptxReader(new MemoryStream(bytes));
        var slides = reader.ReadAllSlides().ToList();
        Assert.True(slides.Count >= 2);
        var comment = slides[1].Comments.FirstOrDefault();
        Assert.NotNull(comment);
        Assert.Contains("复制来的批注", comment!.Text);
    }
    #endregion

    #region S12 动画
    [Fact, DisplayName("S12-01 动画：timing XML 写入")]
    public void Animation_WriteTimingXml()
    {
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "动画目标" } } });
            w.Slides[0].Animations.Add(new Animation { TargetType = "textBox", TargetIndex = 0, Category = AnimationCategory.Entrance, Effect = "fade", DurationMs = 500 });
        });

        var xml = ReadZipEntry(bytes, "ppt/slides/slide1.xml");
        Assert.Contains("<p:timing>", xml);
        Assert.Contains("animEffect", xml);
        Assert.Contains("fade", xml);
    }

    [Fact, DisplayName("S12-02 动画：触发/延迟/时序写入")]
    public void Animation_TriggerTiming()
    {
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "目标" } } });
            w.Slides[0].Animations.Add(new Animation
            {
                TargetType = "textBox",
                TargetIndex = 0,
                Category = AnimationCategory.Entrance,
                Effect = "flyIn",
                Trigger = AnimationTrigger.AfterPrevious,
                DelayMs = 300,
                DurationMs = 800,
            });
        });

        var xml = ReadZipEntry(bytes, "ppt/slides/slide1.xml");
        Assert.Contains("flyIn", xml);
        // 延迟 300ms 与时长 800ms
        Assert.Contains("delay=\"300\"", xml);
        Assert.Contains("dur=\"800\"", xml);
    }

    [Fact, DisplayName("S12-03 动画：读取回读解析")]
    public void Animation_ReadBack()
    {
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].TextBoxes.Add(new TextBox { Left = 100000, Top = 100000, Width = 5000000, Height = 400000, Runs = { new Run { Text = "目标" } } });
            w.Slides[0].Animations.Add(new Animation { TargetType = "textBox", TargetIndex = 0, Category = AnimationCategory.Entrance, Effect = "fade", DurationMs = 500 });
            w.Slides[0].Animations.Add(new Animation { TargetType = "textBox", TargetIndex = 0, Category = AnimationCategory.Emphasis, Effect = "pulse", DurationMs = 400 });
        });

        using var reader = new PptxReader(new MemoryStream(bytes));
        var doc = reader.ReadDocument();
        Assert.NotNull(doc);
        var anims = doc.Slides[0].Animations;
        Assert.Equal(2, anims.Count);
        Assert.Contains(anims, a => a.Effect.Contains("fade"));
        Assert.Contains(anims, a => a.Effect.Contains("pulse"));
    }
    #endregion

    #region S13 批注/连接器/页眉页脚
    [Fact, DisplayName("S13-01 批注：comment XML 写入")]
    public void Comment_WriteXml()
    {
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].Comments.Add(new Comment { Index = 1, Author = "张三", Text = "这是一个批注", X = 0.2f, Y = 0.3f });
        });

        var xml = ReadZipEntry(bytes, "ppt/comments/comment1.xml");
        Assert.Contains("<p:cm", xml);
        Assert.Contains("这是一个批注", xml);
        // 作者记录在 commentAuthors.xml
        var authorsXml = ReadZipEntry(bytes, "ppt/comments/commentAuthors.xml");
        Assert.Contains("张三", authorsXml);
    }

    [Fact, DisplayName("S13-02 连接器：cxnSp XML 写入（类型/箭头/虚线）")]
    public void Connector_WriteXml()
    {
        var bytes = BuildPptx(w =>
        {
            w.Slides[0].Connectors.Add(new ConnectionShape
            {
                Left = 1000000, Top = 1000000, Width = 3000000, Height = 1000000,
                ConnectorType = "elbow",
                LineColor = "FF0000",
                EndArrow = "arrow",
                DashStyle = "dash",
            });
        });

        var xml = ReadZipEntry(bytes, "ppt/slides/slide1.xml");
        Assert.Contains("<p:cxnSp>", xml);
        Assert.Contains("elbow", xml);
        Assert.Contains("FF0000", xml);
        Assert.Contains("arrow", xml);
        Assert.Contains("dash", xml);
    }

    [Fact, DisplayName("S13-03 全局页眉页脚：presentation.xml p:hf 写入")]
    public void HeaderFooter_WriteXml()
    {
        var bytes = BuildPptx(w =>
        {
            w.HeaderFooter = new HeaderFooter
            {
                ShowFooter = true,
                FooterText = "内部资料",
                ShowPageNumber = true,
                ShowDate = true,
                DateFormat = "yyyy/MM/dd",
            };
        });

        var xml = ReadZipEntry(bytes, "ppt/presentation.xml");
        Assert.Contains("<p:hf", xml);
        Assert.Contains("内部资料", xml);
        Assert.Contains("showSlideNum=\"1\"", xml);
        Assert.Contains("dfmt=\"yyyy/MM/dd\"", xml);
    }
    #endregion
}
