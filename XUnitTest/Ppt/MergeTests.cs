using System.ComponentModel;
using System.IO.Compression;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PptxWriter.Merge 合并保真测试（S05-02）</summary>
/// <remarks>
/// 验证合并后图表/批注/备注/嵌入部件完整保留、关系重映射正确、幻灯片尺寸保留。
/// </remarks>
public class MergeTests
{
    /// <summary>构建含指定内容的 pptx 字节</summary>
    private static Byte[] BuildFile(Action<PptxWriter> build)
    {
        using var writer = new PptxWriter();
        build(writer);
        using var ms = new MemoryStream();
        writer.Save(ms);
        return ms.ToArray();
    }

    /// <summary>合并多个文件，返回可读流</summary>
    private static MemoryStream Merge(params Byte[][] files)
    {
        var ms = new MemoryStream();
        PptxWriter.Merge(files, ms);
        ms.Position = 0;
        return ms;
    }

    /// <summary>读取合并结果的 ZIP 条目名集合</summary>
    private static HashSet<String> ZipEntryNames(Stream merged)
    {
        merged.Position = 0;
        using var zip = new ZipArchive(merged, ZipArchiveMode.Read, leaveOpen: true);
        return [.. zip.Entries.Select(e => e.FullName)];
    }

    [Fact, DisplayName("PPT—合并两个文件幻灯片数量与文本")]
    public void Merge_TwoFiles_SlidesCombined()
    {
        var a = BuildFile(w => { w.AddSlide(); w.AddTextBox(0, "甲页内容", 2, 2, 10, 2); });
        var b = BuildFile(w => { w.AddSlide(); w.AddTextBox(0, "乙页内容", 2, 2, 10, 2); });

        using var ms = Merge(a, b);
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Equal(2, slides.Count);
        Assert.Contains("甲页内容", slides[0].TextBoxes[0].Text);
        Assert.Contains("乙页内容", slides[1].TextBoxes[0].Text);
    }

    [Fact, DisplayName("PPT—合并保留各文件图表（部件重命名无冲突）")]
    public void Merge_PreservesCharts_FromAllFiles()
    {
        var a = BuildFile(w =>
        {
            w.AddSlide();
            var c = w.AddBarChart(0, ["分类A", "分类B"]);
            c.Series.Add(new ChartSeries { Name = "甲系列", Values = [10, 20] });
        });
        var b = BuildFile(w =>
        {
            w.AddSlide();
            var c = w.AddPieChart(0, ["分类C", "分类D"]);
            c.Series.Add(new ChartSeries { Name = "乙系列", Values = [30, 40] });
        });

        using var ms = Merge(a, b);
        // 图表部件重命名不冲突：chart1.xml + chart2.xml
        var entries = ZipEntryNames(ms);
        Assert.Contains("ppt/charts/chart1.xml", entries);
        Assert.Contains("ppt/charts/chart2.xml", entries);
        // [Content_Types] 覆盖两个图表
        using var ctZip = new ZipArchive(new MemoryStream(ms.ToArray()), ZipArchiveMode.Read);
        var ct = ctZip.GetEntry("[Content_Types].xml");
        String ctXml;
        using (var sr = new StreamReader(ct!.Open())) ctXml = sr.ReadToEnd();
        Assert.Contains("/ppt/charts/chart1.xml", ctXml);
        Assert.Contains("/ppt/charts/chart2.xml", ctXml);

        // 读回：每张幻灯片各有一个图表，系列数据完整
        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Equal(2, slides.Count);
        Assert.Single(slides[0].Charts);
        Assert.Single(slides[1].Charts);
        Assert.Equal("甲系列", slides[0].Charts[0].Series[0].Name);
        Assert.Equal(20, slides[0].Charts[0].Series[0].Values[1]);
        Assert.Equal("乙系列", slides[1].Charts[0].Series[0].Name);
        Assert.Equal(40, slides[1].Charts[0].Series[0].Values[1]);
    }

    [Fact, DisplayName("PPT—合并保留批注（按输出幻灯片编号命名）")]
    public void Merge_PreservesComments()
    {
        var a = BuildFile(w =>
        {
            w.AddSlide();
            w.AddTextBox(0, "甲页", 2, 2, 10, 2);
            w.Slides[0].Comments.Add(new Comment { Author = "测试者", Text = "甲页批注" });
        });
        var b = BuildFile(w => { w.AddSlide(); w.AddTextBox(0, "乙页", 2, 2, 10, 2); });

        using var ms = Merge(a, b);
        var entries = ZipEntryNames(ms);
        // 批注按输出幻灯片编号命名 comment1.xml（甲页为输出第 1 张）
        Assert.Contains("ppt/comments/comment1.xml", entries);
        Assert.Contains("ppt/comments/commentAuthors.xml", entries);

        // 读回批注
        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Single(slides[0].Comments);
        Assert.Contains("甲页批注", slides[0].Comments[0].Text);
        Assert.Empty(slides[1].Comments);
    }

    [Fact, DisplayName("PPT—合并保留各文件备注文本")]
    public void Merge_PreservesNotes()
    {
        // 注：当前 Writer 将备注以内嵌 <p:notes> 写在幻灯片 XML 内（非独立 notesSlide 部件），
        // 合并时随幻灯片整体透传，故通过 Reader 读回验证备注保留
        var a = BuildFile(w => { w.AddSlide(); w.SetNotes(0, "备注甲"); });
        var b = BuildFile(w => { w.AddSlide(); w.SetNotes(0, "备注乙"); });

        using var ms = Merge(a, b);
        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Equal(2, slides.Count);
        Assert.Contains("备注甲", slides[0].Notes);
        Assert.Contains("备注乙", slides[1].Notes);
    }

    [Fact, DisplayName("PPT—合并保留首文件幻灯片尺寸")]
    public void Merge_PreservesFirstFileSlideSize()
    {
        // 首文件 4:3（25.4cm × 19.05cm）
        var a = BuildFile(w =>
        {
            w.SetSlideSize(25.4, 19.05);
            w.AddSlide();
        });
        var b = BuildFile(w => { w.AddSlide(); }); // 默认 16:9

        using var ms = Merge(a, b);
        ms.Position = 0;
        using var reader = new PptxReader(ms);
        Assert.Equal(9144000, reader.SlideWidth);
        Assert.Equal(6858000, reader.SlideHeight);
    }

    [Fact, DisplayName("PPT—两文件图表+批注混合合并无冲突")]
    public void Merge_ChartAndComment_TwoFiles_NoCollision()
    {
        var a = BuildFile(w =>
        {
            w.AddSlide();
            var c = w.AddBarChart(0, ["A", "B"]);
            c.Series.Add(new ChartSeries { Name = "甲", Values = [1, 2] });
            w.Slides[0].Comments.Add(new Comment { Author = "甲", Text = "批注一" });
        });
        var b = BuildFile(w =>
        {
            w.AddSlide();
            var c = w.AddPieChart(0, ["C", "D"]);
            c.Series.Add(new ChartSeries { Name = "乙", Values = [3, 4] });
            w.Slides[0].Comments.Add(new Comment { Author = "乙", Text = "批注二" });
        });

        using var ms = Merge(a, b);
        var entries = ZipEntryNames(ms);
        Assert.Contains("ppt/charts/chart1.xml", entries);
        Assert.Contains("ppt/charts/chart2.xml", entries);
        Assert.Contains("ppt/comments/comment1.xml", entries);
        Assert.Contains("ppt/comments/comment2.xml", entries);
        Assert.Contains("ppt/comments/commentAuthors.xml", entries);

        // 读回完整性
        ms.Position = 0;
        using var reader = new PptxReader(ms);
        var slides = reader.ReadAllSlides().ToList();
        Assert.Equal(2, slides.Count);
        Assert.Single(slides[0].Charts);
        Assert.Single(slides[1].Charts);
        Assert.Single(slides[0].Comments);
        Assert.Single(slides[1].Comments);
        Assert.Contains("批注一", slides[0].Comments[0].Text);
        Assert.Contains("批注二", slides[1].Comments[0].Text);
    }
}
