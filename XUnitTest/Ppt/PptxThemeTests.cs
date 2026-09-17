using System.IO.Compression;
using System.Text;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PptxWriter 主题字体与图表主题色测试</summary>
public class PptxThemeTests
{
    private static String ReadEntry(String path, String entryName)
    {
        using var za = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read);
        var entry = za.GetEntry(entryName);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    [Fact(DisplayName = "PPT—设置主题字体写入theme1.xml（含东亚字体与Hans脚本）")]
    public void SetThemeFonts_WritesThemeFontScheme()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.SetThemeFonts("Segoe UI", "微软雅黑");
            writer.AddSlide();
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "ppt/theme/theme1.xml");
            Assert.Contains("<a:latin typeface=\"Segoe UI\"/>", xml);
            Assert.Contains("<a:ea typeface=\"微软雅黑\"/>", xml);
            Assert.Contains("<a:font script=\"Hans\" typeface=\"微软雅黑\"/>", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—未设置主题字体时保持Calibri默认")]
    public void SetThemeFonts_NotSet_KeepsCalibri()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "ppt/theme/theme1.xml");
            Assert.Contains("<a:latin typeface=\"Calibri\"/>", xml);
            Assert.Contains("<a:ea typeface=\"\"/>", xml);
            Assert.DoesNotContain("Hans", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—图表系列未指定颜色时回退主题强调色")]
    public void Chart_SeriesFallback_UsesAccentColors()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.SetAccentColors("FF0000", "00FF00", "0000FF");
            writer.AddSlide();
            var chart = writer.AddBarChart(0, ["A", "B"]);
            chart.Series.Add(new ChartSeries { Name = "S1", Values = [1, 2] });
            writer.Save(tempFile);

            var xml = ReadEntry(tempFile, "ppt/charts/chart1.xml");
            Assert.Contains("FF0000", xml);
            Assert.DoesNotContain("4F81BD", xml);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
}
