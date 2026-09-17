using System.IO.Compression;
using System.Text;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel 写出器默认字体测试（styles.xml font[0] 注入）</summary>
public class WriterDefaultFontTests
{
    private static String ReadEntry(Stream stream, String entryName)
    {
        stream.Position = 0;
        using var za = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var entry = za.GetEntry(entryName);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    [Fact(DisplayName = "ExcelWriter—设置默认字体后写入styles.xml font[0]")]
    public void DefaultFont_InjectedIntoStyles()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.DefaultFontName = "微软雅黑";
            writer.DefaultFontSize = 11;
            writer.WriteHeader("Sheet1", ["姓名", "分数"]);
            writer.WriteRow("Sheet1", new Object[] { "张三", 85 });
            writer.Save();
        }

        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.Contains("<name val=\"微软雅黑\"/>", xml);
        Assert.Contains("<sz val=\"11\"/>", xml);
    }

    [Fact(DisplayName = "ExcelWriter—未设置默认字体时输出空font元素（保持兼容）")]
    public void DefaultFont_NotSet_KeepsEmptyFont()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.WriteHeader("Sheet1", ["A"]);
            writer.WriteRow("Sheet1", new Object[] { 1 });
            writer.Save();
        }

        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.Contains("<font></font>", xml);
        Assert.DoesNotContain("微软雅黑", xml);
    }
}
