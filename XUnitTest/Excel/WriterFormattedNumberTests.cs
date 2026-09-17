using System.IO.Compression;
using System.Text;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel 写出器数值字符串自动识别测试（千分位与货币 → 真数值并保留显示格式）</summary>
public class WriterFormattedNumberTests
{
    private static String ReadEntry(Stream stream, String entryName)
    {
        stream.Position = 0;
        using var za = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var entry = za.GetEntry(entryName);
        using var sr = new StreamReader(entry!.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    [Fact(DisplayName = "ExcelWriter—千分位整数写入真数值并保留千分位格式")]
    public void Thousands_Integer_WrittenAsNumber()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.WriteHeader("Sheet1", ["渠道", "营收"]);
            writer.WriteRow("Sheet1", new Object[] { "官网直销", "182,400" });
            writer.Save();
        }

        ms.Position = 0;
        using (var reader = new ExcelReader(ms, Encoding.UTF8))
        {
            var rows = reader.ReadRows("Sheet1").ToList();
            Assert.Equal("官网直销", rows[1][0]);
            Assert.Equal(182400m, Convert.ToDecimal(rows[1][1]));
        }

        // 数字格式保留千分位
        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.Contains("formatCode=\"#,##0\"", xml);
    }

    [Fact(DisplayName = "ExcelWriter—千分位小数保留原始精度")]
    public void Thousands_Decimal_KeepsPrecision()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.WriteHeader("Sheet1", ["单价"]);
            writer.WriteRow("Sheet1", new Object[] { "1,234.50" });
            writer.Save();
        }

        ms.Position = 0;
        using (var reader = new ExcelReader(ms, Encoding.UTF8))
        {
            var rows = reader.ReadRows("Sheet1").ToList();
            Assert.Equal(1234.50m, Convert.ToDecimal(rows[1][0]));
        }

        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.Contains("formatCode=\"#,##0.00\"", xml);
    }

    [Fact(DisplayName = "ExcelWriter—货币符号数值化并保留货币格式")]
    public void Currency_WrittenAsNumber()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.WriteHeader("Sheet1", ["金额"]);
            writer.WriteRow("Sheet1", new Object[] { "¥1,234.5", "-￥2,000" });
            writer.Save();
        }

        ms.Position = 0;
        using (var reader = new ExcelReader(ms, Encoding.UTF8))
        {
            var rows = reader.ReadRows("Sheet1").ToList();
            Assert.Equal(1234.5m, Convert.ToDecimal(rows[1][0]));
            Assert.Equal(-2000m, Convert.ToDecimal(rows[1][1]));
        }

        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.Contains("formatCode=\"¥#,##0.0\"", xml);
        Assert.Contains("formatCode=\"￥#,##0\"", xml);
    }

    [Fact(DisplayName = "ExcelWriter—普通文本与编号保持字符串")]
    public void PlainText_StaysString()
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.WriteHeader("Sheet1", ["编号", "月份", "备注"]);
            writer.WriteRow("Sheet1", new Object[] { "1,2", "2026-07", "约 1,000 人工时" });
            writer.Save();
        }

        ms.Position = 0;
        using (var reader = new ExcelReader(ms, Encoding.UTF8))
        {
            var rows = reader.ReadRows("Sheet1").ToList();
            Assert.Equal("1,2", rows[1][0]);
            Assert.Equal("2026-07", rows[1][1]);
            Assert.Equal("约 1,000 人工时", rows[1][2]);
        }

        // 不应生成千分位数字格式
        var xml = ReadEntry(ms, "xl/styles.xml");
        Assert.DoesNotContain("#,##0", xml);
    }
}
