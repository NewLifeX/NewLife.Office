using System.ComponentModel;
using System.Globalization;
using System.Text;
using NewLife.Office.Csv;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelDataFrame 的 CSV 导入导出测试</summary>
public class ExcelDataFrameCsvTests
{
    #region 工具
    private static ExcelDataFrame FromText(String csv, CsvOptions? options = null)
        => ExcelDataFrame.FromCsv(new MemoryStream(new UTF8Encoding(false).GetBytes(csv)), options);
    #endregion

    [Fact, DisplayName("CSV 导入：首行为列名")]
    public void FromCsv_UsesHeaderRow()
    {
        var frame = FromText("地区,金额\r\n华东,100\r\n华南,200\r\n");

        Assert.Equal(["地区", "金额"], frame.Columns);
        Assert.Equal(2, frame.RowCount);
        Assert.Equal("华东", frame.GetValue(0, "地区"));
    }

    [Fact, DisplayName("CSV 导入：数字列推断后按数值排序")]
    public void FromCsv_InferTypes_SortsNumerically()
    {
        var frame = FromText("name,amount\r\nA,9\r\nB,10\r\nC,2\r\n");

        Assert.IsType<Int64>(frame.GetValue(0, "amount"));

        var sorted = frame.OrderBy("amount", descending: true);
        Assert.Equal("B", sorted.GetValue(0, "name"));
        Assert.Equal(10L, sorted.GetValue(0, "amount"));
    }

    [Fact, DisplayName("CSV 导入：日期列被识别为 DateTime")]
    public void FromCsv_InferTypes_DetectsDates()
    {
        var frame = FromText("date,amount\r\n2024-01-31,1\r\n2024-02-01,2\r\n");

        Assert.IsType<DateTime>(frame.GetValue(0, "date"));
        Assert.Equal(new DateTime(2024, 1, 31), frame.GetValue(0, "date"));
    }

    [Fact, DisplayName("CSV 导入：前导零与超长数字保持字符串")]
    public void FromCsv_InferTypes_KeepsLeadingZerosAndLongDigits()
    {
        var frame = FromText("code,order\r\n00123,12345678901234567890\r\n00456,12345678901234567891\r\n");

        Assert.IsType<String>(frame.GetValue(0, "code"));
        Assert.Equal("00123", frame.GetValue(0, "code"));
        Assert.IsType<String>(frame.GetValue(0, "order"));
    }

    [Fact, DisplayName("CSV 导入：关闭推断后全部保持字符串")]
    public void FromCsv_WithoutInferTypes_KeepsText()
    {
        var frame = FromText("amount\r\n9\r\n10\r\n", new CsvOptions { InferTypes = false });

        Assert.IsType<String>(frame.GetValue(0, "amount"));
        Assert.Equal("9", frame.GetValue(0, "amount"));
    }

    [Fact, DisplayName("CSV 导入：无表头时列为 Column1..N")]
    public void FromCsv_WithoutHeader_GeneratesColumnNames()
    {
        var frame = FromText("1,2\r\n3,4\r\n", new CsvOptions { HasHeader = false });

        Assert.Equal(["Column1", "Column2"], frame.Columns);
        Assert.Equal(2, frame.RowCount);
    }

    [Fact, DisplayName("CSV 导入：重复列名抛异常")]
    public void FromCsv_DuplicateColumns_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => FromText("a,a\r\n1,2\r\n"));

        Assert.Contains("a", ex.Message);
    }

    [Fact, DisplayName("CSV 导入：空内容抛异常")]
    public void FromCsv_EmptyContent_Throws()
    {
        Assert.Throws<ArgumentException>(() => FromText(""));
    }

    [Fact, DisplayName("CSV 导出：往返保持列名与数据")]
    public void ToCsv_RoundTrips()
    {
        var source = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            ["华东", 100L],
            ["华南", null],
        ]);

        using var ms = new MemoryStream();
        source.ToCsv(ms);

        ms.Position = 0;
        var back = ExcelDataFrame.FromCsv(ms);

        Assert.Equal(source.Columns, back.Columns);
        Assert.Equal(source.RowCount, back.RowCount);
        Assert.Equal(100L, back.GetValue(0, "金额"));
        Assert.Null(back.GetValue(1, "金额"));
    }

    [Fact, DisplayName("CSV 导出：数值不受当前区域影响")]
    public void ToCsv_UsesInvariantCulture()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["name", "value"],
            ["a", 1.5],
        ]);

        var origin = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            using var ms = new MemoryStream();
            frame.ToCsv(ms);

            var text = Encoding.UTF8.GetString(ms.ToArray());
            Assert.Contains("1.5", text);
            Assert.DoesNotContain("1,5", text);
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = origin;
        }
    }

    [Fact, DisplayName("行数据入口：FromRows/ToRows 往返一致")]
    public void FromRows_ToRows_RoundTrips()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            ["华东", 100L],
        ]);

        var rows = frame.ToRows().ToList();

        Assert.Equal(2, rows.Count);
        Assert.Equal(["地区", "金额"], rows[0]);
        Assert.Equal("华东", rows[1][0]);
        Assert.Equal(100L, rows[1][1]);

        var body = frame.ToRows(includeHeader: false).ToList();
        Assert.Single(body);
    }
}
