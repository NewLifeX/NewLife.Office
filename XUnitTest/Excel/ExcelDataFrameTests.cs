using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelDataFrame 表格分析 API 测试</summary>
public class ExcelDataFrameTests
{
    private static Worksheet BuildSheet()
    {
        var sheet = new Worksheet { Name = "销售" };
        sheet.Rows.Add(["地区", "产品", "销售额"]);
        sheet.Rows.Add(["华东", "苹果", 100d]);
        sheet.Rows.Add(["华东", "香蕉", 50d]);
        sheet.Rows.Add(["华南", "苹果", 200d]);
        sheet.Rows.Add(["华南", "梨", null]);
        return sheet;
    }

    [Fact(DisplayName = "表格分析：选列筛选排序不修改源数据")]
    public void SelectWhereOrderBy_ProcessesRows()
    {
        var source = BuildSheet();
        var frame = ExcelDataFrame.FromWorksheet(source);

        var result = frame.Where(row => row["地区"]!.ToString() == "华东")
            .Select("产品", "销售额")
            .OrderBy("销售额", descending: true);

        Assert.Equal(["产品", "销售额"], result.Columns);
        Assert.Equal(2, result.RowCount);
        Assert.Equal("苹果", result.GetValue(0, "产品"));
        Assert.Equal(100d, result.GetValue(0, "销售额"));
        Assert.Equal("苹果", source.Rows[1][1]);
        Assert.Equal(3, frame.ColumnCount);

        var sorted = frame.OrderBy("销售额", descending: true);
        Assert.Null(sorted.GetValue(sorted.RowCount - 1, "销售额"));
    }

    [Fact(DisplayName = "表格分析：分组汇总空键与空数值")]
    public void GroupBy_AggregatesRowsAndSkipsMissingValues()
    {
        var sheet = BuildSheet();
        sheet.Rows.Add([null, "其他", 25d]);
        sheet.Rows.Add(["华北", "无数据", null]);
        var frame = ExcelDataFrame.FromWorksheet(sheet);

        var sum = frame.GroupBy("地区").Sum("销售额");
        var count = frame.GroupBy("地区").Count();

        Assert.Equal(4, sum.RowCount);
        Assert.Equal("华东", sum.GetValue(0, "地区"));
        Assert.Equal(150d, sum.GetValue(0, "Sum_销售额"));
        Assert.Equal(2, count.GetValue(1, "Count"));
        Assert.Null(sum.GetValue(2, "地区"));
        Assert.Equal(25d, sum.GetValue(2, "Sum_销售额"));
        Assert.Equal(0d, sum.GetValue(3, "Sum_销售额"));

        var average = frame.GroupBy("地区").Average("销售额");
        var minimum = frame.GroupBy("地区").Min("销售额");
        var maximum = frame.GroupBy("地区").Max("销售额");
        var valueCount = frame.GroupBy("地区").Count("销售额");
        Assert.Equal(75d, average.GetValue(0, "Average_销售额"));
        Assert.Equal(50d, minimum.GetValue(0, "Min_销售额"));
        Assert.Equal(200d, maximum.GetValue(1, "Max_销售额"));
        Assert.Equal(1, valueCount.GetValue(1, "Count_销售额"));
        Assert.Null(average.GetValue(3, "Average_销售额"));
    }

    [Fact(DisplayName = "表格分析：描述统计含四分位数和缺失数")]
    public void Describe_ReturnsStatistics()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        var stats = frame.Describe("销售额");

        Assert.Equal(9, stats.RowCount);
        Assert.Equal(3, stats.GetValue(0, "Value"));
        Assert.Equal(350d / 3, stats.GetValue(1, "Value"));
        Assert.Equal(75d, stats.GetValue(4, "Value"));
        Assert.Equal(100d, stats.GetValue(5, "Value"));
        Assert.Equal(1, stats.GetValue(8, "Value"));
    }

    [Fact(DisplayName = "表格分析：缺少列名和非数字统计值明确失败")]
    public void InvalidColumnsAndNumericValues_Throw()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentException>(() => frame.Select("不存在"));
        Assert.Throws<ArgumentException>(() => frame.Describe("产品"));
        Assert.Throws<ArgumentException>(() => frame.GroupBy("地区").Sum("产品"));
        Assert.Throws<ArgumentException>(() => ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "重复列",
            Rows = [new Object?[] { "值", "值" }],
        }));
    }

    [Fact(DisplayName = "表格分析：输出工作表可写回并重新读取")]
    public void ToWorksheet_RoundTrips()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet())
            .GroupBy("地区")
            .Sum("销售额");
        var document = new ExcelDocument();
        document.Worksheets.Add(frame.ToWorksheet("汇总"));

        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var reader = new ExcelReader(stream, System.Text.Encoding.UTF8);
        var saved = reader.ReadExcel().GetSheet("汇总");

        Assert.Equal("地区", saved.Rows[0][0]);
        Assert.Equal("Sum_销售额", saved.Rows[0][1]);
        Assert.Equal("华东", saved.Rows[1][0]);
        Assert.Equal(150d, Convert.ToDouble(saved.Rows[1][1]));
    }
}
