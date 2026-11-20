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

    [Fact(DisplayName = "表格分析：多列分组并一次生成多个指标")]
    public void GroupBy_AggregatesMultipleKeysAndMeasures()
    {
        var sheet = new Worksheet { Name = "明细" };
        sheet.Rows.Add(["地区", "月份", "产品", "金额", "数量"]);
        sheet.Rows.Add(["华东", "一月", "苹果", 100d, 2]);
        sheet.Rows.Add(["华东", "一月", "梨", 50d, 1]);
        sheet.Rows.Add(["华东", "二月", "苹果", 80d, 2]);

        var result = ExcelDataFrame.FromWorksheet(sheet)
            .GroupBy("地区", "月份")
            .Aggregate(
                new ExcelDataFrame.AggregateSpec("金额", ExcelDataFrame.AggregateFunction.Sum, "总额"),
                new ExcelDataFrame.AggregateSpec("数量", ExcelDataFrame.AggregateFunction.Average, "平均数量"),
                ExcelDataFrame.AggregateSpec.CountRows("行数"));

        Assert.Equal(2, result.RowCount);
        Assert.Equal(["地区", "月份", "总额", "平均数量", "行数"], result.Columns);
        Assert.Equal("一月", result.GetValue(0, "月份"));
        Assert.Equal(150d, result.GetValue(0, "总额"));
        Assert.Equal(1.5d, result.GetValue(0, "平均数量"));
        Assert.Equal(2, result.GetValue(0, "行数"));
    }

    [Fact(DisplayName = "表格分析：派生列、空值填补和去重")]
    public void AddColumn_FillMissingAndDropDuplicates()
    {
        var sheet = new Worksheet { Name = "数据" };
        sheet.Rows.Add(["名称", "数量", "单价"]);
        sheet.Rows.Add(["A", 2d, 5d]);
        sheet.Rows.Add(["A", 2d, 5d]);
        sheet.Rows.Add(["B", null, 3d]);

        var result = ExcelDataFrame.FromWorksheet(sheet)
            .DropDuplicates("名称")
            .FillMissing(0, "数量")
            .AddColumn("金额", row => Convert.ToDouble(row["数量"]) * Convert.ToDouble(row["单价"]));

        Assert.Equal(2, result.RowCount);
        Assert.Equal(10d, result.GetValue(0, "金额"));
        Assert.Equal(0, result.GetValue(1, "数量"));
        Assert.Equal(0d, result.GetValue(1, "金额"));
        Assert.Equal(3, ExcelDataFrame.FromWorksheet(sheet).RowCount);
    }

    [Fact(DisplayName = "表格分析：清理缺失行并多列排序取前几项")]
    public void DropMissingAndTop_SortsByMultipleColumns()
    {
        var sheet = new Worksheet { Name = "排名" };
        sheet.Rows.Add(["地区", "销售额", "利润"]);
        sheet.Rows.Add(["华东", 100d, 10d]);
        sheet.Rows.Add(["华南", 100d, 20d]);
        sheet.Rows.Add(["华北", 90d, 30d]);
        sheet.Rows.Add(["未知", null, 99d]);

        var result = ExcelDataFrame.FromWorksheet(sheet)
            .DropMissing("销售额")
            .Top(2, ("销售额", true), ("利润", true));

        Assert.Equal(2, result.RowCount);
        Assert.Equal("华南", result.GetValue(0, "地区"));
        Assert.Equal("华东", result.GetValue(1, "地区"));
        Assert.Equal(3, ExcelDataFrame.FromWorksheet(sheet).DropMissing("销售额").RowCount);
    }
}
