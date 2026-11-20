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

    [Fact(DisplayName = "表格分析：相关矩阵成对忽略空值并处理常量列")]
    public void CorrelationMatrix_UsesPairwiseCompleteRows()
    {
        var sheet = new Worksheet { Name = "指标" };
        sheet.Rows.Add(["A", "B", "C", "常量", "单值"]);
        sheet.Rows.Add([1d, 2d, 4d, 5d, 10d]);
        sheet.Rows.Add([2d, null, 3d, 5d, null]);
        sheet.Rows.Add([3d, 6d, 2d, 5d, null]);
        sheet.Rows.Add([4d, 8d, 1d, 5d, null]);
        var frame = ExcelDataFrame.FromWorksheet(sheet);

        var result = frame.CorrelationMatrix("A", "B", "C", "常量", "单值");

        Assert.Equal(5, result.RowCount);
        Assert.Equal(["Column", "A", "B", "C", "常量", "单值"], result.Columns);
        Assert.Equal("A", result.GetValue(0, "Column"));
        Assert.Equal(1d, result.GetValue(0, "A"));
        Assert.Equal(1d, result.GetValue(0, "B"));
        Assert.Equal(-1d, result.GetValue(0, "C"));
        Assert.Null(result.GetValue(3, "常量"));
        Assert.Null(result.GetValue(4, "A"));
        Assert.Null(result.GetValue(4, "单值"));
    }

    [Fact(DisplayName = "表格分析：相关矩阵拒绝无效列并处理非数字数据")]
    public void CorrelationMatrix_ValidatesColumnsAndValues()
    {
        var sheet = new Worksheet { Name = "指标" };
        sheet.Rows.Add(["A", "B"]);
        sheet.Rows.Add([1d, 2d]);
        sheet.Rows.Add(["文本", 3d]);
        var frame = ExcelDataFrame.FromWorksheet(sheet);

        Assert.Throws<ArgumentException>(() => frame.CorrelationMatrix());
        Assert.Throws<ArgumentException>(() => frame.CorrelationMatrix("A", "不存在"));
        Assert.Throws<ArgumentException>(() => frame.CorrelationMatrix("A", "A"));
        Assert.Throws<ArgumentException>(() => frame.CorrelationMatrix("A", "B"));
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

    [Fact(DisplayName = "表格分析：表间关联支持多键、重复键和左右外连接")]
    public void Merge_CombinesRowsByCompositeKeys()
    {
        var leftSheet = new Worksheet { Name = "订单" };
        leftSheet.Rows.Add(["地区", "月份", "金额"]);
        leftSheet.Rows.Add(["华东", "一月", 10d]);
        leftSheet.Rows.Add(["华东", "一月", 20d]);
        leftSheet.Rows.Add(["华南", "一月", 30d]);
        leftSheet.Rows.Add([null, "一月", 40d]);

        var rightSheet = new Worksheet { Name = "区域" };
        rightSheet.Rows.Add(["地区", "月份", "经理"]);
        rightSheet.Rows.Add(["华东", "一月", "张三"]);
        rightSheet.Rows.Add(["华东", "一月", "李四"]);
        rightSheet.Rows.Add(["华北", "一月", "王五"]);
        rightSheet.Rows.Add([null, "一月", "空键不匹配"]);

        var left = ExcelDataFrame.FromWorksheet(leftSheet);
        var right = ExcelDataFrame.FromWorksheet(rightSheet);
        var inner = left.Merge(right, ["地区", "月份"], ["地区", "月份"]);
        var leftJoin = left.Merge(right, ["地区", "月份"], ["地区", "月份"], ExcelDataFrame.JoinKind.Left);
        var rightJoin = left.Merge(right, ["地区", "月份"], ["地区", "月份"], ExcelDataFrame.JoinKind.Right);
        var full = left.Merge(right, ["地区", "月份"], ["地区", "月份"], ExcelDataFrame.JoinKind.Full);

        Assert.Equal(4, inner.RowCount);
        Assert.Equal(["地区", "月份", "金额", "经理"], inner.Columns);
        Assert.Equal("张三", inner.GetValue(0, "经理"));
        Assert.Equal("李四", inner.GetValue(1, "经理"));
        Assert.Equal(6, leftJoin.RowCount);
        Assert.Null(leftJoin.GetValue(4, "经理"));
        Assert.Equal(6, rightJoin.RowCount);
        Assert.Equal("华北", rightJoin.GetValue(4, "地区"));
        Assert.Equal(8, full.RowCount);
        Assert.Equal("华北", full.GetValue(6, "地区"));
        Assert.Equal("空键不匹配", full.GetValue(7, "经理"));
    }

    [Fact(DisplayName = "表格分析：关联列名冲突时自动加后缀并拒绝冲突结果")]
    public void Merge_ResolvesAndRejectsDuplicateOutputColumns()
    {
        var left = ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "左",
            Rows = [new Object?[] { "Id", "名称" }, new Object?[] { 1, "左侧" }],
        });
        var right = ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "右",
            Rows = [new Object?[] { "Key", "名称" }, new Object?[] { 1, "右侧" }],
        });

        var result = left.Merge(right, "Id", "Key");
        Assert.Equal("名称_right", result.Columns[2]);
        Assert.Equal("右侧", result.GetValue(0, "名称_right"));

        var leftConflict = ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "左侧冲突",
            Rows = [new Object?[] { "Id", "名称", "名称_right" }, new Object?[] { 1, "左侧", "既有列" }],
        });
        Assert.Throws<ArgumentException>(() => leftConflict.Merge(right, "Id", "Key"));
    }

    [Fact(DisplayName = "表格分析：透视表按重复组合汇总并可反透视")]
    public void PivotAndMelt_TransformWideAndLongData()
    {
        var sheet = new Worksheet { Name = "月报" };
        sheet.Rows.Add(["地区", "月份", "销售额"]);
        sheet.Rows.Add(["华东", "一月", 10d]);
        sheet.Rows.Add(["华东", "一月", 5d]);
        sheet.Rows.Add(["华东", "二月", 20d]);
        sheet.Rows.Add(["华南", "一月", 30d]);
        sheet.Rows.Add(["华南", null, 99d]);

        var source = ExcelDataFrame.FromWorksheet(sheet);
        var pivot = source.Pivot("地区", "月份", "销售额");
        var melted = pivot.Melt(["地区"], ["一月", "二月"]);

        Assert.Equal(["地区", "一月", "二月"], pivot.Columns);
        Assert.Equal(2, pivot.RowCount);
        Assert.Equal(15d, pivot.GetValue(0, "一月"));
        Assert.Equal(20d, pivot.GetValue(0, "二月"));
        Assert.Equal(30d, pivot.GetValue(1, "一月"));
        Assert.Null(pivot.GetValue(1, "二月"));
        Assert.Equal(4, melted.RowCount);
        Assert.Equal("一月", melted.GetValue(0, "Variable"));
        Assert.Equal(15d, melted.GetValue(0, "Value"));
        Assert.Null(melted.GetValue(3, "Value"));
    }

    [Fact(DisplayName = "表格分析：交叉表统计全部行并按首次出现顺序保留空类别")]
    public void Crosstab_CountsRowsAndPreservesFirstSeenCategories()
    {
        var sheet = new Worksheet { Name = "分类" };
        sheet.Rows.Add(["地区", "状态", "数值"]);
        sheet.Rows.Add(["华东", "开启", null]);
        sheet.Rows.Add(["华南", "关闭", null]);
        sheet.Rows.Add(["华东", "开启", null]);
        sheet.Rows.Add([null, "关闭", null]);
        sheet.Rows.Add(["华东", null, null]);
        sheet.Rows.Add(["华北", "开启", null]);
        sheet.Rows.Add(["", "  ", null]);

        var result = ExcelDataFrame.FromWorksheet(sheet).Crosstab("地区", "状态");

        Assert.Equal(["地区", "开启", "关闭", "(空)"], result.Columns);
        Assert.Equal(4, result.RowCount);
        Assert.Equal("华东", result.GetValue(0, "地区"));
        Assert.Equal(2, result.GetValue(0, "开启"));
        Assert.Equal(0, result.GetValue(0, "关闭"));
        Assert.Equal(1, result.GetValue(0, "(空)"));
        Assert.Equal("华南", result.GetValue(1, "地区"));
        Assert.Equal(1, result.GetValue(1, "关闭"));
        Assert.Null(result.GetValue(2, "地区"));
        Assert.Equal(1, result.GetValue(2, "关闭"));
        Assert.Equal(1, result.GetValue(2, "(空)"));
        Assert.Equal("华北", result.GetValue(3, "地区"));
        Assert.Equal(1, result.GetValue(3, "开启"));

        var emptyValues = ExcelDataFrame.FromRows(
        [
            ["地区", "状态", "数值"],
            ["华东", "开启", null],
            ["华东", "开启", null],
        ]);
        var crosstab = emptyValues.Crosstab("地区", "状态");
        var pivotCount = emptyValues.Pivot("地区", "状态", "数值", ExcelDataFrame.AggregateFunction.Count);
        Assert.Equal(2, crosstab.GetValue(0, "开启"));
        Assert.Equal(0, pivotCount.GetValue(0, "开启"));
    }

    [Fact(DisplayName = "表格分析：交叉表拒绝缺列、相同键和冲突结果列")]
    public void Crosstab_ValidatesKeysAndResultColumnNames()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "状态"],
            ["华东", "地区"],
        ]);

        Assert.Throws<ArgumentException>(() => frame.Crosstab("不存在", "状态"));
        Assert.Throws<ArgumentException>(() => frame.Crosstab("地区", "不存在"));
        Assert.Throws<ArgumentException>(() => frame.Crosstab("地区", "地区"));
        Assert.Throws<ArgumentException>(() => frame.Crosstab("地区", "状态"));

        var duplicateHeaders = ExcelDataFrame.FromRows(
        [
            ["地区", "状态"],
            ["华东", "A"],
            ["华南", "a"],
        ]);
        Assert.Throws<ArgumentException>(() => duplicateHeaders.Crosstab("地区", "状态"));
    }

    [Fact(DisplayName = "表格分析：反透视保留空值并检查结果列冲突")]
    public void Melt_PreservesMissingValuesAndRejectsNameCollisions()
    {
        var frame = ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "数据",
            Rows = [new Object?[] { "Id", "一月", "二月" }, new Object?[] { 1, null, 2 }],
        });

        var result = frame.Melt(["Id"], ["一月", "二月"]);

        Assert.Equal(2, result.RowCount);
        Assert.Null(result.GetValue(0, "Value"));
        Assert.Equal(2, result.GetValue(1, "Value"));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id"], ["一月"], "Id", "Value"));
    }

    [Fact(DisplayName = "表格分析：累计和滚动统计按行处理空值")]
    public void CumulativeAndRollingStatistics_RespectMissingValues()
    {
        var frame = ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "序列",
            Rows =
            [
                new Object?[] { "值" },
                new Object?[] { 1d },
                new Object?[] { 2d },
                new Object?[] { null },
                new Object?[] { 4d },
            ],
        });

        var cumulative = frame.CumulativeSum("值");
        var rollingAverage = frame.RollingAverage("值", 3, minPeriods: 2);
        var rollingSum = frame.RollingSum("值", 3, minPeriods: 2);

        Assert.Equal(1d, cumulative.GetValue(0, "CumulativeSum_值"));
        Assert.Equal(3d, cumulative.GetValue(1, "CumulativeSum_值"));
        Assert.Null(cumulative.GetValue(2, "CumulativeSum_值"));
        Assert.Equal(7d, cumulative.GetValue(3, "CumulativeSum_值"));
        Assert.Null(rollingAverage.GetValue(0, "RollingAverage_值"));
        Assert.Equal(1.5d, rollingAverage.GetValue(2, "RollingAverage_值"));
        Assert.Equal(3d, rollingAverage.GetValue(3, "RollingAverage_值"));
        Assert.Equal(6d, rollingSum.GetValue(3, "RollingSum_值"));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.RollingAverage("值", 2, minPeriods: 3));
    }
}
