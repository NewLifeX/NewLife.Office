using NewLife.Office.Excel;
using Xunit;
using static NewLife.Office.Excel.AggregateSpec;

namespace XUnitTest.Excel;

/// <summary>ExcelDataFrame 聚合定义与简写 API 测试</summary>
public partial class ExcelDataFrameTests
{
    private static ExcelDataFrame BuildAggregateSheet() => ExcelDataFrame.FromRows(
    [
        ["地区", "月份", "金额", "数量"],
        ["华东", "一月", 100d, 2],
        ["华东", "一月", 50d, 1],
        ["华东", "二月", 80d, null],
        ["华南", "一月", 20d, 3],
    ]);

    [Fact(DisplayName = "聚合定义：工厂方法创建各函数并保留结果列名")]
    public void AggregateSpec_Factories_CreateDefinitions()
    {
        AssertSpec(AggregateSpec.Sum("金额"), AggregateFunction.Sum, "金额", null);
        AssertSpec(AggregateSpec.Average("金额", "均值"), AggregateFunction.Average, "金额", "均值");
        AssertSpec(AggregateSpec.Min("金额"), AggregateFunction.Min, "金额", null);
        AssertSpec(AggregateSpec.Max("金额"), AggregateFunction.Max, "金额", null);
        AssertSpec(AggregateSpec.Count("数量", "非空"), AggregateFunction.Count, "数量", "非空");
        AssertSpec(AggregateSpec.CountRows("行数"), AggregateFunction.Count, null, "行数");
        AssertSpec(AggregateSpec.CountRows(), AggregateFunction.Count, null, "Count");
    }

    [Fact(DisplayName = "聚合定义：缺少列或空结果列名时明确失败")]
    public void AggregateSpec_ValidatesColumnAndResultName()
    {
        Assert.Throws<ArgumentException>(() => new AggregateSpec(null, AggregateFunction.Sum));
        Assert.Throws<ArgumentException>(() => new AggregateSpec("金额", AggregateFunction.Sum, ""));
        Assert.Throws<ArgumentException>(() => AggregateSpec.Sum(""));
        Assert.Throws<ArgumentException>(() => AggregateSpec.Average(" "));
        Assert.Throws<ArgumentException>(() => AggregateSpec.Min(null!));
        Assert.Throws<ArgumentException>(() => AggregateSpec.Max("数量", ""));

        var rows = new AggregateSpec(null, AggregateFunction.Count);
        Assert.Null(rows.ColumnName);
        Assert.Equal(AggregateFunction.Count, rows.Function);
    }

    [Fact(DisplayName = "分组汇总：工厂与元组简写生成汇总表")]
    public void GroupBy_Aggregate_AcceptsShorthand()
    {
        var result = BuildAggregateSheet()
            .GroupBy("地区", "月份")
            .Aggregate(
                Sum("金额", "金额合计"),
                ("数量", AggregateFunction.Average, "平均数量"),
                ("金额", AggregateFunction.Count));

        Assert.Equal(["地区", "月份", "金额合计", "平均数量", "Count_金额"], result.Columns);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(150d, result.GetValue(0, "金额合计"));
        Assert.Equal(1.5d, result.GetValue(0, "平均数量"));
        Assert.Equal(2, result.GetValue(0, "Count_金额"));
        Assert.Equal(80d, result.GetValue(1, "金额合计"));
        Assert.Null(result.GetValue(1, "平均数量"));
        Assert.Equal(20d, result.GetValue(2, "金额合计"));
    }

    [Fact(DisplayName = "分组汇总：单指标可指定结果列名")]
    public void GroupBy_SingleMetric_SupportsResultName()
    {
        var frame = BuildAggregateSheet().Where(row => (String)row["地区"]! == "华东");

        Assert.Equal(230d, frame.GroupBy("地区").Sum("金额", "合计").GetValue(0, "合计"));
        Assert.Equal(230d / 3, frame.GroupBy("地区").Average("金额", "均值").GetValue(0, "均值"));
        Assert.Equal(50d, frame.GroupBy("地区").Min("金额", "最小").GetValue(0, "最小"));
        Assert.Equal(100d, frame.GroupBy("地区").Max("金额", "最大").GetValue(0, "最大"));
        Assert.Equal(2, frame.GroupBy("地区").Count("数量", "非空").GetValue(0, "非空"));
        Assert.Equal(3, frame.GroupBy("地区").Count().GetValue(0, "Count"));
    }

    [Fact(DisplayName = "分组汇总：默认结果列名与重复结果列校验")]
    public void GroupBy_Aggregate_UsesDefaultNamesAndRejectsDuplicates()
    {
        var frame = BuildAggregateSheet();
        var grouping = frame.GroupBy("地区");

        var defaults = grouping.Aggregate(AggregateSpec.Sum("金额"), AggregateSpec.Count("数量"));
        Assert.Equal(["地区", "Sum_金额", "Count_数量"], defaults.Columns);

        Assert.Throws<ArgumentException>(() => grouping.Aggregate());
        Assert.Throws<ArgumentNullException>(() => grouping.Aggregate(null!));
        Assert.Throws<ArgumentException>(() => grouping.Aggregate([null!]));
        Assert.Throws<ArgumentException>(() => grouping.Aggregate(AggregateSpec.Sum("金额"), AggregateSpec.Average("金额", "sum_金额")));
        Assert.Throws<ArgumentException>(() => grouping.Aggregate(AggregateSpec.Sum("金额", "地区")));
        Assert.Throws<ArgumentException>(() => grouping.Aggregate(new AggregateSpec("金额", AggregateFunction.Sum, " ")));
        Assert.Throws<ArgumentException>(() => frame.GroupBy("地区", "地区"));
    }

    [Fact(DisplayName = "分组汇总：字符串数字参与求和并按数值相等分组")]
    public void GroupBy_ConvertsTextNumbersAndMatchesNumericKeys()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["键", "金额"],
            [1, "100"],
            [1.0d, "50"],
        ]);

        var result = frame.GroupBy("键").Sum("金额");

        Assert.Equal(1, result.RowCount);
        Assert.Equal(150d, result.GetValue(0, "Sum_金额"));
    }

    [Fact(DisplayName = "分组汇总：内容相同的字符串归为同一组")]
    public void GroupBy_ComparesTextByValue()
    {
        var first = new String(['华', '东']);
        var second = new String(['华', '东']);
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            [first, 1d],
            [second, 2d],
        ]);

        var result = frame.GroupBy("地区").Sum("金额");

        Assert.Equal(1, result.RowCount);
        Assert.Equal(3d, result.GetValue(0, "Sum_金额"));
    }

    [Fact(DisplayName = "分组汇总：最小最大值返回原始极值")]
    public void GroupBy_MinMax_ReturnsExtremes()
    {
        var result = BuildAggregateSheet()
            .GroupBy("地区")
            .Aggregate(AggregateSpec.Min("金额"), AggregateSpec.Max("金额"));

        Assert.Equal(50d, result.GetValue(0, "Min_金额"));
        Assert.Equal(100d, result.GetValue(0, "Max_金额"));
    }

    private static void AssertSpec(AggregateSpec spec, AggregateFunction function, String? columnName, String? resultName)
    {
        Assert.Equal(function, spec.Function);
        Assert.Equal(columnName, spec.ColumnName);
        Assert.Equal(resultName, spec.ResultName);
    }
}
