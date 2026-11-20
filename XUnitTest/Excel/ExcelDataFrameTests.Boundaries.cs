using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelDataFrame 公开 API 边界与异常路径测试</summary>
public partial class ExcelDataFrameTests
{
    #region 载入与导出
    [Fact(DisplayName = "载入：列名行越界、空工作表与自定义列名行")]
    public void FromWorksheet_ValidatesHeaderAndColumns()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelDataFrame.FromWorksheet(null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelDataFrame.FromWorksheet(BuildSheet(), 99));
        Assert.Throws<ArgumentException>(() => ExcelDataFrame.FromWorksheet(new Worksheet
        {
            Name = "空",
            Rows = [Array.Empty<Object?>()],
        }));

        var sheet = new Worksheet { Name = "自定义表头" };
        sheet.Rows.Add(["标题"]);
        sheet.Rows.Add(["地区", "金额"]);
        sheet.Rows.Add(["华东", 100d]);
        var frame = ExcelDataFrame.FromWorksheet(sheet, headerRow: 1);
        Assert.Equal(["地区", "金额"], frame.Columns);
        Assert.Equal(1, frame.RowCount);

        var blankHeaders = new Worksheet { Name = "空列名" };
        blankHeaders.Rows.Add([null, "", "金额"]);
        blankHeaders.Rows.Add([1d, 2d, 3d]);
        var blank = ExcelDataFrame.FromWorksheet(blankHeaders);
        Assert.Equal(["Column1", "Column2", "金额"], blank.Columns);
    }

    [Fact(DisplayName = "载入：空行集合、空表头与短行补齐")]
    public void FromRows_ValidatesEmptyInputAndPadsRows()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelDataFrame.FromRows(null!));
        Assert.Throws<ArgumentException>(() => ExcelDataFrame.FromRows([]));
        Assert.Throws<ArgumentException>(() => ExcelDataFrame.FromRows([null!, null!]));

        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            ["华东"],
        ]);
        Assert.Equal(1, frame.RowCount);
        Assert.Null(frame.GetValue(0, "金额"));

        var noHeader = ExcelDataFrame.FromRows(
        [
            [1d, 2d],
            [3d, 4d],
        ], hasHeader: false);
        Assert.Equal(["Column1", "Column2"], noHeader.Columns);
        Assert.Equal(2, noHeader.RowCount);
    }

    [Fact(DisplayName = "导出：ToRows 返回副本并支持不含表头")]
    public void ToRows_ReturnsDetachedCopies()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["金额"],
            [1d],
        ]);

        var rows = frame.ToRows(includeHeader: false).ToList();
        rows[0][0] = 99d;

        Assert.Equal(1d, frame.GetValue(0, "金额"));
        Assert.Single(frame.ToRows(includeHeader: false));
    }

    [Fact(DisplayName = "查询：行列索引越界与未知列名明确失败")]
    public void GetValueAndRow_ValidateIndexesAndNames()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentOutOfRangeException>(() => frame.GetValue(-1, "销售额"));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.GetValue(99, "销售额"));
        Assert.Throws<ArgumentNullException>(() => frame.Where(row => row[null!] == null).RowCount);
        Assert.Throws<ArgumentException>(() => frame.Where(row => row["不存在"] == null).RowCount);
    }

    [Fact(DisplayName = "选列：空参数、重复列与大小写不敏感")]
    public void Select_ValidatesColumns()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentNullException>(() => frame.Select(null!));
        Assert.Throws<ArgumentException>(() => frame.Select());
        Assert.Throws<ArgumentException>(() => frame.Select("销售额", "销售额"));
        Assert.Equal(["销售额", "产品"], frame.Select("销售额", "产品").Columns);
    }

    [Fact(DisplayName = "筛选：空条件明确失败")]
    public void Where_ValidatesPredicate()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentNullException>(() => frame.Where(null!));
    }

    [Fact(DisplayName = "描述统计：全空列与单值列的统计量")]
    public void Describe_HandlesMissingAndSingleValueColumns()
    {
        var empty = ExcelDataFrame.FromRows(
        [
            ["值"],
            [null],
            [DBNull.Value],
        ]).Describe("值");
        Assert.Equal(0, empty.GetValue(0, "Value"));
        Assert.Null(empty.GetValue(1, "Value"));
        Assert.Null(empty.GetValue(2, "Value"));
        Assert.Null(empty.GetValue(3, "Value"));
        Assert.Null(empty.GetValue(4, "Value"));
        Assert.Equal(2, empty.GetValue(8, "Value"));

        var single = ExcelDataFrame.FromRows(
        [
            ["值"],
            ["5"],
        ]).Describe("值");
        Assert.Equal(5d, single.GetValue(1, "Value"));
        Assert.Null(single.GetValue(2, "Value"));
    }

    [Fact(DisplayName = "导出工作表：空表名明确失败")]
    public void ToWorksheet_ValidatesName()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentNullException>(() => frame.ToWorksheet(null!));
    }
    #endregion

    #region 排序与清洗
    [Fact(DisplayName = "排序：支持日期、布尔、同类型可比较值和混合类型")]
    public void OrderBy_ComparesSupportedValueTypes()
    {
        var dates = ExcelDataFrame.FromRows(
        [
            ["值"],
            [new DateTime(2026, 1, 2)],
            [new DateTime(2026, 1, 1)],
        ]);
        Assert.Equal(new DateTime(2026, 1, 1), dates.OrderBy("值").GetValue(0, "值"));

        var flags = ExcelDataFrame.FromRows(
        [
            ["值"],
            [true],
            [false],
        ]);
        Assert.Equal(false, flags.OrderBy("值").GetValue(0, "值"));

        var first = Guid.Parse("11111111-1111-1111-1111-111111111111");
        var second = Guid.Parse("22222222-2222-2222-2222-222222222222");
        var guids = ExcelDataFrame.FromRows(
        [
            ["值"],
            [second],
            [first],
        ]);
        Assert.Equal(first, guids.OrderBy("值").GetValue(0, "值"));

        var mixed = ExcelDataFrame.FromRows(
        [
            ["值"],
            ["abc"],
            [10d],
        ]);
        Assert.Equal(10d, mixed.OrderBy("值").GetValue(0, "值"));
    }

    [Fact(DisplayName = "排序：空参数、重复列与等键稳定排序")]
    public void OrderBy_ValidatesArgumentsAndKeepsStableOrder()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["组", "值"],
            ["A", 1d],
            ["A", 1d],
            ["A", 2d],
        ]);

        Assert.Throws<ArgumentNullException>(() => frame.OrderBy((ValueTuple<String, Boolean>[])null!));
        Assert.Throws<ArgumentException>(() => frame.OrderBy());
        Assert.Throws<ArgumentException>(() => frame.OrderBy(("组", false), ("组", true)));
        Assert.Throws<ArgumentException>(() => frame.OrderBy(("不存在", false)));

        var sorted = frame.OrderBy(("组", false), ("值", false));
        Assert.Equal(1d, sorted.GetValue(0, "值"));
        Assert.Equal(2d, sorted.GetValue(2, "值"));
    }

    [Fact(DisplayName = "派生列：空列名、空选择器与重名列")]
    public void AddColumn_ValidatesArguments()
    {
        var frame = ExcelDataFrame.FromWorksheet(BuildSheet());

        Assert.Throws<ArgumentException>(() => frame.AddColumn("", row => null));
        Assert.Throws<ArgumentException>(() => frame.AddColumn(" ", row => null));
        Assert.Throws<ArgumentNullException>(() => frame.AddColumn("新列", null!));
        Assert.Throws<ArgumentException>(() => frame.AddColumn("销售额", row => null));
    }

    [Fact(DisplayName = "缺失值：未指定列、重复列与空列集合")]
    public void DropMissingAndFillMissing_ValidateColumns()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            ["华东", 1d],
            [null, 2d],
            ["华南", null],
        ]);

        Assert.Throws<ArgumentNullException>(() => frame.DropMissing(null!));
        Assert.Throws<ArgumentException>(() => frame.DropMissing("金额", "金额"));
        Assert.Throws<ArgumentException>(() => frame.DropMissing("不存在"));

        Assert.Equal(1, frame.DropMissing().RowCount);
        Assert.Equal(2, frame.DropMissing("金额").RowCount);

        var filled = frame.FillMissing(0);
        Assert.Equal(0, filled.GetValue(1, "地区"));
        Assert.Equal(0, filled.GetValue(2, "金额"));

        var unchanged = frame.FillMissing(null);
        Assert.Null(unchanged.GetValue(2, "金额"));
        Assert.Throws<ArgumentException>(() => frame.FillMissing(0, "金额", "金额"));
    }

    [Fact(DisplayName = "去重：未指定列、数值相等与空值归一")]
    public void DropDuplicates_ComparesAllColumns()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["键", "值"],
            [1, null],
            [1.0d, DBNull.Value],
            ["2", 3d],
        ]);

        Assert.Equal(2, frame.DropDuplicates().RowCount);
        Assert.Equal(2, frame.DropDuplicates("键").RowCount);
        Assert.Throws<ArgumentException>(() => frame.DropDuplicates("键", "键"));
    }
    #endregion

    #region 关联
    [Fact(DisplayName = "关联：空参数、键长度、连接类型与后缀校验")]
    public void Merge_ValidatesArguments()
    {
        var left = ExcelDataFrame.FromRows(
        [
            ["Id", "名称"],
            [1, "左"],
        ]);
        var right = ExcelDataFrame.FromRows(
        [
            ["Key", "名称"],
            [1, "右"],
        ]);

        Assert.Throws<ArgumentNullException>(() => left.Merge(null!, "Id", "Key"));
        Assert.Throws<ArgumentNullException>(() => left.Merge(right, (String[])null!, ["Key"]));
        Assert.Throws<ArgumentNullException>(() => left.Merge(right, ["Id"], (String[])null!));
        Assert.Throws<ArgumentException>(() => left.Merge(right, ["Id", "名称"], ["Key"]));
        Assert.Throws<ArgumentOutOfRangeException>(() => left.Merge(right, "Id", "Key", (JoinKind)99));
        Assert.Throws<ArgumentException>(() => left.Merge(right, "Id", "Key", JoinKind.Inner, ""));
        Assert.Throws<ArgumentException>(() => left.Merge(right, ["Id", "Id"], ["Key", "Key"]));
    }

    [Fact(DisplayName = "关联：右侧键列不重复进入结果")]
    public void Merge_ExcludesRightKeyColumn()
    {
        var left = ExcelDataFrame.FromRows(
        [
            ["Id", "名称"],
            [1, "左"],
        ]);
        var right = ExcelDataFrame.FromRows(
        [
            ["Key", "名称"],
            [1, "右"],
        ]);

        var result = left.Merge(right, "Id", "Key");

        Assert.Equal(["Id", "名称", "名称_right"], result.Columns);
        Assert.Equal("右", result.GetValue(0, "名称_right"));
    }
    #endregion

    #region 透视与反透视
    [Fact(DisplayName = "透视：非法函数、空白列类别与结果列冲突")]
    public void Pivot_ValidatesFunctionAndCategories()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "状态", "值"],
            ["华东", "开启", 1d],
        ]);

        Assert.Throws<ArgumentNullException>(() => frame.Pivot((String[])null!, "状态", "值"));
        Assert.Throws<ArgumentException>(() => frame.Pivot([], "状态", "值"));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.Pivot("地区", "状态", "值", (AggregateFunction)99));

        var blank = ExcelDataFrame.FromRows(
        [
            ["地区", "状态", "值"],
            ["华东", "  ", 1d],
        ]);
        Assert.Throws<ArgumentException>(() => blank.Pivot("地区", "状态", "值"));

        var conflict = ExcelDataFrame.FromRows(
        [
            ["地区", "状态", "值"],
            ["华东", "地区", 1d],
        ]);
        Assert.Throws<ArgumentException>(() => conflict.Pivot("地区", "状态", "值"));
    }

    [Fact(DisplayName = "透视：支持全部聚合函数并跳过空值")]
    public void Pivot_SupportsAllAggregateFunctions()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "状态", "值"],
            ["华东", "开启", 10d],
            ["华东", "开启", 5d],
            ["华东", "关闭", null],
        ]);

        Assert.Equal(2, frame.Pivot("地区", "状态", "值", AggregateFunction.Count).GetValue(0, "开启"));
        Assert.Equal(15d, frame.Pivot("地区", "状态", "值", AggregateFunction.Sum).GetValue(0, "开启"));
        Assert.Equal(7.5d, frame.Pivot("地区", "状态", "值", AggregateFunction.Average).GetValue(0, "开启"));
        Assert.Equal(5d, frame.Pivot("地区", "状态", "值", AggregateFunction.Min).GetValue(0, "开启"));
        Assert.Equal(10d, frame.Pivot("地区", "状态", "值", AggregateFunction.Max).GetValue(0, "开启"));
    }

    [Fact(DisplayName = "反透视：空参数、重复列与结果列名冲突")]
    public void Melt_ValidatesArguments()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["Id", "一月"],
            [1, 2d],
        ]);

        Assert.Throws<ArgumentNullException>(() => frame.Melt(null!, ["一月"]));
        Assert.Throws<ArgumentNullException>(() => frame.Melt(["Id"], null!));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id"], []));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id"], ["一月"], variableColumn: " "));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id"], ["一月"], valueColumn: " "));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id"], ["一月"], valueColumn: "Id"));
        Assert.Throws<ArgumentException>(() => frame.Melt(["Id", "Id"], ["一月"]));
    }
    #endregion

    #region 相关与交叉表
    [Fact(DisplayName = "相关矩阵：标签列重名时自动编号")]
    public void CorrelationMatrix_RenamesConflictingLabelColumn()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["Column", "B"],
            [1d, 2d],
            [2d, 4d],
            [3d, 6d],
        ]);

        var result = frame.CorrelationMatrix("Column", "B");

        Assert.Equal("Column2", result.Columns[0]);
        Assert.Equal("Column", result.GetValue(0, "Column2"));
        Assert.Equal(1d, result.GetValue(0, "Column"));
        Assert.Throws<ArgumentNullException>(() => frame.CorrelationMatrix(null!));
        Assert.Throws<ArgumentException>(() => frame.CorrelationMatrix());
    }

    [Fact(DisplayName = "交叉表：非字符串空白类别不能作为列标题")]
    public void Crosstab_RejectsWhitespaceCategory()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "状态"],
            ["华东", ' '],
        ]);

        Assert.Throws<ArgumentException>(() => frame.Crosstab("地区", "状态"));
    }
    #endregion

    #region 窗口与重采样
    [Fact(DisplayName = "滚动统计：窗口与最少期数校验并支持自定义结果列")]
    public void Rolling_ValidatesWindowAndSupportsCustomName()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["值"],
            [1d],
            [2d],
        ]);

        Assert.Throws<ArgumentOutOfRangeException>(() => frame.RollingAverage("值", 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.RollingAverage("值", 2, minPeriods: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.RollingSum("值", 2, minPeriods: 3));

        var custom = frame.RollingSum("值", 2, resultColumn: "合计");
        Assert.Equal(3d, custom.GetValue(1, "合计"));
        Assert.Equal(0, ExcelDataFrame.FromRows([["值"]]).RollingAverage("值", 2).RowCount);
    }

    [Fact(DisplayName = "重采样：聚合定义与结果列名校验")]
    public void Resample_ValidatesAggregations()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["时间", "值"],
            [new DateTime(2026, 1, 1), 1d],
        ]);

        Assert.Throws<ArgumentNullException>(() => frame.Resample("时间", TimeBucket.Day, (AggregateSpec[])null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.Resample("时间", TimeBucket.Day, new AggregateSpec("值", (AggregateFunction)99)));
        Assert.Throws<ArgumentException>(() => frame.Resample("时间", TimeBucket.Day, new AggregateSpec("值", AggregateFunction.Sum, " ")));
        Assert.Throws<ArgumentException>(() => frame.Resample("时间", TimeBucket.Day, [null!]));
        Assert.Throws<ArgumentException>(() => frame.Resample("时间", TimeBucket.Day, AggregateSpec.Sum("值"), AggregateSpec.Sum("值")));
        Assert.Throws<ArgumentException>(() => frame.Resample("时间", TimeBucket.Day, AggregateSpec.Sum("值", "时间")));

        var tuple = frame.Resample("时间", TimeBucket.Day, ("值", AggregateFunction.Sum, "合计"));
        Assert.Equal(1d, tuple.GetValue(0, "合计"));
    }
    #endregion
}
