using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>ExcelDataFrame 行内类型化取值测试</summary>
public partial class ExcelDataFrameTests
{
    [Fact(DisplayName = "行取值：Number 解析数字与文本数字并保留空值")]
    public void Row_Number_ReadsNumericValues()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["地区", "金额"],
            ["华东", 100d],
            ["华东", "150"],
            ["华东", null],
            ["华东", DBNull.Value],
        ]);

        var result = frame.AddColumn("数值", row => row.Number("金额"));

        Assert.Equal(100d, result.GetValue(0, "数值"));
        Assert.Equal(150d, result.GetValue(1, "数值"));
        Assert.Null(result.GetValue(2, "数值"));
        Assert.Null(result.GetValue(3, "数值"));
        Assert.Equal(2, frame.Where(row => row.Number("金额") > 0).RowCount);
    }

    [Fact(DisplayName = "行取值：Number 遇到非数字或未知列明确失败")]
    public void Row_Number_ValidatesValuesAndColumns()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["金额"],
            ["abc"],
        ]);

        Assert.Throws<ArgumentException>(() => frame.Where(row => row.Number("金额") > 0).RowCount);
        Assert.Throws<ArgumentException>(() => frame.Where(row => row.Number("不存在") > 0).RowCount);
    }

    [Fact(DisplayName = "行取值：Text 转换各类值并保留空值")]
    public void Row_Text_ReadsTextValues()
    {
        var frame = ExcelDataFrame.FromRows(
        [
            ["名称", "数值"],
            ["华东", 100d],
            [null, null],
            [DBNull.Value, DBNull.Value],
        ]);

        var result = frame.AddColumn("文本", row => row.Text("数值"));

        Assert.Equal("100", result.GetValue(0, "文本"));
        Assert.Null(result.GetValue(1, "文本"));
        Assert.Null(result.GetValue(2, "文本"));
        Assert.Equal(1, frame.Where(row => row.Text("名称") == "华东").RowCount);
        Assert.Throws<ArgumentException>(() => frame.Where(row => row.Text("不存在") == null).RowCount);
    }
}
