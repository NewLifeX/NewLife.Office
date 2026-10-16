using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>数据透视表字段回读测试（E17）</summary>
/// <remarks>
/// PivotBuilder 写入透视表 → ExcelReader.ReadPivotTables 回读字段定义（行列/数据/筛选标记、
/// 汇总方式、数据字段显示名、字段名从透视缓存还原），覆盖透视表模型的读写闭环。
/// </remarks>
public class PivotTableReadTests
{
    /// <summary>构造含行/列/筛选/数据四类字段的透视表 xlsx 字节</summary>
    private static Byte[] BuildPivotXlsx()
    {
        using var ms = new MemoryStream();
        var builder = new PivotBuilder();
        builder.SetSourceData(
            new[] { "Region", "Product", "Quarter", "Sales" },
            new List<Object?[]>
            {
                new Object?[] { "East", "A", "Q1", 100 },
                new Object?[] { "West", "B", "Q2", 200 },
                new Object?[] { "East", "C", "Q1", 150 },
            });
        builder.AddRowField("Product");
        builder.AddColumnField("Region");
        builder.AddFilterField("Quarter");
        builder.AddDataField("Sales", PivotSummaryFunction.Sum);
        builder.Save(ms);
        return ms.ToArray();
    }

    [Fact, DisplayName("E17 透视表：字段定义回读（行列/数据/筛选 + 汇总方式）")]
    public void PivotTable_ReadDefinition()
    {
        var bytes = BuildPivotXlsx();

        using var reader = new ExcelReader(new MemoryStream(bytes), Encoding.UTF8);
        var tables = reader.ReadPivotTables("Pivot");

        Assert.Single(tables);
        var pt = tables[0];
        Assert.Equal("PivotTable1", pt.Name);
        Assert.False(String.IsNullOrEmpty(pt.Location), "应回读透视表位置");
        Assert.True(pt.CacheId > 0, "应回读透视缓存 ID");
        Assert.Equal(4, pt.Fields.Count);

        // 字段名从透视缓存还原，每列一个字段并保留角色标记
        var product = pt.Fields.First(f => f.Name == "Product");
        Assert.True(product.IsRowField);
        Assert.False(product.IsColumnField || product.IsDataField || product.IsFilterField);

        var region = pt.Fields.First(f => f.Name == "Region");
        Assert.True(region.IsColumnField);

        var quarter = pt.Fields.First(f => f.Name == "Quarter");
        Assert.True(quarter.IsFilterField);

        var sales = pt.Fields.First(f => f.Name == "Sales");
        Assert.True(sales.IsDataField);
        Assert.Equal(PivotSummaryFunction.Sum, sales.SummaryFunction);
        Assert.Equal("Sum of Sales", sales.Caption);
    }

    [Fact, DisplayName("E17 透视表：汇总方式回读（Average 数据字段）")]
    public void PivotTable_ReadSummaryFunction()
    {
        using var ms = new MemoryStream();
        var builder = new PivotBuilder();
        builder.SetSourceData(
            new[] { "Product", "Sales" },
            new List<Object?[]>
            {
                new Object?[] { "A", 100 },
                new Object?[] { "B", 200 },
            });
        builder.AddRowField("Product");
        builder.AddDataField("Sales", PivotSummaryFunction.Average, "平均销售额");
        builder.Save(ms);
        ms.Position = 0;

        using var reader = new ExcelReader(ms, Encoding.UTF8);
        var tables = reader.ReadPivotTables("Pivot");

        Assert.Single(tables);
        var sales = tables[0].Fields.First(f => f.Name == "Sales");
        Assert.True(sales.IsDataField);
        Assert.Equal(PivotSummaryFunction.Average, sales.SummaryFunction);
        Assert.Equal("平均销售额", sales.Caption);
    }
}
