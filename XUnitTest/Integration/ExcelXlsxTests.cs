using System.ComponentModel;
using System.Data;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>Excel xlsx 格式集成测试</summary>
public class ExcelXlsxTests : IntegrationTestBase
{
    [Fact, DisplayName("Excel_xlsx_复杂写入再读取往返")]
    public void Excel_Xlsx_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.xlsx");

        using (var w = new ExcelWriter(path))
        {
            var headerStyle = new CellStyle
            {
                Bold = true,
                FontSize = 12,
                BackgroundColor = "1F4E79",
                FontColor = "FFFFFF",
                HAlign = HorizontalAlignment.Center,
                Border = CellBorderStyle.Thin,
            };

            // Sheet1: 员工数据
            w.WriteHeader("员工表", new[] { "编号", "姓名", "年龄", "入职日期", "薪资", "在职" }, headerStyle);
            var dataStyle = new CellStyle { Border = CellBorderStyle.Thin };
            w.WriteRow("员工表", new Object?[] { 1, "张三", 28, new DateTime(2020, 1, 15), 8500.50m, true }, dataStyle);
            w.WriteRow("员工表", new Object?[] { 2, "李四", 35, new DateTime(2018, 6, 1), 12000m, true }, dataStyle);
            w.WriteRow("员工表", new Object?[] { 3, "王五", 42, new DateTime(2015, 3, 20), 15000.75m, false }, dataStyle);
            w.WriteRow("员工表", new Object?[] { 4, "赵六", 25, new DateTime(2023, 9, 1), 7000m, true }, dataStyle);
            w.WriteRow("员工表", new Object?[] { 5, "孙七", 30, new DateTime(2021, 4, 10), 9200.25m, true }, dataStyle);

            w.FreezePane("员工表", 1);
            w.SetAutoFilter("员工表", "A1:F1");
            w.SetColumnWidth("员工表", 0, 8);
            w.SetColumnWidth("员工表", 1, 12);
            w.SetColumnWidth("员工表", 3, 16);
            w.SetColumnWidth("员工表", 4, 12);
            w.MergeCell("员工表", "A7:F7");
            w.WriteRow("员工表", new Object?[] { });
            w.WriteRow("员工表", new Object?[] { "汇总行（合并单元格）" });
            w.AddHyperlink("员工表", 2, 1, "https://example.com/zhangsan", "张三主页");
            w.AddDropdownValidation("员工表", "F2:F100", new[] { "TRUE", "FALSE" });
            w.AddConditionalFormat("员工表", "E2:E6", ConditionalFormatType.GreaterThan, "10000", "92D050");
            w.SetPageSetup("员工表", PageOrientation.Landscape, PaperSize.A4);

            // Sheet2: 产品
            w.WriteHeader("产品表", new[] { "产品", "价格", "库存", "状态" }, headerStyle);
            w.WriteRow("产品表", new Object?[] { "笔记本电脑", 5999.99m, 100, "在售" }, dataStyle);
            w.WriteRow("产品表", new Object?[] { "智能手机", 3999m, 500, "在售" }, dataStyle);
            w.WriteRow("产品表", new Object?[] { "平板电脑", 2999.50m, 200, "停产" }, dataStyle);
            w.WriteRow("产品表", new Object?[] { "耳机", 299.90m, 1000, "在售" }, dataStyle);
            w.FreezePane("产品表", 1);

            // Sheet3: DataTable
            var dt = new DataTable();
            dt.Columns.Add("区域", typeof(String));
            dt.Columns.Add("销售额", typeof(Decimal));
            dt.Columns.Add("季度", typeof(String));
            dt.Rows.Add("华东", 150000m, "Q1");
            dt.Rows.Add("华西", 120000m, "Q2");
            dt.Rows.Add("华北", 90000m, "Q3");
            dt.Rows.Add("华南", 180000m, "Q4");
            w.WriteDataTable("销售统计", dt, CellStyle.Header);
            w.AddConditionalFormat("销售统计", "B2:B5", ConditionalFormatType.DataBar, null, "4472C4");

            w.Save();
        }

        Assert.True(File.Exists(path));

        // 读取验证
        using var reader = new ExcelReader(path);
        var sheets = reader.Sheets?.ToList();
        Assert.NotNull(sheets);
        Assert.Equal(3, sheets!.Count);
        Assert.Contains("员工表", sheets);
        Assert.Contains("产品表", sheets);
        Assert.Contains("销售统计", sheets);

        // 验证员工表数据
        var rows = reader.ReadRows("员工表").ToList();
        Assert.True(rows.Count >= 6); // header + 5 data + blank + summary
        var header = rows[0].Select(e => e + "").ToArray();
        Assert.Equal("编号", header[0]);
        Assert.Equal("姓名", header[1]);

        // 验证产品表数据
        var productRows = reader.ReadRows("产品表").ToList();
        Assert.Equal(5, productRows.Count); // header + 4 data
        Assert.Equal("笔记本电脑", productRows[1][0] + "");

        // 通过工厂创建读取器
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<ExcelReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }
}
