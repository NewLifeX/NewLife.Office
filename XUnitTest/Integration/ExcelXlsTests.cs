using System.ComponentModel;
using NewLife.Office;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>Excel xls (BIFF) 格式集成测试</summary>
public class ExcelXlsTests : IntegrationTestBase
{
    [Fact, DisplayName("Excel_xls_BIFF写入再读取往返")]
    public void Excel_Xls_WriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_biff.xls");

        // BiffWriter 构造时自动创建默认 "Sheet1"，设置 SheetName 会新增工作表
        using (var w = new BiffWriter())
        {
            w.SheetName = "数据表";
            w.WriteHeader(new[] { "编号", "姓名", "分数", "日期", "通过" });
            w.WriteRow(new Object?[] { 1, "Alice", 95.5, new DateTime(2024, 1, 15), true });
            w.WriteRow(new Object?[] { 2, "Bob", 82.0, new DateTime(2024, 2, 20), true });
            w.WriteRow(new Object?[] { 3, "Charlie", 58.5, new DateTime(2024, 3, 10), false });
            w.WriteRow(new Object?[] { 4, "Diana", 91.0, new DateTime(2024, 4, 5), true });

            w.SheetName = "汇总";
            w.WriteHeader(new[] { "科目", "平均分" });
            w.WriteRow(new Object?[] { "数学", 88.5 });
            w.WriteRow(new Object?[] { "英语", 76.0 });

            w.Save(path);
        }

        Assert.True(File.Exists(path));

        // 读取验证：BiffWriter 构造时有默认 Sheet1，加上显式创建的 2 个共 3 个
        using var reader = new BiffReader(path);
        Assert.True(reader.SheetNames.Count >= 2);
        Assert.Contains("数据表", reader.SheetNames);
        Assert.Contains("汇总", reader.SheetNames);

        var rows = reader.ReadSheet("数据表").ToList();
        Assert.Equal(5, rows.Count); // header + 4 data
        Assert.Equal("Alice", rows[1][1] + "");

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<BiffReader>(factoryReader);
        (factoryReader as IDisposable)?.Dispose();
    }

    [Fact, DisplayName("Excel_xls转xlsx格式")]
    public void Excel_Xls_To_Xlsx()
    {
        var xlsPath = Path.Combine(OutputDir, "convert_source.xls");
        var xlsxPath = Path.Combine(OutputDir, "converted_from_xls.xlsx");

        using (var w = new BiffWriter())
        {
            w.WriteHeader(new[] { "Name", "Score" });
            w.WriteRow(new Object?[] { "Tom", 90 });
            w.WriteRow(new Object?[] { "Jerry", 85 });
            w.Save(xlsPath);
        }

        // 读取 xls 并转为 xlsx
        using var biff = new BiffReader(xlsPath);
        using var writer = new ExcelWriter(xlsxPath);
        foreach (var sheetName in biff.SheetNames)
        {
            var rows = biff.ReadSheet(sheetName).ToList();
            if (rows.Count > 0)
            {
                writer.WriteHeader(sheetName, rows[0].Select(e => e + "").ToArray());
                for (var i = 1; i < rows.Count; i++)
                {
                    writer.WriteRow(sheetName, rows[i]);
                }
            }
        }
        writer.Save();

        Assert.True(File.Exists(xlsxPath));

        // 验证 xlsx
        using var xlsxReader = new ExcelReader(xlsxPath);
        var xlsxRows = xlsxReader.ReadRows().ToList();
        Assert.Equal(3, xlsxRows.Count);
        Assert.Equal("Tom", xlsxRows[1][0] + "");
    }
}
