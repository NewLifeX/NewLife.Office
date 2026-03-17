using System.ComponentModel;
using NewLife.Office;
using NewLife.Office.Ods;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>ODS 格式集成测试</summary>
public class OdsTests : IntegrationTestBase
{
    [Fact, DisplayName("ODS_复杂写入再读取往返")]
    public void Ods_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.ods");

        var writer = new OdsWriter
        {
            Title = "ODS测试",
            Author = "NewLife Office",
        };

        writer.AddSheet("数据表", new[]
        {
            new[] { "编号", "姓名", "分数" },
            new[] { "1", "张三", "95.5" },
            new[] { "2", "李四", "82.0" },
            new[] { "3", "王五", "71.5" },
            new[] { "4", "赵六", "88.0" },
        });

        writer.AddSheet("汇总", new[]
        {
            new[] { "科目", "平均分", "最高分" },
            new[] { "语文", "85.5", "98" },
            new[] { "数学", "78.0", "100" },
        });

        writer.Save(path);

        Assert.True(File.Exists(path));

        // 读取验证
        var sheets = OdsReader.ReadFile(path);
        Assert.Equal(2, sheets.Count);
        Assert.Equal("数据表", sheets[0].Name);
        Assert.Equal("汇总", sheets[1].Name);

        Assert.Equal(5, sheets[0].Rows.Count);
        Assert.Equal("张三", sheets[0].Rows[1][1]);
        Assert.Equal("95.5", sheets[0].Rows[1][2]);

        Assert.Equal(3, sheets[1].Rows.Count);
        Assert.Equal("语文", sheets[1].Rows[1][0]);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<List<OdsSheet>>(factoryReader);
    }

    [Fact, DisplayName("ODS转XLSX格式")]
    public void Ods_To_Xlsx()
    {
        var odsPath = Path.Combine(OutputDir, "convert_ods.ods");
        var xlsxPath = Path.Combine(OutputDir, "converted_from_ods.xlsx");

        var writer = new OdsWriter();
        writer.AddSheet("Sheet1", new[]
        {
            new[] { "Name", "Value" },
            new[] { "Alpha", "100" },
            new[] { "Beta", "200" },
        });
        writer.Save(odsPath);

        // 读取 ODS 并转为 XLSX
        var sheets = OdsReader.ReadFile(odsPath);
        using var xlsxWriter = new ExcelWriter(xlsxPath);
        foreach (var sheet in sheets)
        {
            if (sheet.Rows.Count > 0)
            {
                xlsxWriter.WriteHeader(sheet.Name, sheet.Rows[0]);
                for (var i = 1; i < sheet.Rows.Count; i++)
                {
                    xlsxWriter.WriteRow(sheet.Name, sheet.Rows[i].Cast<Object?>().ToArray());
                }
            }
        }
        xlsxWriter.Save();

        Assert.True(File.Exists(xlsxPath));

        using var xlsxReader = new ExcelReader(xlsxPath);
        var rows = xlsxReader.ReadRows().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("Alpha", rows[1][0] + "");
    }
}
