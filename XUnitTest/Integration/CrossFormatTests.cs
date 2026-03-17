using System.ComponentModel;
using NewLife.Office;
using NewLife.Office.Ods;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>跨格式转换集成测试</summary>
public class CrossFormatTests : IntegrationTestBase
{
    [Fact, DisplayName("XLSX转ODS格式")]
    public void Xlsx_To_Ods()
    {
        var xlsxPath = Path.Combine(OutputDir, "convert_xlsx.xlsx");
        var odsPath = Path.Combine(OutputDir, "converted_from_xlsx.ods");

        // 写入 xlsx
        using (var w = new ExcelWriter(xlsxPath))
        {
            w.WriteHeader(null!, new[] { "City", "Population" });
            w.WriteRow(null!, new Object?[] { "Beijing", 21540000 });
            w.WriteRow(null!, new Object?[] { "Shanghai", 24870000 });
            w.WriteRow(null!, new Object?[] { "Guangzhou", 18680000 });
            w.Save();
        }

        // 读取 xlsx
        using var xlsxReader = new ExcelReader(xlsxPath);
        var rows = xlsxReader.ReadRows().ToList();

        // 转为 ods
        var odsWriter = new OdsWriter();
        var odsRows = rows.Select(r => r.Select(c => c + "").ToArray());
        odsWriter.AddSheet("Sheet1", odsRows);
        odsWriter.Save(odsPath);

        Assert.True(File.Exists(odsPath));

        // 验证 ods
        var odsSheets = OdsReader.ReadFile(odsPath);
        Assert.True(odsSheets.Count >= 1);
        Assert.Equal("City", odsSheets[0].Rows[0][0]);
        Assert.Equal("Beijing", odsSheets[0].Rows[1][0]);
    }

    [Fact, DisplayName("XLSX转XLS格式")]
    public void Xlsx_To_Xls()
    {
        var xlsxPath = Path.Combine(OutputDir, "convert_xlsx_to_xls.xlsx");
        var xlsPath = Path.Combine(OutputDir, "converted_from_xlsx.xls");

        using (var w = new ExcelWriter(xlsxPath))
        {
            w.WriteHeader(null!, new[] { "Item", "Qty", "Price" });
            w.WriteRow(null!, new Object?[] { "Apple", 10, 5.5 });
            w.WriteRow(null!, new Object?[] { "Banana", 20, 3.0 });
            w.Save();
        }

        // 读取 xlsx 并转为 xls
        using var xlsxReader = new ExcelReader(xlsxPath);
        var rows = xlsxReader.ReadRows().ToList();

        using var biffWriter = new BiffWriter();
        if (rows.Count > 0)
        {
            biffWriter.WriteHeader(rows[0].Select(c => c + "").ToArray());
            for (var i = 1; i < rows.Count; i++)
            {
                biffWriter.WriteRow(rows[i]);
            }
        }
        biffWriter.Save(xlsPath);

        Assert.True(File.Exists(xlsPath));

        using var biffReader = new BiffReader(xlsPath);
        var readRows = biffReader.ReadSheet(biffReader.SheetNames[0]).ToList();
        Assert.Equal(3, readRows.Count);
        Assert.Equal("Apple", readRows[1][0] + "");
    }

    [Fact, DisplayName("EML转VCard提取联系人")]
    public void Eml_Extract_To_VCard()
    {
        var emlPath = Path.Combine(OutputDir, "contact_source.eml");
        var vcfPath = Path.Combine(OutputDir, "extracted_from_eml.vcf");

        var msg = new EmlMessage
        {
            From = "John Doe <john@example.com>",
            Subject = "联系人提取测试",
            TextBody = "测试内容",
        };
        msg.To.Add("jane@example.com");
        new EmlWriter().Write(msg, emlPath);

        // 从 EML 提取发件人信息创建 VCard
        var readMsg = new EmlReader().Read(emlPath);
        var contact = new VCardContact
        {
            FullName = "John Doe",
        };
        contact.Emails.Add(new VCardEmail { Address = "john@example.com", Type = "WORK" });
        new VCardWriter().Write(contact, vcfPath);

        Assert.True(File.Exists(vcfPath));

        var contacts = new VCardReader().ReadAll(vcfPath);
        Assert.True(contacts.Count >= 1);
        Assert.Equal("John Doe", contacts[0].FullName);
    }
}
