using System.ComponentModel;
using System.Text;
using NewLife.Office.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>xls（BIFF8）交叉兼容：与 NPOI 双向读写</summary>
/// <remarks>
/// BIFF8 的 XF 记录长度、SST + CONTINUE 的分段规则都很容易被写错，
/// 单靠本库自读自写无法发现（自己写的偏移自己也按同样规则读）。
/// 这里用 NPOI 作为独立实现双向验证。
/// </remarks>
public class XlsCrossCompatibilityTests
{
    [Fact, DisplayName("X1 NPOI 读取本库生成的 xls")]
    public void X1_Npoi_ReadsOurXls()
    {
        using var ms = new MemoryStream();
        using (var w = new BiffWriter())
        {
            w.SheetName = "数据";
            w.WriteHeader(["编号", "名称", "金额"]);
            w.WriteRow([1, "甲", 12.5]);
            w.WriteRow([2, "乙", -3.25]);
            w.Save(ms);
        }

        ms.Position = 0;
        using var wb = new HSSFWorkbook(ms);
        var sheet = wb.GetSheet("数据");
        Assert.NotNull(sheet);

        Assert.Equal("编号", sheet.GetRow(0).GetCell(0).StringCellValue);
        var row1 = sheet.GetRow(1);
        Assert.Equal(CellType.Numeric, row1.GetCell(0).CellType);
        Assert.Equal(1, (Int32)row1.GetCell(0).NumericCellValue);
        Assert.Equal("甲", row1.GetCell(1).StringCellValue);
        Assert.Equal(12.5, row1.GetCell(2).NumericCellValue);
        Assert.Equal(-3.25, sheet.GetRow(2).GetCell(2).NumericCellValue);
    }

    [Fact, DisplayName("X2 NPOI 读取本库生成的长字符串 xls")]
    public void X2_Npoi_ReadsOurLongStrings()
    {
        var longText = new String('A', 10_000) + "尾部";
        using var ms = new MemoryStream();
        using (var w = new BiffWriter())
        {
            w.SheetName = "长文本";
            w.WriteHeader(["内容"]);
            w.WriteRow([longText]);
            w.Save(ms);
        }

        ms.Position = 0;
        using var wb = new HSSFWorkbook(ms);
        var actual = wb.GetSheet("长文本").GetRow(1).GetCell(0).StringCellValue;
        Assert.Equal(longText, actual);
    }

    [Fact, DisplayName("X3 本库读取 NPOI 生成的长字符串 xls")]
    public void X3_ReadNpoiLongStrings()
    {
        var longText = new String('B', 10_000) + "尾部";
        using var ms = new MemoryStream();
        using (var wb = new HSSFWorkbook())
        {
            var sheet = wb.CreateSheet("长文本");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(longText);
            wb.Write(ms, false);
        }

        ms.Position = 0;
        using var reader = new BiffReader(ms);
        var rows = reader.ReadSheet("长文本").ToList();
        Assert.Single(rows);
        Assert.Equal(longText, rows[0][0] + "");
    }

    [Fact, DisplayName("X4 本库读取 NPOI 生成的多单元格 xls")]
    public void X4_ReadNpoiCells()
    {
        using var ms = new MemoryStream();
        using (var wb = new HSSFWorkbook())
        {
            var sheet = wb.CreateSheet("混合");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("名称");
            header.CreateCell(1).SetCellValue("数量");
            header.CreateCell(2).SetCellValue("比例");
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("丙");
            row.CreateCell(1).SetCellValue(7);
            row.CreateCell(2).SetCellValue(0.5);
            wb.Write(ms, false);
        }

        ms.Position = 0;
        using var reader = new BiffReader(ms);
        var rows = reader.ReadSheet("混合").ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("名称", rows[0][0] + "");
        Assert.Equal("丙", rows[1][0] + "");
        Assert.Equal(7, Convert.ToInt32(rows[1][1]));
        Assert.Equal(0.5, Convert.ToDouble(rows[1][2]), 6);
    }
}
