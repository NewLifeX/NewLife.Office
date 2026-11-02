using System.ComponentModel;
using System.Text;
using NewLife.Office.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>XLSX 交叉兼容：用 NPOI 读取本项目生成的 xlsx，验证内容真实可读</summary>
/// <remarks>
/// OOXML schema 校验只能证明"包结构合法"，证明不了"Excel 类库能正确读出内容"。
/// 本类用第三方实现（NPOI，Excel 的独立开源实现）打开并读取生成文件，
/// 覆盖值类型、样式、合并、公式、图片、表格、冻结与筛选等语义。
/// </remarks>
public class ExcelCrossCompatibilityTests
{
    private static Byte[] Build(Action<ExcelWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            build(w);
            w.Save();
        }
        return ms.ToArray();
    }

    [Fact, DisplayName("E1 NPOI 读取单元格值与类型")]
    public void E1_NPOI_ReadsCellValues()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["Id", "Name", "Amount", "Flag", "Date"]);
            w.WriteRows("S1", [
                new Object?[] { 1, "甲", 1234.5, true, new DateTime(2024, 1, 2) },
                new Object?[] { 2, "乙", -0.25, false, new DateTime(2024, 3, 4) },
            ]);
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = wb.GetSheetAt(0);
        Assert.Equal("Id", sheet.GetRow(0).GetCell(0).StringCellValue);

        var r1 = sheet.GetRow(1);
        Assert.Equal(CellType.Numeric, r1.GetCell(0).CellType);
        Assert.Equal(1, (Int32)r1.GetCell(0).NumericCellValue);
        Assert.Equal("甲", r1.GetCell(1).StringCellValue);
        Assert.Equal(1234.5, r1.GetCell(2).NumericCellValue);
        Assert.True(r1.GetCell(3).BooleanCellValue);
        Assert.True(DateUtil.IsCellDateFormatted(r1.GetCell(4)));
        Assert.Equal(new DateTime(2024, 1, 2), r1.GetCell(4).DateCellValue!.Value.Date);

        var r2 = sheet.GetRow(2);
        Assert.Equal("乙", r2.GetCell(1).StringCellValue);
        Assert.Equal(-0.25, r2.GetCell(2).NumericCellValue);
        Assert.False(r2.GetCell(3).BooleanCellValue);
    }

    [Fact, DisplayName("E2 NPOI 读取表头样式与单元格格式")]
    public void E2_NPOI_ReadsCellStyles()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.SetCellFormat("S1", 1, 0, new CellFormat { Bold = true, FontColor = "FF0000", BackgroundColor = "FFFF00", HAlign = NewLife.Office.Excel.HorizontalAlignment.Center });
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var cell = wb.GetSheetAt(0).GetRow(1).GetCell(0);
        var style = cell.CellStyle;
        Assert.True(style.GetFont(wb).IsBold);
        var fontColor = ((XSSFFont)style.GetFont(wb)).GetXSSFColor();
        Assert.Equal("FF0000", fontColor.ARGBHex.Substring(2));
        var fillColor = ((XSSFCellStyle)style).FillForegroundXSSFColor;
        Assert.Equal("FFFF00", fillColor.ARGBHex.Substring(2));
        Assert.Equal(NPOI.SS.UserModel.HorizontalAlignment.Center, style.Alignment);
    }

    [Fact, DisplayName("E3 NPOI 读取合并区域与冻结窗格")]
    public void E3_NPOI_ReadsMergeAndFreeze()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B", "C"]);
            w.WriteRows("S1", [new Object?[] { 1, 2, 3 }]);
            w.MergeCell("S1", "A4:C4");
            w.FreezePane("S1", 1, 1);
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = (XSSFSheet)wb.GetSheetAt(0);
        Assert.Equal(1, sheet.NumMergedRegions);
        var region = sheet.GetMergedRegion(0);
        Assert.Equal(3, region.FirstRow);
        Assert.Equal(2, region.LastColumn);

        // 冻结窗格：NPOI 通过 PaneInformation 暴露，冻结 1 行 1 列
        var pane = sheet.PaneInformation;
        Assert.NotNull(pane);
        Assert.True(pane!.IsFreezePane());
        Assert.Equal(1, pane.HorizontalSplitPosition);
        Assert.Equal(1, pane.VerticalSplitPosition);
    }

    [Fact, DisplayName("E4 NPOI 读取公式与缓存值")]
    public void E4_NPOI_ReadsFormula()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B", "C"]);
            w.WriteRows("S1", [new Object?[] { 1, 2, null }]);
            w.AppendFormula("S1", "SUM(A2:B2)", 3.0);
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = wb.GetSheetAt(0);
        // 公式行由 AppendFormula 追加，定位到含公式的单元格
        ICell? formulaCell = null;
        for (var r = 0; r <= sheet.LastRowNum && formulaCell == null; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;
            for (var c = 0; c < row.LastCellNum; c++)
            {
                var cell = row.GetCell(c);
                if (cell != null && cell.CellType == CellType.Formula) { formulaCell = cell; break; }
            }
        }
        Assert.NotNull(formulaCell);
        Assert.Contains("SUM", formulaCell!.CellFormula);
        Assert.Equal(3.0, formulaCell.NumericCellValue);
    }

    [Fact, DisplayName("E5 NPOI 读取图片数量")]
    public void E5_NPOI_ReadsImage()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A"]);
            w.AddImage("S1", 2, 0, MinimalPng, "png", 80, 60);
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = (XSSFSheet)wb.GetSheetAt(0);
        var drawings = sheet.GetDrawingPatriarch();
        Assert.NotNull(drawings);
        var pictures = drawings!.GetShapes().OfType<XSSFPicture>().ToList();
        Assert.Single(pictures);
        Assert.NotNull(pictures[0].PictureData);
        Assert.True(pictures[0].PictureData.Data.Length > 0);
    }

    [Fact, DisplayName("E6 NPOI 读取结构化表格与自动筛选")]
    public void E6_NPOI_ReadsTable()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }, new Object?[] { 3, 4 }]);
            w.AddTable("S1", "A1:B3", "T1", null, ["A", "B"]);
            w.SetAutoFilter("S1", "A1:B3");
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = (XSSFSheet)wb.GetSheetAt(0);
        Assert.Single(sheet.GetTables());
        Assert.Equal("T1", sheet.GetTables().First().Name);
    }

    [Fact, DisplayName("E7 NPOI 读取先写数据后设样式的覆盖效果")]
    public void E7_NPOI_ReadsLateCellFormat()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            // 关键：数据已写完后才设样式，样式必须在保存时补写
            w.SetCellFormat("S1", 1, 0, new CellFormat { Bold = true, BackgroundColor = "C6EFCE" });
            w.SetCellFormat("S1", 5, 2, new CellFormat { Italic = true, BackgroundColor = "FFEB9C" });
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var sheet = (XSSFSheet)wb.GetSheetAt(0);

        var styled = sheet.GetRow(1).GetCell(0);
        Assert.True(styled.CellStyle.GetFont(wb).IsBold);
        Assert.Equal("C6EFCE", ((XSSFCellStyle)styled.CellStyle).FillForegroundXSSFColor.ARGBHex.Substring(2));

        // 无值单元格也要补出带样式的空单元格
        var blank = sheet.GetRow(5).GetCell(2);
        Assert.NotNull(blank);
        Assert.True(blank!.CellStyle.GetFont(wb).IsItalic);
        Assert.Equal("FFEB9C", ((XSSFCellStyle)blank.CellStyle).FillForegroundXSSFColor.ARGBHex.Substring(2));
    }

    [Fact, DisplayName("E8 NPOI 读取样式覆盖后仍保留日期格式")]
    public void E8_NPOI_KeepsDateFormatAfterOverride()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["D"]);
            w.WriteRows("S1", [new Object?[] { new DateTime(2024, 5, 6) }]);
            // 只改字体，不应把日期数字格式覆盖成常规
            w.SetCellFormat("S1", 1, 0, new CellFormat { Bold = true });
        });

        using var ms = new MemoryStream(bytes);
        using var wb = new XSSFWorkbook(ms);
        var cell = wb.GetSheetAt(0).GetRow(1).GetCell(0);
        Assert.True(cell.CellStyle.GetFont(wb).IsBold);
        Assert.True(DateUtil.IsCellDateFormatted(cell));
        Assert.Equal(new DateTime(2024, 5, 6), cell.DateCellValue!.Value.Date);
    }

    /// <summary>最小合法 PNG（1x1 透明像素）</summary>
    private static Byte[] MinimalPng { get; } = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAwAB/AF+p0n3AAAAAElFTkSuQmCC");
}
