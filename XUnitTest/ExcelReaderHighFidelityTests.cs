using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>ExcelReader 高保真补齐测试 — TextRotation/Indent/ShrinkToFit/outlineLevel/tabColor/workbookProtection/calcPr/IconSet</summary>
public class ExcelReaderHighFidelityTests
{
    #region 辅助
    private static String SaveAndOpen(Action<ExcelWriter> buildWriter)
    {
        var path = Path.GetTempFileName() + ".xlsx";
        using (var w = new ExcelWriter(path))
        {
            buildWriter(w);
            w.Save();
        }
        return path;
    }
    #endregion

    #region TextRotation / ShrinkToFit / Indent 读取
    [Fact(DisplayName = "Excel读—TextRotation读取")]
    public void ReadCellStyle_TextRotation()
    {
        var path = SaveAndOpen(writer =>
        {
            var cs = new ExcelCellStyle { TextRotation = 45 };
            writer.WriteRow(null, new Object?[] { "旋转文字" }, cs);
        });

        using var reader = new ExcelReader(path);
        var styles = reader.ReadCellStyles("Sheet1");
        var (_, style) = styles.First();
        Assert.Equal(45, style.TextRotation);
    }

    [Fact(DisplayName = "Excel读—Indent读取")]
    public void ReadCellStyle_Indent()
    {
        var path = SaveAndOpen(writer =>
        {
            var cs = new ExcelCellStyle { Indent = 3 };
            writer.WriteRow(null, new Object?[] { "缩进文字" }, cs);
        });

        using var reader = new ExcelReader(path);
        var styles = reader.ReadCellStyles("Sheet1");
        var (_, style) = styles.First();
        Assert.Equal(3, style.Indent);
    }

    [Fact(DisplayName = "Excel读—ShrinkToFit读取")]
    public void ReadCellStyle_ShrinkToFit()
    {
        var path = SaveAndOpen(writer =>
        {
            var cs = new ExcelCellStyle { ShrinkToFit = true };
            writer.WriteRow(null, new Object?[] { "很长很长很长很长的文字内容" }, cs);
        });

        using var reader = new ExcelReader(path);
        var styles = reader.ReadCellStyles("Sheet1");
        var (_, style) = styles.First();
        Assert.True(style.ShrinkToFit);
    }
    #endregion

    #region 行列大纲级别读取
    [Fact(DisplayName = "Excel读—列大纲级别")]
    public void ReadColumnOutlines()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.SetColumnOutlineLevel("Sheet1", 0, 1);
            writer.WriteRow(null, new Object?[] { "分组1" });
            writer.SetColumnOutlineLevel("Sheet1", 3, 2);
            writer.WriteRow(null, new Object?[] { "普通", "普通", "普通", "深层嵌套" });
        });

        using var reader = new ExcelReader(path);
        var colOutlines = reader.ReadColumnOutlines("Sheet1");
        Assert.Equal(1, colOutlines[0]);
        Assert.Equal(2, colOutlines[3]);
        Assert.Equal(0, colOutlines.GetValueOrDefault(1));
    }

    [Fact(DisplayName = "Excel读—行大纲级别")]
    public void ReadRowOutlines()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.SetRowOutlineLevel("Sheet1", 1, 1);
            writer.WriteRow(null, new Object?[] { "行分组" });
            writer.SetRowOutlineLevel("Sheet1", 4, 2);
            writer.WriteRow(null, new Object?[] { "普通", "普通", "普通" });
            writer.WriteRow(null, new Object?[] { "普通2" });
            writer.WriteRow(null, new Object?[] { "深层嵌套行" });
        });

        using var reader = new ExcelReader(path);
        var rowOutlines = reader.ReadRowOutlines("Sheet1");
        Assert.Equal(1, rowOutlines[0]);
        Assert.Equal(2, rowOutlines[3]);
    }
    #endregion

    #region 标签颜色读取
    [Fact(DisplayName = "Excel读—工作表标签颜色")]
    public void ReadTabColor()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.SetSheetTabColor("Sheet1", "FF0000");
            writer.WriteRow(null, new Object?[] { "红色标签" });
        });

        using var reader = new ExcelReader(path);
        var colors = reader.ReadTabColors();
        Assert.Single(colors);
        Assert.Contains("FF0000", colors.Values.First());
    }
    #endregion

    #region 工作簿保护读取
    [Fact(DisplayName = "Excel读—工作簿保护")]
    public void ReadWorkbookProtection()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.ProtectWorkbook(password: "test123", lockStructure: true, lockWindows: false);
            writer.WriteRow(null, new Object?[] { "受保护" });
        });

        using var reader = new ExcelReader(path);
        var prot = reader.ReadWorkbookProtection();
        Assert.NotNull(prot);
        Assert.True(prot.Value.LockStructure);
        Assert.False(prot.Value.LockWindows);
        Assert.NotNull(prot.Value.PasswordHash);
    }
    #endregion

    #region calcPr 读取
    [Fact(DisplayName = "Excel读—计算选项calcPr")]
    public void ReadCalcPr()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.WriteRow(null, new Object?[] { "测试" });
        });

        using var reader = new ExcelReader(path);
        var calcPr = reader.ReadCalcPr();
        Assert.NotNull(calcPr);
        Assert.True(calcPr.Value.CalcId > 0);
        Assert.True(calcPr.Value.FullCalcOnLoad);
    }
    #endregion

    #region 条件格式 IconSet 读取
    [Fact(DisplayName = "Excel读—IconSet条件格式")]
    public void ReadConditionalFormat_IconSet()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.WriteRow(null, new Object?[] { 10, 50, 90 });
            writer.AddIconSetConditionalFormat("Sheet1", "A1:A3", "3Arrows");
        });

        using var reader = new ExcelReader(path);
        var formats = reader.ReadConditionalFormats("Sheet1").ToList();
        Assert.NotEmpty(formats);
        var iconSet = formats.FirstOrDefault(f => f.Type == ExcelConditionalFormatType.IconSet);
        Assert.NotNull(iconSet);
        Assert.Equal("3Arrows", iconSet.IconSetType);
    }

    [Fact(DisplayName = "Excel读—Expression条件格式")]
    public void ReadConditionalFormat_Expression()
    {
        var path = SaveAndOpen(writer =>
        {
            writer.WriteRow(null, new Object?[] { 10, 50, 90 });
            writer.AddExpressionConditionalFormat("Sheet1", "B1:B3", "B1>30", "FF0000");
        });

        using var reader = new ExcelReader(path);
        var formats = reader.ReadConditionalFormats("Sheet1").ToList();
        Assert.NotEmpty(formats);
        var expr = formats.FirstOrDefault(f => f.Type == ExcelConditionalFormatType.Expression);
        Assert.NotNull(expr);
        Assert.Contains("B1>30", expr.Formula);
    }
    #endregion
}
