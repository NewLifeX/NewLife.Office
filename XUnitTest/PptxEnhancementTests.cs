using NewLife.Office;
using Xunit;

namespace XUnitTest;

/// <summary>PPT 模块竞品超越测试 — 表格动态操作/上标下标/渐变背景/SVG/TableStyleGuid</summary>
public class PptxEnhancementTests
{
    #region 表格动态添加/删除行列 (S11-03)
    [Fact(DisplayName = "PPT—表格添加行")]
    public void PptTable_AddRow()
    {
        var tbl = new PptTable();
        tbl.Rows.Add(new[] { "列A", "列B" });
        tbl.AddRow(new[] { "数据1", "数据2" });
        Assert.Equal(2, tbl.Rows.Count);
        Assert.Equal("数据1", tbl.Rows[1][0]);
    }

    [Fact(DisplayName = "PPT—表格删除行")]
    public void PptTable_RemoveRow()
    {
        var tbl = new PptTable();
        tbl.Rows.Add(new[] { "A", "B" });
        tbl.Rows.Add(new[] { "C", "D" });
        tbl.RemoveRow(0);
        Assert.Single(tbl.Rows);
        Assert.Equal("C", tbl.Rows[0][0]);
    }

    [Fact(DisplayName = "PPT—表格插入列")]
    public void PptTable_AddColumn()
    {
        var tbl = new PptTable();
        tbl.Rows.Add(new[] { "A", "B" });
        tbl.AddColumn(1, "新列");
        Assert.Equal(3, tbl.Rows[0].Length);
        Assert.Equal("新列", tbl.Rows[0][1]);
    }
    #endregion

    #region 表格动态操作写入+读取往返 (S11-03)
    [Fact(DisplayName = "PPT—表格动态操作往返")]
    public void PptTable_DynamicOperations_RoundTrip()
    {
        var tempFile = Path.GetTempFileName() + ".pptx";
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tbl = new PptTable
            {
                Left = 1000000, Top = 1000000, Width = 8000000, Height = 4000000,
                FirstRowHeader = true
            };
            tbl.Rows.Add(new[] { "名称", "数量" });
            tbl.AddRow(new[] { "苹果", "100" });
            tbl.AddRow(new[] { "橙子", "200" });
            // 插入列
            tbl.AddColumn(1, "分类");
            tbl.Rows[1][1] = "水果";
            tbl.Rows[2][1] = "水果";
            writer.Slides[0].Tables.Add(tbl);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotNull(doc);
            var slide = doc.Slides[0];
            Assert.NotEmpty(slide.Tables);
            Assert.Equal(3, slide.Tables[0].Rows[0].Length); // 3 columns now
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 上标/下标 (S15-06)
    [Fact(DisplayName = "PPT—上标写入+读取")]
    public void Superscript_WriteAndRead()
    {
        var tempFile = Path.GetTempFileName() + ".pptx";
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tb = new PptTextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new PptTextRun { Text = "E=mc", FontSize = 18 });
            tb.Runs.Add(new PptTextRun { Text = "2", FontSize = 12, Superscript = true });
            writer.Slides[0].TextBoxes.Add(tb);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var runs = doc.Slides[0].TextBoxes[0].Runs;
            Assert.True(runs[1].Superscript);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—下标写入+读取")]
    public void Subscript_WriteAndRead()
    {
        var tempFile = Path.GetTempFileName() + ".pptx";
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tb = new PptTextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new PptTextRun { Text = "H", FontSize = 18 });
            tb.Runs.Add(new PptTextRun { Text = "2", FontSize = 12, Subscript = true });
            tb.Runs.Add(new PptTextRun { Text = "O", FontSize = 18 });
            writer.Slides[0].TextBoxes.Add(tb);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var runs = doc.Slides[0].TextBoxes[0].Runs;
            Assert.True(runs[1].Subscript);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 表格样式主题引用 (S11-04)
    [Fact(DisplayName = "PPT—TableStyleGuid 自定义")]
    public void TableStyleGuid_Custom()
    {
        var tbl = new PptTable { TableStyleGuid = "{D2719A1E-8E4F-4A8D-8AA2-D3E4E5B6F7A8}" };
        Assert.Equal("{D2719A1E-8E4F-4A8D-8AA2-D3E4E5B6F7A8}", tbl.TableStyleGuid);
    }

    [Fact(DisplayName = "PPT—TableStyleGuid 默认 null")]
    public void TableStyleGuid_DefaultNull()
    {
        var tbl = new PptTable();
        Assert.Null(tbl.TableStyleGuid);
    }
    #endregion

    #region 背景渐变 (S15-04)
    [Fact(DisplayName = "PPT—背景渐变写入+读取")]
    public void BackgroundGradient_WriteAndRead()
    {
        var tempFile = Path.GetTempFileName() + ".pptx";
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var slide = writer.Slides[0];
            slide.BackgroundGradientType = "linear";
            slide.BackgroundGradientColor1 = "FF0000";
            slide.BackgroundGradientColor2 = "0000FF";
            var tb = new PptTextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new PptTextRun { Text = "渐变背景", FontSize = 18 });
            slide.TextBoxes.Add(tb);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var s = doc.Slides[0];
            Assert.Equal("linear", s.BackgroundGradientType);
            Assert.Equal("FF0000", s.BackgroundGradientColor1);
            Assert.Equal("0000FF", s.BackgroundGradientColor2);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region SVG 图片 (S15-03)
    [Fact(DisplayName = "PPT—SVG图片 IsSvg属性")]
    public void SvgImage_IsSvg()
    {
        var img = new PptImage { IsSvg = true, Extension = "svg" };
        Assert.True(img.IsSvg);
    }

    [Fact(DisplayName = "PPT—SVG图片写入+读取")]
    public void SvgImage_WriteAndRead()
    {
        var tempFile = Path.GetTempFileName() + ".pptx";
        try
        {
            // 最小有效 SVG 数据
            var svgData = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\"><rect width=\"100\" height=\"100\" fill=\"red\"/></svg>"u8.ToArray();
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Slides[0].Images.Add(new PptImage
            {
                Data = svgData,
                Extension = "svg",
                IsSvg = true,
                Left = 1000000, Top = 1000000, Width = 3000000, Height = 3000000
            });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.Slides[0].Images);
            Assert.Equal("svg", doc.Slides[0].Images[0].Extension);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion
}
