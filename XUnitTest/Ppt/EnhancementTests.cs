using NewLife.Office;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPT 模块竞品超越测试 — 表格动态操作/上标下标/渐变背景/SVG/TableStyleGuid</summary>
public class PptxEnhancementTests
{
    #region 表格动态添加/删除行列 (S11-03)
    [Fact(DisplayName = "PPT—表格添加行")]
    public void PptTable_AddRow()
    {
        var tbl = new Table();
        tbl.Rows.Add(new[] { "列A", "列B" });
        tbl.AddRow(new[] { "数据1", "数据2" });
        Assert.Equal(2, tbl.Rows.Count);
        Assert.Equal("数据1", tbl.Rows[1][0]);
    }

    [Fact(DisplayName = "PPT—表格删除行")]
    public void PptTable_RemoveRow()
    {
        var tbl = new Table();
        tbl.Rows.Add(new[] { "A", "B" });
        tbl.Rows.Add(new[] { "C", "D" });
        tbl.RemoveRow(0);
        Assert.Single(tbl.Rows);
        Assert.Equal("C", tbl.Rows[0][0]);
    }

    [Fact(DisplayName = "PPT—表格插入列")]
    public void PptTable_AddColumn()
    {
        var tbl = new Table();
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
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tbl = new Table
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
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tb = new TextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new Run { Text = "E=mc", FontSize = 18 });
            tb.Runs.Add(new Run { Text = "2", FontSize = 12, Superscript = true });
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
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tb = new TextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new Run { Text = "H", FontSize = 18 });
            tb.Runs.Add(new Run { Text = "2", FontSize = 12, Subscript = true });
            tb.Runs.Add(new Run { Text = "O", FontSize = 18 });
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
        var tbl = new Table { TableStyleGuid = "{D2719A1E-8E4F-4A8D-8AA2-D3E4E5B6F7A8}" };
        Assert.Equal("{D2719A1E-8E4F-4A8D-8AA2-D3E4E5B6F7A8}", tbl.TableStyleGuid);
    }

    [Fact(DisplayName = "PPT—TableStyleGuid 默认 null")]
    public void TableStyleGuid_DefaultNull()
    {
        var tbl = new Table();
        Assert.Null(tbl.TableStyleGuid);
    }
    #endregion

    #region 背景渐变 (S15-04)
    [Fact(DisplayName = "PPT—背景渐变写入+读取")]
    public void BackgroundGradient_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var slide = writer.Slides[0];
            slide.BackgroundGradientType = "linear";
            slide.BackgroundGradientColor1 = "FF0000";
            slide.BackgroundGradientColor2 = "0000FF";
            var tb = new TextBox { Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000 };
            tb.Runs.Add(new Run { Text = "渐变背景", FontSize = 18 });
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
        var img = new Picture { IsSvg = true, Extension = "svg" };
        Assert.True(img.IsSvg);
    }

    [Fact(DisplayName = "PPT—SVG图片写入+读取")]
    public void SvgImage_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            // 最小有效 SVG 数据
            var svgData = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\"><rect width=\"100\" height=\"100\" fill=\"red\"/></svg>"u8.ToArray();
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Slides[0].Images.Add(new Picture
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

    [Fact(DisplayName = "PPT—SVG图片 asvg:svgBlip XML生成验证")]
    public void SvgImage_AsvgSvgBlipXml()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            var slide = writer.AddSlide();
            var svgData = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"50\" height=\"50\"><circle r=\"20\" fill=\"blue\"/></svg>"u8.ToArray();
            slide.Images.Add(new Picture
            {
                Data = svgData,
                Extension = "svg",
                IsSvg = true,
                Left = 0, Top = 0, Width = 2000000, Height = 2000000
            });
            writer.Save(tempFile);

            // 验证 PPTX 文件中包含 asvg:svgBlip 元素
            using var archive = System.IO.Compression.ZipFile.OpenRead(tempFile);
            var slideEntry = archive.GetEntry("ppt/slides/slide1.xml");
            Assert.NotNull(slideEntry);
            using var sr = new StreamReader(slideEntry!.Open());
            var slideXml = sr.ReadToEnd();
            Assert.Contains("asvg:svgBlip", slideXml);

            // 同时验证往返读取
            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.True(doc.Slides[0].Images[0].IsSvg);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region Alt Text (S15-new)
    [Fact(DisplayName = "PPT—形状Alt Text写入+读取")]
    public void AltText_Shape_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Slides[0].Shapes.Add(new Shape
            {
                ShapeType = "rect",
                Left = 1000000, Top = 1000000, Width = 3000000, Height = 2000000,
                AltText = "红色矩形装饰",
                Text = null // No text → treated as shape, not text box
            });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.Slides[0].Shapes);
            Assert.Equal("红色矩形装饰", doc.Slides[0].Shapes[0].AltText);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—文本框Alt Text写入+读取")]
    public void AltText_TextBox_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var tb = new TextBox
            {
                Left = 1000000, Top = 1000000, Width = 5000000, Height = 1000000,
                AltText = "标题文本框"
            };
            tb.Runs.Add(new Run { Text = "标题", FontSize = 24 });
            writer.Slides[0].TextBoxes.Add(tb);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.Slides[0].TextBoxes);
            Assert.Equal("标题文本框", doc.Slides[0].TextBoxes[0].AltText);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—圆角矩形CornerRadius写入+读取")]
    public void RoundRect_CornerRadius_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Slides[0].Shapes.Add(new Shape
            {
                ShapeType = "roundRect",
                Left = 1000000, Top = 1000000, Width = 5000000, Height = 3000000,
                CornerRadius = 300000,
                Text = null
            });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.Slides[0].Shapes);
            Assert.True(doc.Slides[0].Shapes[0].CornerRadius > 0);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—文本框东亚竖排(eaVert) TextDirection")]
    public void TextDirection_EaVert_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            var slide = writer.AddSlide();
            var tb = writer.AddTextBox(0, "竖排文本", 1, 1, 10, 5);
            tb.TextDirection = "eaVert";
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotEmpty(doc.Slides[0].TextBoxes);
            Assert.Equal("eaVert", doc.Slides[0].TextBoxes[0].TextDirection);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—文本框垂直旋转270°(vert270)")]
    public void TextDirection_Vert270_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            var slide = writer.AddSlide();
            var tb = writer.AddTextBox(0, "旋转文本", 1, 1, 10, 5);
            tb.TextDirection = "vert270";
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Equal("vert270", doc.Slides[0].TextBoxes[0].TextDirection);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—文本框默认水平方向(null)")]
    public void TextDirection_DefaultHorz()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            var slide = writer.AddSlide();
            writer.AddTextBox(0, "默认水平", 1, 1, 10, 2);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Null(doc.Slides[0].TextBoxes[0].TextDirection);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—Section写入读取往返")]
    public void Section_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.AddSlide();
            writer.AddSlide();
            writer.Sections = new List<Section>
            {
                new() { Name = "第一章", SlideIndices = [0, 1] },
                new() { Name = "第二章", SlideIndices = [2] }
            };
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.NotNull(doc.Sections);
            Assert.Equal(2, doc.Sections!.Count);
            Assert.Equal("第一章", doc.Sections[0].Name);
            Assert.Equal([0, 1], doc.Sections[0].SlideIndices);
            Assert.Equal("第二章", doc.Sections[1].Name);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—无Section时Sections为null")]
    public void Section_NoneIsNull()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            Assert.Null(doc.Sections);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 组内图片与组内形状属性往返 (S07-02)
    [Fact(DisplayName = "PPT—组内图片写入+读取往返")]
    public void GroupImage_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var grp = writer.GroupShapes(0, 2, 2, 10, 6);
            grp.Images.Add(new Picture
            {
                Data = [0x89, 0x50, 0x4E, 0x47, 1, 2, 3],
                Extension = "png",
                Left = 100000, Top = 200000, Width = 2000000, Height = 1500000,
                Rotation = 5400000,
            });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var slide = doc.Slides[0];
            Assert.Single(slide.Groups);
            Assert.Empty(slide.Groups[0].Shapes);
            Assert.Single(slide.Groups[0].Images);
            var img = slide.Groups[0].Images[0];
            Assert.Equal("png", img.Extension);
            Assert.Equal(5400000, img.Rotation);
            Assert.Equal(2000000, img.Width);
            Assert.Equal(1500000, img.Height);
            Assert.Equal(7, img.Data.Length);
            Assert.True(img.Data[0] == 0x89 && img.Data[1] == 0x50);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—组内形状属性（旋转/翻转/渐变/虚线/文本）往返")]
    public void GroupShape_RichProperties_WriteAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var grp = writer.GroupShapes(0, 2, 2, 12, 8);
            grp.Shapes.Add(new Shape
            {
                ShapeType = "roundRect",
                Left = 100000, Top = 100000, Width = 4000000, Height = 2000000,
                Rotation = 1800000,
                FlipVertical = true,
                CornerRadius = 400000,
                LineColor = "#00FF00",
                LineWidth = 12700,
                DashStyle = "dash",
                Text = "组内形状",
                FontSize = 18,
                Bold = true,
                GradientType = "linear",
                GradientColor1 = "#FF0000",
                GradientColor2 = "#0000FF",
                GradientAngle = 45,
            });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var slide = doc.Slides[0];
            Assert.Single(slide.Groups);
            var sp = Assert.Single(slide.Groups[0].Shapes);
            Assert.Equal("roundRect", sp.ShapeType);
            Assert.Equal("组内形状", sp.Text);
            Assert.Equal(18, sp.FontSize);
            Assert.True(sp.Bold);
            Assert.Equal(1800000, sp.Rotation);
            Assert.True(sp.FlipVertical);
            Assert.Equal(12700, sp.LineWidth);
            Assert.Equal("00FF00", sp.LineColor);
            Assert.Equal("dash", sp.DashStyle);
            Assert.Equal("linear", sp.GradientType);
            Assert.Equal("0000FF", sp.GradientColor2);
            Assert.Equal(45, sp.GradientAngle);
            // 圆角按 adj 存储，≤2 EMU 容差
            Assert.True(Math.Abs(sp.CornerRadius - 400000) <= 2, $"CornerRadius: {sp.CornerRadius}");
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 图表增强 (S19 / Phase4)
    [Fact(DisplayName = "PPT—面积图创建+读取")]
    public void AreaChart_CreateAndRead()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var chart = writer.AddAreaChart(0, ["一月", "二月", "三月"]);
            chart.Series.Add(new ChartSeries { Name = "收入", Values = [100, 150, 200] });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var c = doc.Slides[0].Charts[0];
            Assert.Equal("area", c.ChartType);
            Assert.Equal("收入", c.Series[0].Name);
            Assert.Equal([100d, 150d, 200d], c.Series[0].Values);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—图表系列颜色往返")]
    public void Chart_SeriesColor_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var chart = writer.AddBarChart(0, ["A", "B"]);
            chart.Series.Add(new ChartSeries { Name = "S1", Values = [1, 2], Color = "FF8800" });
            chart.Series.Add(new ChartSeries { Name = "S2", Values = [3, 4] }); // 无颜色→默认配色
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var cs = doc.Slides[0].Charts[0].Series;
            Assert.Equal("FF8800", cs[0].Color);
            Assert.Equal("C0504D", cs[1].Color); // 默认配色第 2 个（无自定义色时按系列索引取默认）
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—散点图XValues往返")]
    public void ScatterChart_XValues_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var chart = writer.AddScatterChart(0, ["A", "B", "C"]);
            chart.Series.Add(new ChartSeries { Name = "散点", Values = [10, 20, 30], XValues = [0.5, 1.5, 2.5] });
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var ser = doc.Slides[0].Charts[0].Series[0];
            Assert.NotNull(ser.XValues);
            Assert.Equal(3, ser.XValues!.Length);
            Assert.Equal(0.5, ser.XValues![0]);
            Assert.Equal(2.5, ser.XValues![2]);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—图表数值轴范围与图例位置往返")]
    public void Chart_AxisRangeAndLegend_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var chart = writer.AddLineChart(0, ["A", "B"]);
            chart.Series.Add(new ChartSeries { Name = "S", Values = [1, 2] });
            chart.AxisMinValue = 0;
            chart.AxisMaxValue = 100;
            chart.LegendPosition = "r";
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var c = doc.Slides[0].Charts[0];
            Assert.Equal(0d, c.AxisMinValue);
            Assert.Equal(100d, c.AxisMaxValue);
            Assert.Equal("r", c.LegendPosition);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion

    #region 形状级超链接 (S20 / Phase5)
    [Fact(DisplayName = "PPT—形状URL超链接往返")]
    public void Shape_HyperlinkUrl_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var sp = writer.AddShape(0, "rect", 2, 2, 10, 5, "FF0000");
            sp.HyperlinkUrl = "https://newlifex.com";
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var s = doc.Slides[0].Shapes[0];
            Assert.Equal("https://newlifex.com", s.HyperlinkUrl);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—形状文件超链接往返")]
    public void Shape_HyperlinkFile_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var sp = writer.AddShape(0, "rect", 2, 2, 10, 5, "00AA00");
            sp.HyperlinkUrl = "C:\\report.pdf"; // 文件跳转（TargetMode=External）
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var s = doc.Slides[0].Shapes[0];
            Assert.Equal("C:\\report.pdf", s.HyperlinkUrl);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }

    [Fact(DisplayName = "PPT—组内形状超链接往返")]
    public void GroupShape_Hyperlink_RoundTrip()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        try
        {
            using var writer = new PptxWriter();
            writer.AddSlide();
            var grp = writer.GroupShapes(0, 2, 2, 12, 8);
            var sp = new Shape { ShapeType = "rect", Left = 100000, Top = 100000, Width = 3000000, Height = 1500000, FillColor = "FF0000", HyperlinkUrl = "https://example.com" };
            grp.Shapes.Add(sp);
            writer.Save(tempFile);

            using var reader = new PptxReader(tempFile);
            var doc = reader.ReadDocument();
            var s = doc.Slides[0].Groups[0].Shapes[0];
            Assert.Equal("https://example.com", s.HyperlinkUrl);
        }
        finally { if (File.Exists(tempFile)) File.Delete(tempFile); }
    }
    #endregion
}
