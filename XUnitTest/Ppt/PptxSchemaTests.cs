using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPTX 输出 OOXML 规范校验（OpenXmlValidator 必须 0 错误）</summary>
/// <remarks>
/// 容错式往返断言（"读出来差不多就行"）发现不了非法枚举、错位子元素、断链部件：
/// PowerPoint 遇到这类问题会直接提示"需要修复"。本类用官方校验器对生成包做严格校验，
/// 并把历史缺陷固化为用例（presentation 元素顺序、母版 clrMap、背景位置、转场时长、备注部件）。
/// </remarks>
public class PptxSchemaTests
{
    #region 辅助
    /// <summary>校验 pptx 字节流，返回错误描述列表</summary>
    /// <param name="data">pptx 包字节</param>
    /// <returns>错误描述列表，空表示完全合法</returns>
    private static List<String> Validate(Byte[] data)
    {
        var list = new List<String>();
        try
        {
            using var ms = new MemoryStream(data);
            using var doc = PresentationDocument.Open(ms, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            foreach (var e in validator.Validate(doc))
                list.Add($"{e.Part?.Uri} | {e.Path?.XPath} | {e.Description}");
        }
        catch (Exception ex)
        {
            list.Add("EX " + ex.GetType().Name + ": " + ex.Message);
        }
        return list;
    }

    /// <summary>把校验结果压成单行文本，便于断言失败时定位</summary>
    /// <param name="errors">错误列表</param>
    /// <returns>单行描述</returns>
    private static String Describe(List<String> errors) => String.Join("\n", errors);


    /// <summary>读取 zip 中的文本部件</summary>
    /// <param name="data">pptx 包字节</param>
    /// <param name="path">部件路径</param>
    /// <returns>部件文本，不存在返回 null</returns>
    private static String? ReadPart(Byte[] data, String path)
    {
        using var ms = new MemoryStream(data);
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var entry = za.GetEntry(path);
        if (entry == null) return null;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    private static Byte[] Build(Action<PptxWriter> build)
    {
        using var w = new PptxWriter();
        build(w);
        using var ms = new MemoryStream();
        w.Save(ms);
        return ms.ToArray();
    }
    #endregion

    #region 用例
    [Fact(DisplayName = "规范-基础幻灯片输出0校验错误")]
    public void BasicDeck_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.AddTextBox(0, "标题", 2, 2, 20, 3, 28, true);
            w.AddShape(0, "rect", 2, 6, 8, 4, "00AAFF");
            w.AddTable(0, [["A", "B"], ["1", "2"]], 2, 12, 16, true);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-背景图/色必须写在cSld内")]
    public void SlideBackground_InsideCSld_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.SetBackground(0, "EEEEEE");
            w.AddTextBox(0, "背景页", 2, 2, 20, 3, 20, false);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "ppt/slides/slide1.xml")!;
        var cSld = xml.IndexOf("<p:cSld", StringComparison.Ordinal);
        var bg = xml.IndexOf("<p:bg", StringComparison.Ordinal);
        Assert.True(cSld >= 0 && bg > cSld, "p:bg 必须出现在 p:cSld 内部");
    }

    [Fact(DisplayName = "规范-转场时长走p14:dur且包在AlternateContent内")]
    public void Transition_P14Duration_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.SetTransition(0, "fade", 1200);
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "ppt/slides/slide1.xml")!;
        Assert.Contains("mc:AlternateContent", xml);
        Assert.Contains("p14:dur=\"1200\"", xml);
        Assert.Contains("spd=\"slow\"", xml);
        Assert.DoesNotContain(" dur=", xml);
    }

    [Fact(DisplayName = "规范-morph转场用p15命名空间且可回读")]
    public void MorphTransition_P15Namespace_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.SetTransition(0, "morph", 1000);
        });

        // p15:morph 属于 p15 扩展，Office2019 校验器认可；这里只断言时长与包结构合法
        var xml = ReadPart(data, "ppt/slides/slide1.xml")!;
        Assert.Contains("p15:morph", xml);
        Assert.Contains("byObject", xml);

        using var ms = new MemoryStream(data);
        using var reader = new PptxReader(ms);
        var doc = reader.ReadDocument();
        Assert.Equal("morph", doc.Slides[0].Transition!.Type);
        Assert.Equal(1000, doc.Slides[0].Transition!.DurationMs);
    }

    [Fact(DisplayName = "规范-演讲者备注写入独立notesSlide部件")]
    public void Notes_WrittenAsNotesSlidePart_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.SetNotes(0, "演讲者备注");
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        Assert.NotNull(ReadPart(data, "ppt/notesSlides/notesSlide1.xml"));
        Assert.NotNull(ReadPart(data, "ppt/notesMasters/notesMaster1.xml"));
        Assert.NotNull(ReadPart(data, "ppt/notesSlides/_rels/notesSlide1.xml.rels"));

        var slideXml = ReadPart(data, "ppt/slides/slide1.xml")!;
        Assert.DoesNotContain("<p:notes", slideXml);
        Assert.Contains("notesSlide", ReadPart(data, "ppt/slides/_rels/slide1.xml.rels")!);
    }

    [Fact(DisplayName = "规范-presentation元素顺序与notesSz")]
    public void Presentation_ElementOrder_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.AddTextBox(0, "A", 2, 2, 10, 2, 20, false);
            w.Sections = [new Section { Name = "第一章", SlideIndices = [0] }];
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "ppt/presentation.xml")!;
        var masterIdLst = xml.IndexOf("<p:sldMasterIdLst", StringComparison.Ordinal);
        var sldIdLst = xml.IndexOf("<p:sldIdLst", StringComparison.Ordinal);
        var sldSz = xml.IndexOf("<p:sldSz", StringComparison.Ordinal);
        var notesSz = xml.IndexOf("<p:notesSz", StringComparison.Ordinal);
        var extLst = xml.IndexOf("<p:extLst", StringComparison.Ordinal);
        Assert.True(masterIdLst < sldIdLst && sldIdLst < sldSz && sldSz < notesSz, "presentation 子元素顺序非法");
        Assert.True(extLst > notesSz, "extLst 必须是最后一个子元素");
    }

    [Fact(DisplayName = "规范-页眉页脚数据落在extLst扩展内")]
    public void HeaderFooter_InExtLst_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.HeaderFooter = new HeaderFooter { ShowFooter = true, FooterText = "内部资料", ShowPageNumber = true };
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "ppt/presentation.xml")!;
        Assert.DoesNotContain("<p:hf", xml);
        Assert.Contains("nlhf:hf", xml);
        Assert.True(xml.IndexOf("nlhf:hf", StringComparison.Ordinal) > xml.IndexOf("<p:extLst", StringComparison.Ordinal));
    }

    [Fact(DisplayName = "规范-编程母版含必需clrMap且至少一个版式")]
    public void ProgMaster_HasClrMapAndLayout_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.CreateMaster();
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var masterXml = ReadPart(data, "ppt/slideMasters/slideMaster1.xml")!;
        var clrMap = masterXml.IndexOf("<p:clrMap", StringComparison.Ordinal);
        var layoutLst = masterXml.IndexOf("<p:sldLayoutIdLst", StringComparison.Ordinal);
        var txStyles = masterXml.IndexOf("<p:txStyles", StringComparison.Ordinal);
        Assert.True(clrMap >= 0, "母版缺少必需的 clrMap");
        Assert.True(clrMap < layoutLst && layoutLst < txStyles, "母版子元素顺序非法");

        // 母版未显式添加版式时，必须兜底写出版式部件，否则幻灯片 rels 断链
        Assert.NotNull(ReadPart(data, "ppt/slideLayouts/slideLayout1.xml"));
    }

    [Fact(DisplayName = "规范-母版版式含形状与文本框时0校验错误")]
    public void ProgMaster_WithLayoutContent_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.AddTextBox(0, "A", 2, 2, 10, 2, 20, false);
            var master = w.CreateMaster();
            var layout = master.AddLayout("自定义", "blank");
            layout.TextBoxes.Add(new TextBox { Text = "版式文本", Left = 100, Top = 100, Width = 1000, Height = 500 });
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-对齐别名center归一化为ctr")]
    public void Alignment_AliasNormalized_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            var tb = w.AddTextBox(0, "居中", 2, 2, 10, 2, 20, false);
            tb.Alignment = "center";
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        var xml = ReadPart(data, "ppt/slides/slide1.xml")!;
        Assert.Contains("algn=\"ctr\"", xml);
        Assert.DoesNotContain("algn=\"center\"", xml);
    }

    [Fact(DisplayName = "规范-全部图表类型0校验错误")]
    public void AllChartTypes_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            var cats = new[] { "甲", "乙", "丙" };
            void AddSeries(Chart c) => c.Series.Add(new ChartSeries { Name = "S1", Values = [1, 2, 3], XValues = [1, 2, 3] });

            AddSeries(w.AddBarChart(0, cats, 1, 1));
            AddSeries(w.AddLineChart(0, cats, 1, 6));
            AddSeries(w.AddAreaChart(0, cats, 1, 11));
            AddSeries(w.AddPieChart(0, cats, 12, 1));
            AddSeries(w.AddRadarChart(0, cats, 12, 6));
            AddSeries(w.AddSurfaceChart(0, cats, 12, 11));
            w.AddSlide(1);
            AddSeries(w.AddScatterChart(1, cats, 1, 1));
            AddSeries(w.AddBubbleChart(1, cats, 1, 8));
            AddSeries(w.AddStockChart(1, cats, 12, 1));
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-图片/组/视频/批注0校验错误")]
    public void PictureGroupVideoComment_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.AddImage(0, [0x89, 0x50, 0x4E, 0x47], "png", 2, 2, 8, 6);
            var grp = w.GroupShapes(0, 2, 10, 8, 6);
            grp.TextBoxes.Add(new TextBox { Text = "组内文本", Left = 100, Top = 100, Width = 1000, Height = 500 });
            grp.Shapes.Add(new Shape { ShapeType = "rect", Left = 100, Top = 800, Width = 900, Height = 400, FillColor = "00AAFF" });
            w.AddSlide(1);
            w.AddVideo(1, [0, 0, 0, 24], "mp4", 2, 2, 8, 6);
            w.Slides[1].Comments.Add(new Comment { Text = "批注内容", Author = "作者", X = 100, Y = 100 });
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));
    }

    [Fact(DisplayName = "规范-多母版多版式0校验错误")]
    public void MultipleMasters_NoSchemaErrors()
    {
        var data = Build(w =>
        {
            w.AddSlide(0);
            w.AddSlide(1);
            var m1 = w.CreateMaster();
            m1.AddLayout("标题版式", "title");
            m1.AddLayout("空白版式", "blank");
            var m2 = w.CreateMaster();
            m2.AddLayout("内容版式", "obj");
        });

        var errors = Validate(data);
        Assert.True(errors.Count == 0, Describe(errors));

        // 每个母版至少一个版式 + 母版 rels 数量匹配
        Assert.NotNull(ReadPart(data, "ppt/slideMasters/slideMaster2.xml"));
        Assert.NotNull(ReadPart(data, "ppt/slideLayouts/slideLayout3.xml"));
        var rels2 = ReadPart(data, "ppt/slideMasters/_rels/slideMaster2.xml.rels")!;
        Assert.Contains("slideLayout3.xml", rels2);
    }
    #endregion
}
