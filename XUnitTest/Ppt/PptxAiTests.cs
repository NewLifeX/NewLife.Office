using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PPT AI 接口测试：结构快照、按锚点改写、幻灯片编排与失败不保存</summary>
public class PptxAiTests
{
    #region 工具
    private static String TempPath(String name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "nloffice-ai", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    // 1x1 透明 PNG
    private const String Png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    private static String BuildPptx()
    {
        var path = TempPath("sample.pptx");
        using var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.AddTextBox(0, "季度报告", 1, 1, 20, 2, 32, true);
        writer.AddTextBox(0, "营收增长 20%", 1, 4, 20, 3);
        writer.AddTable(0, [["地区", "营收"], ["华东", "100"]]);
        writer.AddImage(0, Convert.FromBase64String(Png1x1), "png", 1, 8, 3, 3);
        writer.SetNotes(0, "汇报要点");
        writer.Save(path);
        return path;
    }

    private static Presentation Read(String path)
    {
        using var reader = new PptxReader(path);
        return reader.ReadDocument();
    }
    #endregion

    [Fact(DisplayName = "锚点-解析幻灯片与元素")]
    public void Anchor_Parse()
    {
        var slide = PptxAnchor.Parse("s#2");
        Assert.Equal(2, slide.SlideIndex);
        Assert.Null(slide.Kind);

        var element = PptxAnchor.Parse("s#2/tb#5");
        Assert.Equal(2, element.SlideIndex);
        Assert.Equal("tb", element.Kind);
        Assert.Equal(5, element.Index);
        Assert.Equal("s#2/tb#5", element.ToString());

        Assert.Throws<ArgumentException>(() => PptxAnchor.Parse("x#1"));
        Assert.Throws<ArgumentException>(() => PptxAnchor.Parse("s#a"));
    }

    [Fact(DisplayName = "快照-幻灯片、文本框、表格与备注")]
    public void Outline_FromPresentation()
    {
        var path = BuildPptx();
        var outline = Read(path).ToOutline();

        Assert.Single(outline.Slides);
        var slide = outline.Slides[0];
        Assert.Equal("s#0", slide.Anchor);
        Assert.Equal(2, slide.TextBoxes.Count);
        Assert.Equal("季度报告", slide.TextBoxes[0].Text);
        Assert.Equal(32, slide.TextBoxes[0].FontSize);
        Assert.True(slide.TextBoxes[0].Bold);
        Assert.Equal("s#0/tb#0", slide.TextBoxes[0].Anchor);

        Assert.Single(slide.Tables);
        Assert.Equal(2, slide.Tables[0].Rows);
        Assert.Equal("华东", slide.Tables[0].Cells[1][0]);
        Assert.Equal("汇报要点", slide.Notes);
        Assert.Single(slide.Pictures);

        using var parsed = System.Text.Json.JsonDocument.Parse(outline.ToJson());
        var root = parsed.RootElement;
        Assert.Equal(1, root.GetProperty("slideCount").GetInt32());
        Assert.Equal("s#0", root.GetProperty("slides")[0].GetProperty("anchor").GetString());
        Assert.Equal("季度报告", root.GetProperty("slides")[0].GetProperty("textBoxes")[0].GetProperty("text").GetString());
        Assert.Equal("华东", root.GetProperty("slides")[0].GetProperty("tables")[0].GetProperty("cells")[1][0].GetString());
    }

    [Fact(DisplayName = "补丁-改写文本表格备注并另存")]
    public void Patch_TextTableNotes_Save()
    {
        var source = BuildPptx();
        var output = TempPath("patched.pptx");
        var doc = Read(source);

        var patch = PptxPatch.ParseJson("""
        { "ops": [
            { "op": "set_shape_text", "anchor": "s#0/tb#0", "text": "Q3 季度报告" },
            { "op": "set_table_cell", "anchor": "s#0/tbl#0", "row": 1, "column": 1, "text": "150" },
            { "op": "set_notes", "anchor": "s#0", "text": "更新备注" },
            { "op": "add_slide", "at": 1, "title": "结论", "text": "保持增长" }
        ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));
        doc.Save(output, source);

        var doc2 = Read(output);
        Assert.Equal(2, doc2.Slides.Count);
        Assert.Equal("结论", doc2.Slides[1].TextBoxes[0].Text);

        var first = doc2.Slides[0];
        Assert.Equal("Q3 季度报告", first.TextBoxes[0].Text);
        Assert.Equal("150", first.Tables[0].Rows[1][1]);
        Assert.Equal("更新备注", first.Notes);
    }

    [Fact(DisplayName = "补丁-编排幻灯片顺序与删除")]
    public void Patch_SlideOrder()
    {
        var source = BuildPptx();
        var output = TempPath("order.pptx");
        var doc = Read(source);

        var patch = PptxPatch.ParseJson("""
        { "ops": [
            { "op": "add_slide", "title": "第二页" },
            { "op": "move_slide", "from": 1, "to": 0 }
        ] }
        """);
        Assert.True(patch.Apply(doc).Success);
        doc.Save(output, source);

        var doc2 = Read(output);
        Assert.Equal("第二页", doc2.Slides[0].TextBoxes[0].Text);

        var patch2 = PptxPatch.ParseJson("""
        { "ops": [ { "op": "delete_slide", "anchor": "s#0" } ] }
        """);
        Assert.True(patch2.Apply(doc2).Success);
        doc2.Save(output, source);

        var doc3 = Read(output);
        Assert.Single(doc3.Slides);
        Assert.Equal("季度报告", doc3.Slides[0].TextBoxes[0].Text);
    }

    [Fact(DisplayName = "补丁-失败不落盘")]
    public void Patch_Failure()
    {
        var source = BuildPptx();
        var doc = Read(source);

        var patch = PptxPatch.ParseJson("""
        { "ops": [
            { "op": "set_shape_text", "anchor": "s#0/tb#0", "text": "改动" },
            { "op": "set_shape_text", "anchor": "s#9/tb#0", "text": "越界" }
        ] }
        """);
        var result = patch.Apply(doc);

        Assert.False(result.Success);
        Assert.Single(result.Failed);
        Assert.Single(result.Applied);
    }

    [Fact(DisplayName = "补丁-替换图片")]
    public void Patch_ReplaceImage()
    {
        var source = BuildPptx();
        var output = TempPath("image.pptx");
        var doc = Read(source);
        var original = doc.Slides[0].Images[0].Data;

        var patch = PptxPatch.ParseJson("""
        { "ops": [ { "op": "replace_image", "anchor": "s#0/pic#0", "base64": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==", "extension": "png" } ] }
        """);
        Assert.True(patch.Apply(doc).Success);
        doc.Save(output, source);

        var doc2 = Read(output);
        Assert.Single(doc2.Slides[0].Images);
        Assert.Equal(original.Length, doc2.Slides[0].Images[0].Data.Length);
    }

    [Fact(DisplayName = "补丁-图表系列长度校验")]
    public void Patch_ChartSeriesLength()
    {
        var doc = new Presentation();
        var slide = new Slide();
        slide.Charts.Add(new Chart
        {
            ChartType = "bar",
            Categories = ["A", "B"],
            Series = { new ChartSeries { Name = "S1", Values = [1, 2, 3] } },
        });
        doc.Slides.Add(slide);

        var patch = PptxPatch.ParseJson("""
        { "ops": [ { "op": "update_chart", "anchor": "s#0/chart#0", "categories": ["A", "B"], "series": [ { "name": "S1", "values": [1, 2, 3] } ] } ] }
        """);
        var result = patch.Apply(doc);

        Assert.False(result.Success);
        Assert.Contains("不一致", result.Failed[0]);
        Assert.Equal(3, doc.Slides[0].Charts[0].Series[0].Values.Length);
    }

    [Fact(DisplayName = "补丁-克隆幻灯片")]
    public void Patch_CloneSlide()
    {
        var source = BuildPptx();
        var output = TempPath("clone.pptx");
        var doc = Read(source);

        var patch = PptxPatch.ParseJson("""
        { "ops": [ { "op": "clone_slide", "from": 0 } ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));
        Assert.Equal(2, doc.Slides.Count);

        doc.Save(output, source);
        var doc2 = Read(output);
        Assert.Equal(2, doc2.Slides.Count);
        Assert.Equal("季度报告", doc2.Slides[1].TextBoxes[0].Text);
    }
}
