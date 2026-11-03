using NewLife.Office.Pdf;
using Xunit;

namespace XUnitTest.Pdf;

/// <summary>PDF AI 接口测试：结构快照、搜索定位、对象级补丁与重新写回</summary>
public class PdfAiTests
{
    #region 工具
    private static String TempPath(String name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "nloffice-ai", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    private static String BuildPdf()
    {
        var path = TempPath("sample.pdf");
        using var writer = new PdfWriter();
        writer.BeginPage();
        writer.AppendLine("Page One");
        writer.AppendLine("Hello World");
        writer.AddTextField("name", 72, 700, 200, 20, "旧值");
        writer.EndPage();
        writer.AddBookmark("First");
        writer.BeginPage();
        writer.AppendLine("Page Two");
        writer.EndPage();
        writer.Save(path);
        return path;
    }
    #endregion

    [Fact(DisplayName = "快照-页数元数据书签表单首行")]
    public void Snapshot_FromPdf()
    {
        var path = BuildPdf();
        using var reader = new PdfReader(path);
        var snapshot = PdfSnapshot.From(reader);

        Assert.Equal(2, snapshot.PageCount);
        Assert.Equal(2, snapshot.Pages.Count);
        Assert.Equal(0, snapshot.Pages[0].Index);
        Assert.Contains("Page One", snapshot.Pages[0].FirstLine);
        Assert.Contains("Page Two", snapshot.Pages[1].FirstLine);
        Assert.Single(snapshot.Bookmarks);
        Assert.Equal("First", snapshot.Bookmarks[0].Title);
        Assert.Single(snapshot.FormFields);
        Assert.Equal("name", snapshot.FormFields[0].Name);
        Assert.False(snapshot.Encrypted);

        using var parsed = System.Text.Json.JsonDocument.Parse(snapshot.ToJson());
        Assert.Equal(2, parsed.RootElement.GetProperty("pageCount").GetInt32());
        Assert.Equal("p#0", parsed.RootElement.GetProperty("pages")[0].GetProperty("anchor").GetString());
    }

    [Fact(DisplayName = "搜索-返回页码与矩形")]
    public void Search_ReturnsRects()
    {
        var path = BuildPdf();
        using var reader = new PdfReader(path);

        var hits = reader.SearchText("Page");
        Assert.Equal(2, hits.Count);
        Assert.Equal(1, hits[0].Page);
        Assert.Equal(2, hits[1].Page);
        Assert.True(hits[0].Width > 0);
        Assert.True(hits[0].X >= 0);

        var lines = reader.ReadPageLines(0);
        Assert.Contains(lines, e => e.Text.Contains("Page One"));
    }

    [Fact(DisplayName = "对象级-重写后仍可读取")]
    public void ObjectDocument_RoundTrip()
    {
        var path = BuildPdf();
        var bytes = File.ReadAllBytes(path);

        var doc = PdfObjectDocument.Load(bytes);
        Assert.Equal(2, doc.GetPageObjectNumbers().Count);

        var rewritten = doc.ToArray();
        var path2 = TempPath("rewritten.pdf");
        File.WriteAllBytes(path2, rewritten);

        using var reader = new PdfReader(path2);
        Assert.Equal(2, reader.GetPageCount());
        Assert.Contains("Page One", reader.ExtractText());
        Assert.Contains("Page Two", reader.ExtractText());
    }

    [Fact(DisplayName = "补丁-元数据高亮旋转与文字叠加")]
    public void Patch_MetadataAnnotationRotateText()
    {
        var path = BuildPdf();
        var doc = PdfObjectDocument.Load(path);

        var patch = PdfPatch.ParseJson("""
        { "ops": [
            { "op": "set_metadata", "title": "年度报告", "author": "AI" },
            { "op": "add_highlight", "page": 1, "x": 72, "y": 760, "width": 100, "height": 12, "text": "重点" },
            { "op": "add_note", "page": 2, "x": 80, "y": 780, "text": "备注" },
            { "op": "rotate_page", "pages": "1", "rotation": 90 },
            { "op": "add_text", "page": 1, "text": "CONFIDENTIAL", "x": 100, "y": 500, "fontSize": 14 }
        ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));

        var output = TempPath("patched.pdf");
        doc.Save(output);

        using var reader = new PdfReader(output);
        Assert.Equal("年度报告", reader.ReadMetadata().Title);
        Assert.Contains("CONFIDENTIAL", reader.ExtractText());

        var annotations = reader.ReadAnnotations(0);
        Assert.Contains(annotations, e => e.Type == PdfAnnotationType.Highlight);
        Assert.Equal(1, reader.ReadAnnotations(1).Count);
    }

    [Fact(DisplayName = "补丁-删页与移动页")]
    public void Patch_PageOrder()
    {
        var path = BuildPdf();
        var doc = PdfObjectDocument.Load(path);

        var move = PdfPatch.ParseJson("""
        { "ops": [ { "op": "move_page", "from": 2, "to": 1 } ] }
        """);
        Assert.True(move.Apply(doc).Success);

        var output = TempPath("moved.pdf");
        doc.Save(output);
        using (var reader = new PdfReader(output))
        {
            Assert.Equal(2, reader.GetPageCount());
            Assert.Contains("Page Two", String.Join(" ", reader.GetPageTextBlocks(0).Select(e => e.Text)));
        }

        var delete = PdfPatch.ParseJson("""
        { "ops": [ { "op": "delete_pages", "pages": "1" } ] }
        """);
        Assert.True(delete.Apply(doc).Success);
        doc.Save(output);
        using var reader2 = new PdfReader(output);
        Assert.Equal(1, reader2.GetPageCount());
        Assert.Contains("Page One", reader2.ExtractText());
    }

    [Fact(DisplayName = "补丁-全页水印")]
    public void Patch_Watermark()
    {
        var path = BuildPdf();
        var doc = PdfObjectDocument.Load(path);

        var patch = PdfPatch.ParseJson("""
        { "ops": [ { "op": "add_watermark", "text": "DRAFT", "fontSize": 40 } ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));

        var output = TempPath("watermark.pdf");
        doc.Save(output);
        using var reader = new PdfReader(output);
        Assert.Contains("DRAFT", reader.ExtractText());
    }

    [Fact(DisplayName = "补丁-表单填充")]
    public void Patch_FillForm()
    {
        var path = BuildPdf();
        var doc = PdfObjectDocument.Load(path);

        var patch = PdfPatch.ParseJson("""
        { "ops": [ { "op": "fill_form", "name": "name", "value": "张三" } ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));

        var output = TempPath("form.pdf");
        doc.Save(output);
        using var reader = new PdfReader(output);
        var form = reader.ReadFormFields();
        Assert.NotNull(form);
        Assert.Contains(form!.Fields, e => e.FullName == "name" && e.Value == "张三");
    }

    [Fact(DisplayName = "补丁-合并与失败不落盘")]
    public void Patch_MergeAndFailure()
    {
        var path = BuildPdf();
        var doc = PdfObjectDocument.Load(path);

        var merge = PdfPatch.ParseJson($$"""
        { "ops": [ { "op": "merge", "path": {{System.Text.Json.JsonSerializer.Serialize(path)}} } ] }
        """);
        Assert.True(merge.Apply(doc).Success);
        Assert.Equal(4, doc.GetPageObjectNumbers().Count);

        var bad = PdfPatch.ParseJson("""
        { "ops": [
            { "op": "rotate_page", "pages": "1", "rotation": 180 },
            { "op": "delete_pages", "pages": "99" }
        ] }
        """);
        var failed = bad.Apply(doc);
        Assert.False(failed.Success);
        Assert.Single(failed.Applied);
        Assert.Single(failed.Failed);
    }
}
