using System.Globalization;
using System.Text.Json;
using NewLife;
using NewLife.Office;
using NewLife.Office.Excel;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Word;
using Xunit;
using WDocument = NewLife.Office.Word.Document;

namespace XUnitTest.Common;

/// <summary>跨格式 AI 工具层测试：格式路由、分页续读、搜索与补丁原子性</summary>
public class OfficeToolTests
{
    #region 工具
    private static String TempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "officetool_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static String BuildDocx(String dir, String text)
    {
        var path = Path.Combine(dir, "sample.docx");
        var doc = new WDocument();
        doc.AddParagraph(text);
        doc.AddParagraph("第二段");
        using var writer = new WordWriter();
        writer.Save(path, doc);
        return path;
    }

    private static String BuildXlsx(String dir)
    {
        var path = Path.Combine(dir, "sample.xlsx");
        var doc = new ExcelDocument();
        var sheet = new Worksheet { Name = "数据" };
        sheet.Rows.Add(["姓名", "部门"]);
        for (var i = 0; i < 9; i++) sheet.Rows.Add([$"成员{i}", "研发"]);
        doc.Worksheets.Add(sheet);
        doc.Save(path);
        return path;
    }

    private static String BuildPptx(String dir)
    {
        var path = Path.Combine(dir, "sample.pptx");
        using var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.AddTextBox(0, "标题", 1, 1, 10, 2, 28, true);
        writer.Save(path);
        return path;
    }

    private static String BuildPdf(String dir)
    {
        var path = Path.Combine(dir, "sample.pdf");
        using var writer = new PdfWriter();
        writer.BeginPage();
        writer.AppendLine("Hello PDF");
        writer.EndPage();
        writer.Save(path);
        return path;
    }
    #endregion

    [Fact(DisplayName = "工具-四格式结构快照路由")]
    public void ReadOutline_RoutesAllFormats()
    {
        var dir = TempDir();
        var docx = OfficeDocumentTool.ReadOutline(BuildDocx(dir, "长文本"));
        var xlsx = OfficeDocumentTool.ReadOutline(BuildXlsx(dir));
        var pptx = OfficeDocumentTool.ReadOutline(BuildPptx(dir));
        var pdf = OfficeDocumentTool.ReadOutline(BuildPdf(dir));

        Assert.True(docx.Success);
        Assert.Equal("docx", docx.Format);
        Assert.Contains("\"blocks\"", docx.Content);

        Assert.True(xlsx.Success);
        Assert.Equal("xlsx", xlsx.Format);
        Assert.Contains("\"sheets\"", xlsx.Content);
        Assert.Contains("数据", xlsx.Content);

        Assert.True(pptx.Success);
        Assert.Contains("\"slides\"", pptx.Content);
        Assert.Contains("s#0/tb#0", pptx.Content);

        Assert.True(pdf.Success);
        Assert.Contains("\"pages\"", pdf.Content);
    }

    [Fact(DisplayName = "工具-文本分段续读可拼回全文")]
    public void ReadText_Paging()
    {
        var dir = TempDir();
        var text = new String('中', 120) + "结尾";
        var path = BuildDocx(dir, text);

        var full = OfficeDocumentTool.ReadText(path, 0, 0);
        Assert.True(full.Success);
        Assert.False(full.Truncated);

        var first = OfficeDocumentTool.ReadText(path, 50, 0);
        Assert.True(first.Truncated);
        Assert.Equal(50, first.NextCursor);
        Assert.Equal(50, first.Content!.Length);

        var second = OfficeDocumentTool.ReadText(path, 1000, first.NextCursor);
        Assert.False(second.Truncated);
        Assert.Equal(full.Content, first.Content + second.Content);
    }

    [Fact(DisplayName = "工具-区域分页读取")]
    public void ReadRange_Paging()
    {
        var dir = TempDir();
        var path = BuildXlsx(dir);

        var page1 = OfficeDocumentTool.ReadRange(path, "数据!A1:B10", 3, 0);
        Assert.True(page1.Success);
        Assert.True(page1.Truncated);
        Assert.Equal(3, page1.NextCursor);
        Assert.Contains("成员0", page1.Content);

        var page2 = OfficeDocumentTool.ReadRange(path, "数据!A1:B10", 3, 3);
        Assert.True(page2.Success);
        Assert.Contains("成员2", page2.Content);
        Assert.DoesNotContain("成员0", page2.Content);

        var tail = OfficeDocumentTool.ReadRange(path, "数据!A1:B10", 10, 6);
        Assert.False(tail.Truncated);
        Assert.Contains("成员8", tail.Content);
    }

    [Fact(DisplayName = "工具-工作簿与PDF搜索")]
    public void Search_WorkbookAndPdf()
    {
        var dir = TempDir();
        var xlsx = OfficeDocumentTool.Search(BuildXlsx(dir), "成员3");
        Assert.True(xlsx.Success);
        Assert.Contains("数据!A5", xlsx.Content);

        var pdf = OfficeDocumentTool.Search(BuildPdf(dir), "Hello");
        Assert.True(pdf.Success);
        Assert.Contains("\"page\": 1", pdf.Content);
        Assert.Contains("Hello", pdf.Content);
    }

    [Fact(DisplayName = "工具-补丁成功写出并可读回")]
    public void ApplyPatch_Success()
    {
        var dir = TempDir();
        var path = BuildXlsx(dir);
        var output = Path.Combine(dir, "out.xlsx");

        var result = OfficeDocumentTool.ApplyPatch(path,
            """{ "ops": [ { "op": "set_cells", "range": "数据!A12", "values": [["新增"]] } ] }""", output);

        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(output));
        Assert.NotEmpty(result.Warnings);

        var check = OfficeDocumentTool.ReadRange(output, "数据!A12:A12");
        Assert.Contains("新增", check.Content);
    }

    [Fact(DisplayName = "工具-补丁失败不写文件并返回错误码")]
    public void ApplyPatch_Failure()
    {
        var dir = TempDir();
        var path = BuildXlsx(dir);
        var output = Path.Combine(dir, "out.xlsx");

        var result = OfficeDocumentTool.ApplyPatch(path,
            """{ "ops": [ { "op": "set_cells", "range": "不存在!A1", "values": [["x"]] } ] }""", output);

        Assert.False(result.Success);
        Assert.True(result.Code == OfficeToolErrorCode.AnchorMissing, $"{result.Code}: {result.Message}");
        Assert.False(File.Exists(output));
    }

    [Fact(DisplayName = "工具-不支持格式与加密文档的错误码")]
    public void ReadOutline_UnsupportedAndEncrypted()
    {
        var dir = TempDir();
        var unsupported = Path.Combine(dir, "file.bin");
        File.WriteAllBytes(unsupported, [1, 2, 3, 4]);
        var result = OfficeDocumentTool.ReadOutline(unsupported);
        Assert.False(result.Success);
        Assert.Equal(OfficeToolErrorCode.Unsupported, result.Code);

        var docx = BuildDocx(dir, "加密");
        var encrypted = Path.Combine(dir, "encrypted.docx");
        File.WriteAllBytes(encrypted, WordEncryptor.Encrypt(File.ReadAllBytes(docx), "123456"));
        var encryptedResult = OfficeDocumentTool.ReadOutline(encrypted);
        Assert.False(encryptedResult.Success);
        Assert.Equal(OfficeToolErrorCode.Encrypted, encryptedResult.Code);
    }

    [Fact(DisplayName = "工具-加密 PDF 返回 Encrypted 错误码")]
    public void ReadOutline_EncryptedPdf()
    {
        var dir = TempDir();
        var path = Path.Combine(dir, "secret.pdf");
        using (var writer = new PdfWriter())
        {
            writer.BeginPage();
            writer.AppendLine("Secret");
            writer.EndPage();
            writer.UserPassword = "123456";
            writer.Save(path);
        }

        var outline = OfficeDocumentTool.ReadOutline(path);
        Assert.False(outline.Success);
        Assert.Equal(OfficeToolErrorCode.Encrypted, outline.Code);

        var text = OfficeDocumentTool.ReadText(path);
        Assert.False(text.Success);
        Assert.Equal(OfficeToolErrorCode.Encrypted, text.Code);
    }

    [Fact(DisplayName = "工具-PDF 补丁跨格式路由")]
    public void ApplyPatch_Pdf()
    {
        var dir = TempDir();
        var path = BuildPdf(dir);
        var output = Path.Combine(dir, "out.pdf");

        var result = OfficeDocumentTool.ApplyPatch(path,
            """{ "ops": [ { "op": "set_metadata", "title": "工具层" } ] }""", output);

        Assert.True(result.Success, result.Message);
        using var reader = new PdfReader(output);
        Assert.Equal("工具层", reader.ReadMetadata().Title);
    }

    [Fact(DisplayName = "工具-正则搜索超时返回资源限制")]
    public void Search_RegexTimeout()
    {
        var dir = TempDir();
        var path = Path.Combine(dir, "regex.xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteHeader("Sheet1", ["值"]);
            w.WriteRow("Sheet1", new Object?[] { "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa!" });
        }

        var result = OfficeDocumentTool.Search(path, "(a+)+$", regex: true);

        Assert.False(result.Success);
        Assert.Equal(OfficeToolErrorCode.ResourceLimit, result.Code);
    }

    [Fact(DisplayName = "工具-补丁 JSON 嵌套过深被拒绝")]
    public void ParseJson_DepthLimit()
    {
        var deep = new String('[', 400) + new String(']', 400);
        var ex = Assert.Throws<FormatException>(() => DocxPatch.ParseJson("{ \"ops\": " + deep + " }"));
        Assert.Contains("256", ex.Message);
    }

    [Fact(DisplayName = "工具-xls 走 BIFF 读取器：快照/区域/搜索/转 xlsx")]
    public void LegacyXls_ReadAndConvert()
    {
        var dir = TempDir();
        var path = Path.Combine(dir, "legacy.xls");
        using (var w = new BiffWriter())
        {
            w.WriteHeader(["名称", "数量"]);
            w.WriteRow(["苹果", 5]);
            w.Save(path);
        }

        var outline = OfficeDocumentTool.ReadOutline(path);
        Assert.True(outline.Success, outline.Message);
        Assert.Contains("名称", outline.Content);

        var range = OfficeDocumentTool.ReadRange(path, "A1:B2");
        Assert.True(range.Success, range.Message);
        Assert.Contains("苹果", range.Content);

        var hits = OfficeDocumentTool.Search(path, "苹果");
        Assert.True(hits.Success, hits.Message);
        Assert.Contains("A2", hits.Content);

        var text = OfficeDocumentTool.ReadText(path);
        Assert.NotEqual(OfficeToolErrorCode.Encrypted, text.Code);

        var output = Path.Combine(dir, "converted.xlsx");
        var patched = OfficeDocumentTool.ApplyPatch(path, """{ "ops": [ { "op": "set_cells", "range": "B2", "values": [[9]] } ] }""", output);
        Assert.True(patched.Success, patched.Message);
        using var reader = new ExcelReader(output);
        Assert.Equal(9d, Convert.ToDouble(reader.ReadExcel().Worksheets[0].Rows[1][1], CultureInfo.InvariantCulture));
    }

    [Fact(DisplayName = "工具-跳行工作表按真实行号读取与搜索")]
    public void SparseRows_UseRealRowNumbers()
    {
        var dir = TempDir();
        var path = Path.Combine(dir, "sparse.xlsx");
        var doc = new ExcelDocument();
        var sheet = new Worksheet { Name = "S", ActualRowNumbers = [0, 4] };
        sheet.Rows.Add(["R1"]);
        sheet.Rows.Add(["R5"]);
        doc.Worksheets.Add(sheet);
        doc.Save(path);

        var range = OfficeDocumentTool.ReadRange(path, "S!A1:A10");
        Assert.True(range.Success, range.Message);
        using (var json = JsonDocument.Parse(range.Content!))
        {
            var rows = json.RootElement.GetProperty("rows");
            Assert.Equal(5, rows.GetArrayLength());
            Assert.Equal("R1", rows[0][0].GetString());
            Assert.Equal("R5", rows[4][0].GetString());
        }

        var hits = OfficeDocumentTool.Search(path, "R5");
        Assert.True(hits.Success, hits.Message);
        Assert.Contains("A5", hits.Content);

        var outline = OfficeDocumentTool.ReadOutline(path);
        using var outlineJson = JsonDocument.Parse(outline.Content!);
        Assert.Equal(5, outlineJson.RootElement.GetProperty("sheets")[0].GetProperty("rows").GetInt32());
    }

    [Fact(DisplayName = "工具-超范围页码补丁快速拒绝")]
    public void PdfPatch_HugePageRangeRejected()
    {
        var dir = TempDir();
        var path = BuildPdf(dir);
        var output = Path.Combine(dir, "out.pdf");

        var result = OfficeDocumentTool.ApplyPatch(path,
            """{ "ops": [ { "op": "rotate_page", "pages": "5-2000000000", "rotation": 90 } ] }""", output);

        Assert.False(result.Success);
        Assert.Equal(OfficeToolErrorCode.AnchorMissing, result.Code);
        Assert.False(File.Exists(output));
    }

    [Fact(DisplayName = "工具-就地补丁（输出=输入）docx 与 xlsx 均可写入")]
    public void ApplyPatch_InPlace()
    {
        var dir = TempDir();
        var docx = BuildDocx(dir, "原地");
        var r1 = OfficeDocumentTool.ApplyPatch(docx, """{ "ops": [ { "op": "set_text", "anchor": "p#0", "text": "就地修改" } ] }""", docx);
        Assert.True(r1.Success, r1.Message);
        Assert.Contains("就地修改", OfficeDocumentTool.ReadText(docx).Content);

        var xlsx = Path.Combine(dir, "inplace.xlsx");
        using (var w = new ExcelWriter(xlsx))
        {
            w.WriteHeader("Sheet1", ["值"]);
            w.WriteRow("Sheet1", new Object?[] { "A" });
        }
        var r2 = OfficeDocumentTool.ApplyPatch(xlsx, """{ "ops": [ { "op": "set_cells", "range": "Sheet1!A2", "values": [["B"]] } ] }""", xlsx);
        Assert.True(r2.Success, r2.Message);
        Assert.Contains("B", OfficeDocumentTool.ReadText(xlsx).Content);
    }

    [Fact(DisplayName = "工具-未知工作表统一报锚点类错误")]
    public void ReadRange_MissingSheet_AnchorMissing()
    {
        var dir = TempDir();
        var path = BuildXlsx(dir);

        var result = OfficeDocumentTool.ReadRange(path, "NoSuch!A1:B2");

        Assert.False(result.Success);
        Assert.Equal(OfficeToolErrorCode.AnchorMissing, result.Code);
    }

    [Fact(DisplayName = "PDF-整型数值精确保留")]
    public void PdfNumber_LargeIntegerExact()
    {
        Assert.Equal("20000000", new PdfNumber(20000000).ToString());
        Assert.Equal("3.5", new PdfNumber(3.5f).ToString());
    }

    [Fact(DisplayName = "工具-补丁失败分类不依赖消息语言")]    public void ApplyPatch_AnchorCodeIsCultureInvariant()
    {
        var dir = TempDir();
        var path = BuildDocx(dir, "锚点分类");
        var output = Path.Combine(dir, "out.docx");
        const String patch = """{ "ops": [ { "op": "set_text", "anchor": "p#99", "text": "x" } ] }""";

        var saved = Localization.Culture;
        try
        {
            Localization.Culture = new CultureInfo("en-US");
            var en = OfficeDocumentTool.ApplyPatch(path, patch, output);
            Assert.Equal(OfficeToolErrorCode.AnchorMissing, en.Code);

            Localization.Culture = new CultureInfo("zh-Hans");
            var zh = OfficeDocumentTool.ApplyPatch(path, patch, output);
            Assert.Equal(OfficeToolErrorCode.AnchorMissing, zh.Code);
        }
        finally
        {
            Localization.Culture = saved;
        }

        Assert.False(File.Exists(output));
    }
}
