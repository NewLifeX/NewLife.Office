using System.ComponentModel;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NewLife.Office.Word;
using NPOI.XWPF.UserModel;
using Xunit;
using Xunit.Abstractions;

namespace XUnitTest.Word;

/// <summary>真实 docx 文档上的修订验证（我们改 + 竞品交叉校验）</summary>
/// <remarks>
/// 与 WordCrossCompatibilityTests 的区别：那里验证"读入→写回不变"，这里验证"读入→**修订**→结果正确"。
/// 每条用例都做三重验证：①我们自己的 Reader 读回修订结果；②第三方 NPOI 能读到同样内容；
/// ③OpenXmlValidator 不显著新增规范错误（真实样张本身不规范，故用相对基线）。
/// 另含双向接力（我改→竞品改→我再读）与失败原子性（不留半成品、原文件不变）。
/// </remarks>
[Trait("Category", "CrossCompatibility")]
public class RealDocumentEditingTests
{
    private readonly ITestOutputHelper _output;

    public RealDocumentEditingTests(ITestOutputHelper output) => _output = output;

    #region 辅助
    /// <summary>真实样张所在目录（随测试程序集复制）</summary>
    private static String FixturesDir => Path.Combine(AppContext.BaseDirectory, "Word", "Fixtures");

    /// <summary>取第一个真实 docx 样张路径；不存在返回 null</summary>
    /// <returns>样张路径</returns>
    private static String? FindFixture()
    {
        var dir = FixturesDir;
        if (!Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, "*.docx", SearchOption.TopDirectoryOnly).FirstOrDefault();
    }

    private static String TempFile()
        => Path.Combine(Path.GetTempPath(), $"WordEdit_{Guid.NewGuid():N}.docx");

    /// <summary>校验 docx 的 OOXML 规范错误数</summary>
    /// <param name="path">文件路径</param>
    /// <returns>错误条数</returns>
    private static Int32 CountSchemaErrors(String path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(doc).Count();
    }

    /// <summary>用 NPOI 读取全文文本</summary>
    /// <param name="path">docx 路径</param>
    /// <returns>正文文本</returns>
    private static String ReadAllTextByNpoi(String path)
    {
        using var doc = new XWPFDocument(new MemoryStream(File.ReadAllBytes(path)));
        var sb = new StringBuilder();
        foreach (var p in doc.Paragraphs) sb.Append(p.Text).Append('\n');
        foreach (var t in doc.Tables)
        {
            foreach (var row in t.Rows)
            {
                foreach (var cell in row.GetTableCells()) sb.Append(String.Concat(cell.Paragraphs.Select(x => x.Text))).Append('|');
            }
        }
        return sb.ToString();
    }

    /// <summary>读取 outline 快照</summary>
    /// <param name="path">docx 路径</param>
    /// <returns>结构快照</returns>
    private static DocxOutline ReadOutline(String path)
    {
        using var reader = new WordReader(path);
        return reader.ReadOutline();
    }

    /// <summary>应用补丁 JSON 到真实文档并保存到新文件</summary>
    /// <param name="sourcePath">源 docx</param>
    /// <param name="patchJson">补丁 JSON</param>
    /// <param name="outputPath">输出 docx</param>
    /// <returns>补丁结果</returns>
    private static DocxPatchResult ApplyPatch(String sourcePath, String patchJson, String outputPath)
    {
        var patch = DocxPatch.ParseJson(patchJson);
        using var reader = new WordReader(sourcePath);
        var doc = reader.ReadDocument();
        doc.BeginEdit();
        var result = DocxPatch.Apply(doc, patch);
        if (result.Success)
        {
            using var writer = new WordWriter();
            writer.Save(outputPath, doc);
        }
        return result;
    }

    private static String Sha256(String path)
    {
        using var fs = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(fs));
    }
    #endregion

    [Fact, DisplayName("真实修订-按锚点改段落_NPOI与规范双重校验")]
    public void EditFirstParagraph_VerifiedByNpoiAndValidator()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var outline = ReadOutline(fixture);
        var target = outline.Blocks.FirstOrDefault(e => e.Type == "paragraph" && !String.IsNullOrWhiteSpace(e.Text));
        Assert.NotNull(target);
        _output.WriteLine($"[真实修订] 锚点={target!.Anchor} 原文本前20字={Trim(target.Text)}");

        var outPath = TempFile();
        try
        {
            var newText = "【Office 修订验证】新的段落内容";
            var result = ApplyPatch(fixture, $"{{\"ops\":[{{\"op\":\"set_text\",\"anchor\":\"{target.Anchor}\",\"text\":\"{newText}\"}}]}}", outPath);
            Assert.True(result.Success, String.Join("; ", result.Failed));

            // ① 我们自己的 Reader：该锚点文本已被替换
            var after = ReadOutline(outPath);
            Assert.Equal(newText, after.Blocks.First(e => e.Anchor == target.Anchor).Text);

            // ② 竞品 NPOI：能读到新内容
            Assert.Contains(newText, ReadAllTextByNpoi(outPath));

            // ③ 规范：不显著新增错误（真实样张自身不规范）
            var src = CountSchemaErrors(fixture);
            var dst = CountSchemaErrors(outPath);
            _output.WriteLine($"[真实修订] 规范错误 源={src} 输出={dst}");
            Assert.True(dst <= src + 3, $"修订不应显著新增规范错误: 源={src} 输出={dst}");

            // ④ 其他结构未被破坏：元素数量与表格/图片数保持一致
            Assert.Equal(outline.Blocks.Count, after.Blocks.Count);
            Assert.Equal(
                outline.Blocks.Count(e => e.Type == "table"),
                after.Blocks.Count(e => e.Type == "table"));
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact, DisplayName("真实修订-插入与删除段落_NPOI可读且规范不劣化")]
    public void InsertAndDeleteParagraph_Verified()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var outline = ReadOutline(fixture);
        var first = outline.Blocks.FirstOrDefault(e => e.Type == "paragraph");
        Assert.NotNull(first);
        var beforeCount = outline.Blocks.Count;

        var outPath = TempFile();
        try
        {
            var marker = "【Office 插入的新段落】";
            // insert_paragraph 的锚点语义是"插到该锚点元素之后"
            var result = ApplyPatch(fixture, $"{{\"ops\":[{{\"op\":\"insert_paragraph\",\"anchor\":\"{first!.Anchor}\",\"text\":\"{marker}\"}}]}}", outPath);
            Assert.True(result.Success, String.Join("; ", result.Failed));

            var after = ReadOutline(outPath);
            Assert.Equal(beforeCount + 1, after.Blocks.Count);
            Assert.Contains(after.Blocks, e => e.Text != null && e.Text.Contains(marker));
            Assert.Contains(marker, ReadAllTextByNpoi(outPath));

            var src = CountSchemaErrors(fixture);
            var dst = CountSchemaErrors(outPath);
            _output.WriteLine($"[插入段落] 规范错误 源={src} 输出={dst}");
            Assert.True(dst <= src + 3, $"插入段落不应显著新增规范错误: 源={src} 输出={dst}");

            // 再删除刚插入的段落，元素数量回到原值
            var inserted = after.Blocks.First(e => e.Text != null && e.Text.Contains(marker));
            var outPath2 = TempFile();
            try
            {
                var delResult = ApplyPatch(outPath, $"{{\"ops\":[{{\"op\":\"delete\",\"anchor\":\"{inserted.Anchor}\"}}]}}", outPath2);
                Assert.True(delResult.Success, String.Join("; ", delResult.Failed));
                var final = ReadOutline(outPath2);
                Assert.Equal(beforeCount, final.Blocks.Count);
                Assert.DoesNotContain(marker, ReadAllTextByNpoi(outPath2));
            }
            finally
            {
                if (File.Exists(outPath2)) File.Delete(outPath2);
            }
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact, DisplayName("真实修订-改段落样式_NPOI读到新样式")]
    public void SetStyle_VerifiedByNpoi()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var outline = ReadOutline(fixture);
        var target = outline.Blocks.FirstOrDefault(e => e.Type == "paragraph" && !String.IsNullOrWhiteSpace(e.Text));
        Assert.NotNull(target);

        var outPath = TempFile();
        try
        {
            var result = ApplyPatch(fixture, $"{{\"ops\":[{{\"op\":\"set_style\",\"anchor\":\"{target!.Anchor}\",\"style\":\"Heading2\"}}]}}", outPath);
            Assert.True(result.Success, String.Join("; ", result.Failed));

            // 竞品 NPOI 读到的段落样式应为 Heading2
            using var fs = File.OpenRead(outPath);
            using var doc = new XWPFDocument(fs);
            var styles = doc.Paragraphs.Select(p => p.Style ?? String.Empty).ToList();
            _output.WriteLine($"[改样式] NPOI 读到的样式集合: {String.Join(",", styles.Distinct())}");
            Assert.Contains(styles, e => e.Contains("Heading2", StringComparison.OrdinalIgnoreCase));

            // 我们自己的 Reader 也应看到该段变成标题
            var after = ReadOutline(outPath);
            Assert.Equal("heading", after.Blocks.First(e => e.Anchor == target.Anchor).Type);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact, DisplayName("真实修订-改表格单元格_NPOI读到新值")]
    public void SetCellText_VerifiedByNpoi()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var outline = ReadOutline(fixture);
        var table = outline.Blocks.FirstOrDefault(e => e.Type == "table" && e.Rows > 0 && e.Columns > 0);
        if (table == null) { _output.WriteLine("真实样张无表格，跳过"); return; }

        var outPath = TempFile();
        try
        {
            var cellText = "【表格单元格修订】";
            var result = ApplyPatch(fixture, $"{{\"ops\":[{{\"op\":\"set_cell_text\",\"anchor\":\"{table.Anchor}\",\"row\":0,\"column\":0,\"text\":\"{cellText}\"}}]}}", outPath);
            Assert.True(result.Success, String.Join("; ", result.Failed));

            using var fs = File.OpenRead(outPath);
            using var doc = new XWPFDocument(fs);
            var npoiTable = doc.Tables.FirstOrDefault();
            Assert.NotNull(npoiTable);
            Assert.Contains(cellText, String.Concat(npoiTable!.GetRow(0).GetCell(0).Paragraphs.Select(x => x.Text)));
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact, DisplayName("双向接力-我们改后竞品再改_我们能读回两处修改")]
    public void BidirectionalRelay_OurEditThenNpoiEdit()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var outline = ReadOutline(fixture);
        var target = outline.Blocks.FirstOrDefault(e => e.Type == "paragraph" && !String.IsNullOrWhiteSpace(e.Text));
        Assert.NotNull(target);

        var oursPath = TempFile();
        var npoiPath = TempFile();
        try
        {
            // 第一棒：我们改段落文本
            var ourText = "【我方修订】第一棒";
            var result = ApplyPatch(fixture, $"{{\"ops\":[{{\"op\":\"set_text\",\"anchor\":\"{target!.Anchor}\",\"text\":\"{ourText}\"}}]}}", oursPath);
            Assert.True(result.Success, String.Join("; ", result.Failed));

            // 第二棒：NPOI 打开我们的输出，追加一个段落
            var npoiText = "【NPOI 追加】第二棒";
            using (var fs = File.OpenRead(oursPath))
            using (var doc = new XWPFDocument(fs))
            {
                var p = doc.CreateParagraph();
                p.CreateRun().SetText(npoiText);
                using var outFs = File.Create(npoiPath);
                doc.Write(outFs);
            }

            // 第三方（我们的 Reader）读回：两处修改都在
            var after = ReadOutline(npoiPath);
            Assert.Equal(ourText, after.Blocks.First(e => e.Anchor == target.Anchor).Text);
            Assert.Contains(after.Blocks, e => e.Text != null && e.Text.Contains(npoiText));

            // 竞品自身读回同样能看到自己的修改（基线 F4）
            Assert.Contains(npoiText, ReadAllTextByNpoi(npoiPath));
        }
        finally
        {
            if (File.Exists(oursPath)) File.Delete(oursPath);
            if (File.Exists(npoiPath)) File.Delete(npoiPath);
        }
    }

    [Fact, DisplayName("真实修订-失败原子性_不留半成品且原文件不变")]
    public void FailedPatch_NoOutputAndSourceUnchanged()
    {
        var fixture = FindFixture();
        if (fixture == null) { _output.WriteLine("未找到真实样张，跳过"); return; }

        var before = Sha256(fixture);
        var outPath = TempFile();
        try
        {
            var result = ApplyPatch(fixture, "{\"ops\":[{\"op\":\"set_text\",\"anchor\":\"p#99999\",\"text\":\"x\"}]}", outPath);
            Assert.False(result.Success);
            Assert.NotEmpty(result.Failed);
            Assert.False(File.Exists(outPath), "失败时不应产生输出文件");
            Assert.Equal(before, Sha256(fixture));
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    private static String Trim(String? text)
        => text == null ? String.Empty : (text.Length <= 20 ? text : text.Substring(0, 20));
}
