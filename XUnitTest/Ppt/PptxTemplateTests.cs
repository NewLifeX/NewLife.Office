using System.ComponentModel;
using NewLife.Office.Ppt;
using Xunit;

namespace XUnitTest.Ppt;

/// <summary>PptxTemplate 模板填充（占位符/表格行扩展）单元测试</summary>
public class PptxTemplateTests
{
    #region 辅助
    private sealed class ReportModel
    {
        public String Title { get; set; } = "";
        public String Author { get; set; } = "";
    }

    /// <summary>创建包含占位符的模板文件，返回路径</summary>
    private static String CreateTemplate(String dir)
    {
        var path = Path.Combine(dir, "template.pptx");
        using var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.AddTextBox(0, "标题：{{Title}}", 1, 1, 20, 2);
        writer.AddTextBox(0, "作者：{{Author}}", 1, 3, 20, 2);
        writer.Save(path);
        return path;
    }

    private static String CreateTableTemplate(String dir)
    {
        var path = Path.Combine(dir, "table_template.pptx");
        using var writer = new PptxWriter();
        writer.AddSlide(0);
        writer.AddTable(0,
        [
            ["姓名", "数值"],
            ["{{#Items}}", ""],
            ["{{Name}}", "{{Value}}"],
            ["{{/Items}}", ""],
        ]);
        writer.Save(path);
        return path;
    }

    private static String ReadAllText(String path)
    {
        using var reader = new PptxReader(path);
        var slides = reader.ReadAllSlides().ToList();
        return String.Join("\n", slides.SelectMany(s => s.TextBoxes).Select(t => t.Text));
    }

    private static void SafeDelete(params String[] paths)
    {
        foreach (var path in paths)
            if (File.Exists(path)) File.Delete(path);
    }
    #endregion

    #region 占位符填充
    [Fact, DisplayName("PptxTemplate_字典占位符替换")]
    public void Fill_Dictionary_ReplacesPlaceholders()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PptxTemplate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var template = CreateTemplate(dir);
        var output = Path.Combine(dir, "out.pptx");
        try
        {
            var tpl = new PptxTemplate(template);
            Assert.Equal(Path.GetFullPath(template), tpl.TemplatePath);

            tpl.Fill(output, new Dictionary<String, Object?>
            {
                ["Title"] = "季度报告",
                ["Author"] = "新生命",
            });

            var text = ReadAllText(output);
            Assert.Contains("标题：季度报告", text);
            Assert.Contains("作者：新生命", text);
            Assert.DoesNotContain("{{", text);
        }
        finally
        {
            SafeDelete(template, output);
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("PptxTemplate_对象模型属性填充")]
    public void Fill_Model_ReplacesPlaceholders()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PptxTemplate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var template = CreateTemplate(dir);
        var output = Path.Combine(dir, "out_model.pptx");
        try
        {
            new PptxTemplate(template).Fill(output, new ReportModel { Title = "月报", Author = "AI" });

            var text = ReadAllText(output);
            Assert.Contains("标题：月报", text);
            Assert.Contains("作者：AI", text);
        }
        finally
        {
            SafeDelete(template, output);
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("PptxTemplate_特殊字符转义且空值替换为空串")]
    public void Fill_SpecialCharacters_AreEscaped()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PptxTemplate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var template = CreateTemplate(dir);
        var output = Path.Combine(dir, "out_escape.pptx");
        try
        {
            new PptxTemplate(template).Fill(output, new Dictionary<String, Object?>
            {
                ["Title"] = "A&B <C> \"D\"",
                ["Author"] = null,
            });

            var text = ReadAllText(output);
            Assert.Contains("标题：A&B <C> \"D\"", text);
            Assert.Contains("作者：", text);
            Assert.DoesNotContain("&amp;", text);
        }
        finally
        {
            SafeDelete(template, output);
            Directory.Delete(dir, true);
        }
    }

    [Fact, DisplayName("PptxTemplate_填充到流")]
    public void Fill_Stream_WritesValidPackage()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PptxTemplate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var template = CreateTemplate(dir);
        try
        {
            using var ms = new MemoryStream();
            new PptxTemplate(template).Fill(ms, new Dictionary<String, Object?> { ["Title"] = "流", ["Author"] = "X" });

            Assert.True(ms.Length > 0);
            ms.Position = 0;
            using var reader = new PptxReader(ms);
            var text = String.Join("\n", reader.ReadAllSlides().SelectMany(s => s.TextBoxes).Select(t => t.Text));
            Assert.Contains("标题：流", text);
        }
        finally
        {
            SafeDelete(template);
            Directory.Delete(dir, true);
        }
    }
    #endregion

    #region 表格行扩展
    [Fact, DisplayName("PptxTemplate_FillTable按列表展开模板行")]
    public void FillTable_ExpandsTemplateRows()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PptxTemplate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var template = CreateTableTemplate(dir);
        var output = Path.Combine(dir, "out_table.pptx");
        try
        {
            var lists = new Dictionary<String, IEnumerable<IDictionary<String, Object?>>>
            {
                ["Items"] =
                [
                    new Dictionary<String, Object?> { ["Name"] = "甲", ["Value"] = "10" },
                    new Dictionary<String, Object?> { ["Name"] = "乙", ["Value"] = "20" },
                ],
            };
            new PptxTemplate(template).FillTable(output, new Dictionary<String, Object?>(), lists);

            using var reader = new PptxReader(output);
            var table = Assert.Single(Assert.Single(reader.ReadAllSlides().ToList()).Tables);

            Assert.Equal(3, table.Rows.Count);
            Assert.Equal("姓名", table.Rows[0][0]);
            Assert.Equal("甲", table.Rows[1][0]);
            Assert.Equal("10", table.Rows[1][1]);
            Assert.Equal("乙", table.Rows[2][0]);
            Assert.Equal("20", table.Rows[2][1]);
            Assert.DoesNotContain(table.Rows, r => r.Any(c => c.Contains("{{", StringComparison.Ordinal)));
        }
        finally
        {
            SafeDelete(template, output);
            Directory.Delete(dir, true);
        }
    }
    #endregion
}
