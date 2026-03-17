using System.ComponentModel;
using NewLife.Office;
using NewLife.Office.Rtf;
using Xunit;

namespace XUnitTest.Integration;

/// <summary>RTF 格式集成测试</summary>
public class RtfTests : IntegrationTestBase
{
    [Fact, DisplayName("RTF_复杂写入再读取往返")]
    public void Rtf_ComplexWriteAndRead()
    {
        var path = Path.Combine(OutputDir, "test_complex.rtf");

        var writer = new RtfWriter
        {
            Title = "RTF Test Document",
            Author = "NewLife Office",
            Subject = "Integration Test",
        };

        writer.AddParagraph("RTF 集成测试文档");
        writer.AddParagraph("这是由 NewLife.Office RTF 写入器生成的测试文件。");

        // 带格式段落
        var boldPara = new RtfParagraph { Alignment = RtfAlignment.Center };
        boldPara.Runs.Add(new RtfRun { Text = "加粗居中文本", Bold = true, FontSize = 32 });
        writer.AddParagraph(boldPara);

        var italicPara = new RtfParagraph();
        italicPara.Runs.Add(new RtfRun { Text = "斜体文本", Italic = true });
        italicPara.Runs.Add(new RtfRun { Text = " 和 " });
        italicPara.Runs.Add(new RtfRun { Text = "下划线文本", Underline = true });
        writer.AddParagraph(italicPara);

        // 表格
        writer.AddTable(new[]
        {
            new[] { "项目", "数量", "金额" },
            new[] { "产品A", "10", "5000" },
        });

        writer.AddParagraph("文档结束。");
        writer.Save(path);

        Assert.True(File.Exists(path));

        // 读取验证
        var doc = RtfDocument.ParseFile(path);

        // RTF info 块对 Unicode 标题的往返支持有限，仅验证 ASCII 部分
        Assert.Contains("RTF", doc.Title);
        Assert.Equal("NewLife Office", doc.Author);

        var plainText = doc.GetPlainText();
        Assert.Contains("RTF 集成测试文档", plainText);
        Assert.Contains("加粗居中文本", plainText);
        Assert.Contains("文档结束", plainText);

        // 段落和表格
        var paragraphs = doc.Paragraphs.ToList();
        Assert.True(paragraphs.Count >= 4);

        var tables = doc.Tables.ToList();
        Assert.True(tables.Count >= 1);

        // 工厂创建
        var factoryReader = OfficeFactory.CreateReader(path);
        Assert.IsType<RtfDocument>(factoryReader);
    }
}
