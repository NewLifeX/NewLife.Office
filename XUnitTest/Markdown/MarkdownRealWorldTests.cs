using System.ComponentModel;
using System.IO;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 真实文档测试（Phase G）</summary>
/// <remarks>
/// 使用贴近日常使用的真实文档样例（GitHub 风格 README、中文技术文档）验证解析、
/// 文本提取与语义往返，发现仅靠程序化构造难以覆盖的解析缺陷。
/// </remarks>
public class MarkdownRealWorldTests
{
    private static String Fixture(String name) => Path.Combine(AppContext.BaseDirectory, "Markdown", "Fixtures", name);

    [Fact]
    [DisplayName("真实文档：GitHub 风格 README 结构解析")]
    public void Readme_Structure()
    {
        var doc = MarkdownDocument.ParseFile(Fixture("readme_style.md"));
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Heading);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.BulletList);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Table);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.CodeBlock);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.BlockQuote);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.ThematicBreak);
    }

    [Fact]
    [DisplayName("真实文档：README 语义往返保真")]
    public void Readme_RoundTrip()
    {
        var doc = MarkdownDocument.ParseFile(Fixture("readme_style.md"));
        var text1 = doc.ExtractText();

        var md = doc.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md);
        var text2 = doc2.ExtractText();

        Assert.Equal(text1, text2);
    }

    [Fact]
    [DisplayName("真实文档：README 关键内容提取")]
    public void Readme_KeyContent()
    {
        var doc = MarkdownDocument.ParseFile(Fixture("readme_style.md"));
        var text = doc.ExtractText() ?? "";
        Assert.Contains("示例项目", text);
        Assert.Contains("已完成任务", text);
        Assert.Contains("Hello, World!", text);
        Assert.Contains("数据表格", text);
    }

    [Fact]
    [DisplayName("真实文档：中文技术文档含脚注/有序列表/表格")]
    public void ChineseTech_Structure()
    {
        var doc = MarkdownDocument.ParseFile(Fixture("chinese_tech.md"));
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.FootnoteDefinition);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.OrderedList);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Table);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.BlockQuote);
    }

    [Fact]
    [DisplayName("真实文档：中文技术文档语义往返")]
    public void ChineseTech_RoundTrip()
    {
        var doc = MarkdownDocument.ParseFile(Fixture("chinese_tech.md"));
        var text1 = doc.ExtractText();

        var md = doc.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md);
        var text2 = doc2.ExtractText();

        Assert.Equal(text1, text2);
    }

    [Fact]
    [DisplayName("真实文档：Fixtures 目录逐文件往返回归（M4）")]
    public void AllFixtures_RoundTrip()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "Markdown", "Fixtures");
        Assert.True(Directory.Exists(dir), $"未找到 Fixtures 目录：{dir}");
        var files = Directory.GetFiles(dir, "*.md");
        Assert.True(files.Length > 0, $"Fixtures 目录下未找到 .md 文件：{dir}");

        var failures = new List<String>();
        foreach (var file in files)
        {
            try
            {
                // 语义往返：解析 → 写回 → 再解析，正文文本必须一致
                var doc1 = MarkdownDocument.ParseFile(file);
                var md1 = doc1.ToMarkdown();
                var doc2 = MarkdownDocument.Parse(md1);
                Assert.Equal(doc1.ExtractText(), doc2.ExtractText());

                // 序列化不动点：二次写回必须与首次一致，保证往返稳定
                var md2 = doc2.ToMarkdown();
                Assert.Equal(md1, md2);
            }
            catch (Exception ex)
            {
                var lines = ex.Message.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                failures.Add($"[{Path.GetFileName(file)}] {lines.FirstOrDefault()?.Trim()}");
            }
        }

        if (failures.Count > 0)
            Assert.Fail($"❌ {failures.Count}/{files.Length} 个真实文档往返失败: {String.Join(" | ", failures)}");
    }
}
