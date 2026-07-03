using System;
using System.ComponentModel;
using System.Linq;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>MD05 高级解析扩展测试</summary>
[Trait("Category", "Markdown")]
public class MarkdownAdvancedTests
{
    #region MD05-01 AutoLinks

    [Fact]
    [DisplayName("裸 URL 自动链接 https")]
    public void AutoLink_Https()
    {
        var doc = MarkdownDocument.Parse("Visit https://example.com/page for info.");

        var p = doc.Blocks[0];
        Assert.Contains(p.Inlines, i => i.Type == MarkdownInlineType.Link);

        var md = doc.ToMarkdown();
        Assert.Contains("https://example.com/page", md);
    }

    [Fact]
    [DisplayName("裸 URL 自动链接 http")]
    public void AutoLink_Http()
    {
        var doc = MarkdownDocument.Parse("See http://test.org");

        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    [Fact]
    [DisplayName("裸 www 前缀自动链接")]
    public void AutoLink_Www()
    {
        var doc = MarkdownDocument.Parse("Go to www.example.com now.");

        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    [Fact]
    [DisplayName("裸 Email 自动链接")]
    public void AutoLink_Email()
    {
        var doc = MarkdownDocument.Parse("Contact user@example.com for help.");

        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    #endregion

    #region MD05-02 YAML Front Matter

    [Fact]
    [DisplayName("YAML Front Matter 解析")]
    public void FrontMatter_Parse()
    {
        var md = "---\ntitle: My Document\nauthor: Stone\ndate: 2026-07-03\n---\n\n# Hello";
        var doc = MarkdownDocument.Parse(md);

        Assert.Equal(3, doc.FrontMatter.Count);
        Assert.Equal("My Document", doc.FrontMatter["title"]);
        Assert.Equal("Stone", doc.FrontMatter["author"]);
        Assert.Equal("2026-07-03", doc.FrontMatter["date"]);

        // 确保 FrontMatter 之后的内容正常解析
        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Heading, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("YAML Front Matter 引号值")]
    public void FrontMatter_QuotedValues()
    {
        var md = "---\ntitle: \"Hello World\"\ndesc: 'Simple desc'\n---\n\nContent";
        var doc = MarkdownDocument.Parse(md);

        Assert.Equal("Hello World", doc.FrontMatter["title"]);
        Assert.Equal("Simple desc", doc.FrontMatter["desc"]);
    }

    [Fact]
    [DisplayName("无 FrontMatter 的文档")]
    public void FrontMatter_None()
    {
        var doc = MarkdownDocument.Parse("# Just a heading");

        Assert.Empty(doc.FrontMatter);
        Assert.Single(doc.Blocks);
    }

    #endregion

    #region MD05-03 Footnotes

    [Fact]
    [DisplayName("脚注定义和引用")]
    public void Footnote_DefinitionAndReference()
    {
        var md = "Text with a footnote[^1].\n\n[^1]: This is the footnote.";
        var doc = MarkdownDocument.Parse(md);

        // 应该有段落+脚注定义 = 2个块
        Assert.True(doc.Blocks.Count >= 2);

        var fnDef = doc.Blocks.OfType<FootnoteDefinitionBlock>().FirstOrDefault();
        Assert.NotNull(fnDef);
        Assert.Equal("1", fnDef!.Id);
    }

    [Fact]
    [DisplayName("脚注引用生成标记")]
    public void Footnote_RefInOutput()
    {
        var md = "See footnote[^note1].\n\n[^note1]: Important note here.";
        var doc = MarkdownDocument.Parse(md);

        var output = doc.ToMarkdown();
        Assert.Contains("[^note1]", output);
    }

    #endregion

    #region MD05-04 Math Formulas

    [Fact]
    [DisplayName("数学公式块 $$")]
    public void Math_Block()
    {
        var md = "$$\nx = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}\n$$";
        var doc = MarkdownDocument.Parse(md);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.MathBlock, doc.Blocks[0].Type);

        var mathBlock = (MathBlock)doc.Blocks[0];
        Assert.Contains("\\frac", mathBlock.Content);

        var output = doc.ToMarkdown();
        Assert.Contains("$$", output);
    }

    [Fact]
    [DisplayName("行内数学公式 $")]
    public void Math_Inline()
    {
        var doc = MarkdownDocument.Parse("The formula $E=mc^2$ is famous.");

        var p = doc.Blocks[0];
        Assert.Contains(p.Inlines, i => i.Type == MarkdownInlineType.MathInline);

        var output = doc.ToMarkdown();
        Assert.Contains("$E=mc^2$", output);
    }

    #endregion

    #region 往返测试

    [Fact]
    [DisplayName("YAML FrontMatter 往返")]
    public void RoundTrip_FrontMatter()
    {
        var original = "---\ntitle: Test\nauthor: Stone\n---\n\n# Hello World\n\nContent here.";
        var doc = MarkdownDocument.Parse(original);

        Assert.Equal("Test", doc.FrontMatter["title"]);
        Assert.Equal("Hello World", ((HeadingBlock)doc.Blocks[0]).GetPlainText().Trim());
    }

    [Fact]
    [DisplayName("Math 公式往返")]
    public void RoundTrip_Math()
    {
        var original = "$$\nx = 1 + 2\n$$\n\nInline $y = 3$ here.";
        var doc = MarkdownDocument.Parse(original);

        var output = doc.ToMarkdown();
        Assert.Contains("$$", output);
        Assert.Contains("$y = 3$", output);
    }

    #endregion

    #region 兼容性

    [Fact]
    [DisplayName("新功能不影响已有解析")]
    public void Compatibility_ExistingParsingStillWorks()
    {
        var md = "# Title\n\n**Bold** and *italic*\n\n- List item\n\n```csharp\nvar x = 1;\n```";
        var doc = MarkdownDocument.Parse(md);

        Assert.Equal(4, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Heading, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[2].Type);
        Assert.Equal(MarkdownBlockType.CodeBlock, doc.Blocks[3].Type);
    }

    #endregion
}
