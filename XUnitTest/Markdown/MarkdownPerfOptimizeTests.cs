using System;
using System.ComponentModel;
using System.Linq;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>MD15 解析分配优化测试：引用块快速路径语义等价 + 引用预扫描行首过滤边界</summary>
[Trait("Category", "Markdown")]
public class MarkdownPerfOptimizeTests
{
    #region MD15-01 引用块快速路径（单行普通内容直接行内解析）

    [Fact]
    [DisplayName("单行引用快速路径：块结构等价")]
    public void BlockQuote_SingleLine_Structure()
    {
        var doc = MarkdownDocument.Parse("> hello world");

        var q = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        var p = Assert.IsType<ParagraphBlock>(q.Children[0]);
        Assert.Equal("hello world", p.GetPlainText());
        Assert.Equal("> hello world\n", doc.ToMarkdown());
    }

    [Fact]
    [DisplayName("单行引用快速路径：行内格式解析")]
    public void BlockQuote_SingleLine_InlineFormat()
    {
        var doc = MarkdownDocument.Parse("> **bold** and *italic* and `code`");

        var q = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        var p = Assert.IsType<ParagraphBlock>(q.Children[0]);
        Assert.Contains(p.Inlines, i => i.Type == MarkdownInlineType.Strong);
        Assert.Contains(p.Inlines, i => i.Type == MarkdownInlineType.Emphasis);
        Assert.Contains(p.Inlines, i => i.Type == MarkdownInlineType.Code);
    }

    [Fact]
    [DisplayName("单行引用快速路径：ToMarkdown 往返一致")]
    public void BlockQuote_SingleLine_Roundtrip()
    {
        const String md = "> quote **with** format\n";
        var doc = MarkdownDocument.Parse(md);
        Assert.Equal(md, doc.ToMarkdown());
    }

    [Fact]
    [DisplayName("多行引用仍走子解析器：段落合并")]
    public void BlockQuote_MultiLine_MergeParagraph()
    {
        var doc = MarkdownDocument.Parse("> line one\n> line two");

        var q = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        var p = Assert.IsType<ParagraphBlock>(q.Children[0]);
        Assert.Equal("line one line two", p.GetPlainText());
    }

    [Fact]
    [DisplayName("引用内引用定义行不进快速路径")]
    public void BlockQuote_ReferenceDef_NotFastPath()
    {
        // 引用定义行（[ 开头）被 IsBlockStart 拦截，必须走子解析器路径；其定义写入共享 _refs，
        // 正文 [ref] 经共享引用解析为链接。若误走快速路径，引用定义会变成普通文本且 [ref] 无法解析
        var doc = MarkdownDocument.Parse("> [ref]: https://newlifex.com\n\nSee [ref]");
        // 引用块内引用定义行不被当作普通段落文本输出
        var q = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        Assert.DoesNotContain(q.Children, c => c.GetPlainText().Contains("[ref]:"));
        // 正文链接解析成功（共享 _refs 生效）
        Assert.Contains(doc.Blocks[1].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    [Fact]
    [DisplayName("空引用行保持原行为")]
    public void BlockQuote_EmptyLine()
    {
        var doc = MarkdownDocument.Parse(">\n");
        Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
    }

    [Fact]
    [DisplayName("嵌套引用快速路径安全")]
    public void BlockQuote_Nested()
    {
        var doc = MarkdownDocument.Parse("> > nested quote");

        var outer = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        var inner = Assert.IsType<BlockQuoteBlock>(outer.Children[0]);
        Assert.Equal("nested quote", inner.Children[0].GetPlainText());
    }

    #endregion

    #region MD15-02 引用预扫描行首过滤

    [Fact]
    [DisplayName("普通引用定义仍识别")]
    public void PreScan_PlainDefinition()
    {
        var doc = MarkdownDocument.Parse("[ref]: https://newlifex.com \"title\"\n\nSee [ref]");
        Assert.Equal("https://newlifex.com", doc.References["ref"]);
        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    [Fact]
    [DisplayName("3 空格缩进引用定义仍识别（CommonMark 允许）")]
    public void PreScan_ThreeSpaceIndent()
    {
        var doc = MarkdownDocument.Parse("   [ref]: https://newlifex.com\n\nSee [ref]");
        Assert.Equal("https://newlifex.com", doc.References["ref"]);
    }

    [Fact]
    [DisplayName("4 空格缩进引用定义不识别（缩进代码块）")]
    public void PreScan_FourSpaceIndent_NotReference()
    {
        var doc = MarkdownDocument.Parse("    [ref]: https://newlifex.com\n\nSee [ref]");
        // 4 空格是代码块，不提取引用定义
        Assert.False(doc.References.ContainsKey("ref"));
        // [ref] 未定义 → 按字面文本
        Assert.DoesNotContain(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.Link);
    }

    [Fact]
    [DisplayName("制表符缩进引用定义（1 制表=4 列）不识别")]
    public void PreScan_TabIndent_NotReference()
    {
        var doc = MarkdownDocument.Parse("\t[ref]: https://newlifex.com\n\nSee [ref]");
        Assert.False(doc.References.ContainsKey("ref"));
    }

    #endregion
}
