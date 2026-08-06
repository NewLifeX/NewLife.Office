using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 引用链接、HTML 实体解码与行内 HTML 测试（Phase C）</summary>
/// <remarks>
/// 覆盖 CommonMark 引用链接三形态（[text][id] / [text][] / [text]）、定义在后使用在前、
/// HTML 实体解码（具名+数字）、行内 HTML 标签与注释。
/// </remarks>
public class MarkdownReferenceTests
{
    #region 引用链接

    [Fact]
    [DisplayName("引用链接：完整引用 [text][id]")]
    public void Reference_Full()
    {
        var doc = MarkdownDocument.Parse("使用 [文档][ref] 查阅\n\n[ref]: https://newlifex.com\n");
        Assert.Equal(1, doc.Blocks.Count);
        var link = Assert.IsType<MarkdownInline>(doc.Blocks[0].Inlines[1]);
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://newlifex.com", link.Href);
        Assert.Equal("文档", link.GetPlainText());
        Assert.True(doc.References.ContainsKey("ref"));
    }

    [Fact]
    [DisplayName("引用链接：折叠引用 [text][]")]
    public void Reference_Collapsed()
    {
        var doc = MarkdownDocument.Parse("[文档][] 见下\n\n[文档]: https://example.com/a\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://example.com/a", link.Href);
    }

    [Fact]
    [DisplayName("引用链接：快捷引用 [text]（仅当已定义）")]
    public void Reference_Shortcut()
    {
        var doc = MarkdownDocument.Parse("[官网] 直达\n\n[官网]: https://newlifex.com\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://newlifex.com", link.Href);

        // 未定义时保持文本
        var doc2 = MarkdownDocument.Parse("[未定义] 保持文本\n");
        var inline = doc2.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Text, inline.Type);
    }

    [Fact]
    [DisplayName("引用链接：定义在后使用在前")]
    public void Reference_DefinedLater()
    {
        var doc = MarkdownDocument.Parse("查看 [目标][t] 内容\n\n[t]: https://example.com/x\n");
        var link = doc.Blocks[0].Inlines[1];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://example.com/x", link.Href);
    }

    [Fact]
    [DisplayName("引用链接：定义标题被解析")]
    public void Reference_WithTitle()
    {
        var doc = MarkdownDocument.Parse("[目标][t]\n\n[t]: https://example.com \"标题文本\"\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal("https://example.com", link.Href);
    }

    [Fact]
    [DisplayName("引用链接：不误伤脚注定义")]
    public void Reference_DoesNotBreakFootnotes()
    {
        var doc = MarkdownDocument.Parse("正文[^1]\n\n[^1]: 脚注内容\n");
        // 应存在脚注定义块
        var hasFootnote = false;
        foreach (var block in doc.Blocks)
        {
            if (block.Type == MarkdownBlockType.FootnoteDefinition) hasFootnote = true;
        }
        Assert.True(hasFootnote);
    }

    [Fact]
    [DisplayName("引用链接：序列化输出解析结果与引用定义")]
    public void Reference_Serialized()
    {
        var doc = MarkdownDocument.Parse("查看 [目标][t]\n\n[t]: https://example.com\n");
        var md = doc.ToMarkdown();
        // 引用在解析时已解析为行内链接
        Assert.Contains("[目标](https://example.com)", md);
        // 引用定义保留输出
        Assert.Contains("[t]: https://example.com", md);
    }

    [Fact]
    [DisplayName("引用链接：行内链接优先于引用")]
    public void Reference_InlineLinkTakesPrecedence()
    {
        var doc = MarkdownDocument.Parse("[目标](https://inline.com)\n\n[目标]: https://ref.com\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://inline.com", link.Href);
    }

    #endregion

    #region HTML 实体解码

    [Fact]
    [DisplayName("实体：具名实体解码")]
    public void Entity_Named()
    {
        var doc = MarkdownDocument.Parse("A &amp; B &lt; C &gt; D");
        var text = doc.Blocks[0].GetPlainText();
        Assert.Equal("A & B < C > D", text);
    }

    [Fact]
    [DisplayName("实体：数字字符引用解码")]
    public void Entity_Numeric()
    {
        var doc = MarkdownDocument.Parse("&#65;&#x42;");
        Assert.Equal("AB", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("实体：未识别实体保持原文")]
    public void Entity_UnknownPreserved()
    {
        var doc = MarkdownDocument.Parse("a &unknown; b");
        Assert.Equal("a &unknown; b", doc.Blocks[0].GetPlainText());
    }

    #endregion

    #region 行内 HTML

    [Fact]
    [DisplayName("行内 HTML：开始标签")]
    public void InlineHtml_OpenTag()
    {
        var doc = MarkdownDocument.Parse("文本 <span class=\"x\">高亮</span> 结束");
        var inlines = doc.Blocks[0].Inlines;
        // Text, RawHtml, Text, RawHtml, Text
        Assert.Equal(5, inlines.Count);
        Assert.Equal(MarkdownInlineType.RawHtml, inlines[1].Type);
        Assert.Equal("<span class=\"x\">", inlines[1].Text);
        Assert.Equal(MarkdownInlineType.RawHtml, inlines[3].Type);
        Assert.Equal("</span>", inlines[3].Text);
    }

    [Fact]
    [DisplayName("行内 HTML：注释")]
    public void InlineHtml_Comment()
    {
        var doc = MarkdownDocument.Parse("a <!-- 注释 --> b");
        Assert.Equal(3, doc.Blocks[0].Inlines.Count);
        Assert.Equal(MarkdownInlineType.RawHtml, doc.Blocks[0].Inlines[1].Type);
        Assert.Equal("<!-- 注释 -->", doc.Blocks[0].Inlines[1].Text);
    }

    [Fact]
    [DisplayName("行内 HTML：非标签保持文本")]
    public void InlineHtml_NotTag_KeptAsText()
    {
        var doc = MarkdownDocument.Parse("a < b > c");
        Assert.Equal(1, doc.Blocks[0].Inlines.Count);
        Assert.Equal(MarkdownInlineType.Text, doc.Blocks[0].Inlines[0].Type);
        Assert.Equal("a < b > c", doc.Blocks[0].Inlines[0].Text);
    }

    #endregion
}
