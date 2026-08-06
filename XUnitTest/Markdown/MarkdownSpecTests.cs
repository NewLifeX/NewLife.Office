using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>CommonMark spec 核心子集测试（Phase G）</summary>
/// <remarks>
/// 从 CommonMark spec.txt 抽取核心用例，覆盖标题/段落/换行/转义/实体/代码段/强调/
/// 链接/图片/列表/引用/分隔线/缩进与围栏代码/HTML 块/硬软换行等核心语法。
/// </remarks>
public class MarkdownSpecTests
{
    private static MarkdownDocument P(String md) => MarkdownDocument.Parse(md);

    #region 标题

    [Fact]
    [DisplayName("Spec：ATX 标题各级别与闭合井号")]
    public void Spec_AtxHeadings()
    {
        Assert.Equal(1, ((HeadingBlock)P("# h1").Blocks[0]).Level);
        Assert.Equal(6, ((HeadingBlock)P("###### h6").Blocks[0]).Level);
        // 闭合井号被移除
        Assert.Equal("foo", P("## foo ##").Blocks[0].GetPlainText());
        // 无空格不算标题
        Assert.Equal(MarkdownBlockType.Paragraph, P("#5 bolt").Blocks[0].Type);
    }

    [Fact]
    [DisplayName("Spec：Setext 标题")]
    public void Spec_SetextHeadings()
    {
        var h1 = P("Foo *bar*\n=========").Blocks[0];
        Assert.Equal(MarkdownBlockType.Heading, h1.Type);
        Assert.Equal(1, ((HeadingBlock)h1).Level);
        var h2 = P("Foo\n---").Blocks[0];
        Assert.Equal(MarkdownBlockType.Heading, h2.Type);
        Assert.Equal(2, ((HeadingBlock)h2).Level);
    }

    #endregion

    #region 段落与换行

    [Fact]
    [DisplayName("Spec：段落续行合并")]
    public void Spec_ParagraphJoining()
    {
        var doc = P("aaa\nbbb\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal("aaa bbb", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Spec：CRLF 换行统一")]
    public void Spec_CrlfNormalized()
    {
        var doc = P("aaa\r\nbbb\r\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal("aaa bbb", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Spec：硬换行（行尾两空格）与软换行")]
    public void Spec_HardAndSoftBreak()
    {
        var doc = P("aaa  \nbbb\n");
        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.HardBreak);
        var doc2 = P("aaa\nbbb\n");
        Assert.Contains(doc2.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.SoftBreak);
    }

    [Fact]
    [DisplayName("Spec：反斜杠转义")]
    public void Spec_BackslashEscapes()
    {
        var doc = P("\\*not emph\\*\n");
        Assert.Equal("*not emph*", doc.Blocks[0].GetPlainText());
    }

    #endregion

    #region 代码

    [Fact]
    [DisplayName("Spec：行内代码基本与多反引号")]
    public void Spec_CodeSpans()
    {
        Assert.Equal("code", P("`code`").Blocks[0].Inlines[0].Text);
        // 多反引号包裹含反引号内容
        var doc = P("`` foo ` bar ``");
        Assert.Equal(MarkdownInlineType.Code, doc.Blocks[0].Inlines[0].Type);
        Assert.Equal("foo ` bar", doc.Blocks[0].Inlines[0].Text);
    }

    [Fact]
    [DisplayName("Spec：缩进代码块")]
    public void Spec_IndentedCode()
    {
        var doc = P("    a simple\n      indented code block\n");
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("a simple\n  indented code block", code.RawText);
    }

    [Fact]
    [DisplayName("Spec：围栏代码块信息串与闭合")]
    public void Spec_FencedCode()
    {
        var doc = P("```ruby\ndef foo(x)\n  return 3\nend\n```\n");
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("ruby", code.Language);
        Assert.Equal("def foo(x)\n  return 3\nend", code.RawText);
    }

    #endregion

    #region 强调

    [Fact]
    [DisplayName("Spec：强调三形态")]
    public void Spec_EmphasisForms()
    {
        Assert.Equal(MarkdownInlineType.Emphasis, P("*foo*").Blocks[0].Inlines[0].Type);
        Assert.Equal(MarkdownInlineType.Strong, P("**foo**").Blocks[0].Inlines[0].Type);
        // CommonMark 规范：***foo*** → <em><strong>foo</strong></em>（嵌套结构）
        var se = P("***foo***").Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Emphasis, se.Type);
        Assert.Single(se.Children);
        Assert.Equal(MarkdownInlineType.Strong, se.Children[0].Type);
        Assert.Equal("foo", se.Children[0].GetPlainText());
        // 下划线
        Assert.Equal(MarkdownInlineType.Emphasis, P("_foo_").Blocks[0].Inlines[0].Type);
    }

    [Fact]
    [DisplayName("Spec：词语内下划线不构成强调")]
    public void Spec_IntrawordUnderscore()
    {
        var doc = P("foo_bar_baz\n");
        Assert.Equal(MarkdownInlineType.Text, doc.Blocks[0].Inlines[0].Type);
        Assert.Equal("foo_bar_baz", doc.Blocks[0].GetPlainText());
    }

    #endregion

    #region 链接与图片

    [Fact]
    [DisplayName("Spec：行内链接与标题")]
    public void Spec_InlineLink()
    {
        var doc = P("[link](/uri \"title\")\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("/uri", link.Href);
        Assert.Equal("title", link.Title);
    }

    [Fact]
    [DisplayName("Spec：自动链接")]
    public void Spec_AutoLink()
    {
        var doc = P("<http://foo.bar.baz>\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("http://foo.bar.baz", link.Href);
    }

    [Fact]
    [DisplayName("Spec：图片")]
    public void Spec_Image()
    {
        var doc = P("![foo](/url \"title\")\n");
        var img = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Image, img.Type);
        Assert.Equal("/url", img.Href);
        Assert.Equal("foo", img.Alt);
    }

    #endregion

    #region 列表

    [Fact]
    [DisplayName("Spec：紧列表与松列表")]
    public void Spec_TightAndLooseLists()
    {
        var tight = P("- a\n- b\n");
        Assert.Equal(2, tight.Blocks[0].Children.Count);

        // 松列表：空行分隔
        var loose = P("- a\n\n- b\n");
        Assert.Equal(2, loose.Blocks[0].Children.Count);
    }

    [Fact]
    [DisplayName("Spec：有序列表起始序号")]
    public void Spec_OrderedStart()
    {
        var doc = P("3. a\n4. b\n");
        var ol = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
        Assert.Equal(3, ol.OrderedStart);
    }

    [Fact]
    [DisplayName("Spec：列表项包含多段与嵌套")]
    public void Spec_ListBlocks()
    {
        var doc = P("- a\n\n  b\n- c\n");
        var item = doc.Blocks[0].Children[0];
        Assert.Equal(2, item.Children.Count);
    }

    #endregion

    #region 引用

    [Fact]
    [DisplayName("Spec：引用块与嵌套")]
    public void Spec_BlockQuote()
    {
        var doc = P("> # Foo\n> bar\n> baz\n");
        var quote = Assert.IsType<BlockQuoteBlock>(doc.Blocks[0]);
        Assert.Equal(MarkdownBlockType.Heading, quote.Children[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, quote.Children[1].Type);

        // 嵌套引用
        var nested = P("> > 嵌套\n");
        var q2 = Assert.IsType<BlockQuoteBlock>(nested.Blocks[0]);
        Assert.Equal(MarkdownBlockType.BlockQuote, q2.Children[0].Type);
    }

    #endregion

    #region 分隔线

    [Fact]
    [DisplayName("Spec：分隔线三种符号")]
    public void Spec_ThematicBreaks()
    {
        Assert.Equal(MarkdownBlockType.ThematicBreak, P("***\n").Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.ThematicBreak, P("---\n").Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.ThematicBreak, P("___\n").Blocks[0].Type);
    }

    #endregion

    #region HTML

    [Fact]
    [DisplayName("Spec：HTML 块与行内 HTML")]
    public void Spec_Html()
    {
        Assert.Equal(MarkdownBlockType.HtmlBlock, P("<div>\n</div>\n").Blocks[0].Type);
        var doc = P("x <span>y</span> z\n");
        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.RawHtml);
    }

    #endregion

    #region 引用链接

    [Fact]
    [DisplayName("Spec：引用链接与折叠")]
    public void Spec_ReferenceLinks()
    {
        var doc = P("[foo][bar]\n\n[bar]: /url \"title\"\n");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("/url", link.Href);

        var doc2 = P("[foo][]\n\n[foo]: /url\n");
        Assert.Equal(MarkdownInlineType.Link, doc2.Blocks[0].Inlines[0].Type);
    }

    #endregion
}
