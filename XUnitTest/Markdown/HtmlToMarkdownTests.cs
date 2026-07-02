using System;
using System.ComponentModel;
using System.Linq;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>HTML → Markdown 反向转换测试</summary>
[Trait("Category", "Markdown")]
public class HtmlToMarkdownTests
{
    #region 标题

    [Fact]
    [DisplayName("H1-H6 标题转换")]
    public void Convert_H1ToH6()
    {
        var html = "<h1>Title 1</h1><h2>Title 2</h2><h3>Title 3</h3><h4>Title 4</h4><h5>Title 5</h5><h6>Title 6</h6>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Equal(6, doc.Blocks.Count);
        Assert.All(doc.Blocks, b => Assert.Equal(MarkdownBlockType.Heading, b.Type));
        Assert.Equal(1, ((HeadingBlock)doc.Blocks[0]).Level);
        Assert.Equal(6, ((HeadingBlock)doc.Blocks[5]).Level);
    }

    [Fact]
    [DisplayName("标题内联格式（粗体/斜体）")]
    public void Convert_Heading_WithInlineFormat()
    {
        var html = "<h2>Hello <strong>World</strong> and <em>Universe</em></h2>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        var h = (HeadingBlock)doc.Blocks[0];
        Assert.Equal(2, h.Level);

        var md = doc.ToMarkdown();
        Assert.Contains("**World**", md);
        Assert.Contains("*Universe*", md);
    }

    #endregion

    #region 段落

    [Fact]
    [DisplayName("段落基本转换")]
    public void Convert_Paragraph_Basic()
    {
        var html = "<p>Hello World</p>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
        Assert.Contains("Hello World", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("多个段落转换")]
    public void Convert_MultipleParagraphs()
    {
        var html = "<p>First paragraph.</p><p>Second paragraph.</p>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Equal(2, doc.Blocks.Count);
        Assert.All(doc.Blocks, b => Assert.Equal(MarkdownBlockType.Paragraph, b.Type));
    }

    [Fact]
    [DisplayName("段落内联格式（粗体/斜体/删除线）")]
    public void Convert_Paragraph_WithInlineFormat()
    {
        var html = "<p>This is <strong>bold</strong>, <em>italic</em>, and <del>deleted</del> text.</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("**bold**", md);
        Assert.Contains("*italic*", md);
        Assert.Contains("~~deleted~~", md);
    }

    [Fact]
    [DisplayName("段落包含链接")]
    public void Convert_Paragraph_WithLink()
    {
        var html = "<p>Visit <a href=\"https://example.com\" title=\"Example\">our site</a> for more.</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("[our site](https://example.com \"Example\")", md);
    }

    [Fact]
    [DisplayName("段落包含图片")]
    public void Convert_Paragraph_WithImage()
    {
        var html = "<p>Image: <img src=\"https://example.com/photo.png\" alt=\"Photo\" title=\"My Photo\" /></p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("![Photo](https://example.com/photo.png \"My Photo\")", md);
    }

    [Fact]
    [DisplayName("换行标签 br")]
    public void Convert_Br_Tag()
    {
        var html = "<p>Line 1<br>Line 2</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("Line 1", md);
        Assert.Contains("Line 2", md);
    }

    #endregion

    #region 代码块

    [Fact]
    [DisplayName("代码块含语言标识")]
    public void Convert_Pre_CodeBlock_WithLanguage()
    {
        var html = "<pre><code class=\"language-csharp\">var x = 1;\nConsole.WriteLine(x);</code></pre>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        var cb = (CodeBlock)doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.CodeBlock, cb.Type);
        Assert.Equal("csharp", cb.Language);

        var md = doc.ToMarkdown();
        Assert.Contains("```csharp", md);
        Assert.Contains("var x = 1;", md);
    }

    [Fact]
    [DisplayName("代码块无语言标识")]
    public void Convert_Pre_CodeBlock_NoLanguage()
    {
        var html = "<pre><code>plain text code</code></pre>";
        var doc = MarkdownDocument.FromHtml(html);

        var cb = (CodeBlock)doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.CodeBlock, cb.Type);
        Assert.Equal("", cb.Language);
    }

    #endregion

    #region 行内代码

    [Fact]
    [DisplayName("行内代码标签")]
    public void Convert_InlineCode()
    {
        var html = "<p>Use the <code>ToString()</code> method.</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("`ToString()`", md);
    }

    #endregion

    #region 引用块

    [Fact]
    [DisplayName("引用块转换")]
    public void Convert_BlockQuote()
    {
        var html = "<blockquote><p>This is a quote.</p><p>Second line.</p></blockquote>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.BlockQuote, doc.Blocks[0].Type);

        var md = doc.ToMarkdown();
        Assert.Contains(">", md);
        Assert.Contains("This is a quote", md);
    }

    #endregion

    #region 列表

    [Fact]
    [DisplayName("无序列表转换")]
    public void Convert_UnorderedList()
    {
        var html = "<ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[0].Type);
        Assert.Equal(3, doc.Blocks[0].Children.Count);

        var md = doc.ToMarkdown();
        Assert.Contains("- Item 1", md);
        Assert.Contains("- Item 2", md);
        Assert.Contains("- Item 3", md);
    }

    [Fact]
    [DisplayName("有序列表转换")]
    public void Convert_OrderedList()
    {
        var html = "<ol><li>First</li><li>Second</li><li>Third</li></ol>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.OrderedList, doc.Blocks[0].Type);

        var md = doc.ToMarkdown();
        Assert.Contains("1. First", md);
        Assert.Contains("2. Second", md);
        Assert.Contains("3. Third", md);
    }

    [Fact]
    [DisplayName("嵌套列表")]
    public void Convert_NestedList()
    {
        var html = "<ul><li>Parent<ul><li>Child 1</li><li>Child 2</li></ul></li></ul>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("- Parent", md);
        Assert.Contains("  - Child 1", md);
    }

    [Fact]
    [DisplayName("列表项内联格式")]
    public void Convert_ListItem_WithInline()
    {
        var html = "<ul><li><strong>Bold item</strong></li><li><em>Italic item</em></li></ul>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("**Bold item**", md);
        Assert.Contains("*Italic item*", md);
    }

    #endregion

    #region 表格

    [Fact]
    [DisplayName("表格转换（含表头）")]
    public void Convert_Table_WithHeader()
    {
        var html = "<table><thead><tr><th>Name</th><th>Age</th></tr></thead><tbody><tr><td>Alice</td><td>30</td></tr><tr><td>Bob</td><td>25</td></tr></tbody></table>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Table, doc.Blocks[0].Type);

        var md = doc.ToMarkdown();
        Assert.Contains("| Name | Age |", md);
        Assert.Contains("| Alice | 30 |", md);
        Assert.Contains("| Bob | 25 |", md);
    }

    [Fact]
    [DisplayName("简单表格（无 thead/tbody）")]
    public void Convert_Table_Simple()
    {
        var html = "<table><tr><th>Key</th><th>Value</th></tr><tr><td>A</td><td>1</td></tr></table>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("| Key | Value |", md);
    }

    #endregion

    #region 分隔线

    [Fact]
    [DisplayName("分隔线 hr 转换")]
    public void Convert_Hr()
    {
        var html = "<p>Above</p><hr><p>Below</p>";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Equal(3, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.ThematicBreak, doc.Blocks[1].Type);

        var md = doc.ToMarkdown();
        Assert.Contains("---", md);
    }

    #endregion

    #region HTML 实体解码

    [Fact]
    [DisplayName("HTML 实体解码")]
    public void Convert_HtmlEntities()
    {
        var html = "<p>Price: &lt; $100 &amp; &gt; $50</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("<", md);
        Assert.Contains("&", md);
        Assert.Contains(">", md);
    }

    [Fact]
    [DisplayName("空白 HTML 返回空文档")]
    public void Convert_EmptyHtml_ReturnsEmptyDocument()
    {
        var doc = MarkdownDocument.FromHtml("");
        Assert.Empty(doc.Blocks);

        doc = MarkdownDocument.FromHtml("   \n  ");
        Assert.Empty(doc.Blocks);
    }

    #endregion

    #region 选项配置

    [Fact]
    [DisplayName("自定义列表符号 *")]
    public void Convert_CustomBulletChar()
    {
        var options = new HtmlToMarkdownOptions { ListBulletChar = "*" };
        var html = "<ul><li>Item</li></ul>";
        var doc = MarkdownDocument.FromHtml(html, options);

        var md = doc.ToMarkdown(options.ListBulletChar);
        Assert.Contains("* Item", md);
        Assert.DoesNotContain("- Item", md);
    }

    [Fact]
    [DisplayName("未知标签策略 Drop")]
    public void Convert_UnknownTag_Drop()
    {
        var options = new HtmlToMarkdownOptions { UnknownTags = UnknownTagStrategy.Drop };
        var html = "<div><p>Keep this</p></div>";
        var doc = MarkdownDocument.FromHtml(html, options);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("未知标签策略 PassThrough")]
    public void Convert_UnknownTag_PassThrough()
    {
        var options = new HtmlToMarkdownOptions { UnknownTags = UnknownTagStrategy.PassThrough };
        var html = "<div class=\"wrapper\"><p>Content</p></div>";
        var doc = MarkdownDocument.FromHtml(html, options);

        // 应该保留 div 作为 HtmlBlock
        var md = doc.ToMarkdown();
        Assert.Contains("<div", md);
    }

    [Fact]
    [DisplayName("URI 白名单过滤")]
    public void Convert_UriWhitelist()
    {
        var options = new HtmlToMarkdownOptions { WhitelistUriSchemes = "https" };
        var html = "<p><a href=\"http://evil.com\">bad</a> <a href=\"https://safe.com\">good</a></p>";
        var doc = MarkdownDocument.FromHtml(html, options);

        var md = doc.ToMarkdown();
        Assert.DoesNotContain("[bad]", md);
        Assert.Contains("[good](https://safe.com)", md);
    }

    #endregion

    #region 往返测试

    [Fact]
    [DisplayName("Markdown → HTML → Markdown 往返")]
    public void RoundTrip_Markdown_Html_Markdown()
    {
        var original = "# Hello\n\nThis is **bold** and *italic*.\n\n- Item 1\n- Item 2\n";
        var doc1 = MarkdownDocument.Parse(original);
        var html = doc1.ToHtml();

        // HTML 转回 Markdown
        var doc2 = MarkdownDocument.FromHtml(html);
        var restored = doc2.ToMarkdown();

        // 验证关键内容保留
        Assert.Contains("# Hello", restored);
        Assert.Contains("**bold**", restored);
        Assert.Contains("*italic*", restored);
        Assert.Contains("- Item 1", restored);
        Assert.Contains("- Item 2", restored);
    }

    #endregion

    #region Edge Cases

    [Fact]
    [DisplayName("无标签纯文本")]
    public void Convert_PlainText_NoTags()
    {
        var html = "Just some plain text without any HTML tags.";
        var doc = MarkdownDocument.FromHtml(html);

        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("自闭合标签")]
    public void Convert_SelfClosingTags()
    {
        var html = "<p>Before<img src=\"test.png\" alt=\"img\" />After</p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("![img](test.png)", md);
    }

    [Fact]
    [DisplayName("大/小写标签名")]
    public void Convert_MixedCaseTags()
    {
        var html = "<P>Paragraph</P><H1>Heading</H1><STRONG>Bold</STRONG>";
        var doc = MarkdownDocument.FromHtml(html);

        // P→段落, H1→标题, STRONG（无外层块包裹）→段落
        Assert.Equal(3, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.Heading, doc.Blocks[1].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[2].Type);
    }

    [Fact]
    [DisplayName("粗斜体嵌套 b/i")]
    public void Convert_BoldItalic_Nested()
    {
        var html = "<p><b><i>Bold and Italic</i></b></p>";
        var doc = MarkdownDocument.FromHtml(html);

        var md = doc.ToMarkdown();
        Assert.Contains("***Bold and Italic***", md);
    }

    [Fact]
    [DisplayName("ConvertToString 快捷方法")]
    public void ConvertToString_Shortcut()
    {
        var converter = new HtmlToMarkdownConverter();
        var md = converter.ConvertToString("<h1>Test</h1>");

        Assert.Contains("# Test", md);
    }

    #endregion
}
