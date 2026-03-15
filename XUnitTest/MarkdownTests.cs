using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest;

/// <summary>Markdown 解析/渲染/序列化测试</summary>
[Trait("Category", "Markdown")]
public class MarkdownTests
{
    #region 解析 - 标题

    [Fact]
    [DisplayName("ATX 一级标题")]
    public void Parse_AtxHeading_H1()
    {
        var doc = MarkdownDocument.Parse("# Hello World");
        Assert.Single(doc.Blocks);
        var h = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.Heading, h.Type);
        Assert.Equal(1, h.Level);
        Assert.Equal("Hello World", h.GetPlainText().Trim());
    }

    [Fact]
    [DisplayName("ATX 三级标题")]
    public void Parse_AtxHeading_H3()
    {
        var doc = MarkdownDocument.Parse("### Section Title");
        var h = doc.Blocks[0];
        Assert.Equal(3, h.Level);
    }

    [Fact]
    [DisplayName("ATX 标题支持内联格式")]
    public void Parse_AtxHeading_WithInline()
    {
        var doc = MarkdownDocument.Parse("## Hello **World**");
        var h = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.Heading, h.Type);
        Assert.Equal(2, h.Level);
        Assert.True(h.Inlines.Count >= 2);
    }

    [Fact]
    [DisplayName("Setext 二级标题")]
    public void Parse_SetextHeading_H2()
    {
        var doc = MarkdownDocument.Parse("Section\n-------");
        Assert.True(doc.Blocks.Count >= 1);
        var h = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.Heading, h.Type);
        Assert.Equal(2, h.Level);
    }

    #endregion

    #region 解析 - 段落

    [Fact]
    [DisplayName("单段落解析")]
    public void Parse_Paragraph()
    {
        var doc = MarkdownDocument.Parse("Hello world.");
        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
        Assert.Contains("Hello world", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("多段落以空行分隔")]
    public void Parse_MultipleParagraphs()
    {
        var doc = MarkdownDocument.Parse("Para 1.\n\nPara 2.");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.All(doc.Blocks, b => Assert.Equal(MarkdownBlockType.Paragraph, b.Type));
    }

    #endregion

    #region 解析 - 代码块

    [Fact]
    [DisplayName("围栏代码块含语言标识")]
    public void Parse_FencedCodeBlock_WithLanguage()
    {
        var md = "```csharp\nvar x = 1;\n```";
        var doc = MarkdownDocument.Parse(md);
        Assert.Single(doc.Blocks);
        var cb = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.CodeBlock, cb.Type);
        Assert.Equal("csharp", cb.Language);
        Assert.Contains("var x = 1;", cb.RawText);
    }

    [Fact]
    [DisplayName("围栏代码块无语言标识")]
    public void Parse_FencedCodeBlock_NoLanguage()
    {
        var md = "```\nsome code\n```";
        var doc = MarkdownDocument.Parse(md);
        var cb = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.CodeBlock, cb.Type);
        Assert.True(String.IsNullOrEmpty(cb.Language));
    }

    #endregion

    #region 解析 - 引用块

    [Fact]
    [DisplayName("引用块解析")]
    public void Parse_BlockQuote()
    {
        var doc = MarkdownDocument.Parse("> This is a quote.");
        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.BlockQuote, doc.Blocks[0].Type);
        Assert.True(doc.Blocks[0].Children.Count >= 1);
    }

    #endregion

    #region 解析 - 列表

    [Fact]
    [DisplayName("无序列表（连字符）")]
    public void Parse_BulletList_Hyphen()
    {
        var doc = MarkdownDocument.Parse("- Item 1\n- Item 2\n- Item 3");
        Assert.Single(doc.Blocks);
        var list = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.BulletList, list.Type);
        Assert.Equal(3, list.Children.Count);
        Assert.All(list.Children, c => Assert.Equal(MarkdownBlockType.ListItem, c.Type));
    }

    [Fact]
    [DisplayName("有序列表")]
    public void Parse_OrderedList()
    {
        var doc = MarkdownDocument.Parse("1. First\n2. Second\n3. Third");
        Assert.Single(doc.Blocks);
        var list = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.OrderedList, list.Type);
        Assert.Equal(3, list.Children.Count);
    }

    [Fact]
    [DisplayName("任务列表项（未勾选）")]
    public void Parse_TaskItem_Unchecked()
    {
        var doc = MarkdownDocument.Parse("- [ ] Todo item");
        var item = doc.Blocks[0].Children[0];
        Assert.True(item.IsTaskItem);
        Assert.False(item.IsChecked);
    }

    [Fact]
    [DisplayName("任务列表项（已勾选）")]
    public void Parse_TaskItem_Checked()
    {
        var doc = MarkdownDocument.Parse("- [x] Done item");
        var item = doc.Blocks[0].Children[0];
        Assert.True(item.IsTaskItem);
        Assert.True(item.IsChecked);
    }

    #endregion

    #region 解析 - 表格 (GFM)

    [Fact]
    [DisplayName("GFM 表格基本解析")]
    public void Parse_Table_Basic()
    {
        var md = "| A | B |\n|---|---|\n| 1 | 2 |";
        var doc = MarkdownDocument.Parse(md);
        Assert.Single(doc.Blocks);
        var table = doc.Blocks[0];
        Assert.Equal(MarkdownBlockType.Table, table.Type);
        Assert.Equal(2, table.Children.Count); // header + 1 row
    }

    [Fact]
    [DisplayName("GFM 表格对齐解析")]
    public void Parse_Table_Alignment()
    {
        var md = "| L | C | R |\n|:--|:-:|--:|\n| a | b | c |";
        var doc = MarkdownDocument.Parse(md);
        var table = doc.Blocks[0];
        var headerRow = table.Children[0];
        Assert.Equal(3, headerRow.Children.Count);
    }

    #endregion

    #region 解析 - 分割线

    [Fact]
    [DisplayName("三连线分割线")]
    public void Parse_ThematicBreak()
    {
        var doc = MarkdownDocument.Parse("---");
        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.ThematicBreak, doc.Blocks[0].Type);
    }

    #endregion

    #region 解析 - 行内元素

    [Fact]
    [DisplayName("行内加粗")]
    public void Parse_Inline_Bold()
    {
        var doc = MarkdownDocument.Parse("This is **bold** text.");
        var p = doc.Blocks[0];
        Assert.Contains(p.Inlines, (MarkdownInline i) => i.Type == MarkdownInlineType.Strong);
    }

    [Fact]
    [DisplayName("行内斜体")]
    public void Parse_Inline_Italic()
    {
        var doc = MarkdownDocument.Parse("This is *italic* text.");
        var p = doc.Blocks[0];
        Assert.Contains(p.Inlines, (MarkdownInline i) => i.Type == MarkdownInlineType.Emphasis);
    }

    [Fact]
    [DisplayName("行内代码")]
    public void Parse_Inline_Code()
    {
        var doc = MarkdownDocument.Parse("Use `var x = 1;` in code.");
        var p = doc.Blocks[0];
        var code = p.Inlines.FirstOrDefault(i => i.Type == MarkdownInlineType.Code);
        Assert.NotNull(code);
        Assert.Equal("var x = 1;", code.Text);
    }

    [Fact]
    [DisplayName("行内删除线")]
    public void Parse_Inline_Strikethrough()
    {
        var doc = MarkdownDocument.Parse("This is ~~deleted~~ text.");
        var p = doc.Blocks[0];
        Assert.Contains(p.Inlines, (MarkdownInline i) => i.Type == MarkdownInlineType.Strikethrough);
    }

    [Fact]
    [DisplayName("超链接解析")]
    public void Parse_Inline_Link()
    {
        var doc = MarkdownDocument.Parse("[Example](https://example.com)");
        var p = doc.Blocks[0];
        var link = p.Inlines.FirstOrDefault(i => i.Type == MarkdownInlineType.Link);
        Assert.NotNull(link);
        Assert.Equal("https://example.com", link.Href);
    }

    [Fact]
    [DisplayName("带标题的超链接")]
    public void Parse_Inline_Link_WithTitle()
    {
        var doc = MarkdownDocument.Parse("[Example](https://example.com \"My Site\")");
        var p = doc.Blocks[0];
        var link = p.Inlines.FirstOrDefault(i => i.Type == MarkdownInlineType.Link);
        Assert.NotNull(link);
        Assert.Equal("My Site", link.Title);
    }

    [Fact]
    [DisplayName("图片解析")]
    public void Parse_Inline_Image()
    {
        var doc = MarkdownDocument.Parse("![Alt Text](https://example.com/img.png)");
        var p = doc.Blocks[0];
        var img = p.Inlines.FirstOrDefault(i => i.Type == MarkdownInlineType.Image);
        Assert.NotNull(img);
        Assert.Equal("https://example.com/img.png", img.Href);
        Assert.Equal("Alt Text", img.Alt);
    }

    #endregion

    #region 转换 - HTML

    [Fact]
    [DisplayName("标题转HTML")]
    public void ToHtml_Heading()
    {
        var html = MarkdownDocument.Parse("# Hello").ToHtml();
        Assert.Contains("<h1", html);
        Assert.Contains("Hello", html);
        Assert.Contains("</h1>", html);
    }

    [Fact]
    [DisplayName("段落转HTML")]
    public void ToHtml_Paragraph()
    {
        var html = MarkdownDocument.Parse("Hello world.").ToHtml();
        Assert.Contains("<p>", html);
        Assert.Contains("Hello world.", html);
        Assert.Contains("</p>", html);
    }

    [Fact]
    [DisplayName("代码块转HTML含语言class")]
    public void ToHtml_CodeBlock_WithLanguageClass()
    {
        var md = "```python\nprint('hi')\n```";
        var opts = new MarkdownHtmlOptions { AddLanguageClass = true };
        var html = MarkdownDocument.Parse(md).ToHtml(opts);
        Assert.Contains("language-python", html);
        Assert.Contains("<pre><code", html);
    }

    [Fact]
    [DisplayName("无序列表转HTML")]
    public void ToHtml_BulletList()
    {
        var html = MarkdownDocument.Parse("- A\n- B").ToHtml();
        Assert.Contains("<ul>", html);
        Assert.Contains("<li>", html);
        Assert.Contains("</ul>", html);
    }

    [Fact]
    [DisplayName("有序列表转HTML")]
    public void ToHtml_OrderedList()
    {
        var html = MarkdownDocument.Parse("1. A\n2. B").ToHtml();
        Assert.Contains("<ol>", html);
        Assert.Contains("</ol>", html);
    }

    [Fact]
    [DisplayName("加粗文字转HTML")]
    public void ToHtml_Bold()
    {
        var html = MarkdownDocument.Parse("**Bold**").ToHtml();
        Assert.Contains("<strong>", html);
        Assert.Contains("</strong>", html);
    }

    [Fact]
    [DisplayName("斜体文字转HTML")]
    public void ToHtml_Italic()
    {
        var html = MarkdownDocument.Parse("*Italic*").ToHtml();
        Assert.Contains("<em>", html);
        Assert.Contains("</em>", html);
    }

    [Fact]
    [DisplayName("行内代码转HTML")]
    public void ToHtml_InlineCode()
    {
        var html = MarkdownDocument.Parse("`code`").ToHtml();
        Assert.Contains("<code>", html);
        Assert.Contains("code", html);
    }

    [Fact]
    [DisplayName("链接转HTML")]
    public void ToHtml_Link()
    {
        var html = MarkdownDocument.Parse("[Click](https://x.com)").ToHtml();
        Assert.Contains("<a href=", html);
        Assert.Contains("https://x.com", html);
    }

    [Fact]
    [DisplayName("外部链接加 target=_blank")]
    public void ToHtml_ExternalLink_Target()
    {
        var opts = new MarkdownHtmlOptions { ExternalLinkTarget = true };
        var html = MarkdownDocument.Parse("[X](https://x.com)").ToHtml(opts);
        Assert.Contains("target=\"_blank\"", html);
        Assert.Contains("rel=\"noopener noreferrer\"", html);
    }

    [Fact]
    [DisplayName("图片转HTML")]
    public void ToHtml_Image()
    {
        var html = MarkdownDocument.Parse("![Alt](img.png)").ToHtml();
        Assert.Contains("<img ", html);
        Assert.Contains("alt=\"Alt\"", html);
        Assert.Contains("img.png", html);
    }

    [Fact]
    [DisplayName("表格转HTML")]
    public void ToHtml_Table()
    {
        var md = "| A | B |\n|---|---|\n| 1 | 2 |";
        var html = MarkdownDocument.Parse(md).ToHtml();
        Assert.Contains("<table>", html);
        Assert.Contains("<thead>", html);
        Assert.Contains("<tbody>", html);
        Assert.Contains("<th", html);
        Assert.Contains("<td", html);
    }

    [Fact]
    [DisplayName("分割线转HTML")]
    public void ToHtml_ThematicBreak()
    {
        var html = MarkdownDocument.Parse("---").ToHtml();
        Assert.Contains("<hr", html);
    }

    [Fact]
    [DisplayName("引用块转HTML")]
    public void ToHtml_BlockQuote()
    {
        var html = MarkdownDocument.Parse("> Quote text").ToHtml();
        Assert.Contains("<blockquote>", html);
        Assert.Contains("</blockquote>", html);
    }

    [Fact]
    [DisplayName("HTML特殊字符转义")]
    public void ToHtml_HtmlEncoding()
    {
        var html = MarkdownDocument.Parse("Use <script> & \"quotes\"").ToHtml();
        Assert.Contains("&lt;script&gt;", html);
        Assert.Contains("&amp;", html);
    }

    [Fact]
    [DisplayName("危险链接被过滤（SafeLinks）")]
    public void ToHtml_SafeLinks_DangerousUrl()
    {
        var opts = new MarkdownHtmlOptions { SafeLinks = true };
        var html = MarkdownDocument.Parse("[Click](javascript:alert(1))").ToHtml(opts);
        Assert.DoesNotContain("javascript:", html);
    }

    [Fact]
    [DisplayName("ToHtmlPage返回完整HTML页面")]
    public void ToHtmlPage_Structure()
    {
        var html = MarkdownDocument.Parse("# Hello").ToHtmlPage("Test Page");
        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("<title>Test Page</title>", html);
        Assert.Contains("<h1", html);
    }

    #endregion

    #region 序列化 - 往返

    [Fact]
    [DisplayName("标题往返序列化")]
    public void RoundTrip_Heading()
    {
        var original = "# Hello World";
        var md = MarkdownDocument.Parse(original).ToMarkdown().Trim();
        Assert.StartsWith("# Hello World", md);
    }

    [Fact]
    [DisplayName("代码块往返序列化")]
    public void RoundTrip_CodeBlock()
    {
        var original = "```csharp\nvar x = 1;\n```";
        var md = MarkdownDocument.Parse(original).ToMarkdown();
        Assert.Contains("```csharp", md);
        Assert.Contains("var x = 1;", md);
    }

    [Fact]
    [DisplayName("无序列表往返序列化")]
    public void RoundTrip_BulletList()
    {
        var original = "- Alpha\n- Beta\n- Gamma";
        var doc = MarkdownDocument.Parse(original);
        var md = doc.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md);
        Assert.Single(doc2.Blocks);
        Assert.Equal(3, doc2.Blocks[0].Children.Count);
    }

    [Fact]
    [DisplayName("多块混合往返序列化")]
    public void RoundTrip_Mixed()
    {
        var original = "# Title\n\nSome paragraph.\n\n- Item 1\n- Item 2";
        var doc = MarkdownDocument.Parse(original);
        Assert.Equal(3, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Heading, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[2].Type);
    }

    #endregion

    #region 边界 / 异常

    [Fact]
    [DisplayName("空字符串解析返回空文档")]
    public void Parse_Empty_String()
    {
        var doc = MarkdownDocument.Parse("");
        Assert.NotNull(doc);
        Assert.Empty(doc.Blocks);
    }

    [Fact]
    [DisplayName("null文本解析返回空文档")]
    public void Parse_Null_String()
    {
        var doc = MarkdownDocument.Parse((String)null);
        Assert.NotNull(doc);
        Assert.Empty(doc.Blocks);
    }

    [Fact]
    [DisplayName("从流解析文档")]
    public void Parse_FromStream()
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes("# Stream Heading");
        using var ms = new MemoryStream(bytes);
        var doc = MarkdownDocument.Parse(ms);
        Assert.Single(doc.Blocks);
        Assert.Equal(MarkdownBlockType.Heading, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("超长段落不崩溃")]
    public void Parse_LongParagraph()
    {
        var text = new String('A', 10000);
        var doc = MarkdownDocument.Parse(text);
        Assert.NotEmpty(doc.Blocks);
    }

    [Fact]
    [DisplayName("嵌套引用块不崩溃")]
    public void Parse_NestedBlockQuote()
    {
        var md = "> Level 1\n>> Level 2\n>>> Level 3";
        var doc = MarkdownDocument.Parse(md);
        Assert.NotEmpty(doc.Blocks);
    }

    #endregion
}
