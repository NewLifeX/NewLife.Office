using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>CommonMark 规范边界用例（MD13 数据驱动回归）</summary>
/// <remarks>
/// 覆盖 MD13 修复的 5 类规范缺陷：
/// ① setext 单个 `-` 是空列表项不构成 underline（例 77）；② 代码跨度闭合须匹配等长反引号 run（例 344）；
/// ③ 表格 delimiter 行须含 `|` 且 header 行也须含 `|`（GFM）；④ 链接文本不能包含嵌套链接（例 509，图片除外）；
/// ⑤ 表格 delimiter 行优先于列表中断（GFM 例 205）。
/// </remarks>
public class MarkdownCommonMarkEdgeTests
{
    private static String H(String md) => MarkdownDocument.Parse(md).ToHtml();

    [Theory]
    [DisplayName("CommonMark：setext 边界（单个-为段落，--为H2）")]
    [InlineData("Foo\n-", "<p>Foo", "-</p>")]
    [InlineData("Foo\n--", "<h2", "")]
    [InlineData("Foo\n---", "<h2", "")]
    [InlineData("Foo\n=", "<h1", "")]
    [InlineData("Foo\n==", "<h1", "")]
    public void Setext_Edge(String md, String startTag, String expectedEnd)
    {
        var html = H(md);
        var ok = html.Contains(startTag) && (expectedEnd.Length == 0 || html.Contains(expectedEnd));
        Assert.True(ok, $"期望 {startTag}..{expectedEnd}\n实际: {html}");
    }

    [Theory]
    [DisplayName("CommonMark：代码跨度等长反引号闭合（§6.4）")]
    [InlineData("`foo `` bar`", "<code>foo `` bar</code>")]
    [InlineData("``foo ``", "<code>foo </code>")]
    [InlineData("`foo`", "<code>foo</code>")]
    [InlineData("`` ` ``", "<code>`</code>")]
    [InlineData("` `", "<code> </code>")]
    [InlineData("`` ``", "<code> </code>")]
    public void Backtick_EqualRun(String md, String expected)
    {
        var html = H(md);
        Assert.True(html.Contains(expected), $"期望包含 {expected}\n实际: {html}");
    }

    [Theory]
    [DisplayName("CommonMark：链接内强调/空目标/大小写引用")]
    [InlineData("[foo *bar*](url)", "<a href=\"url\">foo <em>bar</em></a>")]
    [InlineData("[foo]()", "<a href=\"\">foo</a>")]
    public void Link_Basic(String md, String expected)
    {
        var html = H(md);
        Assert.True(html.Contains(expected), $"期望包含 {expected}\n实际: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：引用链接大小写不敏感")]
    public void Ref_CaseInsensitive()
    {
        var html = H("[foo][BAR]\n\n[bar]: /url");
        Assert.True(html.Contains("<a href=\"/url\">foo</a>"), $"实际: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：链接文本不能包含嵌套链接（内层优先，外层字面量）")]
    public void Link_NestedInnerWins()
    {
        // 例 509：内层链接优先，外层 `[foo ` 与 `](/uri)` 为字面文本，不产生非法嵌套 <a>
        var html = H("[foo [bar](/uri)](/uri)");
        Assert.True(html.Contains("<a href=\"/uri\">bar</a>") && html.Contains("[foo "), $"实际: {html}");
        Assert.False(html.Contains("<a href=\"/uri\">foo <a"), $"非法嵌套链接: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：普通嵌套方括号不阻止外层链接（例 504）")]
    public void Link_BracketTextAllowed()
    {
        var html = H("[link [foo [bar]]](/uri)");
        Assert.True(html.Contains("<a href=\"/uri\">link [foo [bar]]</a>"), $"实际: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：图片链接允许嵌套（[![alt](img)](url) 是标准用法）")]
    public void Link_ImageNestedAllowed()
    {
        var html = H("[![alt](/img/a.png)](https://x.com)");
        Assert.True(html.Contains("<a href=\"https://x.com\"><img") && html.Contains("src=\"/img/a.png\""), $"实际: {html}");
    }

    [Theory]
    [DisplayName("CommonMark：嵌套强调三明治")]
    [InlineData("foo***bar***baz", "foo<em><strong>bar</strong></em>baz")]
    [InlineData("*foo**bar**baz*", "<em>foo<strong>bar</strong>baz</em>")]
    public void Emphasis_Nested(String md, String expected)
    {
        var html = H(md);
        Assert.True(html.Contains(expected), $"期望包含 {expected}\n实际: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：HTML 块中断段落（type6 到空行）")]
    public void HtmlBlock_InterruptsParagraph()
    {
        var html = H("aaa\n<div>\nbbb");
        Assert.True(html.Contains("<p>aaa</p>") && html.Contains("<div>\nbbb"), $"实际: {html}");
    }

    [Theory]
    [DisplayName("CommonMark：标题与转义")]
    [InlineData("# foo #", "<h1", "foo</h1>")]
    [InlineData("# \\# 转义", "<h1", "# 转义</h1>")]
    [InlineData("# foo", "<h1", "foo</h1>")]
    public void Heading_Escape(String md, String startTag, String expectedContent)
    {
        var html = H(md);
        Assert.True(html.Contains(startTag) && html.Contains(expectedContent), $"期望 {startTag}..{expectedContent}\n实际: {html}");
    }

    [Fact]
    [DisplayName("CommonMark：代码块内空行保真")]
    public void Fence_BlankLines()
    {
        var doc = MarkdownDocument.Parse("```\na\n\nb\n```");
        Assert.Equal("a\n\nb", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("CommonMark：不同 bullet 字符是不同列表")]
    public void List_MixedBullets()
    {
        var html = H("- a\n* b");
        Assert.True(html.Contains("<ul>") && html.Split("<ul>").Length - 1 >= 2, $"实际: {html}");
    }

    [Theory]
    [DisplayName("GFM：表格无首尾管道（header 与 delimiter 均含 |）")]
    [InlineData("a | b\n- | -\n1 | 2", "<table>")]
    [InlineData("abc | def\n--- | ---", "<table>")]
    public void Table_NoOuterPipe(String md, String expected)
    {
        var html = H(md);
        Assert.True(html.Contains(expected), $"期望 {expected}\n实际: {html}");
    }

    [Fact]
    [DisplayName("GFM：header 行无管道符不是表格（段落+列表）")]
    public void Table_HeaderNeedsPipe()
    {
        // `Foo` 无 | 不是合法表头 → `- | -` 作为列表项中断段落
        var html = H("Foo\n- | -\nbar | baz");
        Assert.True(!html.Contains("<table>") && html.Contains("<p>Foo</p>"), $"实际: {html}");
    }
}
