using System.ComponentModel;
using System.Collections;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>CommonMark 0.31.2 规范套件（MD12 数据驱动回归）</summary>
/// <remarks>
/// 从 CommonMark spec.txt 精选高频用例，按「纯文本 / 结构 / HTML 片段」三组数据驱动断言。
/// 注意：纯文本断言对列表/表格/引用含换行分隔；图片 GetPlainText = Alt；代码块 = 内容。
/// </remarks>
public class MarkdownCommonMarkSuiteTests
{
    private static MarkdownDocument P(String md) => MarkdownDocument.Parse(md);

    public static IEnumerable<Object[]> PlainTextCases()
    {
        yield return new Object[] { "*foo bar*", "foo bar" };
        yield return new Object[] { "**foo bar**", "foo bar" };
        yield return new Object[] { "***foo bar***", "foo bar" };
        yield return new Object[] { "_foo bar_", "foo bar" };
        yield return new Object[] { "__foo bar__", "foo bar" };
        yield return new Object[] { "foo*bar*", "foobar" };
        yield return new Object[] { "5*6*78", "5678" };
        yield return new Object[] { "foo_bar_", "foo_bar_" };
        yield return new Object[] { "5_6_78", "5_6_78" };
        yield return new Object[] { "aa_\"bb\"_cc", "aa_\"bb\"_cc" };   // 首 _ 非左 flanking → 字面量
        yield return new Object[] { "_foo*", "_foo*" };
        yield return new Object[] { "*foo bar *", "*foo bar *" };
        yield return new Object[] { "**foo*", "*foo" };                  // rule-of-3 余量
        yield return new Object[] { "*foo**", "foo*" };
        yield return new Object[] { "# 标题", "标题" };
        yield return new Object[] { "## 二级", "二级" };
        yield return new Object[] { "正文\n===", "正文" };               // Setext H1
        yield return new Object[] { "正文\n---", "正文" };               // Setext H2
        yield return new Object[] { "`code`", "code" };
        yield return new Object[] { "`` ` ``", "`" };
        yield return new Object[] { "[链接](https://x.com)", "链接" };
        yield return new Object[] { "![图](/img/a.png)", "图" };
        yield return new Object[] { "[a][id]\n\n[id]: /x", "a" };
        yield return new Object[] { "<https://x.com>", "https://x.com" };
        yield return new Object[] { "<foo@bar.com>", "foo@bar.com" };
        yield return new Object[] { "a &amp; b", "a & b" };
        yield return new Object[] { "&#65;&#x42;", "AB" };
        yield return new Object[] { "\\*not em\\*", "*not em*" };
        yield return new Object[] { "a  \nb", "a b" };                   // 硬换行
        yield return new Object[] { "a\\\nb", "a b" };                   // 反斜杠硬换行
    }

    [Theory]
    [MemberData(nameof(PlainTextCases))]
    [DisplayName("CommonMark：纯文本提取")]
    public void CommonMark_PlainText(String md, String expected)
    {
        var doc = P(md);
        Assert.True(doc.Blocks.Count > 0, $"空块: {md}");
        Assert.Equal(expected, doc.Blocks[0].GetPlainText());
    }

    public static IEnumerable<Object[]> StructureCases()
    {
        // (markdown, 首个内联类型, 期望子结构描述)
        yield return new Object[] { "*foo **bar** baz*", "Emphasis(Text,Strong(Text),Text)" };
        yield return new Object[] { "**foo *bar* baz**", "Strong(Text,Emphasis(Text),Text)" };
        yield return new Object[] { "***foo***", "Emphasis(Strong(Text))" };
        yield return new Object[] { "**foo**", "Strong(Text)" };
        yield return new Object[] { "*foo*", "Emphasis(Text)" };
    }

    [Theory]
    [MemberData(nameof(StructureCases))]
    [DisplayName("CommonMark：强调嵌套结构")]
    public void CommonMark_Structure(String md, String expectedShape)
    {
        var inline = P(md).Blocks[0].Inlines[0];
        Assert.Equal(expectedShape, Shape(inline));
    }

    private static String Shape(MarkdownInline inline)
    {
        if (inline.Children.Count == 0) return inline.Type.ToString();
        var inner = String.Join(",", inline.Children.Select(c => Shape(c)));
        return $"{inline.Type}({inner})";
    }

    [Theory]
    [DisplayName("CommonMark：HTML 片段")]
    [InlineData("# 标题", "<h1")]
    [InlineData("**粗体**", "<strong>粗体</strong>")]
    [InlineData("*斜体*", "<em>斜体</em>")]
    [InlineData("~~删除~~", "<del>删除</del>")]
    [InlineData("`code`", "<code>code</code>")]
    [InlineData("[链接](https://x.com)", "<a href=\"https://x.com\">链接</a>")]
    [InlineData("![图](/img/a.png)", "<img src=\"/img/a.png\" alt=\"图\"")]
    [InlineData("---", "<hr />")]
    [InlineData("> 引用", "<blockquote")]
    [InlineData("```cs\nx\n```", "language-cs")]
    [InlineData("- a\n- b", "<ul>")]
    [InlineData("1. a\n2. b", "<ol")]
    [InlineData("| a | b |\n| - | - |\n| 1 | 2 |", "<table>")]
    [InlineData("- [x] 完成", "<input type=\"checkbox\" disabled=\"\" checked=\"\"")]
    [InlineData("https://newlifex.com", "<a href=\"https://newlifex.com\">")]
    public void CommonMark_Html(String md, String expectedFragment)
    {
        var html = P(md).ToHtml();
        Assert.Contains(expectedFragment, html);
    }
}
