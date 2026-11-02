using System.ComponentModel;
using System.Text;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>深层嵌套健壮性：合法的深嵌套输入不得导致栈溢出或解析崩溃</summary>
/// <remarks>
/// 递归下降解析器遇到极深嵌套（HTML 标签、列表、引用块）容易栈溢出导致整个进程崩溃，
/// 这比抛异常严重得多。真实网页/机器生成文档确实会出现几十层嵌套，因此这里用
/// 100~500 层输入验证解析路径是"稳妥失败或正常完成"，而不是把进程打挂。
/// </remarks>
public class DeepNestingTests
{
    [Theory, DisplayName("嵌套-深层 HTML 标签不影响 Markdown 转换")]
    [InlineData(100)]
    [InlineData(500)]
    public void DeepHtmlTags_Convert(Int32 depth)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < depth; i++) sb.Append("<blockquote>\n");
        sb.Append("<p>最深处的文字</p>\n");
        for (var i = 0; i < depth; i++) sb.Append("</blockquote>\n");

        var doc = MarkdownDocument.FromHtml(sb.ToString());
        Assert.Contains("最深处的文字", doc.ToMarkdown());
    }

    [Theory, DisplayName("嵌套-深层 div 包裹不影响 Markdown 转换")]
    [InlineData(200)]
    [InlineData(1000)]
    public void DeepDivs_Convert(Int32 depth)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < depth; i++) sb.Append("<div>");
        sb.Append("文本");
        for (var i = 0; i < depth; i++) sb.Append("</div>");

        var doc = MarkdownDocument.FromHtml(sb.ToString());
        Assert.Contains("文本", doc.ToMarkdown());
    }

    [Fact, DisplayName("嵌套-深层列表不影响 Markdown 解析")]
    public void DeepList_Parse()
    {
        const Int32 depth = 200;
        var sb = new StringBuilder();
        for (var i = 0; i < depth; i++) sb.Append(new String(' ', i * 2)).Append("- 第").Append(i).Append("层\n");

        var doc = MarkdownDocument.Parse(sb.ToString());
        Assert.NotNull(doc);
        Assert.NotEmpty(doc.Blocks);
    }

    [Fact, DisplayName("嵌套-深层引用块不影响 Markdown 解析")]
    public void DeepQuote_Parse()
    {
        const Int32 depth = 300;
        var sb = new StringBuilder();
        for (var i = 0; i < depth; i++) sb.Append(new String('>', i + 1)).Append(' ').Append("层").Append(i).Append('\n');

        var doc = MarkdownDocument.Parse(sb.ToString());
        Assert.NotNull(doc);
        Assert.NotEmpty(doc.Blocks);
    }
}
