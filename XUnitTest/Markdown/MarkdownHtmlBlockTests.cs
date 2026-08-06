using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown HTML 块类型与表格转义管道符测试（Phase D）</summary>
/// <remarks>
/// 覆盖 CommonMark 6 种 HTML 块类型（①特殊标签 ②注释 ③处理指令 ④声明 ⑤CDATA ⑥块级标签 ⑦单行标签）
/// 与 GFM 表格转义管道符的解析与往返。
/// </remarks>
public class MarkdownHtmlBlockTests
{
    #region HTML 块类型

    [Fact]
    [DisplayName("HTML块：块级标签 div 到空行")]
    public void HtmlBlock_Div()
    {
        var doc = MarkdownDocument.Parse("<div>\n内容\n</div>\n\n段落\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Equal("<div>\n内容\n</div>", ((HtmlBlock)doc.Blocks[0]).RawText);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
    }

    [Fact]
    [DisplayName("HTML块：script 跨空行到闭合标签")]
    public void HtmlBlock_Script()
    {
        var doc = MarkdownDocument.Parse("<script>\nvar a = 1;\n\nalert(a);\n</script>\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Contains("alert(a)", ((HtmlBlock)doc.Blocks[0]).RawText);
    }

    [Fact]
    [DisplayName("HTML块：注释跨行到结束标记")]
    public void HtmlBlock_Comment()
    {
        var doc = MarkdownDocument.Parse("<!-- 注释\n跨行 -->\n\n段落\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Contains("跨行", ((HtmlBlock)doc.Blocks[0]).RawText);
    }

    [Fact]
    [DisplayName("HTML块：CDATA 到结束标记")]
    public void HtmlBlock_CData()
    {
        var doc = MarkdownDocument.Parse("<![CDATA[\ncontent\n]]>\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Contains("content", ((HtmlBlock)doc.Blocks[0]).RawText);
    }

    [Fact]
    [DisplayName("HTML块：DOCTYPE 声明")]
    public void HtmlBlock_Doctype()
    {
        var doc = MarkdownDocument.Parse("<!DOCTYPE html>\n\n段落\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Equal("<!DOCTYPE html>", ((HtmlBlock)doc.Blocks[0]).RawText);
    }

    [Fact]
    [DisplayName("HTML块：单行完整标签+空行")]
    public void HtmlBlock_SingleTag()
    {
        var doc = MarkdownDocument.Parse("<span>\n\n段落\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[0].Type);
        Assert.Equal("<span>", ((HtmlBlock)doc.Blocks[0]).RawText);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
    }

    [Fact]
    [DisplayName("HTML块：非标签的 < 行保持段落")]
    public void HtmlBlock_NotHtml_KeptParagraph()
    {
        var doc = MarkdownDocument.Parse("< 不是标签\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
    }

    #endregion

    #region 表格转义管道符

    [Fact]
    [DisplayName("表格：转义管道符解析为字面量")]
    public void Table_EscapedPipe()
    {
        var doc = MarkdownDocument.Parse("| A\\|B | C |\n|---|---|\n| 1 | 2 |\n");
        var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
        var header = table.Children[0];
        Assert.Equal(2, header.Children.Count);
        Assert.Equal("A|B", header.Children[0].GetPlainText());
        Assert.Equal("C", header.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("表格：转义管道符往返")]
    public void Table_EscapedPipe_RoundTrip()
    {
        var doc = MarkdownDocument.Parse("| C1 | C2 |\n|---|---|\n| 1 | A\\|B |\n");
        var md = doc.ToMarkdown();
        Assert.Contains("A\\|B", md);

        var doc2 = MarkdownDocument.Parse(md);
        var table2 = Assert.IsType<TableBlock>(doc2.Blocks[0]);
        // 数据行第 2 列还原为 A|B
        Assert.Equal("A|B", table2.Children[1].Children[1].GetPlainText());
        // 表格仍为 2 列结构
        Assert.Equal(2, table2.Children[0].Children.Count);
    }

    [Fact]
    [DisplayName("表格：未转义管道符正常拆分")]
    public void Table_NormalPipe()
    {
        var doc = MarkdownDocument.Parse("| A | B | C |\n|---|---|---|\n");
        var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
        Assert.Equal(3, table.Children[0].Children.Count);
    }

    #endregion
}
