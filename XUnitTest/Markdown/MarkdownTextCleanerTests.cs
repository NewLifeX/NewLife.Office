using System.ComponentModel;
using System.Linq;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 文本规范化与块构建辅助测试</summary>
[Trait("Category", "Markdown")]
public class MarkdownTextCleanerTests
{
    #region MarkdownTextCleaner

    [Fact, DisplayName("Clean_剔除控制字符保留换行制表")]
    public void Clean_RemovesControlChars()
    {
        var result = MarkdownTextCleaner.Clean("a\u0001b\u0002c\n\t d");
        Assert.Equal("abc\n\t d", result);
    }

    [Fact, DisplayName("Clean_NBSP转换为普通空格")]
    public void Clean_NbspToSpace()
    {
        var result = MarkdownTextCleaner.Clean("a\u00A0b");
        Assert.Equal("a b", result);
    }

    [Fact, DisplayName("Clean_剥离软连字符")]
    public void Clean_StripsSoftHyphen()
    {
        var result = MarkdownTextCleaner.Clean("a\u00ADb");
        Assert.Equal("ab", result);
    }

    [Fact, DisplayName("Clean_保留ZWNJ与ZWJ")]
    public void Clean_PreservesZwnjZwj()
    {
        var result = MarkdownTextCleaner.Clean("a\u200Cb\u200Dc");
        Assert.Equal("a\u200Cb\u200Dc", result);
    }

    [Fact, DisplayName("Clean_空值返回空串，空格保留不trim")]
    public void Clean_NullOrEmpty()
    {
        Assert.Equal("", MarkdownTextCleaner.Clean(null));
        Assert.Equal("", MarkdownTextCleaner.Clean(""));
        // Clean 只做规范化，不 trim 普通空格
        Assert.Equal("  ", MarkdownTextCleaner.Clean("  "));
    }

    [Fact, DisplayName("CleanCell_换行转换为空格")]
    public void CleanCell_NewlineToSpace()
    {
        var result = MarkdownTextCleaner.CleanCell("第一行\r\n第二行");
        Assert.Equal("第一行 第二行", result);
    }

    #endregion

    #region MarkdownBlockBuilder

    [Fact, DisplayName("Table_单元格竖线转义")]
    public void Table_EscapesPipe()
    {
        var table = MarkdownBlockBuilder.Table(new[] { new[] { "名称", "值" }, new[] { "A|B", "1" } }, true);
        var md = new MarkdownDocument();
        md.Blocks.Add(table);
        var text = md.ToMarkdown();

        Assert.Contains("A\\|B", text);
        Assert.DoesNotContain("| A|B |", text);
    }

    [Fact, DisplayName("Table_首行表头标记")]
    public void Table_FirstRowHeader()
    {
        var table = MarkdownBlockBuilder.Table(new[] { new[] { "列1", "列2" }, new[] { "A", "B" } }, true);
        var row0 = (TableRowBlock)table.Children[0];
        var cell0 = (TableCellBlock)row0.Children[0];
        Assert.True(cell0.IsHeader);

        var row1 = (TableRowBlock)table.Children[1];
        var cell1 = (TableCellBlock)row1.Children[0];
        Assert.False(cell1.IsHeader);
    }

    [Fact, DisplayName("Heading_自动清洗文本")]
    public void Heading_CleansText()
    {
        var heading = MarkdownBlockBuilder.Heading(2, "标题\u0001内容");
        var doc = new MarkdownDocument();
        doc.Blocks.Add(heading);
        Assert.Contains("标题内容", doc.ToMarkdown());
    }

    [Fact, DisplayName("BlockQuote_创建引用块")]
    public void BlockQuote_CreatesQuote()
    {
        var quote = MarkdownBlockBuilder.BlockQuote("引用内容");
        var doc = new MarkdownDocument();
        doc.Blocks.Add(quote);
        var text = doc.ToMarkdown();
        Assert.Contains("> 引用内容", text);
    }

    #endregion
}
