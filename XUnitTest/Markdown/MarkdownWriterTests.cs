using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 写入器测试：转义正确性、往返修改跟踪、源码保真</summary>
/// <remarks>
/// 覆盖写入侧正确性：特殊字符转义、行内代码反引号、表格管道符、URL 尖括号包裹、
/// 行首标记保护、往返模式修改检测与块间空行保真。
/// </remarks>
public class MarkdownWriterTests
{
    #region 文本转义

    [Fact]
    [DisplayName("写入：文本含特殊字符自动转义，往返语义不变")]
    public void Write_EscapesSpecialChars()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("a * b _ c [ d ] ` e")]));

        var md = doc.ToMarkdown();
        Assert.Contains("\\*", md);
        Assert.Contains("\\_", md);
        Assert.Contains("\\[", md);
        Assert.Contains("\\]", md);
        Assert.Contains("\\`", md);

        // 往返语义不变
        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal("a * b _ c [ d ] ` e", doc2.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("写入：文本含反斜杠正确转义")]
    public void Write_EscapesBackslash()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("a\\b")]));

        var md = doc.ToMarkdown();
        Assert.Contains("\\\\", md);

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal("a\\b", doc2.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("写入：链接标题含引号与反斜杠转义")]
    public void Write_LinkTitleEscapesQuote()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([
            MarkdownInline.CreateLink("https://newlifex.com", "标题\"引号\\斜杠", [MarkdownInline.CreateText("链接")])
        ]));

        var md = doc.ToMarkdown();
        Assert.Contains("标题\\\"引号", md);

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal("标题\"引号\\斜杠", doc2.Blocks[0].Inlines[0].Title);
    }

    [Fact]
    [DisplayName("写入：段首特殊字符不被误解析为列表/标题/引用")]
    public void Write_ProtectsLineStartMarkers()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("- 列表内容")]));
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("# 标题内容")]));
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("> 引用内容")]));
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText("--- 分隔线内容")]));

        // 重新解析应仍是 4 个段落
        var md = doc.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal(4, doc2.Blocks.Count);
        for (var i = 0; i < 4; i++)
        {
            Assert.Equal(MarkdownBlockType.Paragraph, doc2.Blocks[i].Type);
        }
        Assert.Equal("- 列表内容", doc2.Blocks[0].GetPlainText());
        Assert.Equal("# 标题内容", doc2.Blocks[1].GetPlainText());
        Assert.Equal("> 引用内容", doc2.Blocks[2].GetPlainText());
        Assert.Equal("--- 分隔线内容", doc2.Blocks[3].GetPlainText());
    }

    #endregion

    #region 行内代码

    [Fact]
    [DisplayName("写入：行内代码含反引号用双反引号包裹")]
    public void Write_CodeWithBacktick()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateCode("a`b")]));

        var md = doc.ToMarkdown();
        Assert.Contains("`` a`b ``", md);

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal(MarkdownInlineType.Code, doc2.Blocks[0].Inlines[0].Type);
        Assert.Equal("a`b", doc2.Blocks[0].Inlines[0].Text);
    }

    [Fact]
    [DisplayName("写入：行内代码首尾空格按 CommonMark 剥离还原")]
    public void Write_CodeWithPaddingSpaces()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateCode(" 代码 ")]));

        var md = doc.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal(MarkdownInlineType.Code, doc2.Blocks[0].Inlines[0].Type);
        Assert.Equal(" 代码 ", doc2.Blocks[0].Inlines[0].Text);
    }

    [Fact]
    [DisplayName("解析：双反引号包裹代码还原含反引号内容")]
    public void Parse_DoubleBacktickCode()
    {
        var doc = MarkdownDocument.Parse("`` a`b ``");
        Assert.Equal(MarkdownInlineType.Code, doc.Blocks[0].Inlines[0].Type);
        Assert.Equal("a`b", doc.Blocks[0].Inlines[0].Text);
    }

    #endregion

    #region 表格与 URL

    [Fact]
    [DisplayName("写入：表格单元格含管道符自动转义")]
    public void Write_TableCellEscapesPipe()
    {
        var doc = new MarkdownDocument();
        var table = new TableBlock();

        var header = new TableRowBlock();
        header.Children.Add(new TableCellBlock([MarkdownInline.CreateText("名称")], isHeader: true));
        header.Children.Add(new TableCellBlock([MarkdownInline.CreateText("值")], isHeader: true));
        table.Children.Add(header);

        var row = new TableRowBlock();
        row.Children.Add(new TableCellBlock([MarkdownInline.CreateText("a|b")], isHeader: false));
        row.Children.Add(new TableCellBlock([MarkdownInline.CreateText("c")], isHeader: false));
        table.Children.Add(row);

        doc.Blocks.Add(table);

        var md = doc.ToMarkdown();
        Assert.Contains("a\\|b", md);
        // 管道符被转义后，表格仍为 2 列结构（重新解析见 Phase D 转义管道支持）
        Assert.DoesNotContain("| a|b |", md);
    }

    [Fact]
    [DisplayName("写入：链接 URL 含空格/括号用尖括号包裹")]
    public void Write_LinkWithSpaceInUrl()
    {
        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlock.CreateParagraph([
            MarkdownInline.CreateLink("my page.html", "", [MarkdownInline.CreateText("链接")]),
            MarkdownInline.CreateText(" "),
            MarkdownInline.CreateLink("https://x.com/a(b)", "", [MarkdownInline.CreateText("括号")]),
        ]));

        var md = doc.ToMarkdown();
        Assert.Contains("](<my page.html>)", md);
        Assert.Contains("](<https://x.com/a(b)>)", md);

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal("my page.html", doc2.Blocks[0].Inlines[0].Href);
        Assert.Equal("https://x.com/a(b)", doc2.Blocks[0].Inlines[2].Href);
    }

    #endregion

    #region 往返模式

    [Fact]
    [DisplayName("往返模式：修改块后序列化反映修改，未修改块保留原始格式")]
    public void Roundtrip_ModifiedReflectsChange()
    {
        var doc = MarkdownDocument.ParseRoundtrip("# 标题\n\n原始段落\n");

        // 修改第二个块（段落）的文本
        var para = doc.Blocks[1];
        para.Inlines.Clear();
        para.Inlines.Add(MarkdownInline.CreateText("新段落"));

        var md = doc.ToMarkdown();
        Assert.Contains("# 标题", md);
        Assert.Contains("新段落", md);
        Assert.DoesNotContain("原始段落", md);
    }

    [Fact]
    [DisplayName("往返模式：未修改文档保留块间空行")]
    public void Roundtrip_PreservesBlankLines()
    {
        var src = "第一段\n\n第二段\n\n第三段\n";
        var doc = MarkdownDocument.ParseRoundtrip(src);

        var md = doc.ToMarkdown();
        Assert.Contains("\n\n", md);

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal(3, doc2.Blocks.Count);
        Assert.Equal("第一段", doc2.Blocks[0].GetPlainText());
        Assert.Equal("第二段", doc2.Blocks[1].GetPlainText());
        Assert.Equal("第三段", doc2.Blocks[2].GetPlainText());
    }

    [Fact]
    [DisplayName("往返模式：未修改的列表保留原始项目符号")]
    public void Roundtrip_PreservesListBullet()
    {
        var src = "* 第一项\n* 第二项\n";
        var doc = MarkdownDocument.ParseRoundtrip(src);

        var md = doc.ToMarkdown();
        // 原始列表用 * 作为项目符号，往返后应保留而非默认 -
        Assert.StartsWith("* ", md.TrimStart());
        Assert.Contains("* 第二项", md);
    }

    [Fact]
    [DisplayName("往返模式：新增块正常序列化")]
    public void Roundtrip_AppendedBlockSerialized()
    {
        var doc = MarkdownDocument.ParseRoundtrip("第一段\n");

        // 追加新块
        doc.Blocks.Add(MarkdownBlock.CreateHeading(2, [MarkdownInline.CreateText("新标题")]));

        var md = doc.ToMarkdown();
        Assert.Contains("## 新标题", md);
        Assert.Contains("第一段", md);
    }

    #endregion
}
