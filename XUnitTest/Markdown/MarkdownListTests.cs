using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 列表与缩进代码块解析测试（Phase B）</summary>
/// <remarks>
/// 覆盖列表续行合并、嵌套列表、项内多段落/引用/代码块、有序列表续行、
/// 缩进代码块（含空行/结束/优先级）与深层嵌套保护。
/// </remarks>
public class MarkdownListTests
{
    #region 列表

    [Fact]
    [DisplayName("列表：多行列表项续行合并为段落")]
    public void List_ContinuationLines_MergeIntoParagraph()
    {
        var doc = MarkdownDocument.Parse("- 第一行\n  第二行\n- 下一项\n");
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[0].Type);
        var list = (BulletListBlock)doc.Blocks[0];
        Assert.Equal(2, list.Children.Count);
        Assert.Equal("第一行 第二行", list.Children[0].GetPlainText());
        Assert.Equal("下一项", list.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("列表：嵌套列表")]
    public void List_NestedList()
    {
        var doc = MarkdownDocument.Parse("- 父项\n  - 子项1\n  - 子项2\n- 兄弟\n");
        var list = (BulletListBlock)doc.Blocks[0];
        Assert.Equal(2, list.Children.Count);
        var parent = list.Children[0];
        // 父项 = 段落 + 嵌套列表 两个子块
        Assert.Equal(2, parent.Children.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, parent.Children[0].Type);
        Assert.Equal(MarkdownBlockType.BulletList, parent.Children[1].Type);
        Assert.Equal("父项", parent.Children[0].GetPlainText());
        var nested = (BulletListBlock)parent.Children[1];
        Assert.Equal(2, nested.Children.Count);
        Assert.Equal("子项1", nested.Children[0].GetPlainText());
        Assert.Equal("子项2", nested.Children[1].GetPlainText());
        Assert.Equal("兄弟", list.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("列表：项内多段落（loose）")]
    public void List_MultiParagraphInItem()
    {
        var doc = MarkdownDocument.Parse("- 第一段\n\n  第二段\n- 下一项\n");
        var list = (BulletListBlock)doc.Blocks[0];
        var item0 = list.Children[0];
        Assert.Equal(2, item0.Children.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, item0.Children[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, item0.Children[1].Type);
        Assert.Equal("第一段", item0.Children[0].GetPlainText());
        Assert.Equal("第二段", item0.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("列表：项内引用块")]
    public void List_BlockQuoteInItem()
    {
        var doc = MarkdownDocument.Parse("- 正文\n  > 引用\n");
        var list = (BulletListBlock)doc.Blocks[0];
        var item0 = list.Children[0];
        Assert.Equal(2, item0.Children.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, item0.Children[0].Type);
        Assert.Equal(MarkdownBlockType.BlockQuote, item0.Children[1].Type);
        Assert.Equal("引用", item0.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("列表：项内缩进代码块")]
    public void List_CodeBlockInItem()
    {
        var doc = MarkdownDocument.Parse("- 正文\n\n      var x = 1;\n");
        var list = (BulletListBlock)doc.Blocks[0];
        var item0 = list.Children[0];
        Assert.Equal(2, item0.Children.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, item0.Children[0].Type);
        Assert.Equal(MarkdownBlockType.CodeBlock, item0.Children[1].Type);
        Assert.Equal("var x = 1;", ((CodeBlock)item0.Children[1]).RawText);
    }

    [Fact]
    [DisplayName("列表：有序列表续行与起始序号")]
    public void List_OrderedContinuation()
    {
        var doc = MarkdownDocument.Parse("3. 第一行\n   续行\n4. 第二项\n");
        var list = (OrderedListBlock)doc.Blocks[0];
        Assert.Equal(3, list.OrderedStart);
        Assert.Equal(2, list.Children.Count);
        Assert.Equal("第一行 续行", list.Children[0].GetPlainText());
        Assert.Equal("第二项", list.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("列表：不同项目符号结束列表")]
    public void List_DifferentBullet_EndsList()
    {
        var doc = MarkdownDocument.Parse("- 连字符\n* 星号\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.BulletList, doc.Blocks[1].Type);
        Assert.Equal(1, doc.Blocks[0].Children.Count);
    }

    [Fact]
    [DisplayName("列表：任务列表项解析")]
    public void List_TaskItems()
    {
        var doc = MarkdownDocument.Parse("- [x] 已完成\n- [ ] 未完成\n");
        var list = (BulletListBlock)doc.Blocks[0];
        var item0 = (ListItemBlock)list.Children[0];
        Assert.True(item0.IsTaskItem);
        Assert.True(item0.IsChecked);
        Assert.Equal("已完成", item0.GetPlainText());
        var item1 = (ListItemBlock)list.Children[1];
        Assert.True(item1.IsTaskItem);
        Assert.False(item1.IsChecked);
        Assert.Equal("未完成", item1.GetPlainText());
    }

    #endregion

    #region 缩进代码块

    [Fact]
    [DisplayName("缩进代码块：基本解析")]
    public void IndentedCode_Basic()
    {
        var doc = MarkdownDocument.Parse("    code line 1\n    code line 2\n");
        Assert.Equal(1, doc.Blocks.Count);
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("code line 1\ncode line 2", code.RawText);
    }

    [Fact]
    [DisplayName("缩进代码块：含空行保留")]
    public void IndentedCode_BlankLinesPreserved()
    {
        var doc = MarkdownDocument.Parse("    a\n\n    b\n");
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("a\n\nb", code.RawText);
    }

    [Fact]
    [DisplayName("缩进代码块：非缩进行结束")]
    public void IndentedCode_EndsAtNonIndented()
    {
        var doc = MarkdownDocument.Parse("    code\nparagraph\n");
        Assert.Equal(2, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.CodeBlock, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
    }

    [Fact]
    [DisplayName("缩进代码块：4 空格内列表标记视为代码")]
    public void IndentedCode_PrecedesListMarker()
    {
        var doc = MarkdownDocument.Parse("    - not a list\n");
        Assert.Equal(1, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.CodeBlock, doc.Blocks[0].Type);
        Assert.Equal("- not a list", ((CodeBlock)doc.Blocks[0]).RawText);
    }

    [Fact]
    [DisplayName("缩进代码块：制表符缩进")]
    public void IndentedCode_TabIndent()
    {
        var doc = MarkdownDocument.Parse("\tcode\n");
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("code", code.RawText);
    }

    #endregion

    #region 健壮性

    [Fact]
    [DisplayName("深度保护：深层嵌套列表不崩溃")]
    public void List_DeepNesting_NoCrash()
    {
        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < 100; i++)
        {
            sb.Append("- item").AppendLine();
            sb.Append(new String(' ', 2 * (i + 1)));
        }
        sb.Append("- leaf\n");
        var doc = MarkdownDocument.Parse(sb.ToString());
        Assert.NotNull(doc);
        Assert.NotEmpty(doc.Blocks);
    }

    [Fact]
    [DisplayName("健壮性：列表续行包含围栏代码块")]
    public void List_FencedCodeInItem()
    {
        var doc = MarkdownDocument.Parse("- 说明\n\n  ```csharp\n  var x = 1;\n  ```\n");
        var list = (BulletListBlock)doc.Blocks[0];
        var item0 = list.Children[0];
        Assert.Equal(2, item0.Children.Count);
        Assert.Equal(MarkdownBlockType.CodeBlock, item0.Children[1].Type);
        Assert.Equal("csharp", ((CodeBlock)item0.Children[1]).Language);
        Assert.Equal("var x = 1;", ((CodeBlock)item0.Children[1]).RawText);
    }

    #endregion
}
