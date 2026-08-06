using System.ComponentModel;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 管线开关测试（Phase E）</summary>
/// <remarks>
/// 验证 MarkdownPipeline 的启用/禁用开关真实影响解析行为
/// （此前 Pipeline 属性从未被解析器读取，开关无效）。
/// </remarks>
public class MarkdownPipelineTests
{
    #region Strict 管线（全部关闭）

    [Fact]
    [DisplayName("Strict管线：表格保持段落")]
    public void Strict_TablesOff()
    {
        var doc = MarkdownDocument.Parse("| A | B |\n|---|---|\n", MarkdownPipeline.CreateStrict());
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("Strict管线：删除线不解析")]
    public void Strict_StrikethroughOff()
    {
        var doc = MarkdownDocument.Parse("~~删除~~\n", MarkdownPipeline.CreateStrict());
        var inline = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Text, inline.Type);
        Assert.Equal("~~删除~~", inline.Text);
    }

    [Fact]
    [DisplayName("Strict管线：脚注不解析")]
    public void Strict_FootnotesOff()
    {
        var doc = MarkdownDocument.Parse("正文[^1]\n\n[^1]: 内容\n", MarkdownPipeline.CreateStrict());
        Assert.Equal(1, doc.Blocks[0].Inlines.Count);
        Assert.Equal("正文[^1]", doc.Blocks[0].Inlines[0].Text);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[1].Type);
    }

    [Fact]
    [DisplayName("Strict管线：数学不解析")]
    public void Strict_MathOff()
    {
        var doc = MarkdownDocument.Parse("$x=1$\n", MarkdownPipeline.CreateStrict());
        Assert.Equal("$x=1$", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Strict管线：Emoji不解析")]
    public void Strict_EmojiOff()
    {
        var doc = MarkdownDocument.Parse(":smile:\n", MarkdownPipeline.CreateStrict());
        Assert.Equal(":smile:", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Strict管线：自动链接不解析")]
    public void Strict_AutoLinksOff()
    {
        var doc = MarkdownDocument.Parse("访问 https://newlifex.com\n", MarkdownPipeline.CreateStrict());
        Assert.Equal("访问 https://newlifex.com", doc.Blocks[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Strict管线：YAML FrontMatter不解析")]
    public void Strict_FrontMatterOff()
    {
        var doc = MarkdownDocument.Parse("---\ntitle: 测试\n---\n\n正文\n", MarkdownPipeline.CreateStrict());
        Assert.Empty(doc.FrontMatter);
        // --- 成为分隔线
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.ThematicBreak);
    }

    [Fact]
    [DisplayName("Strict管线：任务列表不解析")]
    public void Strict_TaskListsOff()
    {
        var doc = MarkdownDocument.Parse("- [x] 完成\n", MarkdownPipeline.CreateStrict());
        var list = (BulletListBlock)doc.Blocks[0];
        var item = (ListItemBlock)list.Children[0];
        Assert.False(item.IsTaskItem);
        Assert.Equal("[x] 完成", item.GetPlainText());
    }

    [Fact]
    [DisplayName("Strict管线：定义列表不解析")]
    public void Strict_DefinitionListsOff()
    {
        var doc = MarkdownDocument.Parse("术语\n: 描述\n", MarkdownPipeline.CreateStrict());
        Assert.DoesNotContain(doc.Blocks, b => b.Type == MarkdownBlockType.DefinitionList);
    }

    #endregion

    #region GFM / 默认管线

    [Fact]
    [DisplayName("默认管线：表格正常解析")]
    public void Default_AllEnabled()
    {
        var doc = MarkdownDocument.Parse("| A | B |\n|---|---|\n", new MarkdownPipeline());
        Assert.Equal(MarkdownBlockType.Table, doc.Blocks[0].Type);
    }

    [Fact]
    [DisplayName("GFM管线：表格/任务/删除线解析，脚注关闭")]
    public void Gfm_EnabledAndDisabled()
    {
        var doc = MarkdownDocument.Parse("| A | B |\n|---|---|\n", MarkdownPipeline.CreateGfm());
        Assert.Equal(MarkdownBlockType.Table, doc.Blocks[0].Type);

        var doc2 = MarkdownDocument.Parse("~~删除~~\n", MarkdownPipeline.CreateGfm());
        Assert.Equal(MarkdownInlineType.Strikethrough, doc2.Blocks[0].Inlines[0].Type);

        var doc3 = MarkdownDocument.Parse("正文[^1]\n\n[^1]: 内容\n", MarkdownPipeline.CreateGfm());
        Assert.Equal(2, doc3.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, doc3.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc3.Blocks[1].Type);
    }

    #endregion
}
