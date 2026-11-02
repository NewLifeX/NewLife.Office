using System.Text;
using NewLife.Office.Word;
using Xunit;

namespace XUnitTest.Word;

/// <summary>静态目录与 Markdown 多级列表测试</summary>
public class TocAndMarkdownTests
{
    #region 静态目录
    [Fact(DisplayName = "静态目录-按标题生成缩进条目并插入到最前")]
    public void AddStaticToc_InsertsEntries()
    {
        var doc = new Document();
        doc.AddHeading("第一章", 1);
        doc.AddHeading("1.1 小节", 2);
        doc.AddParagraph("正文内容");

        doc.BeginEdit();
        var title = doc.AddStaticTableOfContents(levels: 3, title: "目录");

        Assert.NotNull(title);
        // 目录标题 + 2 条条目 + 原有 3 段
        Assert.Equal(6, doc.Elements.Count);
        Assert.Equal("目录", doc.Elements[0].Paragraph!.Runs[0].Text);
        Assert.Equal("第一章", doc.Elements[1].Paragraph!.Runs[0].Text);
        Assert.Null(doc.Elements[1].Paragraph!.IndentLeft);
        Assert.Equal("1.1 小节", doc.Elements[2].Paragraph!.Runs[0].Text);
        Assert.Equal(360, doc.Elements[2].Paragraph!.IndentLeft);
    }

    [Fact(DisplayName = "静态目录-按 levels 过滤深层标题")]
    public void AddStaticToc_FiltersByLevel()
    {
        var doc = new Document();
        doc.AddHeading("一级", 1);
        doc.AddHeading("三级", 3);
        doc.BeginEdit();

        doc.AddStaticTableOfContents(levels: 1, title: null);

        var texts = doc.Elements.Where(e => e.Paragraph != null).Select(e => e.Paragraph!.Runs.Count > 0 ? e.Paragraph.Runs[0].Text : "").ToList();
        Assert.Equal(3, doc.Elements.Count);          // 1 条条目 + 原有两个标题
        Assert.Equal("一级", texts[0]);
        Assert.Equal("三级", texts[2]);
        Assert.Null(doc.Elements[0].Paragraph!.IndentLeft);
    }

    [Fact(DisplayName = "静态目录-无标题时返回 null 且不插入")]
    public void AddStaticToc_NoHeadings_ReturnsNull()
    {
        var doc = new Document();
        doc.AddParagraph("只有正文");
        doc.BeginEdit();

        Assert.Null(doc.AddStaticTableOfContents());
        Assert.Single(doc.Elements);
    }
    #endregion

    #region Markdown 多级列表
    [Fact(DisplayName = "调试-直接构造列表块可渲染")]
    public void Debug_DirectListBlock()
    {
        var md = new NewLife.Office.Markdown.MarkdownDocument();
        md.Blocks.Add(NewLife.Office.Markdown.MarkdownBlock.CreateBulletList(
        [
            NewLife.Office.Markdown.MarkdownBlock.CreateListItem([NewLife.Office.Markdown.MarkdownInline.CreateText("一级条目")])
        ]));

        Assert.Contains("一级条目", md.ToMarkdown());
    }

    [Fact(DisplayName = "Markdown-多级无序列表按层级嵌套")]
    public void Markdown_NestedBulletList()
    {
        var doc = new Document();
        var p1 = doc.AddParagraph("一级条目");
        p1.Paragraph!.IsBullet = true;
        p1.Paragraph!.ListLevel = 0;

        var p2 = doc.AddParagraph("二级条目");
        p2.Paragraph!.IsBullet = true;
        p2.Paragraph!.ListLevel = 1;

        var p3 = doc.AddParagraph("回到一级");
        p3.Paragraph!.IsBullet = true;
        p3.Paragraph!.ListLevel = 0;

        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
            writer.Save(ms, doc);

        var bytes = ms.ToArray();

        // 重新读取校验：列表标记与层级需保住
        Document back;
        using (var r = new WordReader(new MemoryStream(bytes)))
            back = r.ReadDocument();
        var paras = back.AllParagraphs();
        Assert.Equal(3, paras.Count);
        Assert.True(paras[0].IsBullet);
        Assert.Equal(0, paras[0].ListLevel);
        Assert.True(paras[1].IsBullet);
        Assert.Equal(1, paras[1].ListLevel);

        String? md;
        using (var r2 = new WordReader(new MemoryStream(bytes)))
            md = r2.ExtractMarkdown();
        Assert.NotNull(md);

        var lines = md!.Replace("\r\n", "\n").Split('\n');
        var top = lines.First(l => l.Contains("一级条目"));
        var child = lines.First(l => l.Contains("二级条目"));
        Assert.True(LeadingSpaces(child) > LeadingSpaces(top), md);
        Assert.Equal(LeadingSpaces(top), LeadingSpaces(lines.First(l => l.Contains("回到一级"))));
    }

    private static Int32 LeadingSpaces(String line)
    {
        var count = 0;
        while (count < line.Length && line[count] == ' ') count++;
        return count;
    }

    [Fact(DisplayName = "Markdown-单级列表行为不变")]
    public void Markdown_FlatBulletList()
    {
        var doc = new Document();
        foreach (var text in new[] { "A", "B" })
        {
            var el = doc.AddParagraph(text);
            el.Paragraph!.IsBullet = true;
        }

        using var ms = new MemoryStream();
        using (var writer = new WordWriter())
            writer.Save(ms, doc);

        ms.Position = 0;
        using var reader = new WordReader(ms);
        var md = reader.ExtractMarkdown();

        var lines = md!.Replace("\r\n", "\n").Split('\n');
        var lineA = lines.First(l => l.Contains("A") && l.Contains("-"));
        var lineB = lines.First(l => l.Contains("B") && l.Contains("-"));
        Assert.Equal(LeadingSpaces(lineA), LeadingSpaces(lineB));
    }
    #endregion
}
