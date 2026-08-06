using System.ComponentModel;
using System.IO;
using System.Text;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>Markdown 编码检测与边界增强测试（Phase F）</summary>
/// <remarks>
/// 覆盖文件编码自动检测（GBK/BOM/UTF-16）、强调嵌套、行首 Email 自动链接、
/// FrontMatter 增强（数组/注释/引号冒号）与 Tab 缩进支持。
/// </remarks>
public class MarkdownEncodingTests
{
    static MarkdownEncodingTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    #region 编码检测

    [Fact]
    [DisplayName("编码：GBK 文件解析")]
    public void Encoding_GbkFile()
    {
        var path = Path.GetTempFileName();
        try
        {
            var gbk = Encoding.GetEncoding("GBK");
            File.WriteAllBytes(path, gbk.GetBytes("中文标题\n\n这是 GBK 编码的段落。\n"));
            var doc = MarkdownDocument.ParseFile(path);
            Assert.Equal("中文标题", doc.Blocks[0].GetPlainText());
            Assert.Equal("这是 GBK 编码的段落。", doc.Blocks[1].GetPlainText());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    [DisplayName("编码：UTF-8 BOM 文件解析")]
    public void Encoding_Utf8BomFile()
    {
        var path = Path.GetTempFileName();
        try
        {
            var bytes = new List<Byte> { 0xEF, 0xBB, 0xBF };
            bytes.AddRange(Encoding.UTF8.GetBytes("# 标题\n\n正文\n"));
            File.WriteAllBytes(path, bytes.ToArray());
            var doc = MarkdownDocument.ParseFile(path);
            Assert.Equal("标题", doc.Blocks[0].GetPlainText());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    [DisplayName("编码：UTF-16 LE 文件解析")]
    public void Encoding_Utf16File()
    {
        var path = Path.GetTempFileName();
        try
        {
            // WriteAllText 带 BOM 写出 UTF-16 LE（符合实际 UTF-16 文件特征）
            File.WriteAllText(path, "# 标题\n\n正文\n", Encoding.Unicode);
            var doc = MarkdownDocument.ParseFile(path);
            Assert.Equal("标题", doc.Blocks[0].GetPlainText());
            Assert.Equal("正文", doc.Blocks[1].GetPlainText());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    #endregion

    #region 强调嵌套

    [Fact]
    [DisplayName("强调：*外层* 嵌套 **粗体**")]
    public void Emphasis_NestedStrongInEm()
    {
        var doc = MarkdownDocument.Parse("*foo **bar** baz*");
        var inline = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Emphasis, inline.Type);
        Assert.Equal(3, inline.Children.Count);
        Assert.Equal(MarkdownInlineType.Strong, inline.Children[1].Type);
        Assert.Equal("bar", inline.Children[1].GetPlainText());
    }

    [Fact]
    [DisplayName("强调：**外层** 嵌套 *斜体*")]
    public void Emphasis_NestedEmInStrong()
    {
        var doc = MarkdownDocument.Parse("**foo *bar* baz**");
        var inline = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Strong, inline.Type);
        Assert.Equal(3, inline.Children.Count);
        Assert.Equal(MarkdownInlineType.Emphasis, inline.Children[1].Type);
        Assert.Equal("bar", inline.Children[1].GetPlainText());
    }

    #endregion

    #region Email 自动链接

    [Fact]
    [DisplayName("自动链接：行首 Email")]
    public void AutoLink_EmailAtStart()
    {
        var doc = MarkdownDocument.Parse("test@example.com 邮箱");
        var inline = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, inline.Type);
        Assert.Equal("mailto:test@example.com", inline.Href);
    }

    #endregion

    #region FrontMatter 增强

    [Fact]
    [DisplayName("FrontMatter：数组值")]
    public void FrontMatter_Array()
    {
        var doc = MarkdownDocument.Parse("---\ntags:\n  - a\n  - b\ntitle: 测试\n---\n\n正文\n");
        Assert.Equal("a, b", doc.FrontMatter["tags"]);
        Assert.Equal("测试", doc.FrontMatter["title"]);
    }

    [Fact]
    [DisplayName("FrontMatter：注释行")]
    public void FrontMatter_Comment()
    {
        var doc = MarkdownDocument.Parse("---\n# 注释\ntitle: 测试\n---\n");
        Assert.Equal("测试", doc.FrontMatter["title"]);
        Assert.Equal(1, doc.FrontMatter.Count);
    }

    [Fact]
    [DisplayName("FrontMatter：引号值内冒号")]
    public void FrontMatter_QuotedValueWithColon()
    {
        var doc = MarkdownDocument.Parse("---\ntitle: \"Hello: World\"\n---\n");
        Assert.Equal("Hello: World", doc.FrontMatter["title"]);
    }

    #endregion

    #region Tab 缩进

    [Fact]
    [DisplayName("Tab：缩进列表续行")]
    public void Tab_ListContinuation()
    {
        var doc = MarkdownDocument.Parse("- 第一行\n\t第二行\n");
        var list = (BulletListBlock)doc.Blocks[0];
        Assert.Equal("第一行 第二行", list.Children[0].GetPlainText());
    }

    [Fact]
    [DisplayName("Tab：制表符缩进代码块")]
    public void Tab_IndentedCode()
    {
        var doc = MarkdownDocument.Parse("\tcode\n");
        var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("code", code.RawText);
    }

    #endregion
}
