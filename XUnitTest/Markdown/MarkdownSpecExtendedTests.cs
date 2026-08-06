using System.ComponentModel;
using System.Text;
using NewLife.Office.Markdown;
using Xunit;

namespace XUnitTest.Markdown;

/// <summary>MD10 全面完善回归测试：实体/引用图片/删除线/HTML块中断/表格/自动链接/强调分隔符栈/往返/HTML→MD/转换中枢</summary>
/// <remarks>覆盖 2026-08-06 读写全面完善（MD10）全部新增与修复行为，防止回归。</remarks>
public class MarkdownSpecExtendedTests
{
    static MarkdownSpecExtendedTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    private static MarkdownDocument P(String md) => MarkdownDocument.Parse(md);

    #region 1.1 HTML 实体

    [Fact, DisplayName("实体：HTML5 全表具名实体解码")]
    public void Entity_FullNamedSet()
    {
        // Latin-1 扩展
        Assert.Equal("é", P("&eacute;").Blocks[0].GetPlainText());
        Assert.Equal("ü", P("&uuml;").Blocks[0].GetPlainText());
        Assert.Equal("ñ", P("&ntilde;").Blocks[0].GetPlainText());
        // 希腊字母
        Assert.Equal("α", P("&alpha;").Blocks[0].GetPlainText());
        Assert.Equal("β", P("&beta;").Blocks[0].GetPlainText());
        // 数学符号
        Assert.Equal("≤", P("&le;").Blocks[0].GetPlainText());
        Assert.Equal("∑", P("&sum;").Blocks[0].GetPlainText());
        // 货币/其他常用（超出旧 12 字符限制的实体内含长名也应解码）
        Assert.Equal("€", P("&euro;").Blocks[0].GetPlainText());
        Assert.Equal("♥", P("&hearts;").Blocks[0].GetPlainText());
    }

    [Fact, DisplayName("实体：补充平面数字引用解码为代理对")]
    public void Entity_SurrogatePair()
    {
        // &#128512; = U+1F600 表情，需代理对
        Assert.Equal("😀", P("&#128512;").Blocks[0].GetPlainText());
        Assert.Equal("😀", P("&#x1F600;").Blocks[0].GetPlainText());
    }

    [Fact, DisplayName("实体：无效数字码点解码为 U+FFFD")]
    public void Entity_InvalidNumeric()
    {
        // 0 / 代理区 / 超出 0x10FFFF → U+FFFD（CommonMark）
        Assert.Equal("\uFFFD", P("&#0;").Blocks[0].GetPlainText());
        Assert.Equal("\uFFFD", P("&#xD800;").Blocks[0].GetPlainText());
        Assert.Equal("\uFFFD", P("&#x110000;").Blocks[0].GetPlainText());
    }

    [Fact, DisplayName("实体：HTML→MD 共享解码器行为一致")]
    public void Entity_HtmlConverterShared()
    {
        var md = HtmlToMarkdownConverter.DecodeHtmlEntities("&copy; &eacute; &#128512;");
        Assert.Equal("© é 😀", md);
        Assert.Equal("&unknown;", HtmlToMarkdownConverter.DecodeHtmlEntities("&unknown;"));
    }

    #endregion

    #region 1.2 引用图片

    [Fact, DisplayName("引用图片：完整引用 ![alt][id]")]
    public void RefImage_Full()
    {
        var doc = P("![Logo][logo]\n\n[logo]: https://x.com/logo.png \"Logo 图\"\n");
        var img = Assert.IsType<MarkdownInline>(doc.Blocks[0].Inlines[0]);
        Assert.Equal(MarkdownInlineType.Image, img.Type);
        Assert.Equal("https://x.com/logo.png", img.Href);
        Assert.Equal("Logo", img.Alt);
    }

    [Fact, DisplayName("引用图片：折叠引用 ![alt][] 与快捷引用 ![alt]")]
    public void RefImage_CollapsedAndShortcut()
    {
        var doc = P("![图][] 与 ![图]\n\n[图]: /img/a.png\n");
        Assert.Equal(3, doc.Blocks[0].Inlines.Count);
        var img1 = doc.Blocks[0].Inlines[0];
        var img2 = doc.Blocks[0].Inlines[2];
        Assert.Equal(MarkdownInlineType.Image, img1.Type);
        Assert.Equal(MarkdownInlineType.Image, img2.Type);
        Assert.Equal("/img/a.png", img1.Href);
        Assert.Equal("/img/a.png", img2.Href);
    }

    [Fact, DisplayName("引用图片：定义在后使用在前 + 未定义保持文本")]
    public void RefImage_DefinedLaterAndUndefined()
    {
        var doc = P("![后定义][r]\n\n[r]: /img/r.png\n");
        Assert.Equal(MarkdownInlineType.Image, doc.Blocks[0].Inlines[0].Type);

        var doc2 = P("![未定义][missing]\n");
        Assert.Equal(MarkdownInlineType.Text, doc2.Blocks[0].Inlines[0].Type);
    }

    #endregion

    #region 1.3 删除线边界

    [Fact, DisplayName("删除线：首尾空白不解析、空内容不解析")]
    public void Strike_Boundary()
    {
        // GFM：~~ foo ~~ 不解析（内容以空白首尾）
        var doc = P("~~ foo ~~");
        Assert.Equal(MarkdownInlineType.Text, doc.Blocks[0].Inlines[0].Type);
        // ~~ ~~ 空内容不解析
        var doc2 = P("~~ ~~");
        Assert.Equal(MarkdownInlineType.Text, doc2.Blocks[0].Inlines[0].Type);
        // ~~foo~~ 正常解析
        var doc3 = P("~~foo~~");
        Assert.Equal(MarkdownInlineType.Strikethrough, doc3.Blocks[0].Inlines[0].Type);
    }

    [Fact, DisplayName("删除线：~~a b~~ 内嵌空白允许、~~~ 是围栏代码块")]
    public void Strike_InteriorAndTriple()
    {
        var doc = P("~~a b~~");
        Assert.Equal(MarkdownInlineType.Strikethrough, doc.Blocks[0].Inlines[0].Type);
        // ~~~ 是围栏代码块起始（非删除线）；同行剩余作为 info string
        var doc2 = P("~~~x~~~");
        Assert.Equal(MarkdownBlockType.CodeBlock, doc2.Blocks[0].Type);
        Assert.StartsWith("x", ((CodeBlock)doc2.Blocks[0]).Language);
    }

    #endregion

    #region 1.4 段落被 HTML 块中断

    [Fact, DisplayName("HTML 块：块级标签中断段落")]
    public void HtmlBlock_InterruptsParagraph()
    {
        var doc = P("第一段\n<div>内容</div>\n\n第二段\n");
        Assert.Equal(3, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[0].Type);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[1].Type);
        Assert.Equal(MarkdownBlockType.Paragraph, doc.Blocks[2].Type);
    }

    [Fact, DisplayName("HTML 块：特殊标签/注释中断段落")]
    public void HtmlBlock_SpecialInterrupts()
    {
        var doc = P("前文\n<!-- 注释 -->\n后文\n");
        Assert.Equal(3, doc.Blocks.Count);
        Assert.Equal(MarkdownBlockType.HtmlBlock, doc.Blocks[1].Type);
    }

    #endregion

    #region 1.5 表格

    [Fact, DisplayName("表格：单元格代码段内管道符不拆分")]
    public void Table_CodePipe()
    {
        var doc = P("| 列 | 值 |\n| --- | --- |\n| `a|b` | 2 |\n");
        var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
        var row = table.Children[1];
        var cell = Assert.IsType<TableCellBlock>(row.Children[0]);
        // 代码段内含 |，应为一个代码内联
        Assert.Single(cell.Inlines);
        Assert.Equal(MarkdownInlineType.Code, cell.Inlines[0].Type);
        Assert.Equal("a|b", cell.Inlines[0].Text);
    }

    [Fact, DisplayName("表格：分隔行单个连字符合法 + 列数以分隔行为准")]
    public void Table_DelimiterRowRules()
    {
        var doc = P("| A | B | C |\n| - | - | - |\n| 1 | 2 |\n");
        var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
        // 表头 3 列
        Assert.Equal(3, table.Children[0].Children.Count);
        // 数据行 3 列（缺列补空）
        var row = table.Children[1];
        Assert.Equal(3, row.Children.Count);
        Assert.Equal("1", row.Children[0].GetPlainText());
        Assert.Equal("", row.Children[2].GetPlainText());
    }

    #endregion

    #region 1.6 自动链接

    [Fact, DisplayName("自动链接：括号平衡")]
    public void AutoLink_BalancedParens()
    {
        var doc = P("https://en.wikipedia.org/wiki/Train_(disambiguation) 查看");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        Assert.Equal("https://en.wikipedia.org/wiki/Train_(disambiguation)", link.Href);
    }

    [Fact, DisplayName("自动链接：www. 前缀与前置字母限制")]
    public void AutoLink_WwwRules()
    {
        // www. 正常（位于文本中间 → 第二个内联）
        var doc = P("访问 www.newlifex.com 网站");
        Assert.Equal(MarkdownInlineType.Link, doc.Blocks[0].Inlines[1].Type);
        Assert.Equal("http://www.newlifex.com", doc.Blocks[0].Inlines[1].Href);
        // 前有字母不识别
        var doc2 = P("abwww.x.com");
        Assert.Equal(MarkdownInlineType.Text, doc2.Blocks[0].Inlines[0].Type);
    }

    [Fact, DisplayName("自动链接：链接内部不识别裸 URL（GFM）")]
    public void AutoLink_NotInsideLink()
    {
        var doc = P("[点击 https://example.com 进入](https://example.com)");
        var link = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Link, link.Type);
        // 链接文本内的 URL 不应再生成嵌套链接（整段保持为一个文本节点）
        Assert.Single(link.Children);
        Assert.Equal(MarkdownInlineType.Text, link.Children[0].Type);
        Assert.Equal("点击 https://example.com 进入", link.Children[0].Text);
    }

    [Fact, DisplayName("自动链接：尖括号自动链接 scheme 校验")]
    public void AutoLink_AngleScheme()
    {
        var doc = P("<https://newlifex.com> 与 <foo>");
        var inlines = doc.Blocks[0].Inlines;
        Assert.Equal(MarkdownInlineType.Link, inlines[0].Type);
        Assert.Equal("https://newlifex.com", inlines[0].Href);
        // <foo> 无合法 scheme/email → 行内 HTML 或文本，不构成自动链接
        Assert.DoesNotContain(inlines, i => i.Type == MarkdownInlineType.Link && i.Href == "foo");
    }

    #endregion

    #region 1.7 强调分隔符栈

    [Fact, DisplayName("强调：intraword 下划线闭合不解析（snake_case）")]
    public void Emphasis_IntrawordClosing()
    {
        // _foo bar_baz：闭合 _ 后跟字母 → 不构成强调
        var doc = P("_foo bar_baz");
        Assert.Equal(MarkdownInlineType.Text, doc.Blocks[0].Inlines[0].Type);
    }

    [Fact, DisplayName("强调：**foo* 与 *foo**（rule of 3 与余量）")]
    public void Emphasis_Unbalanced()
    {
        // **foo* → *<em>foo</em>（开 2 闭 1，用 1）
        var doc = P("**foo*");
        Assert.Equal(2, doc.Blocks[0].Inlines.Count);
        Assert.Equal(MarkdownInlineType.Emphasis, doc.Blocks[0].Inlines[1].Type);
        // *foo** → <em>foo</em>*
        var doc2 = P("*foo**");
        Assert.Equal(2, doc2.Blocks[0].Inlines.Count);
        Assert.Equal(MarkdownInlineType.Emphasis, doc2.Blocks[0].Inlines[0].Type);
    }

    [Fact, DisplayName("强调：*** 三形态嵌套（Em>Strong）")]
    public void Emphasis_TripleNesting()
    {
        var se = P("***foo***").Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Emphasis, se.Type);
        Assert.Single(se.Children);
        Assert.Equal(MarkdownInlineType.Strong, se.Children[0].Type);
        Assert.Equal("foo", se.Children[0].GetPlainText());
    }

    [Fact, DisplayName("强调：复杂嵌套 *foo **bar** baz* 保持结构")]
    public void Emphasis_ComplexNesting()
    {
        var doc = P("*foo **bar** baz*");
        var em = doc.Blocks[0].Inlines[0];
        Assert.Equal(MarkdownInlineType.Emphasis, em.Type);
        Assert.Equal(3, em.Children.Count);
        Assert.Equal(MarkdownInlineType.Strong, em.Children[1].Type);
    }

    [Fact, DisplayName("强调：分隔符栈序列化往返一致")]
    public void Emphasis_RoundtripText()
    {
        var md = "**粗体** *斜体* ***粗斜体*** ~~删除~~ *嵌套 **粗** 内容*";
        var doc = P(md);
        var outMd = doc.ToMarkdown();
        var doc2 = P(outMd);
        Assert.Equal(doc.Blocks[0].GetPlainText(), doc2.Blocks[0].GetPlainText());
        // 重新解析后内联类型结构一致
        Assert.Equal(doc.Blocks[0].Inlines.Count, doc2.Blocks[0].Inlines.Count);
    }

    #endregion

    #region 2.x 写入与往返

    [Fact, DisplayName("往返：引用定义原位保留（不统一移到底部）")]
    public void Roundtrip_ReferenceInPlace()
    {
        var md = "# 标题\n\n正文 [链接][a] 继续\n\n[a]: https://example.com \"标题\"\n\n末尾段落\n";
        var doc = MarkdownDocument.ParseRoundtrip(md);
        var outMd = doc.ToMarkdown();
        // 引用定义应在原始位置（第三段之前），而非文档末尾
        var refPos = outMd.IndexOf("[a]: https://example.com");
        var lastPos = outMd.IndexOf("末尾段落");
        Assert.True(refPos >= 0, "引用定义应存在");
        Assert.True(refPos < lastPos, "引用定义应保留在原位（末尾段落之前）");
    }

    [Fact, DisplayName("往返：完全未修改时输出与输入一致")]
    public void Roundtrip_Identity()
    {
        var md = "# 标题\n\n- 项目一\n- 项目二\n\n[a]: /ref\n\n> 引用块\n";
        var doc = MarkdownDocument.ParseRoundtrip(md);
        Assert.Equal(md, doc.ToMarkdown());
    }

    [Fact, DisplayName("往返：修改某块后其余保持原格式")]
    public void Roundtrip_ModifiedBlockOnly()
    {
        var md = "# 标题\n\n- 项目一\n- 项目二\n";
        var doc = MarkdownDocument.ParseRoundtrip(md);
        // 修改第二个列表项文本
        var list = (BulletListBlock)doc.Blocks[1];
        var item = (ListItemBlock)list.Children[1];
        item.Inlines.Clear();
        item.Inlines.Add(MarkdownInline.CreateText("已修改"));

        var outMd = doc.ToMarkdown();
        Assert.Contains("- 已修改", outMd);
        Assert.Contains("- 项目一", outMd);
        Assert.StartsWith("# 标题", outMd);
    }

    [Fact, DisplayName("往返：ParseRoundtrip 公开 API 开箱即用")]
    public void Roundtrip_PublicApi()
    {
        var doc = MarkdownDocument.ParseRoundtrip("# 标题");
        Assert.True(doc.Roundtrip);
        var doc2 = MarkdownDocument.Parse("# 标题");
        Assert.False(doc2.Roundtrip);
    }

    [Fact, DisplayName("FrontMatter：特殊值序列化转义并往返一致")]
    public void FrontMatter_EscapedRoundtrip()
    {
        var doc = new MarkdownDocument();
        doc.FrontMatter["title"] = "a: b";
        doc.FrontMatter["note"] = "say \"hi\"";
        doc.FrontMatter["empty"] = "";
        doc.FrontMatter["hash"] = "#hash";
        var md = doc.ToMarkdown();

        var doc2 = MarkdownDocument.Parse(md);
        Assert.Equal("a: b", doc2.FrontMatter["title"]);
        Assert.Equal("say \"hi\"", doc2.FrontMatter["note"]);
        Assert.Equal("", doc2.FrontMatter["empty"]);
        Assert.Equal("#hash", doc2.FrontMatter["hash"]);
    }

    [Fact, DisplayName("写入：表格列数不一致自动补齐")]
    public void Writer_TablePadColumns()
    {
        var doc = new MarkdownDocument();
        var table = new TableBlock();
        var hr = new TableRowBlock();
        hr.Children.Add(new TableCellBlock([MarkdownInline.CreateText("A")], isHeader: true));
        hr.Children.Add(new TableCellBlock([MarkdownInline.CreateText("B")], isHeader: true));
        table.Children.Add(hr);
        var dr = new TableRowBlock();
        dr.Children.Add(new TableCellBlock([MarkdownInline.CreateText("1")]));
        table.Children.Add(dr);
        doc.Blocks.Add(table);

        var md = doc.ToMarkdown();
        var table2 = P(md);
        Assert.IsType<TableBlock>(table2.Blocks[0]);
        var parsedTable = (TableBlock)table2.Blocks[0];
        Assert.Equal(2, parsedTable.Children[1].Children.Count);
        Assert.Equal("", parsedTable.Children[1].Children[1].GetPlainText());
    }

    #endregion

    #region 3.x HTML→MD

    [Fact, DisplayName("HTML→MD：dl/dt/dd 定义列表")]
    public void HtmlToMd_DefinitionList()
    {
        var doc = MarkdownDocument.FromHtml("<dl><dt>术语</dt><dd>定义一</dd><dd>定义二</dd></dl>");
        Assert.Single(doc.Blocks);
        var dl = Assert.IsType<DefinitionListBlock>(doc.Blocks[0]);
        Assert.Equal(3, dl.Children.Count);
        Assert.IsType<DefinitionTermBlock>(dl.Children[0]);
        Assert.IsType<DefinitionDescriptionBlock>(dl.Children[1]);
        Assert.Equal("术语", dl.Children[0].GetPlainText());
        // 序列化为 Markdown 定义列表语法
        var md = doc.ToMarkdown();
        Assert.Contains("术语\n: 定义一", md);
    }

    [Fact, DisplayName("HTML→MD：任务列表 checkbox")]
    public void HtmlToMd_TaskList()
    {
        var html = "<ul><li><input type=\"checkbox\" checked> 已完成</li><li><input type=\"checkbox\"> 未完成</li></ul>";
        var doc = MarkdownDocument.FromHtml(html);
        var list = Assert.IsType<BulletListBlock>(doc.Blocks[0]);
        var item1 = Assert.IsType<ListItemBlock>(list.Children[0]);
        var item2 = Assert.IsType<ListItemBlock>(list.Children[1]);
        Assert.True(item1.IsTaskItem);
        Assert.True(item1.IsChecked);
        Assert.True(item2.IsTaskItem);
        Assert.False(item2.IsChecked);
        // 序列化输出 GFM 任务列表
        var md = doc.ToMarkdown();
        Assert.Contains("- [x] 已完成", md);
        Assert.Contains("- [ ] 未完成", md);
    }

    [Fact, DisplayName("HTML→MD：details/summary 折叠容器内容保留")]
    public void HtmlToMd_Details()
    {
        var html = "<details><summary>更多信息</summary><p>隐藏内容</p></details>";
        var doc = MarkdownDocument.FromHtml(html);
        var md = doc.ToMarkdown();
        Assert.Contains("更多信息", md);
        Assert.Contains("隐藏内容", md);
        // summary 转为粗体段落
        Assert.Contains("**更多信息**", md);
    }

    #endregion

    #region 4.x 转换中枢

    [Fact, DisplayName("转换中枢：.md 文件直通读取")]
    public void Convert_MdPassthrough()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nlo_md_{Guid.NewGuid():N}.md");
        try
        {
            var text = "# 直通测试\n\n原样保留 **格式**。\n";
            File.WriteAllText(path, text, new UTF8Encoding(false));
            var md = FormatToMarkdown.Convert(path);
            Assert.Equal(text, md);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("转换中枢：扩展名未知时按内容识别（文本→Markdown）")]
    public void Convert_ContentDetection()
    {
        // 无扩展名/未知扩展名但内容是 Markdown → 内容识别成功
        var path = Path.Combine(Path.GetTempPath(), $"nlo_md_{Guid.NewGuid():N}.dat");
        try
        {
            File.WriteAllText(path, "# 内容识别\n\n正文", new UTF8Encoding(false));
            var md = FormatToMarkdown.Convert(path);
            Assert.NotNull(md);
            Assert.Contains("# 内容识别", md);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    #endregion

    #region 7.x 非目标验证（Mermaid 保留等）

    [Fact, DisplayName("非目标：mermaid 代码块原样保留（language-mermaid）")]
    public void NonGoal_MermaidPreserved()
    {
        var md = "```mermaid\ngraph TD\n  A-->B\n```\n";
        var doc = P(md);
        var cb = Assert.IsType<CodeBlock>(doc.Blocks[0]);
        Assert.Equal("mermaid", cb.Language);
        Assert.Equal("graph TD\n  A-->B", cb.RawText);
        // 往返保留
        var outMd = doc.ToMarkdown();
        Assert.Contains("```mermaid", outMd);
        Assert.Contains("graph TD", outMd);
        // HTML 输出带 language-mermaid 类（供前端渲染）
        Assert.Contains("language-mermaid", doc.ToHtml());
    }

    #endregion

    #region MD11 优化回归（硬换行/直通编码/紧松列表）

    [Fact, DisplayName("硬换行：标记两空格不进入文本令牌")]
    public void HardBreak_MarkerSpacesStripped()
    {
        // CommonMark：`  \n` 的两空格是换行标记，不属于内容
        var doc = P("a  \nb");
        Assert.Equal("a b", doc.Blocks[0].GetPlainText());
        Assert.Contains(doc.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.HardBreak);
        // 反斜杠硬换行
        var doc2 = P("a\\\nb");
        Assert.Equal("a b", doc2.Blocks[0].GetPlainText());
        Assert.Contains(doc2.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.HardBreak);
        // 软换行保持
        var doc3 = P("a\nb");
        Assert.Contains(doc3.Blocks[0].Inlines, i => i.Type == MarkdownInlineType.SoftBreak);
    }

    [Fact, DisplayName("转换中枢：.md 直通 GBK 编码文件不乱码")]
    public void Convert_MdPassthrough_Encoding()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nlo_md_gbk_{Guid.NewGuid():N}.md");
        try
        {
            var text = "# GBK 测试\n\n中文内容。";
            // 用 GBK 编码写文件（无 BOM），File.ReadAllText 会乱码，ReadFileText 应正确解码
            File.WriteAllText(path, text, Encoding.GetEncoding(936));
            var md = FormatToMarkdown.Convert(path);
            Assert.Contains("GBK 测试", md);
            Assert.Contains("中文内容", md);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact, DisplayName("紧/松列表：项间空行判定 IsLoose")]
    public void List_TightLooseDetection()
    {
        // 紧列表：无空行
        var tight = P("- a\n- b\n- c");
        var bl = Assert.IsType<BulletListBlock>(tight.Blocks[0]);
        Assert.False(bl.IsLoose);

        // 松列表：项间空行
        var loose = P("- a\n\n- b\n- c");
        var bl2 = Assert.IsType<BulletListBlock>(loose.Blocks[0]);
        Assert.True(bl2.IsLoose);

        // 有序列表
        var ol = P("1. a\n\n2. b");
        Assert.True(Assert.IsType<OrderedListBlock>(ol.Blocks[0]).IsLoose);
    }

    [Fact, DisplayName("紧/松列表：HTML 渲染差异（松列表项包裹 p）")]
    public void List_TightLooseHtml()
    {
        var tightHtml = P("- a\n- b").ToHtml();
        Assert.Contains("<li>a</li>", tightHtml);
        Assert.DoesNotContain("<li><p>a</p>", tightHtml);

        var looseHtml = P("- a\n\n- b").ToHtml();
        Assert.Contains("<li><p>a</p></li>", looseHtml);
        Assert.Contains("<li><p>b</p></li>", looseHtml);
    }

    [Fact, DisplayName("紧/松列表：松散列表序列化保留空行（往返保真）")]
    public void List_LooseRoundtrip()
    {
        var md = "- a\n\n- b\n";
        var doc = MarkdownDocument.ParseRoundtrip(md);
        Assert.Equal(md, doc.ToMarkdown());
        // 程序化松列表序列化带空行
        var prog = new MarkdownDocument();
        var loose = new BulletListBlock([
            MarkdownBlock.CreateListItem([MarkdownInline.CreateText("a")]),
            MarkdownBlock.CreateListItem([MarkdownInline.CreateText("b")]),
        ]) { IsLoose = true };
        prog.Blocks.Add(loose);
        var outMd = prog.ToMarkdown();
        Assert.Contains("- a\n\n- b", outMd);
        // 再解析回松列表
        Assert.True(Assert.IsType<BulletListBlock>(P(outMd).Blocks[0]).IsLoose);
    }

    #endregion
}
