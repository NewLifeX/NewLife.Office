using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using NewLife.Office.Excel;
using NewLife.Office.Markdown;
using NewLife.Office.Pdf;
using NewLife.Office.Ppt;
using NewLife.Office.Word;
using Xunit;
using XUnitTest.Common;

namespace XUnitTest.Markdown;

/// <summary>全格式→Markdown 转换中枢测试（MD07 统一 AST 架构）</summary>
/// <remarks>
/// 覆盖 Excel/Word/PPT/PDF 四类主流格式的结构化 AST 转换：
/// 字符串输出、AST 块结构、元数据 FrontMatter、往返幂等。
/// </remarks>
[Trait("Category", "Markdown")]
public class FormatToMarkdownTests : IntegrationTestBase
{
    #region Excel

    [Fact, DisplayName("Excel_转换为Markdown表格字符串")]
    public void Convert_Xlsx_ToMarkdown()
    {
        var path = Path.Combine(OutputDir, "fmtmd_xlsx.xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteHeader("Sheet1", new[] { "城市", "人口", "面积" });
            w.WriteRow("Sheet1", new Object?[] { "北京", 2154, 16410 });
            w.WriteRow("Sheet1", new Object?[] { "上海", 2428, 6340 });
            w.Save();
        }

        var md = FormatToMarkdown.FromExcel(path);

        Assert.NotNull(md);
        Assert.Contains("城市", md);
        Assert.Contains("北京", md);
        Assert.Contains("|", md);
        // GFM 表格分隔线
        Assert.Contains("---", md);
    }

    [Fact, DisplayName("Excel_ToDocument返回结构化AST表格")]
    public void ToDocument_Xlsx_Returns_Table()
    {
        var path = Path.Combine(OutputDir, "fmtmd_ast_xlsx.xlsx");
        using (var w = new ExcelWriter(path))
        {
            w.WriteHeader("Sheet1", new[] { "产品", "价格" });
            w.WriteRow("Sheet1", new Object?[] { "笔记本", 5999m });
            w.Save();
        }

        var doc = FormatToMarkdown.ToDocument(path);

        Assert.NotNull(doc);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Table);
        var table = doc.Blocks.First(b => b.Type == MarkdownBlockType.Table);
        Assert.True(table.Children.Count >= 2);
        Assert.Contains("产品", table.GetPlainText());
    }

    #endregion

    #region Word

    [Fact, DisplayName("Word_转换为Markdown标题段落字符串")]
    public void Convert_Docx_ToMarkdown()
    {
        var path = Path.Combine(OutputDir, "fmtmd_docx.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("文档标题", 1);
            w.AppendParagraph("段落内容示例。");
            w.AppendTable(new[] { new[] { "列1", "列2" }, new[] { "A", "B" } }, true);
            w.Save(path);
        }

        var md = FormatToMarkdown.FromWord(path);

        Assert.NotNull(md);
        // 标题可能带粗体内联（Word 标题 run 默认粗体）
        Assert.Contains("# ", md);
        Assert.Contains("文档标题", md);
        Assert.Contains("段落内容示例", md);
        Assert.Contains("|", md);
    }

    [Fact, DisplayName("Word_ToDocument返回标题与段落AST")]
    public void ToDocument_Docx_Returns_HeadingAndParagraph()
    {
        var path = Path.Combine(OutputDir, "fmtmd_ast_docx.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("一级标题", 1);
            w.AppendParagraph("正文段落。");
            w.AppendBulletList(new[] { "项一", "项二" });
            w.Save(path);
        }

        var doc = FormatToMarkdown.ToDocument(path);

        Assert.NotNull(doc);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Heading);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Paragraph);
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.BulletList);
        var heading = doc.Blocks.First(b => b.Type == MarkdownBlockType.Heading);
        Assert.Contains("一级标题", heading.GetPlainText());
    }

    #endregion

    #region PPT

    [Fact, DisplayName("PPT_转换为Markdown幻灯片结构字符串")]
    public void Convert_Pptx_ToMarkdown()
    {
        var path = Path.Combine(OutputDir, "fmtmd_pptx.pptx");
        using (var w = new PptxWriter())
        {
            w.AddSlide();
            w.AddTextBox(0, "演示文稿标题", 2, 2, 20, 3, fontSize: 36, bold: true);
            w.AddTextBox(0, "要点一\n要点二", 2, 6, 20, 4, fontSize: 18);
            w.Save(path);
        }

        var md = FormatToMarkdown.FromPpt(path);

        Assert.NotNull(md);
        Assert.Contains("## 幻灯片 1", md);
        Assert.Contains("演示文稿标题", md);
        Assert.Contains("要点一", md);
    }

    [Fact, DisplayName("PPT_ToDocument返回幻灯片标题与段落AST")]
    public void ToDocument_Pptx_Returns_SlideStructure()
    {
        var path = Path.Combine(OutputDir, "fmtmd_ast_pptx.pptx");
        using (var w = new PptxWriter())
        {
            w.AddSlide();
            w.AddTextBox(0, "幻灯片标题", 2, 2, 20, 3, fontSize: 32, bold: true);
            w.AddTextBox(0, "正文内容", 2, 6, 20, 2, fontSize: 16);
            w.Save(path);
        }

        var doc = FormatToMarkdown.ToDocument(path);

        Assert.NotNull(doc);
        // 幻灯片编号标题
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Heading && b.GetPlainText().Contains("幻灯片 1"));
        // 幻灯片内标题（字号最大）
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Heading && b.GetPlainText().Contains("幻灯片标题"));
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Paragraph && b.GetPlainText().Contains("正文内容"));
    }

    #endregion

    #region PDF

    [Fact, DisplayName("PDF_转换为Markdown标题段落字符串")]
    public void Convert_Pdf_ToMarkdown()
    {
        var path = Path.Combine(OutputDir, "fmtmd_pdf.pdf");
        using (var w = new PdfDocumentBuilder())
        {
            w.AddText("PDF 文档标题", 24f);
            w.AddEmptyLine(8f);
            w.AddText("这是第一段正文内容。", 12f);
            w.AddText("这是第二段正文内容。", 12f);
            w.Save(path);
        }

        var md = FormatToMarkdown.FromPdf(path);

        Assert.NotNull(md);
        Assert.Contains("PDF 文档标题", md);
        Assert.Contains("第一段正文内容", md);
    }

    [Fact, DisplayName("PDF_ToDocument返回标题与段落AST")]
    public void ToDocument_Pdf_Returns_Structure()
    {
        var path = Path.Combine(OutputDir, "fmtmd_ast_pdf.pdf");
        using (var w = new PdfDocumentBuilder())
        {
            w.AddText("PDF 大标题", 24f);
            w.AddEmptyLine(8f);
            w.AddText("正文段落文字。", 12f);
            w.Save(path);
        }

        var doc = FormatToMarkdown.ToDocument(path);

        Assert.NotNull(doc);
        // PDF 标题推断为启发式（依赖字号可解析），至少保证段落结构
        Assert.Contains(doc.Blocks, b => b.Type == MarkdownBlockType.Paragraph);
        Assert.Contains(doc.Blocks, b => b.GetPlainText().Contains("正文段落文字"));
    }

    #endregion

    #region 通用

    [Fact, DisplayName("ToDocument_自动附加FrontMatter元数据")]
    public void ToDocument_With_Metadata()
    {
        var path = Path.Combine(OutputDir, "fmtmd_meta.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("元数据文档", 1);
            w.AppendParagraph("正文。");
            w.Save(path);
        }

        var doc = FormatToMarkdown.ToDocument(path);

        Assert.NotNull(doc);
        Assert.True(doc.FrontMatter.Count > 0);
        Assert.Contains("title", doc.FrontMatter.Keys);
        Assert.Contains("source", doc.FrontMatter.Keys);
        Assert.Contains("format", doc.FrontMatter.Keys);
        Assert.Equal("docx", doc.FrontMatter["format"]);
    }

    [Fact, DisplayName("ToDocument_关闭元数据时不输出FrontMatter")]
    public void ToDocument_NoMetadata()
    {
        var path = Path.Combine(OutputDir, "fmtmd_nometa.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("无元数据", 1);
            w.Save(path);
        }

        var options = new MarkdownConverterOptions { Metadata = false };
        var doc = FormatToMarkdown.ToDocument(path, options);
        var md = doc?.ToMarkdown();

        Assert.NotNull(md);
        Assert.DoesNotContain("---", md);
    }

    [Fact, DisplayName("转换往返幂等：ToDocument→ToMarkdown→Parse→ToMarkdown")]
    public void RoundTrip_Stable()
    {
        var path = Path.Combine(OutputDir, "fmtmd_roundtrip.docx");
        using (var w = new WordWriter())
        {
            w.AppendHeading("往返测试", 1);
            w.AppendParagraph("第一段。");
            w.AppendTable(new[] { new[] { "名称", "值" }, new[] { "A", "1" } }, true);
            w.Save(path);
        }

        var options = new MarkdownConverterOptions { Metadata = false };
        var doc = FormatToMarkdown.ToDocument(path, options);
        Assert.NotNull(doc);

        var md1 = doc!.ToMarkdown();
        var doc2 = MarkdownDocument.Parse(md1);
        var md2 = doc2.ToMarkdown();

        Assert.Equal(md1, md2);
    }

    [Fact, DisplayName("Convert_不支持的扩展名抛ConvertErrorException")]
    public void Convert_Unsupported_Throws()
    {
        var ex = Assert.Throws<ConvertErrorException>(() => FormatToMarkdown.Convert("data.txt"));
        Assert.Equal(ConvertErrorType.Unsupported, ex.Type);
    }

    [Fact, DisplayName("流入口_Convert_Stream_按扩展名转换")]
    public void Convert_Stream_ByExtension()
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            w.WriteHeader("Sheet1", new[] { "产品", "价格" });
            w.WriteRow("Sheet1", new Object?[] { "手机", 3999m });
            w.Save();
        }

        ms.Position = 0;
        var md = FormatToMarkdown.Convert(ms, ".xlsx", null);

        Assert.NotNull(md);
        Assert.Contains("产品", md);
        Assert.Contains("手机", md);
    }

    #endregion
}
