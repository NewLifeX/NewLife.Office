using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NewLife.Office.Word;
using Xunit;
using WDocument = NewLife.Office.Word.Document;
using WElement = NewLife.Office.Word.Element;
using WParagraph = NewLife.Office.Word.Paragraph;
using WRun = NewLife.Office.Word.Run;
using WRunProperties = NewLife.Office.Word.RunProperties;
using WTable = NewLife.Office.Word.Table;

namespace XUnitTest.Word;

/// <summary>DOCX 语义生效验证：用 OpenXml SDK 读取模型属性是否真正落到目标元素</summary>
/// <remarks>
/// 本项目在 XLSX 侧曾出现"设置 API 静默无效"的缺陷（样式未回填）。
/// 本类把同一类风险在 DOCX 侧铺开：页面设置、页边距、页面边框、段落与字符格式、
/// 表格结构与样式、文档属性，逐项用第三方解析器确认真的写进了文件。
/// </remarks>
public class DocxSemanticTests
{
    /// <summary>构造一个含单 Run 的段落</summary>
    /// <param name="text">文本</param>
    /// <returns>段落</returns>
    private static WParagraph Para(String text)
    {
        var p = new WParagraph();
        p.Runs.Add(new WRun { Text = text });
        return p;
    }

    private static Byte[] Build(Action<WDocument> build)
    {
        var doc = new WDocument();
        doc.Elements.Add(new WElement { Type = ElementType.Paragraph, Paragraph = Para("正文第一段") });
        build(doc);
        using var ms = new MemoryStream();
        new WordWriter().Save(ms, doc);
        return ms.ToArray();
    }

    /// <summary>打开 docx 并返回（内存流 + 文档）</summary>
    /// <param name="bytes">docx 字节</param>
    /// <returns>元组</returns>
    private static (MemoryStream Stream, WordprocessingDocument Doc) Open(Byte[] bytes)
    {
        var ms = new MemoryStream(bytes);
        var doc = WordprocessingDocument.Open(ms, false);
        return (ms, doc);
    }

    [Fact, DisplayName("WS1 页面尺寸与页边距真正生效")]
    public void WS1_PageSetup()
    {
        var bytes = Build(doc =>
        {
            doc.PageSettings.PageWidth = 15840;
            doc.PageSettings.PageHeight = 12240;
            doc.PageSettings.Landscape = true;
            doc.PageSettings.MarginLeft = 2160;
            doc.PageSettings.MarginTop = 1440;
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var sectPr = wd.MainDocumentPart!.Document.Body!.Elements<SectionProperties>().First();
            var size = sectPr.GetFirstChild<PageSize>()!;
            // PageWidth/PageHeight 语义为"纵向纸张尺寸"，Landscape 时写盘前互换
            Assert.Equal(12240u, size.Width!.Value);
            Assert.Equal(15840u, size.Height!.Value);
            Assert.Equal(PageOrientationValues.Landscape, size.Orient!.Value);

            var margin = sectPr.GetFirstChild<PageMargin>()!;
            Assert.Equal(2160, (Int32)margin.Left!.Value);
            Assert.Equal(1440, (Int32)margin.Top!.Value);
        }
    }

    [Fact, DisplayName("WS2 段落对齐与样式真正生效")]
    public void WS2_ParagraphFormat()
    {
        var bytes = Build(doc =>
        {
            var p = Para("格式化段落");
            p.Alignment = "center";
            p.StyleId = "Heading1";
            p.IndentLeft = 720;
            p.SpaceBefore = 240;
            doc.Elements.Add(new WElement { Type = ElementType.Paragraph, Paragraph = p });
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var p = wd.MainDocumentPart!.Document.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .First(e => e.InnerText.Contains("格式化段落"));
            var pPr = p.ParagraphProperties!;
            Assert.Equal(JustificationValues.Center, pPr.Justification!.Val!.Value);
            Assert.Equal("Heading1", pPr.ParagraphStyleId!.Val!.Value);
            Assert.Equal("720", pPr.Indentation!.Left!.Value);
            Assert.Equal("240", pPr.SpacingBetweenLines!.Before!.Value);
        }
    }

    [Fact, DisplayName("WS3 字符格式真正生效")]
    public void WS3_RunFormat()
    {
        var bytes = Build(doc =>
        {
            var p = new WParagraph();
            p.Runs.Add(new WRun
            {
                Text = "字符格式",
                Properties = new WRunProperties
                {
                    Bold = true,
                    Italic = true,
                    FontSize = 16,
                    FontName = "Arial",
                    ForeColor = "FF0000",
                    Underline = true,
                },
            });
            doc.Elements.Add(new WElement { Type = ElementType.Paragraph, Paragraph = p });
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var run = wd.MainDocumentPart!.Document.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                .First(e => e.InnerText.Contains("字符格式"));
            var rPr = run.RunProperties!;
            Assert.NotNull(rPr.Bold);
            Assert.NotNull(rPr.Italic);
            Assert.NotNull(rPr.Underline);
            Assert.Equal("32", rPr.FontSize!.Val!.Value);
            Assert.Equal("Arial", rPr.RunFonts!.Ascii!.Value);
            Assert.Equal("FF0000", rPr.Color!.Val!.Value);
        }
    }

    [Fact, DisplayName("WS4 表格结构与样式真正生效")]
    public void WS4_Table()
    {
        var bytes = Build(doc =>
        {
            var table = new WTable { StyleId = "TableGrid" };
            table.SetCellText(0, 0, "表头1");
            table.SetCellText(1, 2, "数据");
            doc.Elements.Add(new WElement { Type = ElementType.Table, Table = table });
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var tbl = wd.MainDocumentPart!.Document.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().First();
            var rows = tbl.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Contains("表头1", tbl.InnerText);
            Assert.Contains("数据", tbl.InnerText);
            Assert.Equal("TableGrid", tbl.GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value);
        }
    }

    [Fact, DisplayName("WS5 页面边框真正生效")]
    public void WS5_PageBorder()
    {
        var bytes = Build(doc =>
        {
            doc.PageSettings.PageBorder = new NewLife.Office.Word.PageBorder
            {
                Top = "single",
                Bottom = "single",
                Left = "single",
                Right = "single",
                Color = "FF0000",
                Size = 8,
            };
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var sectPr = wd.MainDocumentPart!.Document.Body!.Elements<SectionProperties>().First();
            var border = sectPr.GetFirstChild<PageBorders>();
            Assert.NotNull(border);
            Assert.Equal(BorderValues.Single, border!.TopBorder!.Val!.Value);
            Assert.Equal("FF0000", border.TopBorder.Color!.Value);
        }
    }

    [Fact, DisplayName("WS6 文档属性真正生效")]
    public void WS6_DocumentProperties()
    {
        var bytes = Build(doc =>
        {
            doc.DocumentProperties.Title = "测试标题";
            doc.DocumentProperties.Author = "测试作者";
            doc.DocumentProperties.Subject = "测试主题";
        });

        var (ms, wd) = Open(bytes);
        using (ms)
        using (wd)
        {
            var props = wd.PackageProperties;
            Assert.Equal("测试标题", props.Title);
            Assert.Equal("测试作者", props.Creator);
            Assert.Equal("测试主题", props.Subject);
        }
    }
}
