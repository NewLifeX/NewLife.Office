using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>XLSX 语义生效验证：用 OpenXml SDK 读取设置类 API 是否真正落到目标属性</summary>
/// <remarks>
/// 本站曾出现 `SetCellFormat` 因"行 XML 已固化"而静默失效的缺陷。
/// 本类把同一类风险铺开检查：列宽、行高、隐藏、冻结、筛选、保护、页面设置、页眉页脚、
/// 打印标题/区域、定义名、数据验证、超链接，逐项用第三方解析器确认真的写进了文件。
/// </remarks>
public class XlsxSemanticTests
{
    private static Byte[] Build(Action<ExcelWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new ExcelWriter(ms))
        {
            build(w);
            w.Save();
        }
        return ms.ToArray();
    }

    /// <summary>打开 xlsx 并返回（内存流 + 文档 + 首个工作表部件）</summary>
    /// <param name="bytes">xlsx 字节</param>
    /// <returns>元组</returns>
    private static (MemoryStream Stream, SpreadsheetDocument Doc, WorksheetPart Sheet) Open(Byte[] bytes)
    {
        var ms = new MemoryStream(bytes);
        var doc = SpreadsheetDocument.Open(ms, false);
        var sheet = doc.WorkbookPart!.WorksheetParts.First();
        return (ms, doc, sheet);
    }

    [Fact, DisplayName("S1 列宽/行高/隐藏/大纲真正生效")]
    public void S1_Layout()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B", "C"]);
            w.WriteRows("S1", [new Object?[] { 1, 2, 3 }, new Object?[] { 4, 5, 6 }]);
            w.SetColumnWidth("S1", 0, 20);
            w.SetRowHeight("S1", 1, 30);
            w.SetRowHidden("S1", 2);
            w.SetColumnHidden("S1", 2);
            w.SetColumnOutlineLevel("S1", 1, 1);
            w.SetRowOutlineLevel("S1", 3, 1);
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var ws = part.Worksheet;
            var cols = ws.GetFirstChild<Columns>();
            Assert.True(cols != null, "未写出 cols 节点：" + ws.OuterXml);
            var col0 = cols!.Elements<Column>().First(e => e.Min!.Value == 1);
            Assert.Equal(20, col0.Width!.Value);
            Assert.True(col0.CustomWidth!.Value);

            var col2 = ws.GetFirstChild<Columns>()!.Elements<Column>().First(e => e.Min!.Value == 3);
            Assert.True(col2.Hidden!.Value);

            var col1 = ws.GetFirstChild<Columns>()!.Elements<Column>().First(e => e.Min!.Value == 2);
            Assert.True(col1.OutlineLevel?.Value == 1, "列 B 大纲级别缺失：" + ws.OuterXml);

            var rows = ws.GetFirstChild<SheetData>()!.Elements<Row>().ToList();
            // 行号参数为 1 基：SetRowHeight/SetRowHidden/SetRowOutlineLevel 分别落在第 1/2/3 行
            var r1 = rows.First(e => e.RowIndex!.Value == 1);
            Assert.Equal(30, r1.Height!.Value);

            var r2 = rows.First(e => e.RowIndex!.Value == 2);
            Assert.True(r2.Hidden!.Value);

            var r3 = rows.First(e => e.RowIndex!.Value == 3);
            Assert.Equal(1, r3.OutlineLevel!.Value);
        }
    }

    [Fact, DisplayName("S2 冻结/网格线/缩放/Tab颜色真正生效")]
    public void S2_ViewOptions()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.FreezePane("S1", 1, 1);
            w.SetGridlines("S1", false);
            w.SetZoomScale("S1", 120);
            w.SetSheetTabColor("S1", "FF0000");
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var ws = part.Worksheet;
            var view = ws.GetFirstChild<SheetViews>()!.Elements<SheetView>().First();
            Assert.False(view.ShowGridLines!.Value);
            Assert.Equal(120u, view.ZoomScale!.Value);

            var pane = view.GetFirstChild<Pane>();
            Assert.NotNull(pane);
            Assert.Equal(1d, pane!.VerticalSplit!.Value);
            Assert.Equal(1d, pane.HorizontalSplit!.Value);
            Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);

            var props = ws.GetFirstChild<SheetProperties>();
            Assert.NotNull(props);
            var tab = props!.GetFirstChild<TabColor>();
            Assert.NotNull(tab);
            Assert.Equal("FFFF0000", tab!.Rgb!.Value);
        }
    }

    [Fact, DisplayName("S3 自动筛选与工作表保护真正生效")]
    public void S3_AutoFilterAndProtection()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.SetAutoFilter("S1", "A1:B2");
            w.ProtectSheet("S1", "123");
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var ws = part.Worksheet;
            var filter = ws.GetFirstChild<AutoFilter>();
            Assert.NotNull(filter);
            Assert.Equal("A1:B2", filter!.Reference!.Value);

            var protection = ws.GetFirstChild<SheetProtection>();
            Assert.NotNull(protection);
            Assert.True(protection!.Sheet!.Value);
            Assert.False(String.IsNullOrEmpty(protection.Password!.Value));
        }
    }

    [Fact, DisplayName("S4 页面设置与页眉页脚真正生效")]
    public void S4_PageSetup()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.SetPageSetup("S1", PageOrientation.Landscape, PaperSize.A4);
            w.SetPageMargins("S1", 2, 2, 1.5, 1.5);
            w.SetHeaderFooter("S1", "&C表头 &D", "&R第 &P 页");
            w.SetPageBreak("S1", 2);
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var ws = part.Worksheet;
            var setup = ws.GetFirstChild<PageSetup>();
            Assert.NotNull(setup);
            Assert.Equal(OrientationValues.Landscape, setup!.Orientation!.Value);

            var margins = ws.GetFirstChild<PageMargins>();
            Assert.NotNull(margins);
            Assert.Equal(2d, margins!.Top!.Value);

            var hf = ws.GetFirstChild<HeaderFooter>();
            Assert.NotNull(hf);
            Assert.Equal("&C表头 &D", hf!.OddHeader!.Text);

            // 分页符必须位于 headerFooter 之后（CT_Worksheet 顺序）
            Assert.NotNull(ws.GetFirstChild<RowBreaks>());
        }
    }

    [Fact, DisplayName("S5 打印标题/打印区域/定义名真正生效")]
    public void S5_DefinedNames()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.SetPrintTitleRows("S1", 1, 1);
            w.SetPrintArea("S1", "A1:B10");
            w.AddDefinedName("MyName", "S1!$A$1");
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var names = doc.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>()
                .Select(e => (Name: e.Name!.Value, Text: e.Text))
                .ToList();

            Assert.Contains(names, e => e.Name == "_xlnm.Print_Titles" && e.Text.Contains("$1:$1"));
            Assert.Contains(names, e => e.Name == "_xlnm.Print_Area" && e.Text.Contains("A1:B10"));
            Assert.Contains(names, e => e.Name == "MyName" && e.Text.Contains("$A$1"));
        }
    }

    [Fact, DisplayName("S6 数据验证与超链接真正生效")]
    public void S6_ValidationAndHyperlink()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.AddDropdownValidation("S1", "B2:B4", ["是", "否"]);
            w.AddHyperlink("S1", 2, 0, "https://example.com", "示例");
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var ws = part.Worksheet;
            var validations = ws.GetFirstChild<DataValidations>();
            Assert.NotNull(validations);
            var dv = validations!.Elements<DocumentFormat.OpenXml.Spreadsheet.DataValidation>().First();
            Assert.Equal("B2:B4", dv.SequenceOfReferences!.InnerText);
            Assert.Contains("是", dv.GetFirstChild<Formula1>()!.Text);

            var hyperlinks = ws.GetFirstChild<Hyperlinks>();
            Assert.NotNull(hyperlinks);
            var link = hyperlinks!.Elements<Hyperlink>().First();
            Assert.Equal("A2", link.Reference!.Value);
            var relId = link.Id!.Value!;
            var rel = part.HyperlinkRelationships.First(e => e.Id == relId);
            Assert.Equal("https://example.com", rel.Uri.OriginalString);
        }
    }

    [Fact, DisplayName("S7 隐藏工作表与单元格合并真正生效")]
    public void S7_HideSheetAndMerge()
    {
        var bytes = Build(w =>
        {
            w.WriteHeader("S1", ["A", "B"]);
            w.WriteRows("S1", [new Object?[] { 1, 2 }]);
            w.WriteHeader("S2", ["X", "Y"]);
            w.MergeCell("S1", "A3:B3");
            w.HideSheet("S2", false);
        });

        var (ms, doc, part) = Open(bytes);
        using (ms)
        using (doc)
        {
            var sheet2 = doc.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().First(e => e.Name == "S2");
            Assert.Equal(SheetStateValues.Hidden, sheet2.State!.Value);

            var merge = part.Worksheet.GetFirstChild<MergeCells>();
            Assert.NotNull(merge);
            Assert.Equal("A3:B3", merge!.Elements<MergeCell>().First().Reference!.Value);
        }
    }
}
