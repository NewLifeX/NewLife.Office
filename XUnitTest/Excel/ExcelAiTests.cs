using NewLife.Office.Excel;
using Xunit;

namespace XUnitTest.Excel;

/// <summary>Excel AI 接口测试：结构快照、A1 区域读写、JSON 补丁、搜索</summary>
public class ExcelAiTests
{
    #region 工具
    private static String TempPath(String name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "nloffice-ai", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    private static ExcelDocument BuildSample()
    {
        var doc = new ExcelDocument();
        var sheet = new Worksheet { Name = "销售" };
        sheet.Rows.Add(["产品", "数量", "单价"]);
        sheet.Rows.Add(["苹果", 10d, 3.5d]);
        sheet.Rows.Add(["香蕉", 20d, 2.0d]);
        sheet.Rows.Add(["合计", null, null]);
        sheet.Formulas[(3, 0)] = "SUM(B2:B3)";
        doc.Worksheets.Add(sheet);
        return doc;
    }
    #endregion

    [Fact(DisplayName = "A1-解析各种区域写法")]
    public void ParseRange_Variants()
    {
        var r1 = ExcelCellRange.Parse("Sheet1!A1:D20");
        Assert.NotNull(r1);
        Assert.Equal("Sheet1", r1!.Sheet);
        Assert.Equal(0, r1.StartRow);
        Assert.Equal(0, r1.StartCol);
        Assert.Equal(19, r1.EndRow);
        Assert.Equal(3, r1.EndCol);
        Assert.Equal("A1:D20", r1.Address);

        var r2 = ExcelCellRange.Parse("B3");
        Assert.NotNull(r2);
        Assert.True(r2!.IsSingleCell);
        Assert.Equal(2, r2.StartRow);
        Assert.Equal(1, r2.StartCol);

        var r3 = ExcelCellRange.Parse("'My Sheet'!$A$1:$B$2", "默认");
        Assert.NotNull(r3);
        Assert.Equal("My Sheet", r3!.Sheet);
        Assert.Equal(1, r3.EndRow);
        Assert.Equal(1, r3.EndCol);

        var r4 = ExcelCellRange.Parse("A:C", "S1");
        Assert.NotNull(r4);
        Assert.Equal(0, r4!.StartRow);
        Assert.Equal(ExcelCellRange.MaxRow, r4.EndRow);
        Assert.Equal(2, r4.EndCol);

        Assert.Null(ExcelCellRange.Parse("not a range"));
    }

    [Fact(DisplayName = "快照-工作表与表头预览")]
    public void Outline_FromWorkbook()
    {
        var outline = BuildSample().ToOutline();

        Assert.Single(outline.Sheets);
        var sheet = outline.Sheets[0];
        Assert.Equal("销售", sheet.Name);
        Assert.Equal(4, sheet.Rows);
        Assert.Equal(3, sheet.Columns);
        Assert.Equal(["产品", "数量", "单价"], sheet.Header);
        Assert.Equal(4, sheet.Preview.Count);
        Assert.Equal(1, sheet.FormulaCount);
        Assert.Equal(3, sheet.Summary.Count);
        Assert.Equal("text", sheet.Summary[0].Type);
        Assert.Equal("number", sheet.Summary[1].Type);
        Assert.Equal(0.333, sheet.Summary[1].NullRate);

        using var parsed = System.Text.Json.JsonDocument.Parse(outline.ToJson());
        var root = parsed.RootElement;
        Assert.Equal(1, root.GetProperty("sheetCount").GetInt32());
        Assert.Equal("销售", root.GetProperty("sheets")[0].GetProperty("anchor").GetString());
        Assert.Equal("产品", root.GetProperty("sheets")[0].GetProperty("header")[0].GetString());
    }

    [Fact(DisplayName = "补丁-写值写公式并往返保存")]
    public void Patch_SetCellsAndSave_RoundTrip()
    {
        var path = TempPath("patch.xlsx");
        BuildSample().Save(path);

        ExcelDocument doc;
        using (var reader = new ExcelReader(path)) doc = reader.ReadExcel();

        var patch = ExcelPatch.ParseJson("""
        {
          "ops": [
            { "op": "set_cells", "range": "销售!A5:C5", "values": [["橙子", 5, 4.2]] },
            { "op": "set_formula", "range": "销售!C6", "formula": "=SUM(C2:C3)" },
            { "op": "set_style", "range": "销售!A1:C1", "bold": true, "backgroundColor": "EEEEEE" },
            { "op": "merge", "range": "销售!A8:C8" }
          ]
        }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));
        doc.Save(path);

        using var reader2 = new ExcelReader(path);
        var doc2 = reader2.ReadExcel();
        var sheet = doc2.GetSheet("销售");

        Assert.Equal("橙子", ExcelSheetEditor.GetValue(sheet, 4, 0));
        Assert.Equal(5d, Convert.ToDouble(ExcelSheetEditor.GetValue(sheet, 4, 1)));
        Assert.Equal("SUM(C2:C3)", ExcelSheetEditor.GetFormula(sheet, 5, 2));
        Assert.True(ExcelSheetEditor.GetStyle(sheet, 0, 0)!.Bold);
        Assert.Equal("EEEEEE", ExcelSheetEditor.GetStyle(sheet, 0, 0)!.BackgroundColor);
        Assert.Contains((7, 0, 7, 2), sheet.Merges);
    }

    [Fact(DisplayName = "补丁-单条失败整体不写文件")]
    public void Patch_FailureNotSaved()
    {
        var path = TempPath("bad.xlsx");
        BuildSample().Save(path);
        var before = File.ReadAllBytes(path);

        ExcelDocument doc;
        using (var reader = new ExcelReader(path)) doc = reader.ReadExcel();

        var patch = ExcelPatch.ParseJson("""
        { "ops": [
            { "op": "set_cells", "range": "销售!A2", "values": [["改值"]] },
            { "op": "delete_sheet", "name": "不存在" }
        ] }
        """);
        var result = patch.Apply(doc);

        Assert.False(result.Success);
        Assert.Single(result.Failed);
        Assert.Single(result.Applied);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact(DisplayName = "补丁-插入删除行平移样式")]
    public void Patch_InsertDeleteRows_ShiftStyles()
    {
        var doc = BuildSample();
        var sheet = doc.GetSheet("销售");
        sheet.CellFormats[(2, 0)] = new CellFormat { Bold = true };

        var patch = ExcelPatch.ParseJson("""
        { "ops": [ { "op": "insert_rows", "range": "销售!2:2" } ] }
        """);
        Assert.True(patch.Apply(doc).Success);

        // 原第 3 行（索引 2）下移一行，样式跟着走
        Assert.True(sheet.CellFormats.ContainsKey((3, 0)));
        Assert.Equal("香蕉", ExcelSheetEditor.GetValue(sheet, 3, 0));

        var patch2 = ExcelPatch.ParseJson("""
        { "ops": [ { "op": "delete_rows", "range": "销售!2:2" } ] }
        """);
        Assert.True(patch2.Apply(doc).Success);
        Assert.Equal("香蕉", ExcelSheetEditor.GetValue(sheet, 2, 0));
        Assert.True(sheet.CellFormats.ContainsKey((2, 0)));
    }

    [Fact(DisplayName = "搜索-命中值与公式")]
    public void Search_FindsValueAndFormula()
    {
        var doc = BuildSample();
        var result = ExcelSearch.Search(doc, "SUM");

        Assert.Single(result.Hits);
        Assert.Equal("销售!A4", result.Hits[0].Address);
        Assert.Equal("=SUM(B2:B3)", result.Hits[0].Formula);

        var result2 = ExcelSearch.Search(doc, "苹果");
        Assert.Single(result2.Hits);
        Assert.Equal("销售!A2", result2.Hits[0].Address);
        Assert.False(result2.Truncated);

        var result3 = ExcelSearch.Search(doc, "香蕉", maxResults: 0);
        Assert.Single(result3.Hits);
    }

    [Fact(DisplayName = "补丁-新增表格与图表")]
    public void Patch_AddTableAndChart()
    {
        var path = TempPath("chart.xlsx");
        var doc = BuildSample();

        var patch = ExcelPatch.ParseJson("""
        { "ops": [
            { "op": "add_table", "range": "销售!A1:C3", "name": "销量表" },
            { "op": "add_chart", "range": "销售!A1:C3", "type": "bar", "title": "销量" }
        ] }
        """);
        var result = patch.Apply(doc);
        Assert.True(result.Success, String.Join("; ", result.Failed));
        doc.Save(path);

        using var reader = new ExcelReader(path);
        var sheet = reader.ReadExcel().GetSheet("销售");
        Assert.Single(sheet.Tables);
        Assert.Single(sheet.Charts);
    }

    [Fact(DisplayName = "区域读取-公式带等号前缀")]
    public void ReadRange_FormulasWithEquals()
    {
        var sheet = BuildSample().GetSheet("销售");
        var range = ExcelCellRange.Parse("A1:C2")!;
        var rows = ExcelSheetEditor.ReadRange(sheet, range);

        Assert.Equal(2, rows.Count);
        Assert.Equal("产品", rows[0][0]);
        Assert.Equal(10d, rows[1][1]);
    }

    [Fact(DisplayName = "补丁-append_rows 跳过尾部纯公式行")]
    public void Patch_AppendRowsSkipsFormulaTail()
    {
        var doc = new ExcelDocument();
        var sheet = new Worksheet { Name = "S" };
        sheet.Rows.Add(["名称"]);
        sheet.Rows.Add([1d]);
        sheet.Formulas[(2, 0)] = "SUM(B2:B2)";
        doc.Worksheets.Add(sheet);

        var patch = ExcelPatch.ParseJson("""{ "ops": [ { "op": "append_rows", "values": [["新行"]] } ] }""");
        var result = patch.Apply(doc);

        Assert.True(result.Success, String.Join("; ", result.Failed));
        Assert.True(doc.Worksheets[0].Formulas.ContainsKey((2, 0)));
        Assert.Equal("新行", doc.Worksheets[0].Rows[3][0]);
    }

    [Fact(DisplayName = "补丁-set_cells 标量铺满区域")]
    public void Patch_SetCellsScalar()
    {
        var doc = BuildSample();
        var patch = ExcelPatch.ParseJson("""{ "ops": [ { "op": "set_cells", "range": "销售!A1:B2", "values": 7 } ] }""");
        var result = patch.Apply(doc);

        Assert.True(result.Success, String.Join("; ", result.Failed));
        var rows = ExcelSheetEditor.ReadRange(doc.GetSheet("销售"), ExcelCellRange.Parse("A1:B2")!);
        Assert.All(rows, row => Assert.All(row, cell => Assert.Equal(7d, cell)));
    }
}
