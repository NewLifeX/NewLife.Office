using System.Globalization;
using System.Text;

namespace NewLife.Office.Excel;

/// <summary>工作簿 JSON 补丁：以 A1 坐标批量修改，任一条失败不写文件</summary>
/// <remarks>
/// 指令格式与 docx 补丁一致：<c>{ "ops": [ ... ] }</c>，按数组顺序执行。
/// 执行只改内存模型，调用方应在 <see cref="ExcelPatchResult.Success"/> 为 true 时再保存。
/// <example>
/// <code>
/// var patch = ExcelPatch.ParseJson("""
/// {
///   "ops": [
///     { "op": "set_cells", "range": "Sheet1!A1:B2", "values": [["姓名","部门"],["张三","研发"]] },
///     { "op": "set_formula", "range": "Sheet1!C2", "formula": "=SUM(A2:B2)" },
///     { "op": "set_style", "range": "Sheet1!A1:B1", "bold": true, "backgroundColor": "EEEEEE" }
///   ]
/// }
/// """);
/// using var reader = new ExcelReader("in.xlsx");
/// var doc = reader.ReadExcel();
/// var result = patch.Apply(doc);
/// if (result.Success) doc.Save("out.xlsx");
/// </code>
/// </example>
/// </remarks>
public class ExcelPatch
{
    #region 属性
    /// <summary>指令集合</summary>
    public List<ExcelPatchOp> Ops { get; } = [];
    #endregion

    #region 解析
    /// <summary>解析 JSON 补丁</summary>
    /// <param name="json">JSON 文本</param>
    /// <returns>补丁对象</returns>
    /// <exception cref="FormatException">JSON 结构非法</exception>
    public static ExcelPatch ParseJson(String json)
    {
        if (json == null) throw new ArgumentNullException(nameof(json));

        var root = JsonLite.Parse(json) as Dictionary<String, Object?>
            ?? throw new FormatException(L.T("Patch root must be a JSON object."));

        var patch = new ExcelPatch();
        if (!root.TryGetValue("ops", out var opsValue) || opsValue is not List<Object?> list)
            throw new FormatException(L.T("Patch must contain an 'ops' array."));

        foreach (var item in list)
        {
            if (item is not Dictionary<String, Object?> obj)
                throw new FormatException(L.T("Each op must be a JSON object."));

            var op = new ExcelPatchOp { Args = obj };
            op.Op = JsonLite.GetString(obj, "op") ?? throw new FormatException(L.T("Each op must have an 'op' name."));
            patch.Ops.Add(op);
        }
        return patch;
    }
    #endregion

    #region 执行
    /// <summary>在工作簿模型上执行补丁</summary>
    /// <param name="document">工作簿模型</param>
    /// <returns>执行结果</returns>
    public ExcelPatchResult Apply(ExcelDocument document)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var result = new ExcelPatchResult();
        for (var i = 0; i < Ops.Count; i++)
        {
            var op = Ops[i];
            try
            {
                ApplyOp(document, op, result);
                result.Applied.Add($"{i}:{op.Op}");
            }
            catch (Exception ex)
            {
                result.Failed.Add($"{i}:{op.Op}: {ex.Message}");
                result.FailedCodes.Add(ex.GetType().Name);
            }
        }
        return result;
    }

    private static void ApplyOp(ExcelDocument document, ExcelPatchOp op, ExcelPatchResult result)
    {
        switch (op.Op)
        {
            case "set_cells":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                range.Sheet = sheet.Name;

                var values = op.GetMatrix("values");
                if (values == null)
                {
                    // 标量按文档约定铺满整个区域
                    if (!op.Args.TryGetValue("values", out var scalar) || scalar == null)
                        throw new ArgumentException(L.T("set_cells requires 'values'."));

                    values = [];
                    var columns = range.EndCol - range.StartCol + 1;
                    for (var r = range.StartRow; r <= range.EndRow; r++)
                    {
                        var row = new Object?[columns];
                        for (var c = 0; c < columns; c++) row[c] = scalar;
                        values.Add(row);
                    }
                }
                var written = ExcelSheetEditor.WriteMatrix(sheet, range, values);
                result.Warnings.Add(L.T("set_cells {0}: wrote {1} row(s) x {2} column(s)", range.FullAddress, written.Rows, written.Columns));
                return;
            }

            case "set_cell":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                ExcelSheetEditor.SetValue(sheet, range.StartRow, range.StartCol, op.GetValue("value"));
                return;
            }

            case "set_formula":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                var formula = op.GetString("formula") ?? throw new ArgumentException(L.T("set_formula requires 'formula'."));
                if (!range.IsSingleCell)
                    result.Warnings.Add(L.T("set_formula {0}: the same formula was written to multiple cells; relative references are not adjusted.", range.FullAddress));
                for (var r = range.StartRow; r <= range.EndRow; r++)
                    for (var c = range.StartCol; c <= range.EndCol; c++)
                        ExcelSheetEditor.SetFormula(sheet, r, c, formula);
                return;
            }

            case "append_rows":
            {
                var name = op.GetString("sheet");
                var sheet = document.GetSheet(name);
                var values = op.GetMatrix("values") ?? throw new ArgumentException(L.T("append_rows requires 'values'."));
                var start = ExcelSheetEditor.AppendRows(sheet, values);
                result.Warnings.Add(L.T("append_rows {0}: appended starting at row {1}", sheet.Name, start + 1));
                return;
            }

            case "insert_rows":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                ExcelSheetEditor.InsertRows(sheet, range.StartRow, range.RowCount);
                return;
            }

            case "delete_rows":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                ExcelSheetEditor.DeleteRows(sheet, range.StartRow, range.RowCount);
                return;
            }

            case "merge":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                if (!sheet.Merges.Contains((range.StartRow, range.StartCol, range.EndRow, range.EndCol)))
                    sheet.Merges.Add((range.StartRow, range.StartCol, range.EndRow, range.EndCol));
                return;
            }

            case "unmerge":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                sheet.Merges.RemoveAll(m => m.StartRow == range.StartRow && m.StartCol == range.StartCol
                    && m.EndRow == range.EndRow && m.EndCol == range.EndCol);
                return;
            }

            case "set_style":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                ExcelSheetEditor.ApplyStyle(sheet, range, format => ApplyStyleFields(op, format));
                return;
            }

            case "set_number_format":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                var format = op.GetString("format") ?? throw new ArgumentException(L.T("set_number_format requires 'format'."));
                ExcelSheetEditor.SetNumberFormat(sheet, range, format);
                return;
            }

            case "set_column_width":
            {
                var sheet = ResolveSheet(document, op);
                var colRange = op.GetString("columns") ?? op.GetString("range") ?? throw new ArgumentException(L.T("set_column_width requires 'columns'."));
                var width = op.GetDouble("width") ?? throw new ArgumentException(L.T("set_column_width requires 'width'."));
                var range = ExcelCellRange.Parse(colRange, sheet.Name) ?? throw new ArgumentException(L.T("Cannot parse column range: {0}", colRange));
                for (var c = range.StartCol; c <= range.EndCol; c++)
                    ExcelSheetEditor.SetColumnWidth(sheet, c, width);
                return;
            }

            case "set_row_height":
            {
                var sheet = ResolveSheet(document, op);
                var rowRange = op.GetString("rows") ?? op.GetString("range") ?? throw new ArgumentException(L.T("set_row_height requires 'rows'."));
                var height = op.GetDouble("height") ?? throw new ArgumentException(L.T("set_row_height requires 'height'."));
                var range = ExcelCellRange.Parse(rowRange, sheet.Name) ?? throw new ArgumentException(L.T("Cannot parse row range: {0}", rowRange));
                for (var r = range.StartRow; r <= range.EndRow; r++)
                    ExcelSheetEditor.SetRowHeight(sheet, r, height);
                return;
            }

            case "freeze":
            {
                var sheet = ResolveSheet(document, op);
                var rows = op.GetInt32("rows") ?? 0;
                var cols = op.GetInt32("cols") ?? 0;
                sheet.FreezePane = rows > 0 || cols > 0 ? (rows, cols) : null;
                return;
            }

            case "autofilter":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                sheet.AutoFilter = range.Address;
                return;
            }

            case "add_sheet":
            {
                var name = op.GetString("name") ?? throw new ArgumentException(L.T("add_sheet requires 'name'."));
                if (document.FindSheet(name) != null) throw new ArgumentException(L.T("Worksheet already exists: {0}", name));
                document.Worksheets.Add(new Worksheet { Name = name });
                return;
            }

            case "rename_sheet":
            {
                var from = op.GetString("from") ?? op.GetString("name") ?? throw new ArgumentException(L.T("rename_sheet requires 'from'."));
                var to = op.GetString("to") ?? op.GetString("newName") ?? throw new ArgumentException(L.T("rename_sheet requires 'to'."));
                if (document.FindSheet(to) != null) throw new ArgumentException(L.T("Worksheet already exists: {0}", to));
                var sheet = document.GetSheet(from);
                sheet.Name = to;
                return;
            }

            case "delete_sheet":
            {
                var name = op.GetString("name") ?? throw new ArgumentException(L.T("delete_sheet requires 'name'."));
                if (document.Worksheets.Count <= 1) throw new ArgumentException(L.T("Cannot delete the last worksheet."));
                var sheet = document.GetSheet(name);
                document.Worksheets.Remove(sheet);
                return;
            }

            case "add_table":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                var table = new Table
                {
                    Name = op.GetString("name") ?? $"Table{sheet.Tables.Count + 1}",
                    Range = range.Address,
                    StyleName = op.GetString("style") ?? "TableStyleMedium9",
                    ColumnNames = op.GetStrings("columns"),
                };
                sheet.Tables.Add(table);
                return;
            }

            case "add_chart":
            {
                var range = ParseRange(document, op, true);
                var sheet = document.GetSheet(range.Sheet);
                sheet.Charts.Add(BuildChart(sheet, range, op));
                return;
            }

            default:
                throw new ArgumentException(L.T("Unknown operation: {0}", op.Op));
        }
    }
    #endregion

    #region 辅助
    private static ExcelCellRange ParseRange(ExcelDocument document, ExcelPatchOp op, Boolean required)
    {
        var text = op.GetString("range") ?? (required ? throw new ArgumentException(L.T("{0} requires 'range'.", op.Op)) : null);
        if (text == null) throw new ArgumentException(L.T("{0} requires 'range'.", op.Op));

        var defaultSheet = op.GetString("sheet") ?? document.Worksheets.FirstOrDefault()?.Name;
        var range = ExcelCellRange.Parse(text, defaultSheet);
        if (range == null) throw new ArgumentException(L.T("Cannot parse range: {0}", text));
        if (range.Sheet == null || range.Sheet.Length == 0) range.Sheet = defaultSheet;
        return range;
    }

    private static Worksheet ResolveSheet(ExcelDocument document, ExcelPatchOp op)
    {
        var rangeText = op.GetString("range");
        if (!rangeText.IsNullOrEmpty())
        {
            var range = ExcelCellRange.Parse(rangeText, op.GetString("sheet"));
            if (range?.Sheet != null) return document.GetSheet(range.Sheet);
        }
        return document.GetSheet(op.GetString("sheet"));
    }

    private static void ApplyStyleFields(ExcelPatchOp op, CellFormat format)
    {
        var bold = op.GetBoolean("bold");
        if (bold.HasValue) format.Bold = bold.Value;
        var italic = op.GetBoolean("italic");
        if (italic.HasValue) format.Italic = italic.Value;
        var underline = op.GetBoolean("underline");
        if (underline.HasValue) format.Underline = underline.Value;
        var strike = op.GetBoolean("strike");
        if (strike.HasValue) format.Strike = strike.Value;
        var wrap = op.GetBoolean("wrapText");
        if (wrap.HasValue) format.WrapText = wrap.Value;

        var fontSize = op.GetDouble("fontSize");
        if (fontSize.HasValue) format.FontSize = fontSize.Value;
        var fontName = op.GetString("fontName");
        if (!fontName.IsNullOrEmpty()) format.FontName = fontName;
        var fontColor = op.GetString("fontColor");
        if (!fontColor.IsNullOrEmpty()) format.FontColor = fontColor;
        var background = op.GetString("backgroundColor");
        if (!background.IsNullOrEmpty()) format.BackgroundColor = background;
        var numberFormat = op.GetString("numberFormat");
        if (!numberFormat.IsNullOrEmpty()) format.NumberFormat = numberFormat;
        var borderColor = op.GetString("borderColor");
        if (!borderColor.IsNullOrEmpty()) format.BorderColor = borderColor;

        var hAlign = op.GetString("hAlign");
        if (!hAlign.IsNullOrEmpty()) format.HAlign = ParseAlignment(hAlign!);
        var vAlign = op.GetString("vAlign");
        if (!vAlign.IsNullOrEmpty()) format.VAlign = ParseVerticalAlignment(vAlign!);
        var border = op.GetString("border");
        if (!border.IsNullOrEmpty()) format.Border = ParseBorder(border!);
    }

    private static HorizontalAlignment ParseAlignment(String value) => value.ToLowerInvariant() switch
    {
        "left" or "l" => HorizontalAlignment.Left,
        "center" or "ctr" or "c" => HorizontalAlignment.Center,
        "right" or "r" => HorizontalAlignment.Right,
        "fill" => HorizontalAlignment.Fill,
        "justify" => HorizontalAlignment.Justify,
        "general" => HorizontalAlignment.General,
        _ => throw new ArgumentException(L.T("Unknown horizontal alignment: {0}", value)),
    };

    private static VerticalAlignment ParseVerticalAlignment(String value) => value.ToLowerInvariant() switch
    {
        "top" or "t" => VerticalAlignment.Top,
        "center" or "ctr" or "m" => VerticalAlignment.Center,
        "bottom" or "b" => VerticalAlignment.Bottom,
        _ => throw new ArgumentException(L.T("Unknown vertical alignment: {0}", value)),
    };

    private static BorderStyle ParseBorder(String value) => value.ToLowerInvariant() switch
    {
        "thin" => BorderStyle.Thin,
        "medium" => BorderStyle.Medium,
        "thick" => BorderStyle.Thick,
        "dashed" => BorderStyle.Dashed,
        "dotted" => BorderStyle.Dotted,
        "double" or "doubleline" => BorderStyle.DoubleLine,
        "none" => BorderStyle.None,
        _ => throw new ArgumentException(L.T("Unknown border style: {0}", value)),
    };

    private static ExcelChart BuildChart(Worksheet sheet, ExcelCellRange range, ExcelPatchOp op)
    {
        // 约定：区域首行为系列名、首列为分类；至少 2 行 2 列
        if (range.RowCount < 2 || range.ColumnCount < 2)
            throw new ArgumentException(L.T("add_chart requires a data range of at least 2 rows and 2 columns (first row: series names, first column: categories)."));

        var chart = new ExcelChart
        {
            Type = op.GetString("type") ?? "bar",
            Title = op.GetString("title"),
            AnchorRow = op.GetInt32("anchorRow") ?? (range.EndRow + 2),
            AnchorCol = op.GetInt32("anchorCol") ?? range.StartCol,
            WidthPx = op.GetInt32("width") ?? 480,
            HeightPx = op.GetInt32("height") ?? 300,
        };

        var categories = new String[range.RowCount - 1];
        for (var r = range.StartRow + 1; r <= range.EndRow; r++)
            categories[r - range.StartRow - 1] = Convert.ToString(ExcelSheetEditor.GetValue(sheet, r, range.StartCol), CultureInfo.InvariantCulture) ?? String.Empty;
        chart.Categories = categories;

        for (var c = range.StartCol + 1; c <= range.EndCol; c++)
        {
            var name = Convert.ToString(ExcelSheetEditor.GetValue(sheet, range.StartRow, c), CultureInfo.InvariantCulture) ?? $"Series{c - range.StartCol}";
            var data = new Double[categories.Length];
            for (var r = range.StartRow + 1; r <= range.EndRow; r++)
            {
                var value = ExcelSheetEditor.GetValue(sheet, r, c);
                data[r - range.StartRow - 1] = value == null ? 0d : Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            chart.Series.Add(new ExcelChartSeries { Name = name, Data = data });
        }

        return chart;
    }
    #endregion
}

/// <summary>单条补丁指令</summary>
public class ExcelPatchOp
{
    /// <summary>指令名</summary>
    public String Op { get; set; } = String.Empty;

    /// <summary>参数字典</summary>
    public Dictionary<String, Object?> Args { get; set; } = [];

    /// <summary>读取字符串参数</summary>
    /// <param name="key">键</param>
    /// <returns>字符串值</returns>
    public String? GetString(String key) => JsonLite.GetString(Args, key);

    /// <summary>读取整数参数</summary>
    /// <param name="key">键</param>
    /// <returns>整数值</returns>
    public Int32? GetInt32(String key) => JsonLite.GetInt32(Args, key);

    /// <summary>读取布尔参数</summary>
    /// <param name="key">键</param>
    /// <returns>布尔值</returns>
    public Boolean? GetBoolean(String key) => JsonLite.GetBoolean(Args, key);

    /// <summary>读取数值参数</summary>
    /// <param name="key">键</param>
    /// <returns>数值</returns>
    public Double? GetDouble(String key)
    {
        if (!Args.TryGetValue(key, out var value) || value == null) return null;
        if (value is Double d) return d;
        if (value is String s && Double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)) return parsed;
        return null;
    }

    /// <summary>读取任意值参数</summary>
    /// <param name="key">键</param>
    /// <returns>值</returns>
    public Object? GetValue(String key) => Args.TryGetValue(key, out var value) ? ExcelSheetEditor.NormalizeValue(value) : null;

    /// <summary>读取二维数组参数</summary>
    /// <param name="key">键</param>
    /// <returns>二维数组，缺失返回 null</returns>
    public List<Object?[]>? GetMatrix(String key)
    {
        if (!Args.TryGetValue(key, out var value) || value is not List<Object?> list) return null;

        var rows = new List<Object?[]>(list.Count);
        foreach (var item in list)
        {
            if (item is List<Object?> cells)
            {
                var row = new Object?[cells.Count];
                for (var i = 0; i < cells.Count; i++) row[i] = ExcelSheetEditor.NormalizeValue(cells[i]);
                rows.Add(row);
            }
            else
            {
                rows.Add([ExcelSheetEditor.NormalizeValue(item)]);
            }
        }
        return rows;
    }

    /// <summary>读取字符串数组参数</summary>
    /// <param name="key">键</param>
    /// <returns>字符串数组，缺失返回 null</returns>
    public String[]? GetStrings(String key)
    {
        if (!Args.TryGetValue(key, out var value) || value is not List<Object?> list) return null;
        return list.Select(e => Convert.ToString(e, CultureInfo.InvariantCulture) ?? String.Empty).ToArray();
    }
}

/// <summary>补丁执行结果</summary>
public class ExcelPatchResult
{
    /// <summary>成功的指令（序号:指令名）</summary>
    public List<String> Applied { get; } = [];

    /// <summary>失败的指令与原因</summary>
    public List<String> Failed { get; } = [];

    internal List<String> FailedCodes { get; } = [];

    /// <summary>执行告警（不影响成功判定）</summary>
    public List<String> Warnings { get; } = [];

    /// <summary>是否全部成功（为 false 时不应保存）</summary>
    public Boolean Success => Failed.Count == 0;

    /// <summary>输出 JSON 结果</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"success\": ").Append(Success ? "true" : "false").Append(',').Append(nl);
        sb.Append(pad).Append("\"applied\": ").Append(OfficeJson.Array(Applied)).Append(',').Append(nl);
        sb.Append(pad).Append("\"failed\": ").Append(OfficeJson.Array(Failed)).Append(',').Append(nl);
        sb.Append(pad).Append("\"warnings\": ").Append(OfficeJson.Array(Warnings)).Append(nl);
        sb.Append('}');
        return sb.ToString();
    }
}
