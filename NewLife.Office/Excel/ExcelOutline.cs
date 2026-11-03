using System.Globalization;
using System.Text;

namespace NewLife.Office.Excel;

/// <summary>工作簿结构快照：面向 AI/Agent 的表格视图</summary>
/// <remarks>
/// 输出工作表清单（已用区域、表头、前若干行预览、合并区、冻结/筛选、表格对象、图表与批注统计），
/// 让模型先看清结构，再用 A1 坐标读写。锚点即工作表名称。
/// <example>
/// <code>
/// using var reader = new ExcelReader("report.xlsx");
/// var outline = reader.ReadExcel().ToOutline();
/// Console.WriteLine(outline.ToJson());
/// </code>
/// </example>
/// </remarks>
public class ExcelOutline
{
    #region 属性
    /// <summary>工作表快照集合</summary>
    public List<ExcelSheetOutline> Sheets { get; } = [];

    /// <summary>工作簿命名范围</summary>
    public Dictionary<String, String> DefinedNames { get; } = [];

    /// <summary>透视表数量（按包内部件统计）</summary>
    public Int32 PivotTableCount { get; set; }

    /// <summary>切片器数量（按包内部件统计）</summary>
    public Int32 SlicerCount { get; set; }

    /// <summary>每个工作表输出的预览行数</summary>
    public Int32 PreviewRows { get; set; } = 5;
    #endregion

    #region 构造
    /// <summary>从工作簿模型构建快照</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="previewRows">每个工作表输出的预览行数（0 表示不输出）</param>
    /// <returns>结构快照</returns>
    public static ExcelOutline From(ExcelDocument document, Int32 previewRows = 5)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var outline = new ExcelOutline { PreviewRows = Math.Max(0, previewRows) };
        for (var i = 0; i < document.Worksheets.Count; i++)
        {
            var sheet = document.Worksheets[i];
            var item = new ExcelSheetOutline
            {
                Index = i,
                Name = sheet.Name,
                Anchor = sheet.Name,
                Rows = ExcelSheetEditor.GetRowCount(sheet),
                Columns = ExcelSheetEditor.GetColumnCount(sheet),
                FreezeRows = sheet.FreezePane?.Rows ?? 0,
                FreezeCols = sheet.FreezePane?.Cols ?? 0,
                AutoFilter = sheet.AutoFilter,
                ChartCount = sheet.Charts.Count,
                ImageCount = sheet.Images.Count,
                CommentCount = sheet.Comments.Count,
                FormulaCount = sheet.Formulas.Count,
            };

            if (outline.PreviewRows > 0)
            {
                var count = Math.Min(outline.PreviewRows, sheet.Rows.Count);
                for (var r = 0; r < count; r++)
                {
                    var row = sheet.Rows[r];
                    var values = new Object?[row.Length];
                    for (var c = 0; c < row.Length; c++)
                    {
                        var formula = ExcelSheetEditor.GetFormula(sheet, r, c);
                        values[c] = formula != null ? "=" + formula : row[c];
                    }
                    item.Preview.Add(values);
                }
            }

            if (item.Preview.Count > 0)
            {
                foreach (var value in item.Preview[0])
                    item.Header.Add(value == null ? String.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? String.Empty);
            }

            item.Summary.AddRange(BuildSummary(sheet, Math.Min(item.Rows, 1000), Math.Min(item.Columns, 30)));

            foreach (var (sr, sc, er, ec) in sheet.Merges)
            {
                var range = new ExcelCellRange(null, sr, sc, er, ec);
                item.Merges.Add(range.Address);
            }

            foreach (var table in sheet.Tables)
                item.Tables.Add(new ExcelTableOutline { Name = table.Name, Range = table.Range });

            outline.Sheets.Add(item);
        }

        foreach (var kv in document.DefinedNames)
            outline.DefinedNames[kv.Key] = kv.Value;

        foreach (var key in document.OtherParts.Keys)
        {
            if (key.StartsWith("xl/pivotTables/", StringComparison.OrdinalIgnoreCase) && key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                outline.PivotTableCount++;
            else if (key.IndexOf("slicer", StringComparison.OrdinalIgnoreCase) >= 0 && key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                outline.SlicerCount++;
        }

        return outline;
    }

    /// <summary>生成列摘要（类型推断、空值率、Top-N），最多采样 1000 行避免大表拖慢</summary>
    private static List<ExcelColumnSummary> BuildSummary(Worksheet sheet, Int32 rows, Int32 columns)
    {
        var list = new List<ExcelColumnSummary>();
        if (rows <= 0 || columns <= 0) return list;

        for (var c = 0; c < columns; c++)
        {
            var summary = new ExcelColumnSummary { Column = ExcelCellRange.ColumnName(c) };
            if (sheet.Rows.Count > 0 && c < sheet.Rows[0].Length)
                summary.Header = Convert.ToString(sheet.Rows[0][c], CultureInfo.InvariantCulture) ?? String.Empty;

            var nonNull = 0;
            var number = 0;
            var date = 0;
            var boolean = 0;
            var text = 0;
            var counts = new Dictionary<String, Int32>(StringComparer.Ordinal);

            var limit = Math.Min(rows, sheet.Rows.Count);
            var startRow = limit > 1 ? 1 : 0;
            var sample = limit - startRow;
            for (var r = startRow; r < limit; r++)
            {
                var value = c < sheet.Rows[r].Length ? sheet.Rows[r][c] : null;
                if (value == null) continue;

                nonNull++;
                if (value is DateTime) date++;
                else if (value is Boolean) boolean++;
                else if (value is Byte or SByte or Int16 or UInt16 or Int32 or UInt32 or Int64 or UInt64
                    or Single or Double or Decimal) number++;
                else text++;

                var key = Convert.ToString(value, CultureInfo.InvariantCulture) ?? String.Empty;
                counts[key] = counts.TryGetValue(key, out var count) ? count + 1 : 1;
            }

            var kinds = 0;
            if (number > 0) kinds++;
            if (date > 0) kinds++;
            if (boolean > 0) kinds++;
            if (text > 0) kinds++;

            summary.Type = kinds switch
            {
                0 => "empty",
                > 1 => "mixed",
                _ when number > 0 => "number",
                _ when date > 0 => "date",
                _ when boolean > 0 => "bool",
                _ => "text",
            };
            summary.NullRate = sample > 0 ? Math.Round(1d - (Double)nonNull / sample, 3) : 0d;

            foreach (var kv in counts.OrderByDescending(e => e.Value).Take(3))
                summary.Top.Add(kv.Key);

            list.Add(summary);
        }
        return list;
    }
    #endregion

    #region JSON
    /// <summary>输出 JSON 快照</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"sheetCount\": ").Append(Sheets.Count).Append(',').Append(nl);
        sb.Append(pad).Append("\"pivotTables\": ").Append(PivotTableCount).Append(',').Append(nl);
        sb.Append(pad).Append("\"slicers\": ").Append(SlicerCount).Append(',').Append(nl);
        sb.Append(pad).Append("\"definedNames\": {");
        var first = true;
        foreach (var kv in DefinedNames)
        {
            if (!first) sb.Append(", ");
            first = false;
            sb.Append(OfficeJson.Escape(kv.Key)).Append(": ").Append(OfficeJson.Escape(kv.Value));
        }
        sb.Append("},").Append(nl);

        sb.Append(pad).Append("\"sheets\": [");
        for (var i = 0; i < Sheets.Count; i++)
        {
            var sheet = Sheets[i];
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape(sheet.Anchor));
            sb.Append(", \"index\": ").Append(sheet.Index);
            sb.Append(", \"name\": ").Append(OfficeJson.Escape(sheet.Name));
            sb.Append(", \"rows\": ").Append(sheet.Rows);
            sb.Append(", \"columns\": ").Append(sheet.Columns);
            if (sheet.Header.Count > 0) sb.Append(", \"header\": ").Append(OfficeJson.Array(sheet.Header));
            if (sheet.Preview.Count > 0) sb.Append(", \"preview\": ").Append(OfficeJson.Matrix(sheet.Preview));
            if (sheet.Merges.Count > 0) sb.Append(", \"merges\": ").Append(OfficeJson.Array(sheet.Merges));
            if (sheet.FreezeRows > 0 || sheet.FreezeCols > 0)
                sb.Append(", \"freeze\": {\"rows\": ").Append(sheet.FreezeRows).Append(", \"cols\": ").Append(sheet.FreezeCols).Append('}');
            if (!sheet.AutoFilter.IsNullOrEmpty()) sb.Append(", \"autoFilter\": ").Append(OfficeJson.Escape(sheet.AutoFilter));
            if (sheet.Tables.Count > 0)
            {
                sb.Append(", \"tables\": [");
                for (var t = 0; t < sheet.Tables.Count; t++)
                {
                    if (t > 0) sb.Append(", ");
                    sb.Append("{\"name\": ").Append(OfficeJson.Escape(sheet.Tables[t].Name))
                      .Append(", \"range\": ").Append(OfficeJson.Escape(sheet.Tables[t].Range)).Append('}');
                }
                sb.Append(']');
            }
            if (sheet.ChartCount > 0) sb.Append(", \"charts\": ").Append(sheet.ChartCount);
            if (sheet.ImageCount > 0) sb.Append(", \"images\": ").Append(sheet.ImageCount);
            if (sheet.CommentCount > 0) sb.Append(", \"comments\": ").Append(sheet.CommentCount);
            if (sheet.FormulaCount > 0) sb.Append(", \"formulas\": ").Append(sheet.FormulaCount);
            if (sheet.Summary.Count > 0)
            {
                sb.Append(", \"summary\": [");
                for (var s = 0; s < sheet.Summary.Count; s++)
                {
                    var column = sheet.Summary[s];
                    if (s > 0) sb.Append(", ");
                    sb.Append("{\"column\": ").Append(OfficeJson.Escape(column.Column));
                    if (!column.Header.IsNullOrEmpty()) sb.Append(", \"header\": ").Append(OfficeJson.Escape(column.Header));
                    sb.Append(", \"type\": ").Append(OfficeJson.Escape(column.Type));
                    sb.Append(", \"nullRate\": ").Append(OfficeJson.Value(column.NullRate));
                    if (column.Top.Count > 0) sb.Append(", \"top\": ").Append(OfficeJson.Array(column.Top));
                    sb.Append('}');
                }
                sb.Append(']');
            }
            sb.Append('}');
            if (i < Sheets.Count - 1) sb.Append(',');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }
    #endregion
}

/// <summary>单个工作表的结构快照</summary>
public class ExcelSheetOutline
{
    /// <summary>锚点（即工作表名称）</summary>
    public String Anchor { get; set; } = String.Empty;

    /// <summary>工作表索引（0 基）</summary>
    public Int32 Index { get; set; }

    /// <summary>工作表名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>已用行数</summary>
    public Int32 Rows { get; set; }

    /// <summary>已用列数</summary>
    public Int32 Columns { get; set; }

    /// <summary>表头（预览首行转字符串）</summary>
    public List<String> Header { get; } = [];

    /// <summary>预览行（默认前 5 行，公式以 <c>=</c> 前缀返回）</summary>
    public List<Object?[]> Preview { get; } = [];

    /// <summary>合并区域（A1 记法）</summary>
    public List<String> Merges { get; } = [];

    /// <summary>冻结行数</summary>
    public Int32 FreezeRows { get; set; }

    /// <summary>冻结列数</summary>
    public Int32 FreezeCols { get; set; }

    /// <summary>自动筛选范围</summary>
    public String? AutoFilter { get; set; }

    /// <summary>结构化表格集合</summary>
    public List<ExcelTableOutline> Tables { get; } = [];

    /// <summary>列摘要（类型推断、空值率、Top-3），最多采样前 1000 行</summary>
    public List<ExcelColumnSummary> Summary { get; } = [];

    /// <summary>图表数量</summary>
    public Int32 ChartCount { get; set; }

    /// <summary>图片数量</summary>
    public Int32 ImageCount { get; set; }

    /// <summary>批注数量</summary>
    public Int32 CommentCount { get; set; }

    /// <summary>公式数量</summary>
    public Int32 FormulaCount { get; set; }
}

/// <summary>结构化表格摘要</summary>
public class ExcelTableOutline
{
    /// <summary>表格名称</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>表格范围（A1 记法）</summary>
    public String Range { get; set; } = String.Empty;
}

/// <summary>列摘要</summary>
public class ExcelColumnSummary
{
    /// <summary>列名（如 A）</summary>
    public String Column { get; set; } = String.Empty;

    /// <summary>表头文本</summary>
    public String? Header { get; set; }

    /// <summary>推断类型：number/date/bool/text/mixed/empty</summary>
    public String Type { get; set; } = "empty";

    /// <summary>空值率（按采样行计算）</summary>
    public Double NullRate { get; set; }

    /// <summary>出现次数最多的前 3 个值</summary>
    public List<String> Top { get; } = [];
}
