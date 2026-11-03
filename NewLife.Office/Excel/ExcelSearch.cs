using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace NewLife.Office.Excel;

/// <summary>工作簿全文搜索：按文本/公式/正则定位到 A1 坐标</summary>
/// <remarks>供 AI 在不知道具体位置时先搜索、再按返回的 <c>Sheet!A1</c> 改写。</remarks>
public static class ExcelSearch
{
    /// <summary>正则匹配超时，防止灾难性回溯拖死调用方</summary>
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

    /// <summary>在整本工作簿中搜索</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="pattern">搜索内容（文本或正则）</param>
    /// <param name="regex">是否按正则表达式匹配</param>
    /// <param name="caseSensitive">是否区分大小写</param>
    /// <param name="maxResults">最大命中数（超出时置 <see cref="ExcelSearchResult.Truncated"/>）</param>
    /// <returns>搜索结果</returns>
    public static ExcelSearchResult Search(ExcelDocument document, String pattern, Boolean regex = false, Boolean caseSensitive = false, Int32 maxResults = 100)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (pattern == null) throw new ArgumentNullException(nameof(pattern));
        if (maxResults <= 0) maxResults = 100;

        var result = new ExcelSearchResult();
        Regex? rx = null;
        if (regex)
        {
            var options = RegexOptions.None;
            if (!caseSensitive) options |= RegexOptions.IgnoreCase;
            rx = new Regex(pattern, options, RegexTimeout);
        }

        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        foreach (var sheet in document.Worksheets)
        {
            var rowCount = ExcelSheetEditor.GetRowCount(sheet);
            for (var r = 0; r < rowCount; r++)
            {
                var cells = sheet.Rows[r];
                var realRow = ExcelSheetEditor.GetActualRow(sheet, r);
                for (var c = 0; c < cells.Length; c++)
                {
                    var value = cells[c];
                    var formula = ExcelSheetEditor.GetFormula(sheet, realRow, c);

                    var valueText = value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
                    var formulaText = formula == null ? null : "=" + formula;

                    var matched = false;
                    if (valueText != null)
                        matched = rx != null ? rx.IsMatch(valueText) : valueText.IndexOf(pattern, comparison) >= 0;
                    if (!matched && formulaText != null)
                        matched = rx != null ? rx.IsMatch(formulaText) : formulaText.IndexOf(pattern, comparison) >= 0;

                    if (!matched) continue;

                    if (result.Hits.Count >= maxResults)
                    {
                        result.Truncated = true;
                        return result;
                    }

                    result.Hits.Add(new ExcelSearchHit
                    {
                        Sheet = sheet.Name,
                        Row = realRow,
                        Column = c,
                        Cell = $"{ExcelCellRange.ColumnName(c)}{realRow + 1}",
                        Value = valueText,
                        Formula = formulaText,
                    });
                }
            }
        }

        return result;
    }
}

/// <summary>搜索命中项</summary>
public class ExcelSearchHit
{
    /// <summary>工作表名称</summary>
    public String Sheet { get; set; } = String.Empty;

    /// <summary>行索引（0 基）</summary>
    public Int32 Row { get; set; }

    /// <summary>列索引（0 基）</summary>
    public Int32 Column { get; set; }

    /// <summary>单元格地址（如 <c>B3</c>）</summary>
    public String Cell { get; set; } = String.Empty;

    /// <summary>单元格值文本</summary>
    public String? Value { get; set; }

    /// <summary>公式（含前导 <c>=</c>）</summary>
    public String? Formula { get; set; }

    /// <summary>带工作表名的完整地址（如 <c>Sheet1!B3</c>）</summary>
    public String Address => $"{ExcelCellRange.QuoteSheet(Sheet)}!{Cell}";
}

/// <summary>搜索结果</summary>
public class ExcelSearchResult
{
    /// <summary>命中列表</summary>
    public List<ExcelSearchHit> Hits { get; } = [];

    /// <summary>是否因超出上限而截断</summary>
    public Boolean Truncated { get; set; }

    /// <summary>输出 JSON</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"count\": ").Append(Hits.Count).Append(',').Append(nl);
        sb.Append(pad).Append("\"truncated\": ").Append(Truncated ? "true" : "false").Append(',').Append(nl);
        sb.Append(pad).Append("\"hits\": [");
        for (var i = 0; i < Hits.Count; i++)
        {
            var hit = Hits[i];
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"address\": ").Append(OfficeJson.Escape(hit.Address));
            sb.Append(", \"sheet\": ").Append(OfficeJson.Escape(hit.Sheet));
            sb.Append(", \"cell\": ").Append(OfficeJson.Escape(hit.Cell));
            sb.Append(", \"row\": ").Append(hit.Row);
            sb.Append(", \"column\": ").Append(hit.Column);
            if (hit.Value != null) sb.Append(", \"value\": ").Append(OfficeJson.Escape(hit.Value));
            if (hit.Formula != null) sb.Append(", \"formula\": ").Append(OfficeJson.Escape(hit.Formula));
            sb.Append('}');
            if (i < Hits.Count - 1) sb.Append(',');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }
}
