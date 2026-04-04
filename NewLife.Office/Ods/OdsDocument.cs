using System.Text;

namespace NewLife.Office.Ods;

/// <summary>ODS 文档包装，封装工作表列表并提供文本/Markdown 提取能力</summary>
public class OdsDocument : ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>工作表列表</summary>
    public List<OdsSheet> Sheets { get; }
    #endregion

    #region 构造
    /// <summary>实例化 ODS 文档包装</summary>
    /// <param name="sheets">工作表列表</param>
    public OdsDocument(List<OdsSheet> sheets) => Sheets = sheets ?? [];
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（CSV 格式，逗号分隔）</summary>
    /// <returns>CSV 格式文本</returns>
    public String? ExtractText()
    {
        if (Sheets == null || Sheets.Count == 0) return null;

        var sb = new StringBuilder();
        for (var si = 0; si < Sheets.Count; si++)
        {
            var sheet = Sheets[si];
            if (Sheets.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheet.Name}");
            }

            foreach (var row in sheet.Rows)
            {
                for (var i = 0; i < row.Length; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(CsvEscape(row[i]));
                }
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }

    /// <summary>提取 Markdown 格式（表格）</summary>
    /// <returns>Markdown 表格字符串</returns>
    public String? ExtractMarkdown()
    {
        if (Sheets == null || Sheets.Count == 0) return null;

        var sb = new StringBuilder();
        for (var si = 0; si < Sheets.Count; si++)
        {
            var sheet = Sheets[si];
            if (Sheets.Count > 1)
            {
                if (si > 0) sb.AppendLine();
                sb.AppendLine($"## {sheet.Name}");
                sb.AppendLine();
            }

            if (sheet.Rows.Count == 0) continue;

            // 第一行作为表头
            var header = sheet.Rows[0];
            sb.Append('|');
            foreach (var cell in header)
            {
                sb.Append(' ').Append(MdEscape(cell)).Append(" |");
            }
            sb.AppendLine();

            // 分隔线
            sb.Append('|');
            for (var i = 0; i < header.Length; i++)
            {
                sb.Append(" --- |");
            }
            sb.AppendLine();

            // 数据行
            for (var ri = 1; ri < sheet.Rows.Count; ri++)
            {
                var row = sheet.Rows[ri];
                sb.Append('|');
                for (var i = 0; i < header.Length; i++)
                {
                    var val = i < row.Length ? row[i] : "";
                    sb.Append(' ').Append(MdEscape(val)).Append(" |");
                }
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }
    #endregion

    #region 辅助
    private static String CsvEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        if (value.IndexOfAny([',', '"', '\n', '\r']) >= 0)
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    private static String MdEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        return value.Replace("|", "\\|").Replace("\n", " ").Replace("\r", "");
    }
    #endregion
}
