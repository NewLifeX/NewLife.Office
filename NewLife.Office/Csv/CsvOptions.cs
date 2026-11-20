using System.Text;

namespace NewLife.Office.Csv;

/// <summary>CSV 读取与导入选项</summary>
/// <remarks>
/// 默认值面向"直接打开别人的 CSV"：自动识别编码与分隔符。
/// 解析与写入由 NewLife.Core 的 <c>CsvFile</c> 完成，本类只描述入口策略。
/// </remarks>
public sealed class CsvOptions
{
    /// <summary>字段分隔符；<c>'\0'</c> 表示按首行自动识别（逗号/分号/制表符）</summary>
    public Char Separator { get; set; } = '\0';

    /// <summary>文本编码；null 表示自动识别（BOM → 严格 UTF-8 → GBK 兜底）</summary>
    public Encoding? Encoding { get; set; }

    /// <summary>首行是否表头；仅 <see cref="Excel.ExcelDataFrame.FromCsv(String, CsvOptions)"/> 使用</summary>
    public Boolean HasHeader { get; set; } = true;

    /// <summary>是否按列内容推断数字/日期/布尔；仅 <see cref="Excel.ExcelDataFrame.FromCsv(String, CsvOptions)"/> 使用</summary>
    public Boolean InferTypes { get; set; } = true;
}
