namespace NewLife.Office;

/// <summary>文本提取接口，表示读取器/文档可提取纯文本</summary>
/// <remarks>
/// 由各格式读取器或文档类实现，支持从 OfficeFactory 统一调用。
/// 表格类（xlsx/xls/ods）输出 CSV 格式；文档类输出无格式纯文本。
/// </remarks>
public interface ITextExtractable
{
    /// <summary>提取纯文本内容</summary>
    /// <returns>纯文本字符串，若不支持或无内容则返回 null</returns>
    String? ExtractText();
}
