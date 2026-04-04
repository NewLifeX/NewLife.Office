namespace NewLife.Office;

/// <summary>Markdown 提取接口，表示读取器/文档可提取 Markdown 格式内容</summary>
/// <remarks>
/// 由各格式读取器或文档类实现，支持从 OfficeFactory 统一调用。
/// 表格类（xlsx/xls/ods）输出 Markdown 表格；文档类保留部分格式结构。
/// </remarks>
public interface IMarkdownExtractable
{
    /// <summary>提取 Markdown 格式内容</summary>
    /// <returns>Markdown 字符串，若不支持或无内容则返回 null</returns>
    String? ExtractMarkdown();
}
