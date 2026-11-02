using NewLife.Office.Markdown;

namespace NewLife.Office;

/// <summary>支持按转换选项提取 Markdown 的读取器</summary>
/// <remarks>
/// 比 <see cref="IMarkdownDocumentExtractable"/> 多一个选项参数，用于控制图片落地、
/// 页眉页脚输出等细节。<see cref="Markdown.FormatToMarkdown"/> 会优先走本接口。
/// </remarks>
public interface IMarkdownOptionExtractable
{
    /// <summary>按选项提取 Markdown 文档（AST）</summary>
    /// <param name="options">转换选项，null 时使用读取器默认行为</param>
    /// <returns>Markdown 文档对象，无内容返回 null</returns>
    MarkdownDocument? ToMarkdownDocument(MarkdownConverterOptions? options);
}
