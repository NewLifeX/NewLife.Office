using NewLife.Office.Markdown;

namespace NewLife.Office;

/// <summary>Markdown 文档提取接口，表示读取器/文档可提取结构化 Markdown 文档对象（AST）</summary>
/// <remarks>
/// 由各格式读取器或文档类实现，支持从 OfficeFactory 统一调用。
/// 与 <see cref="IMarkdownExtractable"/> 返回字符串不同，本接口直接返回
/// <see cref="MarkdownDocument"/> 结构化对象模型，供 AI 知识库等场景做
/// 结构化后处理（元数据 FrontMatter、图片策略、按块截断控制 token 预算）。
/// <para>对标 anydoc 统一 Document 模型：各格式前端解析为共享模型，
/// 再由单一序列化器输出 GFM Markdown。</para>
/// </remarks>
public interface IMarkdownDocumentExtractable
{
    /// <summary>提取 Markdown 文档对象（AST）</summary>
    /// <returns>Markdown 文档对象，若不支持或无内容则返回 null</returns>
    MarkdownDocument? ToMarkdownDocument();
}
