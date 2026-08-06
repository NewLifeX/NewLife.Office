namespace NewLife.Office.Markdown;

/// <summary>Markdown 提取扩展方法</summary>
/// <remarks>
/// 提供统一入口：优先走 <see cref="IMarkdownDocumentExtractable"/> 结构化 AST，
/// 回退到 <see cref="IMarkdownExtractable"/> 字符串实现，保持两种接口并存兼容。
/// </remarks>
public static class MarkdownExtension
{
    /// <summary>优先从结构化 AST 提取 Markdown 字符串，回退到旧字符串实现</summary>
    /// <param name="extractable">支持 Markdown 提取的对象</param>
    /// <returns>Markdown 字符串，不支持或无内容返回 null</returns>
    public static String? ToMarkdown(this IMarkdownExtractable extractable)
    {
        if (extractable == null) return null;

        if (extractable is IMarkdownDocumentExtractable de)
        {
            var doc = de.ToMarkdownDocument();
            if (doc != null) return doc.ToMarkdown();
        }
        return extractable.ExtractMarkdown();
    }
}
