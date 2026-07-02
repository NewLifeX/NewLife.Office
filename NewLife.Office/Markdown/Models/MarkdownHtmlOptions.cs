namespace NewLife.Office.Markdown;

/// <summary>Markdown → HTML 转换选项</summary>
public sealed class MarkdownHtmlOptions
{
    /// <summary>是否为代码块添加 language-xxx CSS 类（默认 true）</summary>
    public Boolean AddLanguageClass { get; set; } = true;

    /// <summary>是否在链接上添加 target="_blank"（默认 false）</summary>
    public Boolean ExternalLinkTarget { get; set; }

    /// <summary>是否对链接添加 rel="noopener noreferrer"（默认 false）</summary>
    public Boolean SafeLinks { get; set; }
}
