namespace NewLife.Office.Markdown;

/// <summary>Markdown → HTML 转换选项</summary>
public sealed class MarkdownHtmlOptions
{
    /// <summary>是否为代码块添加 language-xxx CSS 类（默认 true）</summary>
    public Boolean AddLanguageClass { get; set; } = true;

    /// <summary>是否在链接上添加 target="_blank"（默认 false）</summary>
    public Boolean ExternalLinkTarget { get; set; }

    /// <summary>是否对链接添加 rel="noopener noreferrer"（默认 false）</summary>
    public Boolean SafeLinks { get; set; } = true;

    /// <summary>是否放行 Markdown 中的原始 HTML（默认 true，与 CommonMark 一致）</summary>
    /// <remarks>
    /// 转换不受信任的 Markdown 时请设为 false，否则原文里的 &lt;script&gt; 等标签会原样进入输出。
    /// </remarks>
    public Boolean AllowRawHtml { get; set; } = true;

    /// <summary>缩写映射表（键为缩写文本，值为全称），用于渲染 &lt;abbr&gt; 标签（可选）</summary>
    public Dictionary<String, String>? Abbreviations { get; set; }
}
