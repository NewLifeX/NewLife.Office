namespace NewLife.Office.Markdown;

/// <summary>HTML → Markdown 转换选项</summary>
public sealed class HtmlToMarkdownOptions
{
    /// <summary>是否使用 GitHub Flavored Markdown 扩展（默认 true）。启用后支持任务列表、删除线、表格等 GFM 语法。</summary>
    public Boolean GithubFlavored { get; set; } = true;

    /// <summary>无序列表项目符号字符（默认 "-"）。可改为 "*" 或 "+"。</summary>
    public String ListBulletChar { get; set; } = "-";

    /// <summary>未知 HTML 标签处理策略</summary>
    public UnknownTagStrategy UnknownTags { get; set; } = UnknownTagStrategy.PassThrough;

    /// <summary>允许的 URI 协议白名单（分号分隔，默认 "http;https;ftp;mailto"）</summary>
    public String WhitelistUriSchemes { get; set; } = "http;https;ftp;mailto";

    /// <summary>是否在链接 URL 前后添加尖括号 &lt;url&gt;（默认 false，仅裸 URL 使用）</summary>
    public Boolean AutoLink { get; set; }

    /// <summary>标题下划线风格。true 使用 Setext 风格（H1= H2-），false 使用 ATX 风格（# ##）</summary>
    public Boolean SetextHeadings { get; set; }

    /// <summary>粗体标记风格。true 使用 **，false 使用 __</summary>
    public Boolean AsteriskEmphasis { get; set; } = true;

    /// <summary>斜体标记风格。true 使用 *，false 使用 _</summary>
    public Boolean AsteriskItalic { get; set; } = true;
}

/// <summary>未知 HTML 标签处理策略</summary>
public enum UnknownTagStrategy
{
    /// <summary>原样透传 HTML 标签</summary>
    PassThrough,

    /// <summary>丢弃标签，仅保留内部文本内容</summary>
    Drop,

    /// <summary>将标签转义为 HTML 实体输出</summary>
    Escape,
}
