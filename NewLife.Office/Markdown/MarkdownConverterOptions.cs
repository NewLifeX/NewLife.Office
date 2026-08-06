namespace NewLife.Office.Markdown;

/// <summary>图片处理策略</summary>
public enum MarkdownImageHandling
{
    /// <summary>跳过图片（AI token 友好，默认）</summary>
    Skip,

    /// <summary>输出图片引用（![alt](文件名)）</summary>
    Reference,

    /// <summary>输出内嵌图片（data URI 或源文件相对路径）</summary>
    Embed,
}

/// <summary>格式转 Markdown 转换选项</summary>
/// <remarks>
/// 控制元数据、图片、噪音过滤等转换行为，供 <see cref="FormatToMarkdown"/> 各入口使用。
/// 默认值面向 AI 知识库场景：输出 FrontMatter 元数据、跳过图片、过滤噪音。
/// </remarks>
public class MarkdownConverterOptions
{
    /// <summary>是否输出 YAML FrontMatter 元数据（标题/作者/日期/源文件），默认 true</summary>
    public Boolean Metadata { get; set; } = true;

    /// <summary>图片处理策略，默认跳过</summary>
    public MarkdownImageHandling ImageHandling { get; set; } = MarkdownImageHandling.Skip;

    /// <summary>是否包含页眉页脚（RTF 等），默认 false（对标 anydoc 固定策略）</summary>
    public Boolean IncludeHeadersFooters { get; set; }

    /// <summary>是否过滤噪音（空段落/重复空白行），默认 true</summary>
    public Boolean CleanNoise { get; set; } = true;
}
