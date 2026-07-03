using System;
using System.Collections.Generic;

namespace NewLife.Office.Markdown;

/// <summary>Markdown 处理管线（MD06-01）</summary>
/// <remarks>
/// 可配置的解析→转换→渲染管线，支持启用/禁用扩展、注册自定义解析器和转换器。
/// <para>使用示例：</para>
/// <code>
/// var pipeline = new MarkdownPipeline();
/// pipeline.EnableAutoLinks = true;
/// pipeline.EnableFootnotes = true;
/// var doc = MarkdownDocument.Parse("text", pipeline);
/// var html = doc.ToHtml(pipeline.ToHtmlOptions());
/// </code>
/// </remarks>
public sealed class MarkdownPipeline
{
    #region 属性
    /// <summary>是否启用自动链接（裸 URL）</summary>
    public Boolean EnableAutoLinks { get; set; } = true;

    /// <summary>是否启用 YAML Front Matter</summary>
    public Boolean EnableFrontMatter { get; set; } = true;

    /// <summary>是否启用脚注</summary>
    public Boolean EnableFootnotes { get; set; } = true;

    /// <summary>是否启用数学公式</summary>
    public Boolean EnableMath { get; set; } = true;

    /// <summary>是否启用 Emoji 短码</summary>
    public Boolean EnableEmoji { get; set; } = true;

    /// <summary>是否启用缩写</summary>
    public Boolean EnableAbbreviations { get; set; } = true;

    /// <summary>是否启用定义列表</summary>
    public Boolean EnableDefinitionLists { get; set; } = true;

    /// <summary>是否启用自定义属性 {.class #id}</summary>
    public Boolean EnableCustomAttributes { get; set; } = true;

    /// <summary>是否启用表格（GFM）</summary>
    public Boolean EnableTables { get; set; } = true;

    /// <summary>是否启用任务列表（GFM）</summary>
    public Boolean EnableTaskLists { get; set; } = true;

    /// <summary>是否启用删除线（GFM）</summary>
    public Boolean EnableStrikethrough { get; set; } = true;

    /// <summary>无序列表项目符号字符</summary>
    public String ListBulletChar { get; set; } = "-";

    /// <summary>已注册的扩展列表</summary>
    public List<IMarkdownExtension> Extensions { get; } = [];
    #endregion

    #region 构造
    /// <summary>创建默认管线（所有扩展启用）</summary>
    public MarkdownPipeline() { }

    /// <summary>创建精简管线（仅 CommonMark 核心语法）</summary>
    public static MarkdownPipeline CreateStrict()
    {
        return new MarkdownPipeline
        {
            EnableAutoLinks = false,
            EnableFrontMatter = false,
            EnableFootnotes = false,
            EnableMath = false,
            EnableEmoji = false,
            EnableAbbreviations = false,
            EnableDefinitionLists = false,
            EnableCustomAttributes = false,
            EnableTables = false,
            EnableTaskLists = false,
            EnableStrikethrough = false,
        };
    }

    /// <summary>创建 GFM 管线（CommonMark + GitHub 扩展）</summary>
    public static MarkdownPipeline CreateGfm()
    {
        return new MarkdownPipeline
        {
            EnableAutoLinks = true,
            EnableTables = true,
            EnableTaskLists = true,
            EnableStrikethrough = true,
            EnableFrontMatter = false,
            EnableFootnotes = false,
            EnableMath = false,
            EnableEmoji = true,
            EnableAbbreviations = false,
            EnableDefinitionLists = false,
            EnableCustomAttributes = false,
        };
    }
    #endregion

    #region 方法
    /// <summary>注册扩展</summary>
    /// <param name="extension">扩展实例</param>
    public MarkdownPipeline Use(IMarkdownExtension extension)
    {
        Extensions.Add(extension);
        extension.Setup(this);
        return this;
    }

    /// <summary>生成对应的 HTML 转换选项</summary>
    public MarkdownHtmlOptions ToHtmlOptions()
    {
        return new MarkdownHtmlOptions
        {
            SafeLinks = true,
            ExternalLinkTarget = true,
        };
    }

    /// <summary>生成对应的 HTML→Markdown 转换选项</summary>
    public HtmlToMarkdownOptions ToHtmlToMarkdownOptions()
    {
        return new HtmlToMarkdownOptions
        {
            GithubFlavored = EnableTables,
            ListBulletChar = ListBulletChar,
        };
    }
    #endregion
}
