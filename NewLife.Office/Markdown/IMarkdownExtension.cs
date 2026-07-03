namespace NewLife.Office.Markdown;

/// <summary>Markdown 扩展接口（MD06-02）</summary>
/// <remarks>
/// 实现此接口以注册自定义解析器、转换器或渲染器到 MarkdownPipeline。
/// </remarks>
public interface IMarkdownExtension
{
    /// <summary>扩展名称</summary>
    String Name { get; }

    /// <summary>配置管线（在解析前调用）</summary>
    /// <param name="pipeline">管线对象</param>
    void Setup(MarkdownPipeline pipeline);
}
