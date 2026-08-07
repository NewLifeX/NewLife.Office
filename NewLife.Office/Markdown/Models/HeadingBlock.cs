namespace NewLife.Office.Markdown;

/// <summary>标题块（H1–H6）</summary>
public sealed class HeadingBlock : MarkdownBlock
{
    /// <summary>标题等级（1-6）</summary>
    public Int32 Level { get; set; }

    /// <summary>创建标题块</summary>
    /// <param name="level">等级（1-6）</param>
    /// <param name="inlines">行内内容</param>
    /// <returns>标题块</returns>
    public HeadingBlock(Int32 level, IEnumerable<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.Heading;
        Level = level;
        Inlines.AddRange(inlines);
    }

    /// <summary>创建标题块（MD24：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="level">等级（1-6）</param>
    /// <param name="inlines">行内列表</param>
    /// <returns>标题块</returns>
    internal HeadingBlock(Int32 level, List<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.Heading;
        Level = level;
        SetInlines(inlines);
    }
}
