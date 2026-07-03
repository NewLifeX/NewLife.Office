namespace NewLife.Office.Markdown;

/// <summary>定义列表块（MD05-07）</summary>
public sealed class DefinitionListBlock : MarkdownBlock
{
    /// <summary>创建定义列表</summary>
    public DefinitionListBlock()
    {
        Type = MarkdownBlockType.DefinitionList;
    }
}

/// <summary>定义术语块（MD05-07）</summary>
public sealed class DefinitionTermBlock : MarkdownBlock
{
    /// <summary>创建定义术语</summary>
    /// <param name="inlines">术语文本</param>
    public DefinitionTermBlock(IEnumerable<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.DefinitionTerm;
        Inlines.AddRange(inlines);
    }
}

/// <summary>定义描述块（MD05-07）</summary>
public sealed class DefinitionDescriptionBlock : MarkdownBlock
{
    /// <summary>创建定义描述</summary>
    /// <param name="inlines">描述文本</param>
    public DefinitionDescriptionBlock(IEnumerable<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.DefinitionDescription;
        Inlines.AddRange(inlines);
    }
}
