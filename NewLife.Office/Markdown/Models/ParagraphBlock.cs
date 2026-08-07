namespace NewLife.Office.Markdown;

/// <summary>段落块</summary>
public sealed class ParagraphBlock : MarkdownBlock
{
    /// <summary>创建段落块</summary>
    /// <param name="inlines">行内内容</param>
    public ParagraphBlock(IEnumerable<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.Paragraph;
        Inlines.AddRange(inlines);
    }

    /// <summary>创建段落块（MD24：直接接管 List，省 AddRange 复制）</summary>
    /// <param name="inlines">行内列表</param>
    internal ParagraphBlock(List<MarkdownInline> inlines)
    {
        Type = MarkdownBlockType.Paragraph;
        SetInlines(inlines);
    }
}
