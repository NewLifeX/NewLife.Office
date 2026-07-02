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
}
