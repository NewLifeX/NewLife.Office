namespace NewLife.Office.Markdown;

/// <summary>分隔线块</summary>
public sealed class ThematicBreakBlock : MarkdownBlock
{
    /// <summary>创建分隔线块</summary>
    public ThematicBreakBlock()
    {
        Type = MarkdownBlockType.ThematicBreak;
    }
}
