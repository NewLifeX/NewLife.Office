namespace NewLife.Office.Markdown;

/// <summary>Markdown 块级元素类型</summary>
public enum MarkdownBlockType
{
    /// <summary>文档根节点</summary>
    Document,
    /// <summary>标题（H1–H6，通过 Level 区分）</summary>
    Heading,
    /// <summary>普通段落</summary>
    Paragraph,
    /// <summary>代码块（围栏式或缩进式）</summary>
    CodeBlock,
    /// <summary>引用块</summary>
    BlockQuote,
    /// <summary>无序列表</summary>
    BulletList,
    /// <summary>有序列表</summary>
    OrderedList,
    /// <summary>列表项</summary>
    ListItem,
    /// <summary>表格</summary>
    Table,
    /// <summary>表格行</summary>
    TableRow,
    /// <summary>表格单元格</summary>
    TableCell,
    /// <summary>分隔线</summary>
    ThematicBreak,
    /// <summary>原始 HTML 块</summary>
    HtmlBlock,
}
