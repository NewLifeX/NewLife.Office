namespace NewLife.Office.Markdown;

/// <summary>Markdown 块级元素（抽象基类）</summary>
/// <remarks>
/// 所有块类型的抽象基类。容器块通过 <see cref="Children"/> 嵌套子块；
/// 叶子块通过 <see cref="Inlines"/> 保存行内内容。
/// 具体类型见子类：<see cref="HeadingBlock"/>、<see cref="ParagraphBlock"/>、
/// <see cref="CodeBlock"/>、<see cref="BlockQuoteBlock"/>、<see cref="BulletListBlock"/>、
/// <see cref="OrderedListBlock"/>、<see cref="ListItemBlock"/>、<see cref="TableBlock"/>、
/// <see cref="TableRowBlock"/>、<see cref="TableCellBlock"/>、<see cref="ThematicBreakBlock"/>、
/// <see cref="HtmlBlock"/>。
/// </remarks>
public abstract class MarkdownBlock
{
    #region 属性
    /// <summary>块类型</summary>
    public MarkdownBlockType Type { get; protected set; }

    /// <summary>子块列表（容器块使用：列表/引用块/列表项/表格行）</summary>
    public List<MarkdownBlock> Children { get; } = [];

    /// <summary>行内元素列表（叶子块使用：段落/标题/列表项文本/单元格）</summary>
    public List<MarkdownInline> Inlines { get; } = [];
    #endregion

    #region 工厂方法（向后兼容）
    /// <summary>创建标题块</summary>
    /// <param name="level">等级（1-6）</param>
    /// <param name="inlines">行内内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateHeading(Int32 level, IEnumerable<MarkdownInline> inlines)
        => new HeadingBlock(level, inlines);

    /// <summary>创建段落块</summary>
    /// <param name="inlines">行内内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateParagraph(IEnumerable<MarkdownInline> inlines)
        => new ParagraphBlock(inlines);

    /// <summary>创建代码块</summary>
    /// <param name="code">代码文本</param>
    /// <param name="language">语言标识（可空）</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateCodeBlock(String code, String language = "")
        => new CodeBlock(code, language);

    /// <summary>创建分隔线</summary>
    /// <returns>块</returns>
    public static MarkdownBlock CreateThematicBreak() => new ThematicBreakBlock();

    /// <summary>创建引用块</summary>
    /// <param name="children">子块</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateBlockQuote(IEnumerable<MarkdownBlock> children)
        => new BlockQuoteBlock(children);

    /// <summary>创建无序列表</summary>
    /// <param name="items">列表项</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateBulletList(IEnumerable<MarkdownBlock> items)
        => new BulletListBlock(items);

    /// <summary>创建有序列表</summary>
    /// <param name="items">列表项</param>
    /// <param name="start">起始序号</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateOrderedList(IEnumerable<MarkdownBlock> items, Int32 start = 1)
        => new OrderedListBlock(items, start);

    /// <summary>创建列表项（简单行内内容）</summary>
    /// <param name="inlines">行内内容</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateListItem(IEnumerable<MarkdownInline> inlines,
        Boolean isTaskItem = false, Boolean isChecked = false)
        => new ListItemBlock(inlines, isTaskItem, isChecked);

    /// <summary>创建列表项（嵌套块内容）</summary>
    /// <param name="children">子块</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateListItemWithBlocks(IEnumerable<MarkdownBlock> children,
        Boolean isTaskItem = false, Boolean isChecked = false)
        => ListItemBlock.CreateWithBlocks(children, isTaskItem, isChecked);

    /// <summary>创建 HTML 块</summary>
    /// <param name="html">原始 HTML 内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateHtmlBlock(String html) => new HtmlBlock(html);
    #endregion

    #region 方法
    /// <summary>获取纯文本内容（递归展开行内元素）</summary>
    /// <returns>纯文本字符串</returns>
    public virtual String GetPlainText()
    {
        if (Inlines.Count > 0)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var inline in Inlines) sb.Append(inline.GetPlainText());
            return sb.ToString();
        }
        if (Children.Count > 0)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var child in Children) sb.AppendLine(child.GetPlainText());
            return sb.ToString().TrimEnd();
        }
        return String.Empty;
    }

    /// <inheritdoc/>
    public override String ToString()
    {
        var text = GetPlainText().Replace('\n', ' ');
        return $"{Type}: {(text.Length > 60 ? text[..60] : text)}";
    }
    #endregion
}
