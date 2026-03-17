namespace NewLife.Office.Markdown;

/// <summary>Markdown 块级元素</summary>
/// <remarks>
/// 表示文档中的块结构，包括标题、段落、代码块、列表等。
/// 容器块（列表、引用块）通过 <see cref="Children"/> 嵌套子块；
/// 叶子块（段落、标题、单元格）通过 <see cref="Inlines"/> 保存行内内容。
/// </remarks>
public sealed class MarkdownBlock
{
    #region 属性
    /// <summary>块类型</summary>
    public MarkdownBlockType Type { get; set; }

    /// <summary>标题等级（Heading 使用，1-6）</summary>
    public Int32 Level { get; set; }

    /// <summary>代码块语言标识符（CodeBlock 使用）</summary>
    public String Language { get; set; } = String.Empty;

    /// <summary>代码/HTML 块的原始文本（CodeBlock/HtmlBlock 使用）</summary>
    public String RawText { get; set; } = String.Empty;

    /// <summary>有序列表起始序号（OrderedList 使用）</summary>
    public Int32 OrderedStart { get; set; } = 1;

    /// <summary>列表项是否为任务项（ListItem 使用）</summary>
    public Boolean IsTaskItem { get; set; }

    /// <summary>任务项是否已勾选（IsTaskItem = true 时有效）</summary>
    public Boolean IsChecked { get; set; }

    /// <summary>表格单元格是否为表头（TableCell 使用）</summary>
    public Boolean IsHeader { get; set; }

    /// <summary>表格单元格对齐方式（TableCell 使用，"left"/"center"/"right"/""）</summary>
    public String Alignment { get; set; } = String.Empty;

    /// <summary>子块列表（容器块使用：列表/引用块/列表项/表格/表格行）</summary>
    public List<MarkdownBlock> Children { get; } = [];

    /// <summary>行内元素列表（叶子块使用：段落/标题/列表项文本/单元格）</summary>
    public List<MarkdownInline> Inlines { get; } = [];
    #endregion

    #region 工厂方法
    /// <summary>创建标题块</summary>
    /// <param name="level">等级（1-6）</param>
    /// <param name="inlines">行内内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateHeading(Int32 level, IEnumerable<MarkdownInline> inlines)
    {
        var b = new MarkdownBlock { Type = MarkdownBlockType.Heading, Level = level };
        b.Inlines.AddRange(inlines);
        return b;
    }

    /// <summary>创建段落块</summary>
    /// <param name="inlines">行内内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateParagraph(IEnumerable<MarkdownInline> inlines)
    {
        var b = new MarkdownBlock { Type = MarkdownBlockType.Paragraph };
        b.Inlines.AddRange(inlines);
        return b;
    }

    /// <summary>创建代码块</summary>
    /// <param name="code">代码文本</param>
    /// <param name="language">语言标识（可空）</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateCodeBlock(String code, String language = "") =>
        new() { Type = MarkdownBlockType.CodeBlock, RawText = code, Language = language ?? String.Empty };

    /// <summary>创建分隔线</summary>
    /// <returns>块</returns>
    public static MarkdownBlock CreateThematicBreak() => new() { Type = MarkdownBlockType.ThematicBreak };

    /// <summary>创建引用块</summary>
    /// <param name="children">子块</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateBlockQuote(IEnumerable<MarkdownBlock> children)
    {
        var b = new MarkdownBlock { Type = MarkdownBlockType.BlockQuote };
        b.Children.AddRange(children);
        return b;
    }

    /// <summary>创建无序列表</summary>
    /// <param name="items">列表项</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateBulletList(IEnumerable<MarkdownBlock> items)
    {
        var b = new MarkdownBlock { Type = MarkdownBlockType.BulletList };
        b.Children.AddRange(items);
        return b;
    }

    /// <summary>创建有序列表</summary>
    /// <param name="items">列表项</param>
    /// <param name="start">起始序号</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateOrderedList(IEnumerable<MarkdownBlock> items, Int32 start = 1)
    {
        var b = new MarkdownBlock { Type = MarkdownBlockType.OrderedList, OrderedStart = start };
        b.Children.AddRange(items);
        return b;
    }

    /// <summary>创建列表项（简单行内内容）</summary>
    /// <param name="inlines">行内内容</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateListItem(IEnumerable<MarkdownInline> inlines,
        Boolean isTaskItem = false, Boolean isChecked = false)
    {
        var b = new MarkdownBlock
        {
            Type = MarkdownBlockType.ListItem,
            IsTaskItem = isTaskItem,
            IsChecked = isChecked,
        };
        b.Inlines.AddRange(inlines);
        return b;
    }

    /// <summary>创建列表项（嵌套块内容）</summary>
    /// <param name="children">子块</param>
    /// <param name="isTaskItem">是否任务项</param>
    /// <param name="isChecked">是否已勾选</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateListItemWithBlocks(IEnumerable<MarkdownBlock> children,
        Boolean isTaskItem = false, Boolean isChecked = false)
    {
        var b = new MarkdownBlock
        {
            Type = MarkdownBlockType.ListItem,
            IsTaskItem = isTaskItem,
            IsChecked = isChecked,
        };
        b.Children.AddRange(children);
        return b;
    }

    /// <summary>创建 HTML 块</summary>
    /// <param name="html">原始 HTML 内容</param>
    /// <returns>块</returns>
    public static MarkdownBlock CreateHtmlBlock(String html) =>
        new() { Type = MarkdownBlockType.HtmlBlock, RawText = html };
    #endregion

    #region 方法
    /// <summary>获取纯文本内容（递归展开行内元素）</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        if (Type == MarkdownBlockType.CodeBlock) return RawText;
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
