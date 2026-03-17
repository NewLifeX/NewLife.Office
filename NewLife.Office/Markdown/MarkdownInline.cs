namespace NewLife.Office.Markdown;

/// <summary>Markdown 行内元素</summary>
/// <remarks>
/// 表示段落/标题等容器块内的行内内容，如文本、粗体、链接、图片等。
/// </remarks>
public sealed class MarkdownInline
{
    #region 属性
    /// <summary>元素类型</summary>
    public MarkdownInlineType Type { get; set; }

    /// <summary>文本内容（Text/Code/RawHtml 使用）</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>链接/图片目标 URL</summary>
    public String Href { get; set; } = String.Empty;

    /// <summary>链接/图片标题（可选）</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>图片替代文本（Image 使用）</summary>
    public String Alt { get; set; } = String.Empty;

    /// <summary>子行内元素（Strong/Emphasis/Link 可嵌套）</summary>
    public List<MarkdownInline> Children { get; } = [];
    #endregion

    #region 工厂方法
    /// <summary>创建文本节点</summary>
    /// <param name="text">文本内容</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateText(String text) =>
        new() { Type = MarkdownInlineType.Text, Text = text };

    /// <summary>创建粗体节点</summary>
    /// <param name="children">子行内元素</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateStrong(IEnumerable<MarkdownInline> children)
    {
        var node = new MarkdownInline { Type = MarkdownInlineType.Strong };
        node.Children.AddRange(children);
        return node;
    }

    /// <summary>创建斜体节点</summary>
    /// <param name="children">子行内元素</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateEmphasis(IEnumerable<MarkdownInline> children)
    {
        var node = new MarkdownInline { Type = MarkdownInlineType.Emphasis };
        node.Children.AddRange(children);
        return node;
    }

    /// <summary>创建粗斜体节点</summary>
    /// <param name="children">子行内元素</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateStrongEmphasis(IEnumerable<MarkdownInline> children)
    {
        var node = new MarkdownInline { Type = MarkdownInlineType.StrongEmphasis };
        node.Children.AddRange(children);
        return node;
    }

    /// <summary>创建行内代码节点</summary>
    /// <param name="code">代码文本</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateCode(String code) =>
        new() { Type = MarkdownInlineType.Code, Text = code };

    /// <summary>创建删除线节点</summary>
    /// <param name="children">子行内元素</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateStrikethrough(IEnumerable<MarkdownInline> children)
    {
        var node = new MarkdownInline { Type = MarkdownInlineType.Strikethrough };
        node.Children.AddRange(children);
        return node;
    }

    /// <summary>创建超链接节点</summary>
    /// <param name="href">目标 URL</param>
    /// <param name="title">标题（可空）</param>
    /// <param name="children">显示内容</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateLink(String href, String title, IEnumerable<MarkdownInline> children)
    {
        var node = new MarkdownInline { Type = MarkdownInlineType.Link, Href = href, Title = title ?? String.Empty };
        node.Children.AddRange(children);
        return node;
    }

    /// <summary>创建图片节点</summary>
    /// <param name="src">图片 URL</param>
    /// <param name="alt">替代文本</param>
    /// <param name="title">标题（可空）</param>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateImage(String src, String alt, String title) =>
        new() { Type = MarkdownInlineType.Image, Href = src, Alt = alt, Title = title ?? String.Empty };

    /// <summary>创建硬换行节点</summary>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateHardBreak() => new() { Type = MarkdownInlineType.HardBreak };

    /// <summary>创建软换行节点</summary>
    /// <returns>行内元素</returns>
    public static MarkdownInline CreateSoftBreak() => new() { Type = MarkdownInlineType.SoftBreak };
    #endregion

    #region 方法
    /// <summary>获取纯文本内容（递归展开所有子节点）</summary>
    /// <returns>纯文本字符串</returns>
    public String GetPlainText()
    {
        if (Type is MarkdownInlineType.Text or MarkdownInlineType.Code or MarkdownInlineType.RawHtml)
            return Text;
        if (Type == MarkdownInlineType.Image) return Alt;
        if (Type is MarkdownInlineType.HardBreak or MarkdownInlineType.SoftBreak) return " ";
        var sb = new System.Text.StringBuilder();
        foreach (var child in Children) sb.Append(child.GetPlainText());
        return sb.ToString();
    }

    /// <inheritdoc/>
    public override String ToString() => $"{Type}: {GetPlainText()}";
    #endregion
}
