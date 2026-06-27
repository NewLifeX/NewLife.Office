namespace NewLife.Office;

/// <summary>PDF 注释类型</summary>
public enum PdfAnnotationType
{
    /// <summary>超链接（外部 URL 或内部页面跳转）</summary>
    Link = 0,

    /// <summary>便签注释（弹出式文本框）</summary>
    Text,

    /// <summary>高亮标记</summary>
    Highlight,

    /// <summary>下划线标记</summary>
    Underline,

    /// <summary>删除线标记</summary>
    StrikeOut,

    /// <summary>自由文本（在页面上直接显示）</summary>
    FreeText,

    /// <summary>矩形注释框</summary>
    Square,

    /// <summary>椭圆注释框</summary>
    Circle,

    /// <summary>直线注释</summary>
    Line,

    /// <summary>图章注释（如"已批准"印章）</summary>
    Stamp,
}

/// <summary>PDF 页面注释</summary>
/// <remarks>
/// 对应 PDF 规范中的 Annotation 对象，可描述超链接、便签、高亮、图章等。
/// 通过 <see cref="PdfDocument.Annotations"/> 集合关联到文档。
/// <example>
/// <code>
/// // 创建超链接注释
/// var link = new PdfAnnotation
/// {
///     Type = PdfAnnotationType.Link,
///     PageIndex = 0,
///     X = 72, Y = 720, Width = 144, Height = 18,
///     Url = "https://newlifex.com",
/// };
///
/// // 创建便签注释
/// var note = new PdfAnnotation
/// {
///     Type = PdfAnnotationType.Text,
///     PageIndex = 2,
///     X = 400, Y = 600,
///     Contents = "此处数据需要核实",
///     Author = "张三",
/// };
/// doc.Annotations.Add(link);
/// doc.Annotations.Add(note);
/// </code>
/// </example>
/// </remarks>
public class PdfAnnotation
{
    #region 属性
    /// <summary>注释类型</summary>
    public PdfAnnotationType Type { get; set; } = PdfAnnotationType.Link;

    /// <summary>所在页面索引（0起始）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>注释区域左边界（PDF 坐标，从左下角量起，单位磅）</summary>
    public Single X { get; set; }

    /// <summary>注释区域下边界（PDF 坐标，从左下角量起，单位磅）</summary>
    public Single Y { get; set; }

    /// <summary>注释区域宽度（单位磅）</summary>
    public Single Width { get; set; }

    /// <summary>注释区域高度（单位磅）</summary>
    public Single Height { get; set; }

    /// <summary>目标 URL（Type=Link 时使用外部链接）</summary>
    public String? Url { get; set; }

    /// <summary>目标页面索引（Type=Link 时使用文档内跳转，-1 表示不跳转页面）</summary>
    public Int32 DestinationPage { get; set; } = -1;

    /// <summary>注释文本内容（便签/高亮/图章的注释文字）</summary>
    public String? Contents { get; set; }

    /// <summary>注释作者</summary>
    public String? Author { get; set; }

    /// <summary>注释主题</summary>
    public String? Subject { get; set; }

    /// <summary>注释颜色（高亮/边框颜色），null 表示使用默认</summary>
    public OfficeColor? Color { get; set; }

    /// <summary>便签是否默认展开显示（Type=Text 时有效）</summary>
    public Boolean Open { get; set; }
    #endregion
}
