namespace NewLife.Office;

/// <summary>PPT 页眉页脚设置</summary>
/// <remarks>
/// 对应 OOXML pptx 中 <c>presentation.xml</c> 的 <c>p:hf</c> 元素，
/// 统一管理所有幻灯片的页脚文本、页码和日期显示。
/// 通过 <see cref="PptDocument.HeaderFooter"/> 属性关联到演示文稿。
/// <example>
/// <code>
/// doc.HeaderFooter = new PptHeaderFooter
/// {
///     ShowFooter = true,
///     FooterText = "NewLife 内部资料",
///     ShowPageNumber = true,
///     ShowDate = true,
///     DateAutomatic = true,
///     DateFormat = "yyyy/MM/dd",
/// };
/// </code>
/// </example>
/// </remarks>
public class PptHeaderFooter
{
    #region 属性
    /// <summary>是否在幻灯片上显示页眉（PPT 标准中页眉仅显示在讲义/备注页）</summary>
    public Boolean ShowHeader { get; set; }

    /// <summary>页眉文本</summary>
    public String? HeaderText { get; set; }

    /// <summary>是否在幻灯片上显示页脚</summary>
    public Boolean ShowFooter { get; set; }

    /// <summary>页脚文本</summary>
    public String? FooterText { get; set; }

    /// <summary>是否显示幻灯片编号</summary>
    public Boolean ShowPageNumber { get; set; }

    /// <summary>是否显示日期</summary>
    public Boolean ShowDate { get; set; }

    /// <summary>日期是否自动更新（true=使用当前日期，false=使用固定日期 FixedDate）</summary>
    public Boolean DateAutomatic { get; set; } = true;

    /// <summary>固定日期文本（DateAutomatic=false 时使用）</summary>
    public String? FixedDate { get; set; }

    /// <summary>自动日期格式字符串（如 "yyyy/MM/dd"），null 表示使用系统默认</summary>
    public String? DateFormat { get; set; }

    /// <summary>幻灯片编号起始值（默认 1）</summary>
    public Int32 FirstSlideNumber { get; set; } = 1;
    #endregion
}
