namespace NewLife.Office.Word;

/// <summary>
/// Word 文档节（Section）模型，对应 OOXML 中由 <c>w:sectPr</c> 分隔的节。
/// 多节文档各节可拥有独立的页面尺寸/方向/边距/页眉页脚/分栏设置。
/// </summary>
/// <remarks>
/// 由 <see cref="WordReader.ReadDocument"/> 填充到 <see cref="Document.Sections"/>；
/// <see cref="Document.Elements"/> 仍保持全文档平铺视图（向后兼容）。
/// <para>节边界：正文中 <c>w:p/w:pPr/w:sectPr</c> 结束一节，body 末尾的 <c>w:sectPr</c> 是最后一节。</para>
/// <example>
/// <code>
/// // 第一页纵向 A4，第二页横向 A4
/// doc.Sections[0].PageSettings.Landscape = false;
/// doc.Sections[1].PageSettings.Landscape = true;
/// </code>
/// </example>
/// </remarks>
public class Section
{
    #region 属性
    /// <summary>本节内容元素列表（段落/表格/图片/内容控件）</summary>
    public List<Element> Elements { get; set; } = [];

    /// <summary>本节页面设置（尺寸/边距/方向/页眉页脚/分栏）</summary>
    public PageSettings PageSettings { get; set; } = new();

    /// <summary>原始 w:sectPr XML，非空时 Writer 直接使用（完整保真）</summary>
    public String? SectPrXml { get; set; }
    #endregion
}
