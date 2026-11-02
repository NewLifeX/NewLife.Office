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

    /// <summary>
    /// 节起始类型（<c>w:sectPr/w:type w:val</c>）：nextPage（默认）/ continuous / evenPage / oddPage。
    /// <remarks>OOXML 的 <c>ST_SectionMark</c> 枚举不含 <c>newPage</c>，写入时会自动归一为 <c>nextPage</c>。</remarks>
    /// 仅程序化创建节时生效；<see cref="SectPrXml"/> 非空时以原始 XML 为准。
    /// </summary>
    public String? Type { get; set; }

    /// <summary>页眉是否链接到上一节（默认 true = 继承上一节页眉）</summary>
    /// <remarks>
    /// 置为 false 表示本节不继承上一节，改用 <see cref="Headers"/> 中的内容；
    /// <see cref="Headers"/> 为空时本节将没有页眉（Word 中“取消链接到前一节”后删除内容的效果）。
    /// 第一节没有上一节，该标志无实际作用。
    /// </remarks>
    public Boolean LinkHeadersToPrevious { get; set; } = true;

    /// <summary>页脚是否链接到上一节（默认 true = 继承上一节页脚）</summary>
    /// <remarks>置为 false 表示本节使用 <see cref="Footers"/> 中的独立页脚；为空时本节没有页脚。</remarks>
    public Boolean LinkFootersToPrevious { get; set; } = true;

    /// <summary>本节独立页眉（<see cref="LinkHeadersToPrevious"/> 为 false 时生效），可按 default/first/even 定义多个</summary>
    public List<Header> Headers { get; set; } = [];

    /// <summary>本节独立页脚（<see cref="LinkFootersToPrevious"/> 为 false 时生效），可按 default/first/even 定义多个</summary>
    public List<Footer> Footers { get; set; } = [];
    #endregion
}
