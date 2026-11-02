namespace NewLife.Office.Word;

/// <summary>Word 页脚模型</summary>
/// <remarks>
/// 比 <see cref="Document.FooterText"/> 更丰富，支持多段落、内联图片和格式化文本。
/// 在 <see cref="Document.Footers"/> 集合中使用，可定义三种类型（default/first/even）。
/// <example>
/// <code>
/// var footer = new Footer
/// {
///     Type = "default",
///     Elements = [new Element { Type = ElementType.Paragraph,
///         Paragraph = new Paragraph { Runs = { new Run { Text = "第 {PAGE} 页" } } } }],
/// };
/// doc.Footers.Add(footer);
/// </code>
/// </example>
/// </remarks>
public class Footer
{
    #region 属性
    /// <summary>页脚类型：default（普通页）/ first（首页）/ even（偶数页）</summary>
    public String Type { get; set; } = "default";

    /// <summary>页脚内容元素列表（段落/图片/表格），与文档正文结构相同</summary>
    public List<Element> Elements { get; set; } = [];

    /// <summary>原始页脚部件的关系 XML（<c>word/_rels/footerN.xml.rels</c>）</summary>
    /// <remarks>
    /// 从文件读取时保留，写回时随部件一并写出。不保留会导致页脚里的图片引用（<c>r:embed</c>）
    /// 没有对应关系，输出包不符合 OPC（Word 提示文档损坏）。程序化构造时为空即可。
    /// </remarks>
    public String? RelsXml { get; set; }
    #endregion
}
