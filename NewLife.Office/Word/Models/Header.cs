namespace NewLife.Office.Word;

/// <summary>Word 页眉模型</summary>
/// <remarks>
/// 比 <see cref="Document.HeaderText"/> 更丰富，支持多段落、内联图片和格式化文本。
/// 在 <see cref="Document.Headers"/> 集合中使用，可定义三种类型（default/first/even）。
/// <example>
/// <code>
/// var header = new Header
/// {
///     Type = "default",
///     Elements = [new Element { Type = ElementType.Paragraph,
///         Paragraph = new Paragraph { Runs = { new Run { Text = "机密" } } } }],
/// };
/// doc.Headers.Add(header);
/// </code>
/// </example>
/// </remarks>
public class Header
{
    #region 属性
    /// <summary>页眉类型：default（普通页）/ first（首页）/ even（偶数页）</summary>
    public String Type { get; set; } = "default";

    /// <summary>页眉内容元素列表（段落/图片/表格），与文档正文结构相同</summary>
    public List<Element> Elements { get; set; } = [];

    /// <summary>原始页眉部件的关系 XML（<c>word/_rels/headerN.xml.rels</c>）</summary>
    /// <remarks>
    /// 从文件读取时保留，写回时随部件一并写出。不保留会导致页眉里的图片引用（<c>r:embed</c>）
    /// 没有对应关系，输出包不符合 OPC（Word 提示文档损坏）。程序化构造时为空即可。
    /// </remarks>
    public String? RelsXml { get; set; }
    #endregion
}
