namespace NewLife.Office;

/// <summary>文档元素联合类型</summary>
public class WordElement
{
    /// <summary>类型</summary>
    public WordElementType Type { get; set; }

    /// <summary>段落（Type=Paragraph 时有效）</summary>
    public WordParagraph? Paragraph { get; set; }

    /// <summary>表格行集合（Type=Table 时有效）</summary>
    public List<List<WordCell>>? TableRows { get; set; }

    /// <summary>首行是否表头（Type=Table 时有效）</summary>
    public Boolean TableFirstRowHeader { get; set; }

    /// <summary>表格样式（Type=Table 时有效）</summary>
    public WordTableStyle? TableStyle { get; set; }

    /// <summary>表格模型（Type=Table 时使用，比 TableRows 更丰富支持行级属性）</summary>
    /// <remarks>Table 与 TableRows 二选一，Writer 优先使用 Table。</remarks>
    public WordTable? Table { get; set; }

    /// <summary>图片（Type=Image 时有效）</summary>
    public WordImage? Image { get; set; }

    /// <summary>
    /// 元素原始 XML，由 Reader 将对应 <c>w:p</c>/<c>w:tbl</c> 的 OuterXml 存入。
    /// Writer 在此不为空时<strong>优先直接写入</strong>，完整保留所有内联格式
    /// （字体/大小/颜色/间距/段落边框/高亮/删除线等未建模属性）。
    /// 程序化创建的元素（AppendParagraph 等）该属性为 null，使用模型生成 XML。
    /// </summary>
    public String? RawXml { get; set; }
}
