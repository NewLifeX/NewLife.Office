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

    /// <summary>图片（Type=Image 时有效）</summary>
    public WordImageElement? Image { get; set; }
}
