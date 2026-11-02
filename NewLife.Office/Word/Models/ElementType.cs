namespace NewLife.Office.Word;

/// <summary>文档元素类型</summary>
public enum ElementType
{
    /// <summary>段落</summary>
    Paragraph,
    /// <summary>表格</summary>
    Table,
    /// <summary>图片</summary>
    Image,
    /// <summary>内容控件（SDT）</summary>
    Sdt,
    /// <summary>自选图形（W12）</summary>
    Shape,
    /// <summary>图表（W13）</summary>
    Chart,
    /// <summary>
    /// 未建模的元素。Reader 遇到无法解析的 body 级子元素（域、批注锚点、扩展块等）时原样收入，
    /// Writer 直接输出其 <see cref="Element.RawXml"/>，保证「读入→修改→保存」不丢内容。
    /// </summary>
    Unknown,
}
