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
}
