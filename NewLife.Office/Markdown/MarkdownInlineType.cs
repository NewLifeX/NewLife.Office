namespace NewLife.Office.Markdown;

/// <summary>Markdown 行内元素类型</summary>
public enum MarkdownInlineType
{
    /// <summary>普通文本</summary>
    Text,
    /// <summary>粗体（**text** 或 __text__）</summary>
    Strong,
    /// <summary>斜体（*text* 或 _text_）</summary>
    Emphasis,
    /// <summary>粗斜体（***text***）</summary>
    StrongEmphasis,
    /// <summary>行内代码（`code`）</summary>
    Code,
    /// <summary>删除线（~~text~~，GFM）</summary>
    Strikethrough,
    /// <summary>超链接 [text](href "title")</summary>
    Link,
    /// <summary>图片 ![alt](src "title")</summary>
    Image,
    /// <summary>硬换行（两个空格 + 换行）</summary>
    HardBreak,
    /// <summary>软换行（段落内单个换行）</summary>
    SoftBreak,
    /// <summary>原始 HTML 行内片段</summary>
    RawHtml,
}
