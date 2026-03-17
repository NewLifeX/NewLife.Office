namespace NewLife.Office;

/// <summary>EPUB 电子书文档模型</summary>
public class EpubDocument
{
    #region 属性

    /// <summary>书名</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>作者</summary>
    public String Author { get; set; } = String.Empty;

    /// <summary>语言（BCP 47，如 zh-CN/en）</summary>
    public String Language { get; set; } = "zh-CN";

    /// <summary>出版商</summary>
    public String Publisher { get; set; } = String.Empty;

    /// <summary>描述/简介</summary>
    public String Description { get; set; } = String.Empty;

    /// <summary>发布日期</summary>
    public String PublishDate { get; set; } = String.Empty;

    /// <summary>ISBN 或其他唯一标识</summary>
    public String Identifier { get; set; } = String.Empty;

    /// <summary>封面图片数据（PNG/JPEG）</summary>
    public Byte[]? Cover { get; set; }

    /// <summary>封面图片 MIME 类型</summary>
    public String CoverMediaType { get; set; } = "image/jpeg";

    /// <summary>章节列表（按顺序）</summary>
    public List<EpubChapter> Chapters { get; set; } = [];

    /// <summary>自定义 CSS 样式表内容</summary>
    public String StyleSheet { get; set; } = String.Empty;

    #endregion
}

/// <summary>EPUB 章节</summary>
public class EpubChapter
{
    #region 属性

    /// <summary>章节标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>章节 HTML 内容（完整 XHTML 片段或正文部分）</summary>
    public String Content { get; set; } = String.Empty;

    /// <summary>文件名（不含路径，如 chapter01.xhtml）</summary>
    public String FileName { get; set; } = String.Empty;

    /// <summary>子章节（嵌套 TOC 支持）</summary>
    public List<EpubChapter> Children { get; set; } = [];

    #endregion
}
