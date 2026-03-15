using System.Text;

namespace NewLife.Office;

/// <summary>PDF 文档书签</summary>
public class PdfBookmark
{
    #region 属性
    /// <summary>书签标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>目标页面索引（0起始）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>子书签</summary>
    public List<PdfBookmark> Children { get; } = [];
    #endregion
}