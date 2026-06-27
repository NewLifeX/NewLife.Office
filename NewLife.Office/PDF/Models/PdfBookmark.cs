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

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>书签颜色（16 进制 RGB，如 "FF0000"），null 表示默认黑色</summary>
    public String? Color { get; set; }

    /// <summary>书签初始展开状态，默认 true</summary>
    public Boolean Expanded { get; set; } = true;
    #endregion
}