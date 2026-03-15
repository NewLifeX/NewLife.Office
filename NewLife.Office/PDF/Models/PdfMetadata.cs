using System.Text;

namespace NewLife.Office;

/// <summary>PDF 元数据</summary>
public class PdfMetadata
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>创建时间字符串（PDF 格式 D:YYYYMMDDHHmmss）</summary>
    public String? CreationDate { get; set; }

    /// <summary>PDF 版本（如 1.4）</summary>
    public String? PdfVersion { get; set; }

    /// <summary>总页数</summary>
    public Int32 PageCount { get; set; }
    #endregion
}