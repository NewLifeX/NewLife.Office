namespace NewLife.Office;

/// <summary>Word 文档属性（由 docProps/core.xml 读取）</summary>
public class WordProperties
{
    #region 属性
    /// <summary>标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述</summary>
    public String? Description { get; set; }

    /// <summary>创建时间</summary>
    public DateTime? Created { get; set; }
    #endregion
}
