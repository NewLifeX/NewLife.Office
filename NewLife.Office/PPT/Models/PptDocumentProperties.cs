namespace NewLife.Office;

/// <summary>PPT 文档属性</summary>
public class PptDocumentProperties
{
    #region 属性
    /// <summary>文档标题</summary>
    public String? Title { get; set; }

    /// <summary>作者</summary>
    public String? Author { get; set; }

    /// <summary>主题</summary>
    public String? Subject { get; set; }

    /// <summary>描述/备注</summary>
    public String? Description { get; set; }

    /// <summary>修改保护密码（null 表示不保护）</summary>
    public String? Password { get; set; }
    #endregion
}
