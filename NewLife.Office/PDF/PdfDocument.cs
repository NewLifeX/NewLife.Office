namespace NewLife.Office;

/// <summary>PDF 文档数据模型，对标 WordDocument / PptDocument / ExcelDocument</summary>
/// <remarks>
/// 纯数据容器（POCO），不含读写逻辑。
/// 通过 <see cref="PdfReader.ReadDocument"/> 读取，通过 <see cref="PdfWriter.Save(String,PdfDocument)"/> 写入。
/// 文档级工具操作（合并、拆分、水印等）请使用 <see cref="PdfHelper"/>。
/// </remarks>
public class PdfDocument
{
    #region 属性
    /// <summary>页面集合</summary>
    public List<PdfPage> Pages { get; set; } = [];

    /// <summary>文档元数据（标题、作者、主题等）</summary>
    public PdfMetadata Metadata { get; set; } = new();

    /// <summary>书签/大纲列表</summary>
    public List<PdfBookmark> Bookmarks { get; set; } = [];

    /// <summary>页眉文本（null 表示不显示）</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本（null 表示不显示）</summary>
    public String? FooterText { get; set; }

    /// <summary>是否在页脚显示页码</summary>
    public Boolean ShowPageNumbers { get; set; }

    /// <summary>用户密码（文档打开密码，null 表示不加密）</summary>
    public String? UserPassword { get; set; }

    /// <summary>所有者密码（权限管理密码，null 表示不限制）</summary>
    public String? OwnerPassword { get; set; }

    /// <summary>权限标志位（-1 表示全部允许）</summary>
    public Int32 Permissions { get; set; } = -1;
    #endregion
}