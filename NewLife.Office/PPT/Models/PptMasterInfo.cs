namespace NewLife.Office;

/// <summary>PPT 幻灯片母版信息（S04-01）</summary>
public class PptMasterInfo
{
    #region 属性
    /// <summary>母版索引（0起始）</summary>
    public Int32 Index { get; set; }

    /// <summary>母版文件名（不含扩展名）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>母版背景色（16进制 RGB），null 表示未设置</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>关联版式 ID 列表</summary>
    public List<String> LayoutIds { get; } = [];

    /// <summary>关联主题名称</summary>
    public String ThemeRef { get; set; } = String.Empty;
    #endregion
}
