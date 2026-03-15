namespace NewLife.Office;

/// <summary>PPT 幻灯片版式信息（S04-02）</summary>
public class PptLayoutInfo
{
    #region 属性
    /// <summary>版式索引（0起始）</summary>
    public Int32 Index { get; set; }

    /// <summary>版式文件名（不含扩展名）</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>版式类型（如 blank、title、twoContent 等）</summary>
    public String LayoutType { get; set; } = String.Empty;

    /// <summary>版式显示名称</summary>
    public String DisplayName { get; set; } = String.Empty;
    #endregion
}
