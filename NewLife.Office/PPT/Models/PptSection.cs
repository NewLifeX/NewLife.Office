namespace NewLife.Office;

/// <summary>PPT 节（Section），用于按节组织幻灯片</summary>
public class PptSection
{
    #region 属性
    /// <summary>节名称</summary>
    public String Name { get; set; } = "默认节";

    /// <summary>该节包含的幻灯片索引列表（0基）</summary>
    public List<Int32> SlideIndices { get; set; } = [];
    #endregion
}
