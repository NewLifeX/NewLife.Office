namespace NewLife.Office;

/// <summary>幻灯片切换动画</summary>
public class PptTransition
{
    #region 属性
    /// <summary>切换类型（fade/push/wipe/zoom/split/cut）</summary>
    public String Type { get; set; } = "fade";

    /// <summary>切换时长（毫秒）</summary>
    public Int32 DurationMs { get; set; } = 500;

    /// <summary>切换方向（l/r/u/d，部分类型使用）</summary>
    public String Direction { get; set; } = "l";

    /// <summary>是否单击时自动切换</summary>
    public Boolean AdvanceOnClick { get; set; } = true;
    #endregion
}
