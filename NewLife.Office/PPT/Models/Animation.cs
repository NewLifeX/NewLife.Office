namespace NewLife.Office.Ppt;

/// <summary>PPT 元素动画类别</summary>
public enum AnimationCategory
{
    /// <summary>进入动画</summary>
    Entrance,

    /// <summary>强调动画</summary>
    Emphasis,

    /// <summary>退出动画</summary>
    Exit,

    /// <summary>动作路径动画</summary>
    MotionPath,
}

/// <summary>PPT 动画触发方式</summary>
public enum AnimationTrigger
{
    /// <summary>单击时触发</summary>
    OnClick,

    /// <summary>与前一动画同时播放</summary>
    WithPrevious,

    /// <summary>前一动画结束后自动触发</summary>
    AfterPrevious,
}

/// <summary>PPT 元素动画（进入/强调/退出/路径）</summary>
/// <remarks>
/// 描述幻灯片上某个元素的动画效果，包括效果名称、触发方式和时间参数。
/// 通过 <see cref="Slide.Animations"/> 集合关联到幻灯片。
/// <example>
/// <code>
/// var anim = new Animation
/// {
///     Category = AnimationCategory.Entrance,
///     Effect = "appear",
///     TargetType = "textBox",
///     TargetIndex = 0,
///     Trigger = AnimationTrigger.OnClick,
///     DurationMs = 500,
/// };
/// slide.Animations.Add(anim);
/// </code>
/// </example>
/// </remarks>
public class Animation
{
    #region 属性
    /// <summary>目标元素类型：textBox/shape/image/chart/table/group/connector</summary>
    public String TargetType { get; set; } = "textBox";

    /// <summary>目标元素在对应集合中的索引（0起始）</summary>
    public Int32 TargetIndex { get; set; }

    /// <summary>动画类别（进入/强调/退出/路径）</summary>
    public AnimationCategory Category { get; set; } = AnimationCategory.Entrance;

    /// <summary>
    /// 动画效果名称（对应 OOXML animRg/animScale 等 preset 名称），常用值：
    /// 进入：appear / fade / flyIn / floatUp / splitInHorizontal / wipe / zoom / grow
    /// 强调：pulse / spin / bold / colorPulse / transparency
    /// 退出：disappear / fadeOut / flyOut / zoom
    /// </summary>
    public String Effect { get; set; } = "appear";

    /// <summary>动画触发方式</summary>
    public AnimationTrigger Trigger { get; set; } = AnimationTrigger.OnClick;

    /// <summary>触发延迟（毫秒），0 表示无延迟</summary>
    public Int32 DelayMs { get; set; }

    /// <summary>动画持续时间（毫秒），默认 500ms</summary>
    public Int32 DurationMs { get; set; } = 500;

    /// <summary>动画在幻灯片内的顺序（从 0 开始）</summary>
    public Int32 Sequence { get; set; }

    /// <summary>动画排序顺序（同 Sequence，供 OrderBy 使用）</summary>
    public Int32 Order { get; set; }

    /// <summary>重复次数，"indefinite" 表示无限循环，null 表示 1 次</summary>
    public String? Repeat { get; set; }

    /// <summary>自动反转（播放后自动倒放），默认 false</summary>
    public Boolean AutoReverse { get; set; }
    #endregion
}
