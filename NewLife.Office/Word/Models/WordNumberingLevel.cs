namespace NewLife.Office;

/// <summary>Word 编号（列表）层级定义</summary>
public class WordNumberingLevel
{
    #region 属性
    /// <summary>层级索引（0-8），0 为最顶层</summary>
    public Int32 Level { get; set; }

    /// <summary>编号格式：decimal / bullet / upperLetter / lowerLetter / upperRoman / lowerRoman / none</summary>
    public String Format { get; set; } = "decimal";

    /// <summary>格式文本，如 "%1." 表示"数字+"."，%1%2表示二级嵌套；bullet 时直接填符号</summary>
    public String? Text { get; set; }

    /// <summary>项目符号字符（Format = bullet 时使用，如 "•" "–" "✓"）</summary>
    public String? BulletChar { get; set; }

    /// <summary>符号字体名称（Symbol/Wingdings 等特殊符号字体）</summary>
    public String? BulletFontName { get; set; }

    /// <summary>起始编号（decimal 格式时有效），默认 1</summary>
    public Int32 StartAt { get; set; } = 1;

    /// <summary>左缩进（缇，twips）</summary>
    public Int32 Indent { get; set; } = 720;

    /// <summary>悬挂缩进（缇，twips，正文相对于编号符号的偏移）</summary>
    public Int32 HangingIndent { get; set; } = 360;

    /// <summary>编号文字格式（字体/大小/粗斜/颜色等）</summary>
    public WordRunProperties? RunProperties { get; set; }
    #endregion
}
