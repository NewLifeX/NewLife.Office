namespace NewLife.Office;

/// <summary>Word 下划线样式常量（对应 OOXML w:u/@w:val）</summary>
/// <remarks>引用此常量可避免手写字符串，如 WordUnderlineStyles.Wave</remarks>
public static class WordUnderlineStyles
{
    /// <summary>单线（默认）</summary>
    public const String Single = "single";

    /// <summary>双线</summary>
    public const String Double = "double";

    /// <summary>粗线</summary>
    public const String Thick = "thick";

    /// <summary>点线</summary>
    public const String Dotted = "dotted";

    /// <summary>点划线</summary>
    public const String DotDash = "dotDash";

    /// <summary>点点划线</summary>
    public const String DotDotDash = "dotDotDash";

    /// <summary>短划线</summary>
    public const String Dash = "dash";

    /// <summary>波浪线</summary>
    public const String Wave = "wave";

    /// <summary>双波浪线</summary>
    public const String WavyDouble = "wavyDouble";

    /// <summary>仅文字下划线（跳过空白字符）</summary>
    public const String Words = "words";

    /// <summary>无下划线</summary>
    public const String None = "none";
}
