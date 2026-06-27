namespace NewLife.Office;

/// <summary>PPT 富文本片段（S10-01）</summary>
/// <remarks>
/// 支持每个片段独立设置字体、粗体、斜体、颜色、超链接。
/// 将多个 <see cref="PptTextRun"/> 添加到 <see cref="PptTextBox.Runs"/> 即可实现富文本效果。
/// </remarks>
public class PptTextRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字号（磅），0 表示继承文本框默认字号</summary>
    public Int32 FontSize { get; set; }

    /// <summary>粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>文字颜色（16进制 RGB），null 表示继承文本框设置</summary>
    public String? FontColor { get; set; }

    /// <summary>拉丁/西文字体名称（如"Montserrat Black"），null 表示继承</summary>
    public String? LatinFontName { get; set; }

    /// <summary>东亚/中文字体名称（如"微软雅黑"），null 表示继承</summary>
    public String? EastAsianFontName { get; set; }

    /// <summary>复杂脚本字体名称（如阿拉伯/泰文），null 表示继承</summary>
    public String? ComplexScriptFontName { get; set; }

    /// <summary>符号字体名称，null 表示继承</summary>
    public String? SymbolFontName { get; set; }

    /// <summary>字体名称（如"微软雅黑"），null 表示继承。兼容属性：getter 返回 EastAsianFontName ?? LatinFontName</summary>
    public String? FontName
    {
        get => EastAsianFontName ?? LatinFontName;
        set => LatinFontName = EastAsianFontName = value;
    }

    /// <summary>下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>渐变颜色停靠点（如 ["FF0000","0000FF"]），null 表示纯色填充。深度不超过 6 个停靠点。</summary>
    public String[]? GradFillColors { get; set; }

    /// <summary>渐变方向（角度，单位 1/60000°，如 5400000 = 90°）</summary>
    public Int32 GradAngle { get; set; }

    /// <summary>超链接 URL，不为 null 时点击该片段跳转</summary>
    public String? HyperlinkUrl { get; set; }
    #endregion
}
