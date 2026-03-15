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

    /// <summary>超链接 URL，不为 null 时点击该片段跳转</summary>
    public String? HyperlinkUrl { get; set; }
    #endregion
}
