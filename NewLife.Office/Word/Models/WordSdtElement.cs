namespace NewLife.Office;

/// <summary>内容控件（SDT）元素，对应 OOXML <c>w:sdt</c></summary>
/// <remarks>
/// 结构化文档标签（Structured Document Tag），用于在 Word 文档中嵌入
/// 表单控件或受约束的内容区域。
/// 当前仅支持读取，写入通过 RawXml 透传兜底。
/// </remarks>
public class WordSdtElement
{
    #region 属性
    /// <summary>内容控件类型</summary>
    public WordSdtType SdtType { get; set; } = WordSdtType.PlainText;

    /// <summary>标签（<c>w:tag</c> 属性值），用于标识控件</summary>
    public String? Tag { get; set; }

    /// <summary>控件标题/别名（<c>w:alias</c> 属性值）</summary>
    public String? Alias { get; set; }

    /// <summary>内容文本</summary>
    public String? Content { get; set; }

    /// <summary>原始 XML（<c>w:sdt</c> 的 OuterXml），用于往返透传</summary>
    public String? RawXml { get; set; }
    #endregion
}
