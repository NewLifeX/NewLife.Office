namespace NewLife.Office.Word;

/// <summary>段落</summary>
public class Paragraph
{
    #region 属性
    /// <summary>原始样式标识符（如 "Heading2"、"2"、自定义样式名），用于精确往返保留</summary>
    /// <remarks>写入时优先使用此值；为 null 时使用 <see cref="Style"/> 枚举值</remarks>
    public String? StyleId { get; set; }

    /// <summary>段落样式（枚举，Normal/Heading1~6），由 StyleId 或新建文档时设置</summary>
    public ParagraphStyle Style { get; set; } = ParagraphStyle.Normal;

    /// <summary>文字段集合</summary>
    public List<Run> Runs { get; } = [];

    /// <summary>
    /// 文本框内容段落集合（w:drawing 内 wps:txbx/w:txbxContent 的段落，W22）。
    /// Reader 解析填充，FindText/ReplaceText 可覆盖；替换后自动同步外层段落 RawXml 保证保存生效。
    /// </summary>
    public List<Paragraph> TextBoxes { get; } = [];

    /// <summary>
    /// 原始 XML（文本框内层 <c>w:p</c> 原文，W22）。
    /// Reader 解析文本框时填充；ReplaceText 修改后清空并由 Writer 从模型重建，保证修改持久化。
    /// </summary>
    public String? RawXml { get; set; }

    /// <summary>对齐方式（left/center/right/both）</summary>
    public String? Alignment { get; set; }

    /// <summary>左缩进（twips）</summary>
    public Int32? IndentLeft { get; set; }

    /// <summary>右缩进（twips）</summary>
    public Int32? IndentRight { get; set; }

    /// <summary>首行缩进（twips，正值=缩进，负值=悬挂缩进）</summary>
    public Int32? FirstLineIndent { get; set; }

    /// <summary>段前间距（twips）</summary>
    public Int32? SpaceBefore { get; set; }

    /// <summary>段后间距（twips）</summary>
    public Int32? SpaceAfter { get; set; }

    /// <summary>行距（百分值，100=单倍, 150=1.5倍, 200=双倍）</summary>
    public Int32? LineSpacingPct { get; set; }

    /// <summary>是否分页符</summary>
    public Boolean IsPageBreak { get; set; }

    /// <summary>是否项目符号列表</summary>
    public Boolean IsBullet { get; set; }
    /// <summary>是否有序（编号）列表</summary>
    public Boolean IsOrderedList { get; set; }
    /// <summary>列表级别（0=一级, 1=二级...），配合 IsBullet 使用，默认 0</summary>
    public Int32 ListLevel { get; set; }

    /// <summary>
    /// 原始编号定义 ID（w:numPr/w:numId w:val，如 5）。
    /// Reader 从真实文档填充；Writer 非空时优先原样写回，保证引用透传的 numbering.xml 中的编号定义，
    /// 避免程序化重新映射为 1/2/3 导致列表编号错乱。
    /// </summary>
    public Int32? NumId { get; set; }

    /// <summary>
    /// 编号格式（numFmt：bullet / decimal / lowerLetter / upperRoman 等，对应段落所在层级的 numFmt）。
    /// Reader 解析 numbering.xml 后填充，用于程序化判断列表类型与展示。
    /// </summary>
    public String? ListFormat { get; set; }

    /// <summary>有序列表起始编号（仅 IsOrderedList=true 时有效），默认 1</summary>
    public Int32? ListStartOverride { get; set; }

    /// <summary>书签名称</summary>
    public String? BookmarkName { get; set; }

    /// <summary>段落背景色（16进制 RGB，如 "FF0000"）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>制表位集合，null 表示未设置（对应 OOXML w:tabs）</summary>
    public List<TabStop>? TabStops { get; set; }

    /// <summary>段落边框，null 表示无边框（对应 OOXML w:pBdr）</summary>
    public ParagraphBorders? Borders { get; set; }

    /// <summary>首字下沉行数（0 或 null 表示不启用首字下沉），对应 w:framePr w:dropCap="drop"</summary>
    public Int32? DropCapLines { get; set; }

    /// <summary>首字下沉字符数（默认 1），对应 w:framePr w:lines="N"</summary>
    public Int32? DropCapChars { get; set; }

    /// <summary>与下一段落保持同页（w:keepNext），防止标题孤立在页尾</summary>
    public Boolean KeepNext { get; set; }

    /// <summary>段落内各行保持同页（w:keepLines），防止段落跨页断裂</summary>
    public Boolean KeepLines { get; set; }

    /// <summary>孤行控制（w:widowControl），true=防止首行孤悬页尾/末行孤悬页首，默认 true</summary>
    public Boolean WidowControl { get; set; } = true;
    #endregion
}
