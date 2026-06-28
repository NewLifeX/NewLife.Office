namespace NewLife.Office;

/// <summary>段落</summary>
public class WordParagraph
{
    #region 属性
    /// <summary>原始样式标识符（如 "Heading2"、"2"、自定义样式名），用于精确往返保留</summary>
    /// <remarks>写入时优先使用此值；为 null 时使用 <see cref="Style"/> 枚举值</remarks>
    public String? StyleId { get; set; }

    /// <summary>段落样式（枚举，Normal/Heading1~6），由 StyleId 或新建文档时设置</summary>
    public WordParagraphStyle Style { get; set; } = WordParagraphStyle.Normal;

    /// <summary>文字段集合</summary>
    public List<WordRun> Runs { get; } = [];

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

    /// <summary>有序列表起始编号（仅 IsOrderedList=true 时有效），默认 1</summary>
    public Int32? ListStartOverride { get; set; }

    /// <summary>书签名称</summary>
    public String? BookmarkName { get; set; }

    /// <summary>段落背景色（16进制 RGB，如 "FF0000"）</summary>
    public String? BackgroundColor { get; set; }

    /// <summary>制表位集合，null 表示未设置（对应 OOXML w:tabs）</summary>
    public List<WordTabStop>? TabStops { get; set; }

    /// <summary>段落边框，null 表示无边框（对应 OOXML w:pBdr）</summary>
    public WordParagraphBorders? Borders { get; set; }

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
