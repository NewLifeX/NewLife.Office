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

    /// <summary>书签名称</summary>
    public String? BookmarkName { get; set; }

    /// <summary>段落背景色（16进制 RGB，如 "FF0000"）</summary>
    public String? BackgroundColor { get; set; }
    #endregion
}
