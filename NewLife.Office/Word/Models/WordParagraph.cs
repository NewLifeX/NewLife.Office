namespace NewLife.Office;

/// <summary>段落</summary>
public class WordParagraph
{
    #region 属性
    /// <summary>段落样式</summary>
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

    /// <summary>行距（percent  100，如 100=单倍, 150=1.5倍, 200=双倍）</summary>
    public Int32? LineSpacingPct { get; set; }

    /// <summary>是否分页符</summary>
    public Boolean IsPageBreak { get; set; }

    /// <summary>是否项目符号列表</summary>
    public Boolean IsBullet { get; set; }

    /// <summary>书签名称（非空时在段落首尾添加书签）</summary>
    public String? BookmarkName { get; set; }
    #endregion
}
