namespace NewLife.Office;

/// <summary>PPT 文本段落（多段文本的最小单元，每段含独立格式和 Run 列表）</summary>
public class PptParagraph
{
    #region 属性
    /// <summary>富文本片段集合</summary>
    public List<PptTextRun> Runs { get; } = [];

    /// <summary>段落对齐（l/ctr/r），默认 l</summary>
    public String Alignment { get; set; } = "l";

    /// <summary>段落缩进级别（0-8，对应 lstStyle 的 lvl1pPr~lvl9pPr）</summary>
    public Int32 Level { get; set; }

    /// <summary>行间距——百分比模式（10万分为单位，如 100000=100%），0 表示不设置</summary>
    public Int32 LineSpacingPct { get; set; }

    /// <summary>行间距——精确磅值模式（1/100 pt），0 表示不设置</summary>
    public Int32 LineSpacingPts { get; set; }

    /// <summary>段前间距（pt）</summary>
    public Int32 SpaceBeforePt { get; set; }

    /// <summary>段后间距（pt）</summary>
    public Int32 SpaceAfterPt { get; set; }

    /// <summary>项目符号字符（如 "•"），null 或空表示不设置</summary>
    public String? BulletChar { get; set; }

    /// <summary>无项目符号（显式取消项目符号）</summary>
    public Boolean BulletNone { get; set; }
    #endregion
}
