namespace NewLife.Office.Markdown;

/// <summary>数学公式块（$$...$$ 围栏）</summary>
public sealed class MathBlock : MarkdownBlock
{
    #region 属性
    /// <summary>LaTeX 公式内容</summary>
    public String Content { get; set; }
    #endregion

    #region 构造
    /// <summary>创建数学公式块</summary>
    /// <param name="content">LaTeX 公式内容</param>
    public MathBlock(String content)
    {
        Type = MarkdownBlockType.MathBlock;
        Content = content;
    }
    #endregion
}
