namespace NewLife.Office.Markdown;

/// <summary>脚注定义块</summary>
public sealed class FootnoteDefinitionBlock : MarkdownBlock
{
    #region 属性
    /// <summary>脚注标识符</summary>
    public String Id { get; set; }

    /// <summary>脚注内容</summary>
    public List<MarkdownInline> Definition { get; }
    #endregion

    #region 构造
    /// <summary>创建脚注定义块</summary>
    /// <param name="id">标识符</param>
    /// <param name="definition">定义内容</param>
    public FootnoteDefinitionBlock(String id, IEnumerable<MarkdownInline> definition)
    {
        Type = MarkdownBlockType.FootnoteDefinition;
        Id = id;
        Definition = [];
        Definition.AddRange(definition);
    }
    #endregion
}
