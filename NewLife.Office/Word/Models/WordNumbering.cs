namespace NewLife.Office;

/// <summary>Word 编号/列表定义</summary>
/// <remarks>
/// 比透传 <see cref="WordDocument.NumberingXml"/> 更易于程序化创建多层列表。
/// 在 <see cref="WordDocument.Numbering"/> 中定义，通过 <see cref="WordParagraph.StyleId"/>
/// 与段落关联（如 "ListParagraph"）或通过编号 ID 引用。
/// <example>
/// <code>
/// var num = new WordNumbering { NumberingId = 1 };
/// num.Levels.Add(new WordNumberingLevel { Level = 0, Format = "decimal", Text = "%1.", Indent = 720 });
/// num.Levels.Add(new WordNumberingLevel { Level = 1, Format = "lowerLetter", Text = "%2)", Indent = 1440 });
/// doc.Numbering = num;
/// </code>
/// </example>
/// </remarks>
public class WordNumbering
{
    #region 属性
    /// <summary>编号定义 ID（文档内唯一）</summary>
    public Int32 NumberingId { get; set; } = 1;

    /// <summary>快捷编号格式（对应 Levels[0].Format），仅单层列表时使用</summary>
    public String Format { get; set; } = "decimal";

    /// <summary>快捷项目符号字符（对应 Levels[0].BulletChar），bullet 格式时使用</summary>
    public String? BulletChar { get; set; }

    /// <summary>多层编号级别定义（最多 9 层，索引 0-8）</summary>
    public List<WordNumberingLevel> Levels { get; set; } = [];
    #endregion
}
