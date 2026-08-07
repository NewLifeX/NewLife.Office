namespace NewLife.Office.Word;

/// <summary>
/// 脚注/尾注模型（W22）。
/// Reader 从 footnotes.xml/endnotes.xml 解析；纯数据容器，用于文本提取与程序化访问。
/// </summary>
/// <remarks>
/// Word 文档中 <c>w:footnoteReference w:id="N"</c> 引用脚注，脚注内容定义在
/// <c>word/footnotes.xml</c>（尾注在 <c>word/endnotes.xml</c>）。本模型按 id 提取内容，
/// 不包含 Word 内置的分隔符/续接标记脚注（w:type="separator"/"continuationSeparator"）。
/// </remarks>
public class Footnote
{
    /// <summary>脚注 ID（与文档中 w:footnoteReference w:id 对应；特殊脚注为负数或 0）</summary>
    public Int32 Id { get; set; }

    /// <summary>脚注类型（null=普通脚注；separator/continuationSeparator/continuationNotice 为内置标记）</summary>
    public String? Type { get; set; }

    /// <summary>脚注纯文本（各段落以换行分隔）</summary>
    public String Text { get; set; } = "";

    /// <summary>脚注内容段落（含格式 Runs，便于程序化访问与后续 Find/Replace 扩展）</summary>
    public List<Paragraph> Paragraphs { get; set; } = [];
}
