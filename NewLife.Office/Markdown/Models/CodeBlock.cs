namespace NewLife.Office.Markdown;

/// <summary>代码块</summary>
public sealed class CodeBlock : MarkdownBlock
{
    /// <summary>代码语言标识符</summary>
    public String Language { get; set; } = String.Empty;

    /// <summary>代码原始文本</summary>
    public String RawText { get; set; } = String.Empty;

    /// <summary>创建代码块</summary>
    /// <param name="code">代码文本</param>
    /// <param name="language">语言标识（可空）</param>
    public CodeBlock(String code, String language = "")
    {
        Type = MarkdownBlockType.CodeBlock;
        RawText = code;
        Language = language ?? String.Empty;
    }

    /// <inheritdoc/>
    public override String GetPlainText() => RawText;
}
