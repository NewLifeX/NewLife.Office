namespace NewLife.Office;

/// <summary>单元格富文本中的一段文字（对应 OOXML &lt;r&gt; 元素）</summary>
/// <remarks>
/// 多个 RichTextRun 组合成一个单元格的富文本内容，每段可有不同的字体/颜色/大小/粗斜体。
/// 通过 <see cref="ExcelCellStyle.RichTextRuns"/> 设置，Writer 生成 inline string &lt;is&gt;&lt;r&gt; 结构。
/// <example>
/// <code>
/// var style = new ExcelCellStyle
/// {
///     RichTextRuns = new List&lt;RichTextRun&gt;
///     {
///         new RichTextRun { Text = "重要提示：", Bold = true, Color = "FF0000" },
///         new RichTextRun { Text = "请认真填写表格", FontSize = 11 },
///     }
/// };
/// writer.WriteRow(null, new Object?[] { null }, style);
/// </code>
/// </example>
/// </remarks>
public class RichTextRun
{
    #region 属性
    /// <summary>文本内容</summary>
    public String Text { get; set; } = String.Empty;

    /// <summary>字体名称，null 表示继承</summary>
    public String? FontName { get; set; }

    /// <summary>字体大小（磅），0 表示继承</summary>
    public Double FontSize { get; set; }

    /// <summary>是否粗体</summary>
    public Boolean Bold { get; set; }

    /// <summary>是否斜体</summary>
    public Boolean Italic { get; set; }

    /// <summary>是否下划线</summary>
    public Boolean Underline { get; set; }

    /// <summary>是否删除线</summary>
    public Boolean Strike { get; set; }

    /// <summary>字体颜色（RGB六位十六进制，如 "FF0000"），null 表示继承</summary>
    public String? Color { get; set; }
    #endregion
}
