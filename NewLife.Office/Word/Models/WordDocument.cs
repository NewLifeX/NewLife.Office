namespace NewLife.Office;

/// <summary>Word 文档完整模型，承载读写往返所需全部信息</summary>
/// <remarks>
/// 纯数据容器（POCO），不含读写逻辑。
/// 由 <see cref="WordReader.ReadDocument"/> 产出，
/// 由 <see cref="WordWriter.Save(String, WordDocument)"/> 消费。
/// <para>典型用法：</para>
/// <code>
/// // 读取
/// using var reader = new WordReader("source.docx");
/// var doc = reader.ReadDocument();
///
/// // 修改
/// doc.DocumentProperties.Title = "新标题";
/// doc.Elements.Insert(0, new WordElement { Type = WordElementType.Paragraph, Paragraph = new WordParagraph { Style = WordParagraphStyle.Heading1, Runs = { new WordRun { Text = "新增标题" } } } });
///
/// // 写入
/// using var writer = new WordWriter();
/// writer.Save("output.docx", doc);
/// </code>
/// </remarks>
public class WordDocument
{
    #region 属性
    /// <summary>文档元素列表（段落/表格/图片），按文档顺序排列</summary>
    public List<WordElement> Elements { get; set; } = [];

    /// <summary>页面设置（尺寸/边距/方向/页眉页脚）</summary>
    public WordPageSettings PageSettings { get; set; } = new();

    /// <summary>文档属性（标题/作者/主题/描述）</summary>
    public WordDocumentProperties DocumentProperties { get; set; } = new();

    /// <summary>图片数据字典，key 为关系ID（如 rId4），value 为(扩展名, 字节数据)</summary>
    public Dictionary<String, (String Extension, Byte[] Data)> Images { get; set; } = [];

    /// <summary>超链接关系列表（关系ID, URL）</summary>
    public List<(String RelId, String Url)> Hyperlinks { get; set; } = [];

    /// <summary>页眉文本（简单纯文本页眉，与 Headers 二选一）</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本（简单纯文本页脚，与 Footers 二选一）</summary>
    public String? FooterText { get; set; }

    /// <summary>富文本页眉列表（default/first/even 三种类型），Writer 优先使用，为空时回退到 HeaderText</summary>
    public List<WordHeader> Headers { get; set; } = [];

    /// <summary>富文本页脚列表（default/first/even 三种类型），Writer 优先使用，为空时回退到 FooterText</summary>
    public List<WordFooter> Footers { get; set; } = [];

    /// <summary>文档批注列表（审阅评论）</summary>
    public List<WordComment> Comments { get; set; } = [];

    /// <summary>内容控件列表（SDT）</summary>
    public List<WordSdtElement> SdtElements { get; set; } = [];

    /// <summary>自定义 XML 部件（customXml/item*.xml），Key 为部件名（如 item1.xml），Value 为 XML 原始字节</summary>
    public Dictionary<String, Byte[]> CustomXmlParts { get; set; } = [];

    /// <summary>编号/列表定义（程序化创建列表时使用，与 NumberingXml 二选一）</summary>
    public WordNumbering? Numbering { get; set; }

    /// <summary>是否启用只读保护</summary>
    public Boolean ProtectionReadOnly { get; set; }

    /// <summary>原始 styles.xml 内容。非空时 Writer 直接使用而不再生成默认样式，确保字体/大小/颜色/间距与源文件一致</summary>
    public String? StylesXml { get; set; }

    /// <summary>原始 numbering.xml 内容。非空时 Writer 直接使用，保留列表编号/项目符号定义</summary>
    public String? NumberingXml { get; set; }

    /// <summary>原始 settings.xml 内容。非空时 Writer 直接使用，保留兼容性设置</summary>
    public String? SettingsXml { get; set; }

    /// <summary>
    /// 原样透传的其他 ZIP 部件（主题/字体表/脚注/尾注/页眉页脚/媒体等）。
    /// Key 为 ZIP 入口路径（如 "word/theme/theme1.xml"），Value 为原始字节。
    /// Reader 填充，Writer 原样写回，确保未解析内容不丢失。
    /// </summary>
    public Dictionary<String, Byte[]> OtherParts { get; set; } = [];

    /// <summary>原始 w:sectPr XML，保留精确的页面设置（含页眉页脚引用）</summary>
    public String? SectPrXml { get; set; }

    /// <summary>
    /// 原始 word/document.xml 全文。非空时 Writer 直接写入而不重建，
    /// 确保所有复杂格式（渐变/修订记录/批注/复杂段落属性等）完全保留。
    /// 设置此属性后，对 <see cref="Elements"/> 的修改将被忽略。
    /// 若要通过模型修改内容，需先将此属性置 null。
    /// </summary>
    public String? DocumentXml { get; set; }

    /// <summary>
    /// word/document.xml 根元素的命名空间声明字符串。
    /// Writer 重建 document.xml 时使用，保留源文件对 mc:/wps:/v: 等扩展命名空间的寄存性。
    /// </summary>
    public String? DocumentXmlNsDecls { get; set; }

    /// <summary>文档变量（settings.xml w:docVars），Key=变量名 Value=变量值，用于企业模板元数据驱动</summary>
    public Dictionary<String, String> DocumentVariables { get; set; } = [];
    #endregion
}
