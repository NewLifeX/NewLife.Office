using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Security;
using System.Text;
using System.Xml;
using NewLife.Office.Excel;

namespace NewLife.Office.Word;

/// <summary>Word docx 写入器</summary>
/// <remarks>
/// 直接操作 Open XML（ZIP+XML）生成 .docx 文件。
/// 支持段落/标题/表格/图片/超链接/列表/页面设置等核心功能。
/// </remarks>
public class WordWriter : IDisposable
{
    #region 属性
    /// <summary>文本编码</summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;

    /// <summary>页面设置</summary>
    public PageSettings PageSettings { get; set; } = new();

    /// <summary>文档属性</summary>
    public DocumentProperties DocumentProperties { get; set; } = new();
    #endregion

    #region 私有字段
    private readonly List<Element> _elements = [];
    private readonly List<(String RelId, String Url)> _hyperlinkRels = [];
    private readonly List<(String RelId, String Ext, Byte[] Data)> _imageRels = [];
    private Int32 _relCounter = 1;
    private Int32 _imgCounter = 1;
    private Int32 _bookmarkId = 1;
    private readonly Dictionary<Int32, Int32> _orderedStartOverrides = []; // level → startValue

    // 原始 XML 透传（非空时覆盖生成默认）
    private String? _stylesXml;
    private String? _numberingXml;
    private Numbering? _numbering;    // 程序化编号定义（与 NumberingXml 二选一）
    private String? _settingsXml;
    private String? _sectPrXml;           // sectPr 原始 XML
    private String? _documentXmlNsDecls;  // document.xml 根元素命名空间声明
    private String? _documentXml;         // word/document.xml 全文（非空时直接写入，跳过重建）

    // 原样透传的所有 ZIP 部件（除 word/document.xml 外）
    private Dictionary<String, Byte[]> _otherParts = [];

    // 富文本页眉/页脚（default/first/even），由 Document.Headers/Footers 导入，非空时覆盖简单文本页眉页脚
    private List<Header> _headers = [];
    private List<Footer> _footers = [];

    // 多节模型（程序化创建时由 Document.Sections 导入，Count>1 时按节输出）
    private List<Section> _sections = [];

    // 页眉页脚重建时保存的原始 [Content_Types].xml / document.xml.rels（用于合并新引用）
    private String? _mergeContentTypes;
    private String? _mergeDocumentRels;

    /// <summary>是否重建页眉页脚（仅用户显式重建 DocumentXml=null 且设置了富文本页眉页脚时）</summary>
    private Boolean _rebuildHeaderFooter;

    /// <summary>是否启用只读保护</summary>
    public Boolean ProtectionReadOnly { get; set; }

    private Dictionary<String, String> _documentVariables = [];

    // 程序化创建的脚注/尾注（非空时生成 footnotes.xml/endnotes.xml 部件，W22）
    private List<Footnote> _footnotes = [];
    private List<Footnote> _endnotes = [];
    private Int32 _nextNoteId = 1;
    #endregion

    #region 构造
    /// <summary>实例化写入器</summary>
    public WordWriter() { }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 段落方法
    /// <summary>追加段落（Paragraph 对象）</summary>
    /// <param name="para">段落对象</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendParagraph(Paragraph para)
    {
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加普通段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="style">段落样式</param>
    /// <returns>段落对象（可进一步设置间距/缩进等属性）</returns>
    public Paragraph AppendParagraph(String text, ParagraphStyle style = ParagraphStyle.Normal)
    {
        var para = new Paragraph { Style = style };
        para.Runs.Add(new Run { Text = text });
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加带格式的段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="style">段落样式</param>
    /// <param name="runProps">文字格式</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendParagraph(String text, ParagraphStyle style, RunProperties runProps)
    {
        var para = new Paragraph { Style = style };
        para.Runs.Add(new Run { Text = text, Properties = runProps });
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加标题</summary>
    /// <param name="text">标题文本</param>
    /// <param name="level">标题级别（1-6）</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendHeading(String text, Int32 level = 1)
    {
        if (level < 1) level = 1;
        if (level > 6) level = 6;
        return AppendParagraph(text, (ParagraphStyle)level);
    }

    /// <summary>追加多格式 Run 的段落</summary>
    /// <param name="runs">Run 集合</param>
    /// <param name="style">段落样式</param>
    /// <param name="alignment">对齐（left/center/right/both）</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendFormattedParagraph(IEnumerable<Run> runs, ParagraphStyle style = ParagraphStyle.Normal, String? alignment = null)
    {
        var para = new Paragraph { Style = style, Alignment = alignment };
        para.Runs.AddRange(runs);
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加超链接段落</summary>
    /// <param name="displayText">显示文本</param>
    /// <param name="url">目标 URL</param>
    /// <param name="runProps">可选文字格式</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendHyperlink(String displayText, String url, RunProperties? runProps = null)
    {
        var relId = $"rHyp{_relCounter++}";
        _hyperlinkRels.Add((relId, url));
        var para = new Paragraph();
        para.Runs.Add(new Run { Text = displayText, Properties = runProps, HyperlinkRelId = relId });
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加带书签的段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="bookmarkName">书签名称</param>
    /// <param name="style">段落样式</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendBookmarkedParagraph(String text, String bookmarkName, ParagraphStyle style = ParagraphStyle.Normal)
    {
        var para = AppendParagraph(text, style);
        para.BookmarkName = bookmarkName;
        return para;
    }

    /// <summary>追加交叉引用域（引用书签，显示为页码或文本）</summary>
    /// <param name="bookmarkName">被引用的书签名称</param>
    /// <param name="displayText">显示文本（通常为页码数字，如 "1"）</param>
    public void AppendCrossRef(String bookmarkName, String displayText = "1")
    {
        var xml = "<w:p>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + $"<w:r><w:instrText xml:space=\"preserve\"> REF {Esc(bookmarkName)} \\h </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + $"<w:r><w:t>{Esc(displayText)}</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "</w:p>";
        _elements.Add(new Element { Type = ElementType.Paragraph, RawXml = xml });
    }

    /// <summary>追加 OMML 数学公式（W13，独立公式块）</summary>
    /// <param name="formula">公式对象（分数/上下标/根式/积分等）</param>
    /// <returns>公式对象</returns>
    /// <exception cref="ArgumentNullException">formula 为空</exception>
    public OmmlFormula AppendFormula(OmmlFormula formula)
    {
        if (formula == null) throw new ArgumentNullException(nameof(formula));
        var xml = "<w:p><m:oMathPara><m:oMath>" + formula.Xml + "</m:oMath></m:oMathPara></w:p>";
        _elements.Add(new Element { Type = ElementType.Paragraph, RawXml = xml });
        return formula;
    }
    /// <summary>追加邮件合并域 MERGEFIELD 段落</summary>
    /// <param name="fieldName">合并域名（如 "FirstName"、"Company"）</param>
    /// <remarks>
    /// 生成标准 MERGEFIELD 域代码，Word 打开后可执行邮件合并填充数据源。
    /// 域显示文本使用 «FieldName» 占位符格式。
    /// </remarks>
    public void AppendMergeField(String fieldName)
    {
        var xml = "<w:p>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + $"<w:r><w:instrText xml:space=\"preserve\"> MERGEFIELD {Esc(fieldName)} </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + $"<w:r><w:t>«{Esc(fieldName)}»</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "</w:p>";
        _elements.Add(new Element { Type = ElementType.Paragraph, RawXml = xml });
    }

    /// <summary>追加带脚注的段落（W22：程序化创建脚注，对标 Word/Aspose）</summary>
    /// <param name="paragraphText">段落正文文本</param>
    /// <param name="footnoteText">脚注内容（显示在页面底部）</param>
    /// <returns>段落对象</returns>
    /// <remarks>
    /// 自动分配脚注 ID 并注册到 <c>word/footnotes.xml</c>（含内置分隔符标记），
    /// 段落末尾生成上标脚注引用 <c>w:footnoteReference</c>。
    /// </remarks>
    public Paragraph AppendFootnote(String paragraphText, String footnoteText)
    {
        if (footnoteText == null) throw new ArgumentNullException(nameof(footnoteText));
        var id = _nextNoteId++;
        _footnotes.Add(new Footnote
        {
            Id = id,
            Text = footnoteText,
            Paragraphs = { new Paragraph { Runs = { new Run { Text = footnoteText } } } },
        });

        var para = new Paragraph();
        if (!String.IsNullOrEmpty(paragraphText))
            para.Runs.Add(new Run { Text = paragraphText });
        var xml = "<w:p>"
            + (String.IsNullOrEmpty(paragraphText) ? "" : $"<w:r><w:t xml:space=\"preserve\">{Esc(paragraphText)}</w:t></w:r>")
            + $"<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"{id}\"/></w:r>"
            + "</w:p>";
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para, RawXml = xml });
        return para;
    }

    /// <summary>追加带尾注的段落（W22：程序化创建尾注，对标 Word/Aspose）</summary>
    /// <param name="paragraphText">段落正文文本</param>
    /// <param name="endnoteText">尾注内容（显示在文档末尾）</param>
    /// <returns>段落对象</returns>
    public Paragraph AppendEndnote(String paragraphText, String endnoteText)
    {
        if (endnoteText == null) throw new ArgumentNullException(nameof(endnoteText));
        var id = _nextNoteId++;
        _endnotes.Add(new Footnote
        {
            Id = id,
            Text = endnoteText,
            Paragraphs = { new Paragraph { Runs = { new Run { Text = endnoteText } } } },
        });

        var para = new Paragraph();
        if (!String.IsNullOrEmpty(paragraphText))
            para.Runs.Add(new Run { Text = paragraphText });
        var xml = "<w:p>"
            + (String.IsNullOrEmpty(paragraphText) ? "" : $"<w:r><w:t xml:space=\"preserve\">{Esc(paragraphText)}</w:t></w:r>")
            + $"<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:endnoteReference w:id=\"{id}\"/></w:r>"
            + "</w:p>";
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para, RawXml = xml });
        return para;
    }

    /// <summary>追加分页符</summary>
    public void AppendPageBreak()
    {
        var para = new Paragraph { IsPageBreak = true };
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
    }

    /// <summary>追加内容控件（SDT）</summary>
    /// <param name="sdt">SDT 内容控件元素</param>
    public void AppendSdt(SdtElement sdt)
    {
        _elements.Add(new Element { Type = ElementType.Sdt, Sdt = sdt });
    }

    /// <summary>追加纯文本内容控件</summary>
    /// <param name="content">控件内容文本</param>
    /// <param name="tag">标签（可选，用于标识控件）</param>
    /// <param name="alias">别名（可选，用于展示名称）</param>
    public void AppendPlainTextSdt(String content, String? tag = null, String? alias = null)
    {
        AppendSdt(new SdtElement
        {
            SdtType = SdtType.PlainText,
            Content = content,
            Tag = tag,
            Alias = alias
        });
    }

    /// <summary>追加日期选择器内容控件</summary>
    /// <param name="dateText">日期显示文本</param>
    /// <param name="dateFormat">日期格式（如 yyyy-MM-dd）</param>
    /// <param name="tag">标签</param>
    public void AppendDateSdt(String dateText, String dateFormat = "yyyy-MM-dd", String? tag = null)
    {
        AppendSdt(new SdtElement
        {
            SdtType = SdtType.Date,
            Content = dateText,
            DateFormat = dateFormat,
            Tag = tag
        });
    }

    /// <summary>追加下拉列表内容控件</summary>
    /// <param name="selectedText">选中项文本</param>
    /// <param name="items">下拉列表项</param>
    /// <param name="tag">标签</param>
    public void AppendDropDownListSdt(String selectedText, IEnumerable<String> items, String? tag = null)
    {
        AppendSdt(new SdtElement
        {
            SdtType = SdtType.DropDownList,
            Content = selectedText,
            ListItems = items.ToList(),
            Tag = tag
        });
    }

    /// <summary>追加富文本内容控件</summary>
    /// <param name="content">富文本内容</param>
    /// <param name="tag">标签（可选）</param>
    /// <param name="alias">别名（可选）</param>
    public void AppendRichTextSdt(String content, String? tag = null, String? alias = null)
    {
        AppendSdt(new SdtElement
        {
            SdtType = SdtType.RichText,
            Content = content,
            Tag = tag,
            Alias = alias
        });
    }

    /// <summary>追加组合框内容控件（可编辑下拉列表）</summary>
    /// <param name="selectedText">当前选中/输入的文本</param>
    /// <param name="items">下拉建议项</param>
    /// <param name="tag">标签</param>
    public void AppendComboBoxSdt(String selectedText, IEnumerable<String> items, String? tag = null)
    {
        AppendSdt(new SdtElement
        {
            SdtType = SdtType.ComboBox,
            Content = selectedText,
            ListItems = items.ToList(),
            Tag = tag
        });
    }

    /// <summary>追加无序列表</summary>
    /// <param name="items">列表项</param>
    public void AppendBulletList(IEnumerable<String> items)
    {
        foreach (var item in items)
        {
            var para = new Paragraph { IsBullet = true };
            para.Runs.Add(new Run { Text = item });
            _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        }
    }

    /// <summary>追加多级嵌套无序列表</summary>
    /// <param name="items">列表项，每项为 (文本, 级别) 元组，级别 0=一级, 1=二级...</param>
    public void AppendMultiLevelBulletList(IEnumerable<(String Text, Int32 Level)> items)
    {
        foreach (var (text, level) in items)
        {
            var para = new Paragraph { IsBullet = true, ListLevel = level };
            para.Runs.Add(new Run { Text = text });
            _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        }
    }

    /// <summary>追加有序列表</summary>
    /// <param name="items">列表项</param>
    public void AppendOrderedList(IEnumerable<String> items)
    {
        foreach (var item in items)
        {
            var para = new Paragraph { IsOrderedList = true };
            para.Runs.Add(new Run { Text = item });
            _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        }
    }
    #endregion

    #region 表格方法
    /// <summary>追加表格（字符串二维数组）</summary>
    /// <param name="rows">行集合，每行为列字符串集合</param>
    /// <param name="firstRowHeader">首行是否表头</param>
    /// <param name="style">表格样式，null=默认黑色边框</param>
    public void AppendTable(IEnumerable<IEnumerable<String>> rows, Boolean firstRowHeader = false, TableStyle? style = null)
    {
        var tableRows = rows.Select(row => row.Select(cellText =>
        {
            var cell = new Cell();
            var para = new Paragraph();
            para.Runs.Add(new Run { Text = cellText });
            cell.Paragraphs.Add(para);
            return cell;
        }).ToList()).ToList();

        _elements.Add(new Element
        {
            Type = ElementType.Table,
            TableRows = tableRows,
            TableFirstRowHeader = firstRowHeader,
            TableStyle = style,
        });
    }

    /// <summary>追加水平分隔线（段落底部边框）</summary>
    /// <param name="colorHex">线条颜色（16进制RGB，默认 000000）</param>
    /// <param name="width">线宽（1/8pt，默认 6 = 0.75pt）</param>
    public void AppendHorizontalRule(String? colorHex = null, Int32 width = 6)
    {
        var para = new Paragraph
        {
            Borders = new ParagraphBorders
            {
                Bottom = new Border
                {
                    Style = BorderStyle.Single,
                    Width = width,
                    Color = colorHex ?? "000000"
                }
            }
        };
        _elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
    }
    /// <summary>将对象集合写入 Word 表格</summary>
    /// <typeparam name="T">对象类型</typeparam>
    /// <param name="data">对象集合</param>
    /// <param name="firstRowHeader">是否将第一行作为表头</param>
    /// <param name="style">表格样式</param>
    public void WriteObjects<T>(IEnumerable<T> data, Boolean firstRowHeader = true, TableStyle? style = null) where T : class
    {
        WriteObjects(data, firstRowHeader, style, 0);
    }

    /// <summary>将对象集合写入 Word 表格，支持嵌套对象属性展开</summary>
    /// <typeparam name="T">对象类型</typeparam>
    /// <param name="data">对象集合</param>
    /// <param name="firstRowHeader">是否将第一行作为表头</param>
    /// <param name="style">表格样式</param>
    /// <param name="maxDepth">嵌套对象展开深度（0=仅扁平属性，1=展开一层嵌套属性，2=展开两层，以此类推）</param>
    /// <remarks>
    /// 当 maxDepth > 0 时，嵌套对象的属性将按 Parent.Child 格式展平为列名。
    /// 例如 Order.Customer.Name 会生成列 "Customer.Name"。
    /// 集合属性（IEnumerable）不会被展开为子表格，仅展平其直接子属性。
    /// </remarks>
    public void WriteObjects<T>(IEnumerable<T> data, Boolean firstRowHeader, TableStyle? style = null, Int32 maxDepth = 0) where T : class
    {
        var headers = new List<String>();
        var propPaths = new List<PropPathEntry>();

        CollectProperties(typeof(T), [], headers, propPaths, maxDepth);

        var allRows = new List<IEnumerable<String>> { headers };
        foreach (var item in data)
        {
            if (item == null) continue;
            allRows.Add(propPaths.Select(pp =>
            {
                var value = ResolveNestedValue(item, pp.Property, pp.Path);
                return Convert.ToString(value) ?? String.Empty;
            }).ToArray());
        }

        AppendTable(allRows, firstRowHeader, style);
    }
    #endregion

    #region 图片方法
    /// <summary>插入图片</summary>
    /// <param name="imageData">图片字节数据</param>
    /// <param name="extension">文件扩展名（png/jpg）</param>
    /// <param name="widthCm">宽度（厘米）</param>
    /// <param name="heightCm">高度（厘米）</param>
    public void InsertImage(Byte[] imageData, String extension = "png", Double widthCm = 10, Double heightCm = 7.5)
    {
        var relId = $"rImg{_imgCounter++}";
        var ext = extension.TrimStart('.').ToLowerInvariant();
        _imageRels.Add((relId, ext, imageData));
        var img = new Image
        {
            ImageData = imageData,
            Extension = ext,
            RelId = relId,
            WidthEmu = (Int64)(widthCm * 360000),
            HeightEmu = (Int64)(heightCm * 360000),
        };
        _elements.Add(new Element { Type = ElementType.Image, Image = img });
    }

    /// <summary>追加自选图形（W12，矩形/椭圆/线条等 DrawingML 形状）</summary>
    /// <param name="shape">形状对象（类型/尺寸/填充/线条/文本）</param>
    /// <returns>形状对象</returns>
    /// <exception cref="ArgumentNullException">shape 为空</exception>
    public WordShape AppendShape(WordShape shape)
    {
        if (shape == null) throw new ArgumentNullException(nameof(shape));
        _elements.Add(new Element { Type = ElementType.Shape, Shape = shape });
        return shape;
    }

    /// <summary>追加图表（W13，柱状/折线/饼图，含嵌入式数据包）</summary>
    /// <param name="chart">图表对象（类型/标题/分类/系列）</param>
    /// <returns>图表对象</returns>
    /// <exception cref="ArgumentNullException">chart 为空</exception>
    public WordChart AppendChart(WordChart chart)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));
        _elements.Add(new Element { Type = ElementType.Chart, Chart = chart });
        return chart;
    }
    #endregion

    #region 保存方法
    /// <summary>保存到文件</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
        Save(fs);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        using var za = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: Encoding);

        // 透传模式：保留所有原始 ZIP 部件（包含 [Content_Types].xml、rels、主题、字体表等），
        // 仅重新生成 word/document.xml 以支持用户对 Elements 的修改。
        if (_otherParts.Count > 0)
        {
            // 有文档变量变更时，从透传中移除 settings.xml 以便重新生成
            if (_documentVariables.Count > 0 && _settingsXml != null && !DocVarsMatch(_settingsXml, _documentVariables))
                _otherParts.Remove("word/settings.xml");
            foreach (var kv in _otherParts)
                using (var e = za.CreateEntry(kv.Key).Open())
                    e.Write(kv.Value, 0, kv.Value.Length);
            if (_documentVariables.Count > 0 && _settingsXml != null && !DocVarsMatch(_settingsXml, _documentVariables))
                WriteSettings(za);
            // 富文本页眉页脚（用户显式重建时）：生成新部件 + 合并更新 content types/rels 引用
            if (_rebuildHeaderFooter)
            {
                if (_headers.Count > 0) WriteHeaderParts(za);
                if (_footers.Count > 0) WriteFooterParts(za);
                if (_mergeContentTypes != null)
                    WriteEntry(za, "[Content_Types].xml", InjectContentTypeOverrides(_mergeContentTypes));
                if (_mergeDocumentRels != null)
                    WriteEntry(za, "word/_rels/document.xml.rels", InjectHeaderFooterRels(_mergeDocumentRels));
            }
            // 脚注/尾注（模型重建且用户修改/新增时）：重新生成部件，覆盖透传的旧文件
            if (_footnotes.Count > 0 || _endnotes.Count > 0)
                WriteNotes(za);
            WriteDocument(za);
            return;
        }

        // 普通模式（程序化创建，没有源 ZIP）：从模型生成全部文件
        WriteContentTypes(za);
        WriteRels(za);
        WriteStyles(za);
        WriteSettings(za);
        WriteDocument(za);
        WriteNumbering(za);
        WriteDocumentRels(za);
        WriteNotes(za);
        var psave = PageSettings;
        var hdrInOther = _otherParts.ContainsKey("word/header1.xml");
        var ftrInOther = _otherParts.ContainsKey("word/footer1.xml");
        if (_headers.Count > 0)
            WriteHeaderParts(za);
        else if ((psave.HeaderText != null || psave.WatermarkText != null) && !hdrInOther)
            WriteHeaderXml(za);
        if (_footers.Count > 0)
            WriteFooterParts(za);
        else if (psave.FooterText != null && !ftrInOther)
            WriteFooterXml(za);
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            WriteCoreProperties(za);
        WriteCustomProperties(za);
        WriteOtherParts(za);
        // 图片媒体写入：按 relId 去重，避免相同字节重复插入时重复 CreateEntry 造成损坏 zip
        var writtenImages = new HashSet<String>();
        foreach (var (relId, ext, data) in _imageRels)
        {
            if (!writtenImages.Add(relId)) continue;
            using var entry = za.CreateEntry($"word/media/{relId}.{ext}").Open();
            entry.Write(data, 0, data.Length);
        }
        // 图表部件（W13）：chart XML + 嵌入式 xlsx + chart 关系
        WriteCharts(za);
    }

    /// <summary>保存文档模型到文件</summary>
    public void Save(String path, Document document)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
        Save(fs, document);
    }

    /// <summary>保存文档模型到流</summary>
    public void Save(Stream stream, Document document)
    {
        _elements.Clear(); _imageRels.Clear(); _hyperlinkRels.Clear();
        _relCounter = 1; _imgCounter = 1; _bookmarkId = 1;
        _orderedStartOverrides.Clear();
        _rebuildHeaderFooter = false;
        _mergeContentTypes = null;
        _mergeDocumentRels = null;
        _stylesXml = document.StylesXml;
        _numberingXml = document.NumberingXml;
        _numbering = document.Numbering;
        _settingsXml = document.SettingsXml;
        _sectPrXml = document.SectPrXml;
        _documentXmlNsDecls = document.DocumentXmlNsDecls;
        _documentXml = document.DocumentXml;
        _otherParts = document.OtherParts.Count > 0 ? new Dictionary<String, Byte[]>(document.OtherParts) : [];
        _headers = document.Headers;
        _footers = document.Footers;
        _sections = document.Sections;
        // 脚注/尾注：模型输出（DocumentXml=null）时由 Writer 生成部件，否则走 OtherParts 透传
        if (document.DocumentXml == null)
        {
            _footnotes = document.Footnotes;
            _endnotes = document.Endnotes;
            // 计算下一个可用 ID（避免与现有脚注冲突）
            foreach (var fn in _footnotes) if (fn.Id >= _nextNoteId) _nextNoteId = fn.Id + 1;
            foreach (var en in _endnotes) if (en.Id >= _nextNoteId) _nextNoteId = en.Id + 1;
        }
        // 自定义 XML 部件写入 OtherParts 以便原样透传
        foreach (var kv in document.CustomXmlParts)
            _otherParts[$"customXml/{kv.Key}"] = kv.Value;
        _elements.AddRange(document.Elements);
        foreach (var kv in document.Images) _imageRels.Add((kv.Key, kv.Value.Extension, kv.Value.Data));
        foreach (var item in document.Hyperlinks) _hyperlinkRels.Add(item);
        PageSettings = document.PageSettings;
        DocumentProperties = document.DocumentProperties;
        ProtectionReadOnly = document.ProtectionReadOnly;
        _documentVariables = document.DocumentVariables.Count > 0
            ? new Dictionary<String, String>(document.DocumentVariables) : [];
        if (document.PageSettings.HeaderText == null && document.HeaderText != null)
            PageSettings.HeaderText = document.HeaderText;
        if (document.PageSettings.FooterText == null && document.FooterText != null)
            PageSettings.FooterText = document.FooterText;

        // 用户显式重建（DocumentXml=null）且设置了富文本页眉页脚/脚注尾注时：
        // 从透传中移除旧部件与其引用，改由 Writer 生成新部件并合并更新
        // [Content_Types].xml 与 word/_rels/document.xml.rels 引用（含脚注/尾注防重复条目）
        if (document.DocumentXml == null
            && (_headers.Count > 0 || _footers.Count > 0 || _footnotes.Count > 0 || _endnotes.Count > 0))
        {
            _rebuildHeaderFooter = true;
            _mergeContentTypes = document.OtherParts.TryGetValue("[Content_Types].xml", out var ct0)
                ? Encoding.UTF8.GetString(ct0) : null;
            _mergeDocumentRels = document.OtherParts.TryGetValue("word/_rels/document.xml.rels", out var rl0)
                ? Encoding.UTF8.GetString(rl0) : null;

            var removeKeys = _otherParts.Keys
                .Where(k => k.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)
                    || k.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)
                    || k.StartsWith("word/_rels/header", StringComparison.OrdinalIgnoreCase)
                    || k.StartsWith("word/_rels/footer", StringComparison.OrdinalIgnoreCase)
                    || (_footnotes.Count > 0 && k.Equals("word/footnotes.xml", StringComparison.OrdinalIgnoreCase))
                    || (_endnotes.Count > 0 && k.Equals("word/endnotes.xml", StringComparison.OrdinalIgnoreCase))
                    || k.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                    || k.Equals("word/_rels/document.xml.rels", StringComparison.OrdinalIgnoreCase))
                .ToList();
            foreach (var k in removeKeys) _otherParts.Remove(k);
            _sectPrXml = null; // 重建 sectPr 以生成新的 headerReference/footerReference
        }

        Save(stream);
    }

    /// <summary>追加文档元素</summary>
    public void AppendDocument(Document document)
    {
        _elements.AddRange(document.Elements);
        foreach (var kv in document.Images) _imageRels.Add((kv.Key, kv.Value.Extension, kv.Value.Data));
        foreach (var item in document.Hyperlinks) _hyperlinkRels.Add(item);
    }
    #endregion

    #region 私有方法
    private void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding);
        sw.Write(content);
    }

    private static String Esc(String? s) => s == null ? String.Empty : (SecurityElement.Escape(s) ?? s);

    /// <summary>将阴影偏移量转换为 OOXML w14:dir 角度（1/60000 度单位，顺时针从顶部量起）</summary>
    private static Int32 ShadowDirToAngle(Int64 dx, Int64 dy)
    {
        // w14:dir 角度：0=向上/顶部，顺时针递增，单位为 1/60000 度
        if (dx == 0 && dy == 0) return 0;
        var angleRad = Math.Atan2(dx, -dy); // dx=正→右, dy=正→下; -dy 让上方为 0
        var angleDeg = angleRad * 180.0 / Math.PI;
        if (angleDeg < 0) angleDeg += 360.0;
        return (Int32)(angleDeg * 60000);
    }

    /// <summary>写入所有透传部件（主题/字体表/脚注/尾注/页眉页脚 raw XML 等）</summary>
    private void WriteOtherParts(ZipArchive za)
    {
        if (_otherParts.Count == 0) return;

        // 已被显式写入的路径（小写比较）
        var written = new HashSet<String>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml", "_rels/.rels",
            "word/document.xml", "word/_rels/document.xml.rels",
            "word/styles.xml", "word/settings.xml", "word/numbering.xml",
            "docProps/core.xml",
        };
        // 如果 Writer 已生成了页眉/页脚（OtherParts中没有，由模型驱动），则不重复写入
        if ((PageSettings.HeaderText != null || PageSettings.WatermarkText != null) && !_otherParts.ContainsKey("word/header1.xml"))
        {
            written.Add("word/header1.xml");
            written.Add("word/_rels/header1.xml.rels");
        }
        if (PageSettings.FooterText != null && !_otherParts.ContainsKey("word/footer1.xml"))
        {
            written.Add("word/footer1.xml");
            written.Add("word/_rels/footer1.xml.rels");
        }
        // Writer 生成脚注/尾注部件时，排除透传的旧文件
        if (_footnotes.Count > 0) written.Add("word/footnotes.xml");
        if (_endnotes.Count > 0) written.Add("word/endnotes.xml");

        foreach (var kv in _otherParts)
        {
            if (written.Contains(kv.Key)) continue;

            var entry = za.CreateEntry(kv.Key);
            using var es = entry.Open();
            es.Write(kv.Value, 0, kv.Value.Length);
        }
    }

    private void WriteContentTypes(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
        sb.Append("<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");
        sb.Append("<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>");
        if (_numberingXml != null || _elements.Any(e => e.Type == ElementType.Paragraph && e.Paragraph?.IsBullet == true))
            sb.Append("<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>");
        var ps = PageSettings;
        if (_headers.Count > 0)
        {
            for (var i = 0; i < _headers.Count; i++)
                sb.Append($"<Override PartName=\"/word/header{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
        }
        else if (ps.HeaderText != null || ps.WatermarkText != null)
        {
            sb.Append("<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
        }
        if (_footers.Count > 0)
        {
            for (var i = 0; i < _footers.Count; i++)
                sb.Append($"<Override PartName=\"/word/footer{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");
        }
        else if (ps.FooterText != null)
        {
            sb.Append("<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");
        }
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        if (DocumentProperties.CustomProperties.Count > 0)
            sb.Append("<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>");
        // 脚注/尾注（W22）
        if (_footnotes.Count > 0)
            sb.Append("<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>");
        if (_endnotes.Count > 0)
            sb.Append("<Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>");
        // 图表（W13）：chart 部件 + 嵌入式 xlsx
        if (_elements.Any(e => e.Type == ElementType.Chart && e.Chart != null))
        {
            sb.Append("<Override PartName=\"/word/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
            sb.Append("<Default Extension=\"xlsx\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"/>");
        }
        // image content types：为所有实际使用的图片扩展名生成 Default（png/jpg/jpeg/gif/svg/webp 等）
        var imageExts = new HashSet<String>(StringComparer.OrdinalIgnoreCase);
        foreach (var (_, ext, _) in _imageRels)
            imageExts.Add(ext.ToLowerInvariant());
        foreach (var ext in imageExts.OrderBy(e => e))
            sb.Append($"<Default Extension=\"{ext}\" ContentType=\"{GetImageContentType(ext)}\"/>");
        sb.Append("</Types>");
        WriteEntry(za, "[Content_Types].xml", sb.ToString());
    }

    /// <summary>根据图片扩展名返回 OOXML content type</summary>
    /// <param name="ext">扩展名（不含点）</param>
    /// <returns>MIME 类型</returns>
    private static String GetImageContentType(String ext)
    {
        return ext switch
        {
            "png" => "image/png",
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "svg" => "image/svg+xml",
            "tif" or "tiff" => "image/tiff",
            "webp" => "image/webp",
            "emf" => "image/x-emf",
            "wmf" => "image/x-wmf",
            "ico" => "image/x-icon",
            _ => $"image/{ext}",
        };
    }

    private void WriteRels(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>");
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            sb.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
        if (DocumentProperties.CustomProperties.Count > 0)
            sb.Append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/>");
        sb.Append("</Relationships>");
        WriteEntry(za, "_rels/.rels", sb.ToString());
    }

    private void WriteDocumentRels(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.Append("<Relationship Id=\"rStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        sb.Append("<Relationship Id=\"rSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>");
        var psRels = PageSettings;
        if (_headers.Count > 0)
        {
            for (var i = 0; i < _headers.Count; i++)
                sb.Append($"<Relationship Id=\"rHdr{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header{i + 1}.xml\"/>");
        }
        else if (psRels.HeaderText != null || psRels.WatermarkText != null)
        {
            sb.Append("<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>");
        }
        if (_footers.Count > 0)
        {
            for (var i = 0; i < _footers.Count; i++)
                sb.Append($"<Relationship Id=\"rFtr{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer{i + 1}.xml\"/>");
        }
        else if (psRels.FooterText != null)
        {
            sb.Append("<Relationship Id=\"rFtr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>");
        }
        // 脚注/尾注关系（W22）
        if (_footnotes.Count > 0)
            sb.Append("<Relationship Id=\"rFootnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>");
        if (_endnotes.Count > 0)
            sb.Append("<Relationship Id=\"rEndnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/>");
        foreach (var (relId, url) in _hyperlinkRels)
        {
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{Esc(url)}\" TargetMode=\"External\"/>");
        }
        foreach (var (relId, ext, _) in _imageRels)
        {
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{relId}.{ext}\"/>");
        }
        // 图表关系（W13）：段落 rChart{n} → charts/chart{n}.xml
        var chartCount = _elements.Count(e => e.Type == ElementType.Chart && e.Chart != null);
        for (var i = 1; i <= chartCount; i++)
        {
            sb.Append($"<Relationship Id=\"rChart{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"charts/chart{i}.xml\"/>");
        }
        sb.Append("</Relationships>");
        WriteEntry(za, "word/_rels/document.xml.rels", sb.ToString());
    }

    #region 图表生成（W13）
    private void WriteCharts(ZipArchive za)
    {
        var charts = _elements.Where(e => e.Type == ElementType.Chart && e.Chart != null).Select(e => e.Chart!).ToList();
        for (var ci = 0; ci < charts.Count; ci++)
        {
            var chart = charts[ci];
            var chartId = ci + 1;

            // chart{n}.xml
            WriteEntry(za, $"word/charts/chart{chartId}.xml", BuildChartSpaceXml(chart));

            // 嵌入式数据包（ExcelWriter 生成最小 xlsx）
            using (var entry = za.CreateEntry($"word/embeddings/Microsoft_Excel_Worksheet{chartId}.xlsx").Open())
            {
                var data = BuildEmbeddedXlsx(chart);
                entry.Write(data, 0, data.Length);
            }

            // chart 关系：关联嵌入式 xlsx
            WriteEntry(za, $"word/charts/_rels/chart{chartId}.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" Target=\"../embeddings/Microsoft_Excel_Worksheet{chartId}.xlsx\"/>" +
                "</Relationships>");
        }
    }

    /// <summary>构建 chartSpace XML</summary>
    private static String BuildChartSpaceXml(WordChart chart)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        sb.Append("<c:chart>");
        sb.Append("<c:plotArea><c:layout/>");

        var n = chart.Categories.Length;
        var seriesCount = chart.Series.Count;
        var type = chart.Type.ToLowerInvariant();
        var chartTag = type switch
        {
            "bar" or "column" => "c:barChart",
            "line" => "c:lineChart",
            "pie" => "c:pieChart",
            _ => "c:barChart",
        };

        sb.Append($"<{chartTag}>");
        if (chartTag == "c:barChart")
            sb.Append($"<c:barDir val=\"{(type == "bar" ? "bar" : "col")}\"/><c:grouping val=\"clustered\"/>");
        if (chartTag == "c:pieChart")
            sb.Append("<c:varyColors val=\"1\"/>");

        for (var si = 0; si < seriesCount; si++)
        {
            var s = chart.Series[si];
            var colName = GetExcelColName(si + 1); // B=1, C=2...
            sb.Append("<c:ser>");
            sb.Append($"<c:idx val=\"{si}\"/><c:order val=\"{si}\"/>");
            sb.Append($"<c:tx><c:strRef><c:f>Sheet1!${colName}$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{Esc(s.Name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>");
            if (chartTag != "c:pieChart")
            {
                // 类别引用（A 列 2..n+1 行）
                sb.Append($"<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n + 1}</c:f><c:strCache><c:ptCount val=\"{n}\"/>");
                for (var i = 0; i < n; i++)
                    sb.Append($"<c:pt idx=\"{i}\"><c:v>{Esc(chart.Categories[i])}</c:v></c:pt>");
                sb.Append("</c:strCache></c:strRef></c:cat>");
            }
            // 数值引用
            sb.Append($"<c:val><c:numRef><c:f>Sheet1!${colName}$2:${colName}${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{n}\"/>");
            for (var i = 0; i < n; i++)
            {
                var v = i < s.Values.Length ? s.Values[i] : 0;
                sb.Append($"<c:pt idx=\"{i}\"><c:v>{v.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture)}</c:v></c:pt>");
            }
            sb.Append("</c:numCache></c:numRef></c:val>");
            sb.Append("</c:ser>");
        }

        if (chartTag == "c:barChart" || chartTag == "c:lineChart")
        {
            sb.Append("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
        }
        sb.Append($"</{chartTag}>");

        if (chartTag == "c:barChart" || chartTag == "c:lineChart")
        {
            sb.Append("<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>");
            sb.Append("<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>");
        }
        sb.Append("</c:plotArea>");

        // 标题
        if (!String.IsNullOrEmpty(chart.Title))
        {
            sb.Append($"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{Esc(chart.Title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>");
        }

        sb.Append("</c:chart>");
        // 嵌入式数据引用
        sb.Append("<c:externalData r:id=\"rId1\"><c:autoUpdate val=\"0\"/></c:externalData>");
        sb.Append("</c:chartSpace>");
        return sb.ToString();
    }

    /// <summary>生成嵌入式 xlsx（第 1 行表头，2..n+1 行数据）</summary>
    private static Byte[] BuildEmbeddedXlsx(WordChart chart)
    {
        using var ms = new MemoryStream();
        using (var writer = new ExcelWriter(ms))
        {
            writer.SheetName = "Sheet1";
            var header = new List<Object?> { "类别" };
            foreach (var s in chart.Series) header.Add(s.Name);
            writer.WriteRow("Sheet1", header.ToArray());
            for (var i = 0; i < chart.Categories.Length; i++)
            {
                var row = new List<Object?> { chart.Categories[i] };
                foreach (var s in chart.Series)
                    row.Add(i < s.Values.Length ? s.Values[i] : 0);
                writer.WriteRow("Sheet1", row.ToArray());
            }
            writer.Save();
        }
        return ms.ToArray();
    }

    /// <summary>列索引 → Excel 列名（0=A, 1=B...）</summary>
    private static String GetExcelColName(Int32 index)
    {
        var sb = new StringBuilder();
        var i = index;
        while (i >= 0)
        {
            sb.Insert(0, (Char)('A' + i % 26));
            i = i / 26 - 1;
        }
        return sb.ToString();
    }
    #endregion

    private void WriteStyles(ZipArchive za)
    {
        if (_stylesXml != null)
        {
            WriteEntry(za, "word/styles.xml", _stylesXml);
            return;
        }

        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:styles xmlns:w=\"{W}\">");
        sb.Append("<w:docDefaults><w:rPrDefault><w:rPr>");
        sb.Append("<w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:eastAsia=\"SimSun\"/>");
        sb.Append("<w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults>");
        sb.Append("<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>");
        int[] headSizes = [40, 32, 28, 26, 24, 22];
        for (var i = 1; i <= 6; i++)
        {
            sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"Heading{i}\"><w:name w:val=\"heading {i}\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"{i - 1}\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"{headSizes[i - 1]}\"/></w:rPr></w:style>");
        }
        sb.Append("<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/>");
        sb.Append("<w:tblPr><w:tblBorders>");
        foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
        {
            sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        }
        sb.Append("</w:tblBorders></w:tblPr></w:style>");
        sb.Append("</w:styles>");
        WriteEntry(za, "word/styles.xml", sb.ToString());
    }

    private void WriteSettings(ZipArchive za)
    {
        if (_settingsXml != null)
        {
            // 当文档变量与源 settings.xml 一致时，直接透传（保持字节精确）
            if (_documentVariables.Count > 0 && !DocVarsMatch(_settingsXml, _documentVariables))
            {
                var injected = InjectDocVars(_settingsXml, _documentVariables);
                WriteEntry(za, "word/settings.xml", injected);
            }
            else
            {
                WriteEntry(za, "word/settings.xml", _settingsXml);
            }
            return;
        }

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        sb.Append("<w:defaultTabStop w:val=\"720\"/>");
        if (ProtectionReadOnly)
            sb.Append("<w:documentProtection w:edit=\"readOnly\" w:enforcement=\"1\"/>");
        if (_documentVariables.Count > 0)
        {
            sb.Append("<w:docVars>");
            foreach (var kv in _documentVariables)
                sb.Append($"<w:docVar w:name=\"{Esc(kv.Key)}\" w:val=\"{Esc(kv.Value)}\"/>");
            sb.Append("</w:docVars>");
        }
        sb.Append("</w:settings>");
        WriteEntry(za, "word/settings.xml", sb.ToString());
    }

    /// <summary>向 settings.xml 注入文档变量（替换已有 w:docVars 或追加到 w:settings 末尾）</summary>
    private static String InjectDocVars(String settingsXml, Dictionary<String, String> vars)
    {
        var docVarXml = new StringBuilder("<w:docVars>");
        foreach (var kv in vars)
            docVarXml.Append($"<w:docVar w:name=\"{Esc(kv.Key)}\" w:val=\"{Esc(kv.Value)}\"/>");
        docVarXml.Append("</w:docVars>");

        // 如果已有 docVars，替换之
        var idx1 = settingsXml.IndexOf("<w:docVars", StringComparison.Ordinal);
        if (idx1 >= 0)
        {
            var idx2 = settingsXml.IndexOf("</w:docVars>", idx1, StringComparison.Ordinal);
            if (idx2 >= 0)
                return settingsXml[..idx1] + docVarXml + settingsXml[(idx2 + "</w:docVars>".Length)..];
        }

        // 无 docVars，在 </w:settings> 前插入
        var endIdx = settingsXml.LastIndexOf("</w:settings>", StringComparison.Ordinal);
        if (endIdx >= 0)
            return settingsXml[..endIdx] + docVarXml + settingsXml[endIdx..];

        return settingsXml; // 格式异常，不做修改
    }

    /// <summary>检查源 settings.xml 中的文档变量是否与给定变量完全一致</summary>
    private static Boolean DocVarsMatch(String settingsXml, Dictionary<String, String> vars)
    {
        if (vars.Count == 0) return true;
        var existing = new Dictionary<String, String>();
        ParseDocumentVariablesStatic(settingsXml, existing);
        if (existing.Count != vars.Count) return false;
        foreach (var kv in vars)
        {
            if (!existing.TryGetValue(kv.Key, out var v) || v != kv.Value)
                return false;
        }
        return true;
    }

    /// <summary>静态解析 settings.xml 中的文档变量（与 WordReader 中逻辑一致）</summary>
    private static void ParseDocumentVariablesStatic(String settingsXml, Dictionary<String, String> vars)
    {
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(settingsXml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var docVars = doc.SelectSingleNode("//w:docVars", ns) as XmlElement;
            if (docVars == null) return;
            foreach (XmlElement dv in docVars.SelectNodes("w:docVar", ns))
            {
                var name = dv.GetAttribute("w:name");
                var val = dv.GetAttribute("w:val");
                if (!String.IsNullOrEmpty(name))
                    vars[name] = val ?? String.Empty;
            }
        }
        catch { /* 解析失败忽略 */ }
    }

    private void WriteNumbering(ZipArchive za)
    {
        if (_numberingXml != null)
        {
            WriteEntry(za, "word/numbering.xml", _numberingXml);
            return;
        }

        // 检查是否有任何列表项需要编号定义
        var hasBullets = _elements.Any(e => e.Type == ElementType.Paragraph && e.Paragraph?.IsBullet == true);
        var hasOrdered = _elements.Any(e => e.Type == ElementType.Paragraph && e.Paragraph?.IsOrderedList == true);
        if (!hasBullets && !hasOrdered) return;

        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // 使用 Numbering 模型生成（程序化创建自定义列表）
        if (_numbering != null)
        {
            WriteNumberingFromModel(za, W);
            return;
        }

        var hasMultiLevel = _elements.Any(e => e.Type == ElementType.Paragraph && e.Paragraph?.ListLevel > 0);
        var maxLevel = hasMultiLevel ? 3 : 1;

        var sb = new StringBuilder();
        sb.Append($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:numbering xmlns:w=\"{W}\">");

        // 抽象编号定义0：bullet列表
        if (hasBullets)
        {
            sb.Append("<w:abstractNum w:abstractNumId=\"0\">");
            sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
            var bullets = new[] { "\uF0B7", "\uF0D8", "\uF0A7" }; // • → ◆ → §
            for (var l = 0; l < maxLevel; l++)
            {
                sb.Append($"<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/>");
                sb.Append($"<w:lvlText w:val=\"{bullets[l]}\"/><w:lvlJc w:val=\"left\"/>");
                var indent = 720 + l * 720;
                sb.Append($"<w:pPr><w:ind w:left=\"{indent}\" w:hanging=\"360\"/></w:pPr>");
                sb.Append("<w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr>");
                sb.Append("</w:lvl>");
            }
            sb.Append("</w:abstractNum>");
            sb.Append("<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");
        }

        // 抽象编号定义1：ordered列表（decimal/lowerLetter/lowerRoman 层级）
        if (hasOrdered)
        {
            sb.Append("<w:abstractNum w:abstractNumId=\"1\">");
            sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
            for (var l = 0; l < maxLevel; l++)
            {
                var fmt = new[] { "decimal", "lowerLetter", "lowerRoman" }[Math.Min(l, 2)];
                sb.Append($"<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{fmt}\"/>");
                sb.Append($"<w:lvlText w:val=\"%{l + 1}.\"/><w:lvlJc w:val=\"left\"/>");
                var indent = 720 + l * 720;
                sb.Append($"<w:pPr><w:ind w:left=\"{indent}\" w:hanging=\"360\"/></w:pPr>");
                sb.Append("</w:lvl>");
            }
            sb.Append("</w:abstractNum>");
            sb.Append("<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>");

            // numId=3: 有序列表（含 startOverride 的变体）
            if (_orderedStartOverrides.Count > 0)
            {
                sb.Append("<w:num w:numId=\"3\"><w:abstractNumId w:val=\"1\"/>");
                foreach (var kv in _orderedStartOverrides)
                {
                    sb.Append($"<w:lvlOverride w:ilvl=\"{kv.Key}\"><w:startOverride w:val=\"{kv.Value}\"/></w:lvlOverride>");
                }
                sb.Append("</w:num>");
            }
        }

        sb.Append("</w:numbering>");
        WriteEntry(za, "word/numbering.xml", sb.ToString());
    }

    /// <summary>从 Numbering 模型生成 numbering.xml</summary>
    private void WriteNumberingFromModel(ZipArchive za, String wNs)
    {
        var num = _numbering!;
        var sb = new StringBuilder();
        sb.Append($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:numbering xmlns:w=\"{wNs}\">");

        var hasBullets = num.Levels.Any(l => l.Format == "bullet");
        var hasOrdered = num.Levels.Any(l => l.Format != "bullet");

        // 抽象编号定义 0：bullet 列表
        if (hasBullets)
        {
            sb.Append("<w:abstractNum w:abstractNumId=\"0\">");
            sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
            foreach (var lvl in num.Levels.Where(l => l.Format == "bullet").OrderBy(l => l.Level))
            {
                var bulletChar = lvl.BulletChar ?? lvl.Text ?? "\uF0B7";
                sb.Append($"<w:lvl w:ilvl=\"{lvl.Level}\"><w:start w:val=\"{lvl.StartAt}\"/><w:numFmt w:val=\"bullet\"/>");
                sb.Append($"<w:lvlText w:val=\"{Esc(bulletChar)}\"/><w:lvlJc w:val=\"left\"/>");
                sb.Append($"<w:pPr><w:ind w:left=\"{lvl.Indent}\" w:hanging=\"{lvl.HangingIndent}\"/></w:pPr>");
                var fontName = lvl.BulletFontName ?? "Symbol";
                sb.Append($"<w:rPr><w:rFonts w:ascii=\"{fontName}\" w:hAnsi=\"{fontName}\" w:hint=\"default\"/></w:rPr>");
                sb.Append("</w:lvl>");
            }
            sb.Append("</w:abstractNum>");
            sb.Append("<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");
        }

        // 抽象编号定义 1：有序列表
        if (hasOrdered)
        {
            sb.Append("<w:abstractNum w:abstractNumId=\"1\">");
            sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
            foreach (var lvl in num.Levels.Where(l => l.Format != "bullet").OrderBy(l => l.Level))
            {
                var lvlText = lvl.Text ?? $"%{lvl.Level + 1}.";
                sb.Append($"<w:lvl w:ilvl=\"{lvl.Level}\"><w:start w:val=\"{lvl.StartAt}\"/><w:numFmt w:val=\"{lvl.Format}\"/>");
                sb.Append($"<w:lvlText w:val=\"{Esc(lvlText)}\"/><w:lvlJc w:val=\"left\"/>");
                sb.Append($"<w:pPr><w:ind w:left=\"{lvl.Indent}\" w:hanging=\"{lvl.HangingIndent}\"/></w:pPr>");
                sb.Append("</w:lvl>");
            }
            sb.Append("</w:abstractNum>");
            sb.Append("<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>");

            // numId=3: 有序列表（含 startOverride 的变体）
            if (_orderedStartOverrides.Count > 0)
            {
                sb.Append("<w:num w:numId=\"3\"><w:abstractNumId w:val=\"1\"/>");
                foreach (var kv in _orderedStartOverrides)
                {
                    sb.Append($"<w:lvlOverride w:ilvl=\"{kv.Key}\"><w:startOverride w:val=\"{kv.Value}\"/></w:lvlOverride>");
                }
                sb.Append("</w:num>");
            }
        }

        sb.Append("</w:numbering>");
        WriteEntry(za, "word/numbering.xml", sb.ToString());
    }

    private void WriteDocument(ZipArchive za)
    {
        // document.xml 直接透传：保留源文件的全部格式和内容
        if (_documentXml != null)
        {
            // 无 BOM 的 UTF-8，与 Word 生成文件保持一致
            using var sw = new StreamWriter(za.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
            sw.Write(_documentXml);
            return;
        }

        // 必要的 OOXML 命名空间，确保 RawXml 中所有前缀都能解析
        // 源文件可能将这些命名空间声明在子元素而非根元素上，需要在此补全
        var required = new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase)
        {
            ["xmlns:w"]   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            ["xmlns:r"]   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ["xmlns:wp"]  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            ["xmlns:a"]   = "http://schemas.openxmlformats.org/drawingml/2006/main",
            ["xmlns:pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture",
            ["xmlns:mc"]  = "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ["xmlns:v"]   = "urn:schemas-microsoft-com:vml",
            ["xmlns:o"]   = "urn:schemas-microsoft-com:office:office",
            ["xmlns:m"]   = "http://schemas.openxmlformats.org/officeDocument/2006/math",
            ["xmlns:wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            ["xmlns:wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            ["xmlns:wpc"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            ["xmlns:w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
            ["xmlns:w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
        };

        // 合并源文件的命名空间声明（它们可能有更多自定义的）
        if (!String.IsNullOrEmpty(_documentXmlNsDecls))
        {
            var matches = System.Text.RegularExpressions.Regex.Matches(
                _documentXmlNsDecls,
                @"(xmlns:[A-Za-z0-9_]+)=\""([^\""]*)\""");
            foreach (System.Text.RegularExpressions.Match m in matches)
                required[m.Groups[1].Value] = m.Groups[2].Value;
        }

        var nsSb = new StringBuilder();
        foreach (var kv in required)
            nsSb.Append($" {kv.Key}=\"{kv.Value}\"");

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:document{nsSb}>");
        sb.Append("<w:body>");

        // 多节程序化创建：Sections.Count>1 且无原始 document.xml 时按节输出
        if (_sections.Count > 1 && _documentXml == null)
        {
            // 逐节输出：内容 + 节分隔 sectPr（前 n-1 节），最后一节 body 级 sectPr
            sb.Clear();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<w:document{nsSb}>");
            sb.Append("<w:body>");
            for (var si = 0; si < _sections.Count; si++)
            {
                var sec = _sections[si];
                var secChart = 0;
                foreach (var el in sec.Elements)
                    AppendElementXml(sb, el, ref secChart);
                if (si < _sections.Count - 1)
                {
                    // 节分隔：空段落携带本节 sectPr（定义已结束的节）
                    sb.Append("<w:p><w:pPr>");
                    sb.Append(sec.SectPrXml ?? BuildSectPrXml(sec.PageSettings));
                    sb.Append("</w:pPr></w:p>");
                }
            }
            // 最后一节的 sectPr 在 body 级
            var lastSec = _sections[^1];
            sb.Append(lastSec.SectPrXml ?? BuildSectPrXml(lastSec.PageSettings));
        }
        else
        {
            var chartCounter = 0; // 图表编号（W13）
            foreach (var el in _elements)
                AppendElementXml(sb, el, ref chartCounter);

            if (_sectPrXml != null)
            {
                // 单节透传：保留原始节属性 XML
                sb.Append(_sectPrXml);
            }
            else
            {
                // 单节程序化：从 PageSettings 生成节属性
                sb.Append(BuildSectPrXml(PageSettings));
            }
        }

        sb.Append("</w:body></w:document>");
        WriteEntry(za, "word/document.xml", sb.ToString());
    }

    /// <summary>从 PageSettings 生成 <c>w:sectPr</c> 节属性 XML</summary>
    /// <param name="ps">页面设置</param>
    /// <returns>完整的 w:sectPr XML 字符串</returns>
    private String BuildSectPrXml(PageSettings ps)
    {
        var sb = new StringBuilder();
        var pgW = ps.Landscape ? ps.PageHeight : ps.PageWidth;
        var pgH = ps.Landscape ? ps.PageWidth : ps.PageHeight;
        sb.Append("<w:sectPr>");
        if (ps.TitlePage) sb.Append("<w:titlePg/>");
        if (ps.EvenAndOddHeaders) sb.Append("<w:evenAndOddHeaders/>");
        // 富文本页眉页脚（按类型引用），否则回退简单文本页眉页脚
        if (_headers.Count > 0)
        {
            for (var i = 0; i < _headers.Count; i++)
                sb.Append($"<w:headerReference w:type=\"{_headers[i].Type}\" r:id=\"rHdr{i + 1}\"/>");
        }
        else if (ps.HeaderText != null || ps.WatermarkText != null)
        {
            sb.Append("<w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>");
        }
        if (_footers.Count > 0)
        {
            for (var i = 0; i < _footers.Count; i++)
                sb.Append($"<w:footerReference w:type=\"{_footers[i].Type}\" r:id=\"rFtr{i + 1}\"/>");
        }
        else if (ps.FooterText != null)
        {
            sb.Append("<w:footerReference w:type=\"default\" r:id=\"rFtr1\"/>");
        }
        // 分栏设置
        if (ps.ColumnCount > 1)
            sb.Append($"<w:cols w:num=\"{ps.ColumnCount}\" w:space=\"{ps.ColumnSpacing}\"/>");
        // 页面边框
        var pb = ps.PageBorder;
        if (pb != null)
        {
            var offset = pb.OffsetFrom == 0 ? "text" : "page";
            sb.Append($"<w:pgBorders w:offsetFrom=\"{offset}\">");
            AppendPgBorderXml(sb, "top", pb.Top, pb);
            AppendPgBorderXml(sb, "bottom", pb.Bottom, pb);
            AppendPgBorderXml(sb, "left", pb.Left, pb);
            AppendPgBorderXml(sb, "right", pb.Right, pb);
            sb.Append("</w:pgBorders>");
        }
        var orientAttr = ps.Landscape ? " w:orient=\"landscape\"" : String.Empty;
        sb.Append($"<w:pgSz w:w=\"{pgW}\" w:h=\"{pgH}\"{orientAttr}/>");
        sb.Append($"<w:pgMar w:top=\"{ps.MarginTop}\" w:right=\"{ps.MarginRight}\" w:bottom=\"{ps.MarginBottom}\" w:left=\"{ps.MarginLeft}\" w:header=\"720\" w:footer=\"720\"/>");
        // 行号
        var ln = ps.LineNumber;
        if (ln != null)
            sb.Append($"<w:lnNumType w:countBy=\"{ln.CountBy}\" w:start=\"{ln.Start}\" w:distance=\"{ln.Distance}\" w:restart=\"{ln.Restart}\"/>");
        sb.Append("</w:sectPr>");
        return sb.ToString();
    }

    private void WriteHeaderXml(ZipArchive za)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const String V = "urn:schemas-microsoft-com:vml";
        var ps = PageSettings;
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:hdr xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:v=\"{V}\">");
        // 水印（VML）
        if (ps.WatermarkText != null)
        {
            sb.Append("<w:p><w:r><w:pict>");
            sb.Append("<v:shape id=\"wm\" type=\"#_x0000_t136\" style=\"position:absolute;margin-left:0;margin-top:0;");
            sb.Append("width:600pt;height:400pt;z-index:-251655168;");
            sb.Append("mso-position-horizontal:center;mso-position-vertical:center\" ");
            sb.Append("fillcolor=\"#C0C0C0\" stroked=\"f\">");
            sb.Append($"<v:textpath string=\"{Esc(ps.WatermarkText)}\" trim=\"t\" on=\"t\" ");
            sb.Append("style=\"font-family:Arial;font-size:1pt;\"/>");
            sb.Append("</v:shape></w:pict></w:r></w:p>");
        }
        // 页眉文字
        if (ps.HeaderText != null)
        {
            sb.Append("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>");
            sb.Append($"<w:r><w:t>{Esc(ps.HeaderText)}</w:t></w:r></w:p>");
        }
        else if (ps.WatermarkText != null)
        {
            // 水印时需要一个空段落撑开页眉区域
            sb.Append("<w:p/>");
        }
        sb.Append("</w:hdr>");
        WriteEntry(za, "word/header1.xml", sb.ToString());
    }

    private void WriteFooterXml(ZipArchive za)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var ps = PageSettings;
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:ftr xmlns:w=\"{W}\" xmlns:r=\"{R}\">");
        sb.Append("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>");
        if (ps.FooterText != null)
            sb.Append($"<w:r><w:t xml:space=\"preserve\">{Esc(ps.FooterText)}  </w:t></w:r>");
        // 页码字段
        sb.Append("<w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple>");
        sb.Append("</w:p></w:ftr>");
        WriteEntry(za, "word/footer1.xml", sb.ToString());
    }

    /// <summary>生成脚注/尾注部件（footnotes.xml/endnotes.xml，含内置分隔符标记，W22）</summary>
    private void WriteNotes(ZipArchive za)
    {
        if (_footnotes.Count > 0)
            WriteNotePart(za, "word/footnotes.xml", "w:footnote", _footnotes);
        if (_endnotes.Count > 0)
            WriteNotePart(za, "word/endnotes.xml", "w:endnote", _endnotes);
    }

    /// <summary>生成单个脚注/尾注部件</summary>
    private void WriteNotePart(ZipArchive za, String path, String elemName, List<Footnote> notes)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<{elemName}s xmlns:w=\"{W}\">");
        // 内置分隔符/续接分隔符标记（Word 必需）
        sb.Append($"<{elemName} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></{elemName}>");
        sb.Append($"<{elemName} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></{elemName}>");
        foreach (var note in notes)
        {
            sb.Append($"<{elemName} w:id=\"{note.Id}\">");
            if (note.Paragraphs.Count > 0)
            {
                foreach (var p in note.Paragraphs)
                    sb.Append(BuildNoteParagraphXml(p));
            }
            else
            {
                sb.Append($"<w:p><w:r><w:t xml:space=\"preserve\">{Esc(note.Text)}</w:t></w:r></w:p>");
            }
            sb.Append($"</{elemName}>");
        }
        sb.Append($"</{elemName}s>");
        WriteEntry(za, path, sb.ToString());
    }

    /// <summary>构建脚注内容段落 XML（Run 模型，保留粗体/斜体）</summary>
    private static String BuildNoteParagraphXml(Paragraph p)
    {
        var sb = new StringBuilder();
        sb.Append("<w:p>");
        foreach (var run in p.Runs)
        {
            sb.Append("<w:r>");
            if (run.Properties != null)
            {
                sb.Append("<w:rPr>");
                if (run.Properties.Bold == true) sb.Append("<w:b/>");
                if (run.Properties.Italic == true) sb.Append("<w:i/>");
                sb.Append("</w:rPr>");
            }
            sb.Append($"<w:t xml:space=\"preserve\">{Esc(run.Text)}</w:t>");
            sb.Append("</w:r>");
        }
        sb.Append("</w:p>");
        return sb.ToString();
    }

    /// <summary>渲染单个文档元素（段落/表格/图片/形状/图表/内容控件），优先 RawXml 透传</summary>
    /// <param name="sb">目标 StringBuilder</param>
    /// <param name="el">元素</param>
    /// <param name="chartCounter">图表计数器（引用传递，保证跨调用递增）</param>
    private void AppendElementXml(StringBuilder sb, Element el, ref Int32 chartCounter)
    {
        if (el.RawXml != null)
        {
            // 有原始 XML：直接写入，100% 保留所有格式
            sb.Append(el.RawXml);
            return;
        }

        switch (el.Type)
        {
            case ElementType.Paragraph when el.Paragraph != null:
                BuildParagraphXml(sb, el.Paragraph);
                break;
            case ElementType.Table:
                // 优先使用 Table 富模型（支持行高/行级表头/单元格宽度/四边边框），否则回退 TableRows
                if (el.Table != null)
                    BuildTableModelXml(sb, el.Table);
                else if (el.TableRows != null)
                    BuildTableXml(sb, el.TableRows, el.TableFirstRowHeader, el.TableStyle);
                break;
            case ElementType.Image when el.Image != null:
                BuildImageXml(sb, el.Image);
                break;
            case ElementType.Shape when el.Shape != null:
                BuildShapeXml(sb, el.Shape);
                break;
            case ElementType.Chart when el.Chart != null:
                BuildChartXml(sb, el.Chart, ++chartCounter);
                break;
            case ElementType.Sdt:
                if (el.RawXml != null)
                    sb.Append(el.RawXml);
                else if (el.Sdt != null)
                    BuildSdtXml(sb, el.Sdt);
                break;
        }
    }

    /// <summary>写入富文本页眉部件（每个 Header 生成 header{N}.xml）</summary>
    private void WriteHeaderParts(ZipArchive za)
    {
        for (var i = 0; i < _headers.Count; i++)
        {
            var xml = BuildHdrFtrXml("w:hdr", _headers[i].Elements);
            WriteEntry(za, $"word/header{i + 1}.xml", xml);
        }
    }

    /// <summary>向 [Content_Types].xml 注入 header/footer/脚注/尾注 Override（去重）</summary>
    private String InjectContentTypeOverrides(String contentTypes)
    {
        for (var i = 0; i < _headers.Count; i++)
            contentTypes = InjectOneOverride(contentTypes, $"/word/header{i + 1}.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
        for (var i = 0; i < _footers.Count; i++)
            contentTypes = InjectOneOverride(contentTypes, $"/word/footer{i + 1}.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
        if (_footnotes.Count > 0)
            contentTypes = InjectOneOverride(contentTypes, "/word/footnotes.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
        if (_endnotes.Count > 0)
            contentTypes = InjectOneOverride(contentTypes, "/word/endnotes.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml");
        return contentTypes;
    }

    private static String InjectOneOverride(String contentTypes, String partName, String contentType)
    {
        if (contentTypes.Contains($"PartName=\"{partName}\"")) return contentTypes;
        var idx = contentTypes.LastIndexOf("</Types>", StringComparison.Ordinal);
        if (idx < 0) return contentTypes;
        return contentTypes[..idx] + $"<Override PartName=\"{partName}\" ContentType=\"{contentType}\"/>" + contentTypes[idx..];
    }

    /// <summary>向 document.xml.rels 注入 header/footer/脚注/尾注 Relationship（移除旧引用）</summary>
    private String InjectHeaderFooterRels(String relsXml)
    {
        // 移除旧 header/footer/脚注/尾注关系（避免 Id/目标冲突）
        relsXml = System.Text.RegularExpressions.Regex.Replace(
            relsXml,
            "<Relationship[^>]*Type=\"[^\"]*(?:/header|/footer|/footnotes|/endnotes)\"[^>]*/>",
            "");
        // 注入新关系
        var sb = new StringBuilder();
        for (var i = 0; i < _headers.Count; i++)
            sb.Append($"<Relationship Id=\"rHdr{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header{i + 1}.xml\"/>");
        for (var i = 0; i < _footers.Count; i++)
            sb.Append($"<Relationship Id=\"rFtr{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer{i + 1}.xml\"/>");
        if (_footnotes.Count > 0)
            sb.Append("<Relationship Id=\"rFootnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>");
        if (_endnotes.Count > 0)
            sb.Append("<Relationship Id=\"rEndnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/>");
        var idx = relsXml.LastIndexOf("</Relationships>", StringComparison.Ordinal);
        if (idx < 0) return relsXml;
        return relsXml[..idx] + sb + relsXml[idx..];
    }

    /// <summary>写入富文本页脚部件（每个 Footer 生成 footer{N}.xml）</summary>
    private void WriteFooterParts(ZipArchive za)
    {
        for (var i = 0; i < _footers.Count; i++)
        {
            var xml = BuildHdrFtrXml("w:ftr", _footers[i].Elements);
            WriteEntry(za, $"word/footer{i + 1}.xml", xml);
        }
    }

    /// <summary>从元素列表构建页眉/页脚 XML（含必要的命名空间声明）</summary>
    /// <param name="rootTag">根元素（w:hdr / w:ftr）</param>
    /// <param name="elements">内容元素</param>
    /// <returns>完整页眉/页脚 XML 字符串</returns>
    private String BuildHdrFtrXml(String rootTag, List<Element> elements)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<{rootTag} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"");
        sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
        sb.Append(" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"");
        sb.Append(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
        sb.Append(" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"");
        sb.Append(" xmlns:v=\"urn:schemas-microsoft-com:vml\"");
        sb.Append(" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"");
        sb.Append(" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"");
        sb.Append(">");
        var chartCounter = 0;
        foreach (var el in elements)
            AppendElementXml(sb, el, ref chartCounter);
        sb.Append($"</{rootTag}>");
        return sb.ToString();
    }

    private void BuildParagraphXml(StringBuilder sb, Paragraph para)
    {
        // 书签 start/end 均作为 <w:p> 的直接子元素配对（start 在 pPr 前，end 在 runs 后），
        // 不能把 bookmarkStart 放在 <w:p> 之外，否则层级非法导致 Word 提示修复
        var bmId = 0;
        if (para.BookmarkName != null)
            bmId = _bookmarkId++;
        sb.Append("<w:p>");
        if (para.BookmarkName != null)
            sb.Append($"<w:bookmarkStart w:id=\"{bmId}\" w:name=\"{Esc(para.BookmarkName)}\"/>");

        // paragraph properties
        var hasPPr = para.StyleId != null || para.Style != ParagraphStyle.Normal || para.Alignment != null
            || para.IndentLeft.HasValue || para.IndentRight.HasValue || para.FirstLineIndent.HasValue
            || para.SpaceBefore.HasValue || para.SpaceAfter.HasValue || para.LineSpacingPct.HasValue
            || para.IsBullet || para.IsOrderedList || para.NumId.HasValue || para.BackgroundColor != null
            || para.Borders != null || (para.TabStops != null && para.TabStops.Count > 0)
            || para.DropCapLines.HasValue || para.KeepNext || para.KeepLines
            || !para.WidowControl;
        if (hasPPr)
        {
            sb.Append("<w:pPr>");
            if (para.KeepNext) sb.Append("<w:keepNext/>");
            if (para.KeepLines) sb.Append("<w:keepLines/>");
            if (!para.WidowControl) sb.Append("<w:widowControl w:val=\"0\"/>");
            if (para.StyleId != null)
                sb.Append($"<w:pStyle w:val=\"{Esc(para.StyleId)}\"/>");
            else if (para.Style != ParagraphStyle.Normal)
                sb.Append($"<w:pStyle w:val=\"Heading{(Int32)para.Style}\"/>");
            if (para.Alignment != null)
                sb.Append($"<w:jc w:val=\"{para.Alignment}\"/>");
            if (para.BackgroundColor != null)
                sb.Append($"<w:shd w:fill=\"{para.BackgroundColor.TrimStart('#')}\" w:val=\"clear\"/>");
            if (para.SpaceBefore.HasValue || para.SpaceAfter.HasValue || para.LineSpacingPct.HasValue)
            {
                sb.Append("<w:spacing");
                if (para.SpaceBefore.HasValue) sb.Append($" w:before=\"{para.SpaceBefore}\"");
                if (para.SpaceAfter.HasValue) sb.Append($" w:after=\"{para.SpaceAfter}\"");
                if (para.LineSpacingPct.HasValue)
                {
                    // 行距: 单倍=240, 1.5倍=360, 双倍=480; lineRule="auto" 表示百分比
                    var lineValue = para.LineSpacingPct.Value * 240 / 100;
                    sb.Append($" w:line=\"{lineValue}\" w:lineRule=\"auto\"");
                }
                sb.Append("/>");
            }
            if (para.IndentLeft.HasValue || para.IndentRight.HasValue || para.FirstLineIndent.HasValue)
            {
                sb.Append("<w:ind");
                if (para.IndentLeft.HasValue) sb.Append($" w:left=\"{para.IndentLeft}\"");
                if (para.IndentRight.HasValue) sb.Append($" w:right=\"{para.IndentRight}\"");
                if (para.FirstLineIndent.HasValue)
                {
                    if (para.FirstLineIndent.Value >= 0)
                        sb.Append($" w:firstLine=\"{para.FirstLineIndent}\"");
                    else
                        sb.Append($" w:hanging=\"{-para.FirstLineIndent.Value}\"");
                }
                sb.Append("/>");
            }
            if (para.IsBullet || para.IsOrderedList || para.NumId.HasValue)
            {
                // 优先保留原始 numId（真实文档往返保真，引用透传的 numbering.xml），
                // 程序化创建时映射 1/2/3（对应 WriteNumbering 生成的编号定义）
                var numId = para.NumId ?? (para.IsBullet ? 1 : (para.ListStartOverride.HasValue ? 3 : 2));
                sb.Append($"<w:numPr><w:ilvl w:val=\"{para.ListLevel}\"/><w:numId w:val=\"{numId}\"/></w:numPr>");
                if (para.IsOrderedList && para.ListStartOverride.HasValue && !para.NumId.HasValue)
                    _orderedStartOverrides[para.ListLevel] = para.ListStartOverride.Value;
            }
            // 段落边框
            if (para.Borders != null)
            {
                sb.Append("<w:pBdr>");
                AppendBorderXml(sb, "top",    para.Borders.Top);
                AppendBorderXml(sb, "left",   para.Borders.Left);
                AppendBorderXml(sb, "bottom", para.Borders.Bottom);
                AppendBorderXml(sb, "right",  para.Borders.Right);
                sb.Append("</w:pBdr>");
            }
            // 制表位
            if (para.TabStops != null && para.TabStops.Count > 0)
            {
                sb.Append("<w:tabs>");
                foreach (var ts in para.TabStops)
                {
                    sb.Append($"<w:tab w:val=\"{Esc(ts.Alignment)}\" w:pos=\"{ts.Position}\"");
                    if (ts.Leader != null) sb.Append($" w:leader=\"{Esc(ts.Leader)}\"");
                    sb.Append("/>");
                }
                sb.Append("</w:tabs>");
            }
            // 首字下沉
            if (para.DropCapLines.HasValue)
            {
                var chars = para.DropCapChars ?? 1;
                sb.Append($"<w:framePr w:dropCap=\"drop\" w:lines=\"{para.DropCapLines}\" w:hSpace=\"144\" w:vSpace=\"0\" w:wrap=\"around\" w:hAnchor=\"text\" w:vAnchor=\"text\"/>");
            }
            sb.Append("</w:pPr>");
        }
        if (para.IsPageBreak)
        {
            sb.Append("<w:r><w:br w:type=\"page\"/></w:r>");
        }
        else
        {
            foreach (var run in para.Runs)
            {
                BuildRunXml(sb, run);
            }
        }
        if (para.BookmarkName != null)
            sb.Append($"<w:bookmarkEnd w:id=\"{bmId}\"/>");
        sb.Append("</w:p>");
    }

    private static void BuildSdtXml(StringBuilder sb, SdtElement sdt)
    {

        sb.Append("<w:sdt>");
        sb.Append("<w:sdtPr>");

        // 别名和标签
        if (sdt.Alias != null)
            sb.Append($"<w:alias w:val=\"{Esc(sdt.Alias)}\"/>");
        if (sdt.Tag != null)
            sb.Append($"<w:tag w:val=\"{Esc(sdt.Tag)}\"/>");

        // 唯一 ID（负值随机数，符合 OOXML 惯例）
        var id = unchecked((Int32)((UInt32)(sdt.GetHashCode() & 0x7FFFFFFF) + 0x80000000));
        sb.Append($"<w:id w:val=\"{id}\"/>");

        // 占位符文本
        var placeholder = sdt.SdtType switch
        {
            SdtType.PlainText => "单击或点击输入文字",
            SdtType.RichText => "单击或点击输入文字",
            SdtType.Date => "单击或点击输入日期",
            SdtType.DropDownList => "选择一项",
            SdtType.ComboBox => "选择或输入一项",
            _ => "单击或点击输入文字"
        };
        sb.Append($"<w:placeholder><w:docPart w:val=\"{Esc(placeholder)}\"/></w:placeholder>");

        // 类型特定属性
        switch (sdt.SdtType)
        {
            case SdtType.PlainText:
                sb.Append("<w:text/>");
                break;
            case SdtType.RichText:
                sb.Append("<w:richText/>");
                break;
            case SdtType.Date:
                sb.Append("<w:date>");
                var fmt = sdt.DateFormat ?? "yyyy-MM-dd";
                sb.Append($"<w:dateFormat w:val=\"{Esc(fmt)}\"/>");
                sb.Append("<w:lid w:val=\"zh-CN\"/>");
                sb.Append("<w:storeMappedDataAs w:val=\"dateTime\"/>");
                sb.Append("<w:calendar w:val=\"gregorian\"/>");
                sb.Append("</w:date>");
                break;
            case SdtType.DropDownList:
                sb.Append("<w:dropDownList>");
                if (sdt.ListItems != null)
                {
                    foreach (var item in sdt.ListItems)
                        sb.Append($"<w:listItem w:displayText=\"{Esc(item)}\" w:value=\"{Esc(item)}\"/>");
                }
                sb.Append("</w:dropDownList>");
                break;
            case SdtType.ComboBox:
                sb.Append("<w:comboBox>");
                if (sdt.ListItems != null)
                {
                    foreach (var item in sdt.ListItems)
                        sb.Append($"<w:listItem w:displayText=\"{Esc(item)}\" w:value=\"{Esc(item)}\"/>");
                }
                sb.Append("</w:comboBox>");
                break;
            case SdtType.CheckBox:
                sb.Append("<w:checkbox><w:checked w:val=\"0\"/><w:checkedState w:val=\"☒\"/><w:uncheckedState w:val=\"☐\"/></w:checkbox>");
                break;
        }

        sb.Append("</w:sdtPr>");
        sb.Append("<w:sdtContent>");

        // 内容段落
        var content = sdt.Content ?? "";
        sb.Append($"<w:p><w:r><w:rPr><w:rFonts w:ascii=\"等线\" w:hAnsi=\"等线\" w:eastAsia=\"等线\"/></w:rPr><w:t xml:space=\"preserve\">{Esc(content)}</w:t></w:r></w:p>");

        sb.Append("</w:sdtContent>");
        sb.Append("</w:sdt>");
    }

    private static void BuildRunXml(StringBuilder sb, Run run)
    {
        if (run.HyperlinkRelId != null)
            sb.Append($"<w:hyperlink r:id=\"{run.HyperlinkRelId}\" w:history=\"1\">");

        sb.Append("<w:r>");
        var p = run.Properties;
        if (p != null)
        {
            sb.Append("<w:rPr>");
            // 三态写回：true=开启，false=显式关闭（w:val="0"），null=不输出（继承样式）
            AppendSwitchXml(sb, "b", p.Bold);
            AppendSwitchXml(sb, "i", p.Italic);
            AppendSwitchXml(sb, "strike", p.Strikethrough);
            AppendSwitchXml(sb, "smallCaps", p.SmallCaps);
            AppendSwitchXml(sb, "caps", p.AllCaps);
            AppendSwitchXml(sb, "vanish", p.Hidden);
            if (p.Superscript == true) sb.Append("<w:vertAlign w:val=\"superscript\"/>");
            else if (p.Subscript == true) sb.Append("<w:vertAlign w:val=\"subscript\"/>");
            else if (p.Superscript == false || p.Subscript == false) sb.Append("<w:vertAlign w:val=\"baseline\"/>");
            // 下划线（支持样式）
            if (p.Underline != null || p.UnderlineStyle != null)
            {
                var uVal = p.Underline == false && p.UnderlineStyle == null ? "none" : (p.UnderlineStyle ?? "single");
                sb.Append($"<w:u w:val=\"{uVal}\"/>");
            }
            if (p.ForeColor != null) sb.Append($"<w:color w:val=\"{p.ForeColor.TrimStart('#')}\"/>");
            if (p.FontSize.HasValue) sb.Append($"<w:sz w:val=\"{(Int32)(p.FontSize.Value * 2)}\"/>");
            if (p.CharacterSpacing.HasValue) sb.Append($"<w:spacing w:val=\"{p.CharacterSpacing.Value}\"/>");
            if (p.CharacterScaling.HasValue) sb.Append($"<w:w w:val=\"{p.CharacterScaling.Value}\"/>");
            // 字体：西文/东亚字体独立设置（eastAsia 单独可控）
            if (p.FontName != null || p.EastAsiaFontName != null)
            {
                var rf = new StringBuilder();
                if (p.FontName != null)
                    rf.Append($" w:ascii=\"{Esc(p.FontName)}\" w:hAnsi=\"{Esc(p.FontName)}\"");
                if (p.EastAsiaFontName != null)
                    rf.Append($" w:eastAsia=\"{Esc(p.EastAsiaFontName)}\"");
                sb.Append($"<w:rFonts{rf}/>");
            }
            if (p.HighlightColor != null)
                sb.Append($"<w:highlight w:val=\"{Esc(p.HighlightColor.TrimStart('#'))}\"/>");
            if (p.Language != null)
                sb.Append($"<w:lang w:val=\"{Esc(p.Language)}\"/>");
            if (run.HyperlinkRelId != null) sb.Append("<w:rStyle w:val=\"Hyperlink\"/><w:color w:val=\"0563C1\"/><w:u w:val=\"single\"/>");
            // 文字发光效果 (w14:glow)
            if (p.GlowColor != null)
            {
                var rad = p.GlowSize ?? 254000; // EMU，默认 10pt
                sb.Append($"<w14:glow w14:rad=\"{rad}\"><w14:srgbClr val=\"{p.GlowColor.TrimStart('#')}\"/></w14:glow>");
            }
            // 文字阴影效果 (w14:shadow)
            if (p.ShadowColor != null)
            {
                var blurRad = 63500; // EMU，默认 2.5pt
                var dist = p.ShadowOffsetX != null || p.ShadowOffsetY != null
                    ? (Int64)Math.Sqrt((Double)((p.ShadowOffsetX ?? 0) * (p.ShadowOffsetX ?? 0) + (p.ShadowOffsetY ?? 0) * (p.ShadowOffsetY ?? 0)))
                    : 25400L;
                var dir = ShadowDirToAngle(p.ShadowOffsetX ?? 25400, p.ShadowOffsetY ?? 25400);
                sb.Append($"<w14:shadow w14:blurRad=\"{blurRad}\" w14:dist=\"{dist}\" w14:dir=\"{dir}\"><w14:srgbClr val=\"{p.ShadowColor.TrimStart('#')}\"/></w14:shadow>");
            }
            sb.Append("</w:rPr>");
        }
        var spaceAttr = (run.Text.Length > 0 && (run.Text[0] == ' ' || run.Text[^1] == ' '))
            ? " xml:space=\"preserve\"" : "";
        sb.Append($"<w:t{spaceAttr}>{Esc(run.Text)}</w:t>");
        sb.Append("</w:r>");

        if (run.HyperlinkRelId != null)
            sb.Append("</w:hyperlink>");
    }

    /// <summary>生成三态开关元素 XML（w:b 等）：true=开启，false=显式关闭（w:val="0"），null=不输出</summary>
    /// <param name="sb">目标 StringBuilder</param>
    /// <param name="tag">元素名（如 "b"、"i"、"strike"）</param>
    /// <param name="value">三态值</param>
    private static void AppendSwitchXml(StringBuilder sb, String tag, Boolean? value)
    {
        if (value == true)
            sb.Append($"<w:{tag}/>");
        else if (value == false)
            sb.Append($"<w:{tag} w:val=\"0\"/>");
    }

    /// <summary>生成单边段落边框 XML（w:top / w:left / w:bottom / w:right）</summary>
    private static void AppendBorderXml(StringBuilder sb, String edge, Border? border)
    {
        if (border == null || border.Style == BorderStyle.None) return;
        var val = border.Style switch
        {
            BorderStyle.Single     => "single",
            BorderStyle.Thick      => "thick",
            BorderStyle.Double     => "double",
            BorderStyle.Dotted     => "dotted",
            BorderStyle.Dashed     => "dashed",
            BorderStyle.DotDash    => "dotDash",
            BorderStyle.DotDotDash => "dotDotDash",
            _                          => "single",
        };
        sb.Append($"<w:{edge} w:val=\"{val}\" w:sz=\"{border.Width}\" w:space=\"1\"");
        if (!String.IsNullOrEmpty(border.Color)) sb.Append($" w:color=\"{border.Color!.TrimStart('#')}\"");
        if (!String.IsNullOrEmpty(border.ThemeColor)) sb.Append($" w:themeColor=\"{border.ThemeColor}\"");
        if (border.Shadow) sb.Append(" w:shadow=\"1\"");
        sb.Append("/>");
    }

    private static void AppendPgBorderXml(StringBuilder sb, String edge, String? style, PageBorder pb)
    {
        if (style == null || style == "none") return;
        sb.Append($"<w:{edge} w:val=\"{style}\" w:sz=\"{pb.Size}\" w:space=\"{pb.Space}\"");
        if (!String.IsNullOrEmpty(pb.Color)) sb.Append($" w:color=\"{pb.Color!.TrimStart('#')}\"");
        sb.Append("/>");
    }

    private void BuildTableXml(StringBuilder sb, List<List<Cell>> tableRows, Boolean firstRowHeader, TableStyle? style = null)
    {
        var ps = PageSettings;
        var borderColor = style?.BorderColor ?? "000000";
        var borderSize = style?.BorderSize ?? 4;

        sb.Append("<w:tbl><w:tblPr>");
        // 如果有自定义样式，直接内联边框；否则用内置 TableGrid
        if (style != null)
        {
            sb.Append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
            sb.Append("<w:tblBorders>");
            foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
            {
                sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"{borderSize}\" w:space=\"0\" w:color=\"{borderColor}\"/>");
            }
            sb.Append("</w:tblBorders>");
        }
        else
        {
            sb.Append("<w:tblStyle w:val=\"TableGrid\"/>");
            sb.Append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
        }
        sb.Append("</w:tblPr>");

        // 列网格：OOXML 规范要求 tbl 子元素顺序为 tblPr → tblGrid → tr（缺失会被 Word 打开时提示修复）
        var gridCols = 0;
        foreach (var r in tableRows)
            if (r.Count > gridCols) gridCols = r.Count;
        if (gridCols < 1) gridCols = 1;
        var availW = ps.PageWidth - ps.MarginLeft - ps.MarginRight;
        sb.Append("<w:tblGrid>");
        for (var c = 0; c < gridCols; c++)
        {
            Int32 colW;
            if (style?.ColumnWidths != null && c < style.ColumnWidths.Length)
                colW = style.ColumnWidths[c];
            else
                colW = availW / gridCols;
            sb.Append($"<w:gridCol w:w=\"{colW}\"/>");
        }
        sb.Append("</w:tblGrid>");

        for (var ri = 0; ri < tableRows.Count; ri++)
        {
            var row = tableRows[ri];
            sb.Append("<w:tr>");
            if (ri == 0 && firstRowHeader)
                sb.Append("<w:trPr><w:tblHeader/></w:trPr>");

            var colCount = row.Count;

            for (var ci = 0; ci < row.Count; ci++)
            {
                var cell = row[ci];
                // 列宽：优先使用 ColumnWidths，其次均分
                Int32 colW;
                if (style?.ColumnWidths != null && ci < style.ColumnWidths.Length)
                    colW = style.ColumnWidths[ci];
                else
                    colW = colCount > 0 ? availW / colCount : availW;

                sb.Append("<w:tc><w:tcPr>");
                // 列宽：优先单元格自身 tcW，其次 TableStyle.ColumnWidths，最后均分
                if (cell.Width.HasValue)
                    sb.Append($"<w:tcW w:w=\"{cell.Width}\" w:type=\"dxa\"/>");
                else
                    sb.Append($"<w:tcW w:w=\"{colW}\" w:type=\"dxa\"/>");
                // 内联边框（自定义样式时）
                if (style != null)
                {
                    sb.Append("<w:tcBorders>");
                    foreach (var edge in new[] { "top", "left", "bottom", "right" })
                    {
                        sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"{borderSize}\" w:space=\"0\" w:color=\"{borderColor}\"/>");
                    }
                    sb.Append("</w:tcBorders>");
                }
                // 背景色：单元格自身 > 表头行 > 斑马纹
                var bgColor = cell.BackgroundColor;
                if (bgColor == null && ri == 0 && firstRowHeader && style?.HeaderBgColor != null)
                    bgColor = style.HeaderBgColor;
                else if (bgColor == null && ri % 2 == 1 && style?.StripeColor != null)
                    bgColor = style.StripeColor;
                if (bgColor != null)
                    sb.Append($"<w:shd w:fill=\"{bgColor.TrimStart('#')}\" w:val=\"clear\"/>");
                if (cell.VerticalAlignment != null)
                    sb.Append($"<w:vAlign w:val=\"{cell.VerticalAlignment}\"/>");
                if (cell.ColSpan > 1)
                    sb.Append($"<w:gridSpan w:val=\"{cell.ColSpan}\"/>");
                // 垂直合并：RowSpan>1 或 -1（reader 的 restart）= 合并起点；0 = 继续合并
                if (cell.RowSpan > 1 || cell.RowSpan == -1)
                    sb.Append("<w:vMerge w:val=\"restart\"/>");
                else if (cell.RowSpan == 0)
                    sb.Append("<w:vMerge/>");
                // 单元格边框 w:tcBorders
                if (cell.Borders != null)
                {
                    sb.Append("<w:tcBorders>");
                    AppendBorderXml(sb, "top", cell.Borders.Top);
                    AppendBorderXml(sb, "left", cell.Borders.Left);
                    AppendBorderXml(sb, "bottom", cell.Borders.Bottom);
                    AppendBorderXml(sb, "right", cell.Borders.Right);
                    sb.Append("</w:tcBorders>");
                }
                sb.Append("</w:tcPr>");

                // 嵌套表格（W46）：单元格内的 w:tbl
                if (cell.NestedTables.Count > 0)
                {
                    foreach (var nested in cell.NestedTables)
                        BuildTableModelXml(sb, nested);
                }

                foreach (var para in cell.Paragraphs)
                {
                    // 表头行加粗：仅在输出时临时生效，构建后恢复原值，避免保存副作用污染模型
                    var headerBold = new List<(Run Run, Boolean? Old)>();
                    if (ri == 0 && firstRowHeader && style is { HeaderBold: true })
                    {
                        foreach (var run in para.Runs)
                        {
                            run.Properties ??= new RunProperties();
                            if (run.Properties.Bold != true)
                                headerBold.Add((run, run.Properties.Bold));
                            run.Properties.Bold = true;
                        }
                    }
                    BuildParagraphXml(sb, para);
                    foreach (var (run, old) in headerBold)
                        run.Properties!.Bold = old;
                }
                sb.Append("</w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
    }

    /// <summary>从 Table 富模型生成表格 XML（支持行高/行级表头/单元格宽度/四边边框/表格对齐/列宽）</summary>
    /// <param name="sb">目标 StringBuilder</param>
    /// <param name="table">表格模型</param>
    private void BuildTableModelXml(StringBuilder sb, Table table)
    {
        var ps = PageSettings;
        var colCount = table.Rows.Count > 0 ? (table.Rows[0].Cells.Count > 0 ? table.Rows[0].Cells.Count : 1) : 1;
        var availW = ps.PageWidth - ps.MarginLeft - ps.MarginRight;

        sb.Append("<w:tbl><w:tblPr>");
        // 表格样式引用（优先原样保留，如 "TableGrid"）
        if (table.StyleId != null)
            sb.Append($"<w:tblStyle w:val=\"{Esc(table.StyleId)}\"/>");
        // 表格宽度：显式或自适应
        if (table.Width.HasValue)
            sb.Append($"<w:tblW w:w=\"{table.Width}\" w:type=\"dxa\"/>");
        else
            sb.Append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
        // 表格对齐
        if (table.Alignment != null)
            sb.Append($"<w:jc w:val=\"{table.Alignment}\"/>");
        // 表格边框：优先 TableBorders，否则默认黑色单线
        sb.Append("<w:tblBorders>");
        if (table.Borders != null)
        {
            AppendBorderXml(sb, "top", table.Borders.Top);
            AppendBorderXml(sb, "left", table.Borders.Left);
            AppendBorderXml(sb, "bottom", table.Borders.Bottom);
            AppendBorderXml(sb, "right", table.Borders.Right);
            AppendBorderXml(sb, "insideH", table.Borders.InsideH);
            AppendBorderXml(sb, "insideV", table.Borders.InsideV);
        }
        else
        {
            foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
                sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        }
        sb.Append("</w:tblBorders>");
        sb.Append("</w:tblPr>");

        // 列网格：优先 ColumnWidths
        sb.Append("<w:tblGrid>");
        if (table.ColumnWidths != null && table.ColumnWidths.Length > 0)
        {
            foreach (var w in table.ColumnWidths)
                sb.Append($"<w:gridCol w:w=\"{w}\"/>");
        }
        else
        {
            var colW = colCount > 0 ? availW / colCount : availW;
            for (var c = 0; c < colCount; c++)
                sb.Append($"<w:gridCol w:w=\"{colW}\"/>");
        }
        sb.Append("</w:tblGrid>");

        for (var ri = 0; ri < table.Rows.Count; ri++)
        {
            var row = table.Rows[ri];
            sb.Append("<w:tr>");
            if (row.Height.HasValue || row.IsHeader || row.CantSplit || row.BackgroundColor != null)
            {
                sb.Append("<w:trPr>");
                if (row.IsHeader) sb.Append("<w:tblHeader/>");
                if (row.CantSplit) sb.Append("<w:cantSplit/>");
                if (row.Height.HasValue) sb.Append($"<w:trHeight w:val=\"{row.Height}\" w:hRule=\"atLeast\"/>");
                if (row.BackgroundColor != null)
                    sb.Append($"<w:shd w:fill=\"{row.BackgroundColor.TrimStart('#')}\" w:val=\"clear\"/>");
                sb.Append("</w:trPr>");
            }

            var cellCount = row.Cells.Count;
            for (var ci = 0; ci < row.Cells.Count; ci++)
            {
                var cell = row.Cells[ci];
                Int32 colW;
                if (cell.Width.HasValue)
                    colW = cell.Width.Value;
                else if (table.ColumnWidths != null && ci < table.ColumnWidths.Length)
                    colW = table.ColumnWidths[ci];
                else
                    colW = cellCount > 0 ? availW / cellCount : availW;

                sb.Append("<w:tc><w:tcPr>");
                sb.Append($"<w:tcW w:w=\"{colW}\" w:type=\"dxa\"/>");
                if (cell.BackgroundColor != null)
                    sb.Append($"<w:shd w:fill=\"{cell.BackgroundColor.TrimStart('#')}\" w:val=\"clear\"/>");
                if (cell.VerticalAlignment != null)
                    sb.Append($"<w:vAlign w:val=\"{cell.VerticalAlignment}\"/>");
                if (cell.ColSpan > 1)
                    sb.Append($"<w:gridSpan w:val=\"{cell.ColSpan}\"/>");
                if (cell.RowSpan > 1 || cell.RowSpan == -1)
                    sb.Append("<w:vMerge w:val=\"restart\"/>");
                else if (cell.RowSpan == 0)
                    sb.Append("<w:vMerge/>");
                // 单元格边框 w:tcBorders
                if (cell.Borders != null)
                {
                    sb.Append("<w:tcBorders>");
                    AppendBorderXml(sb, "top", cell.Borders.Top);
                    AppendBorderXml(sb, "left", cell.Borders.Left);
                    AppendBorderXml(sb, "bottom", cell.Borders.Bottom);
                    AppendBorderXml(sb, "right", cell.Borders.Right);
                    sb.Append("</w:tcBorders>");
                }
                sb.Append("</w:tcPr>");

                // 嵌套表格（W46）：单元格内的 w:tbl
                if (cell.NestedTables.Count > 0)
                {
                    foreach (var nested in cell.NestedTables)
                        BuildTableModelXml(sb, nested);
                }

                foreach (var para in cell.Paragraphs)
                {
                    // 表头行加粗：仅在输出时临时生效，构建后恢复原值，避免保存副作用污染模型
                    var headerBold = new List<(Run Run, Boolean? Old)>();
                    if (row.IsHeader && table.Style is { HeaderBold: true })
                    {
                        foreach (var run in para.Runs)
                        {
                            run.Properties ??= new RunProperties();
                            if (run.Properties.Bold != true)
                                headerBold.Add((run, run.Properties.Bold));
                            run.Properties.Bold = true;
                        }
                    }
                    BuildParagraphXml(sb, para);
                    foreach (var (run, old) in headerBold)
                        run.Properties!.Bold = old;
                }
                sb.Append("</w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
    }

    private static void BuildImageXml(StringBuilder sb, Image img)
    {
        // 避免 Math.Abs(int.MinValue) 溢出，用掩码取非负
        var id = (img.RelId?.GetHashCode() ?? 0) & 0x7FFFFFFF;
        sb.Append("<w:p><w:r><w:drawing>");

        // 浮动锚定图片：wp:anchor（含位置/环绕），否则 wp:inline
        if (img.AnchorType == "anchor")
        {
            sb.Append("<wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251658240\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">");
            sb.Append("<wp:simplePos x=\"0\" y=\"0\"/>");
            var posH = img.AnchorPosH ?? "offset";
            var posV = img.AnchorPosV ?? "offset";
            sb.Append($"<wp:positionH relativeFrom=\"column\">");
            if (posH == "offset")
                sb.Append($"<wp:posOffset>{img.AnchorOffsetX ?? 0}</wp:posOffset>");
            else
                sb.Append($"<wp:align>{posH}</wp:align>");
            sb.Append("</wp:positionH>");
            sb.Append($"<wp:positionV relativeFrom=\"paragraph\">");
            if (posV == "offset")
                sb.Append($"<wp:posOffset>{img.AnchorOffsetY ?? 0}</wp:posOffset>");
            else
                sb.Append($"<wp:align>{posV}</wp:align>");
            sb.Append("</wp:positionV>");
            // 环绕方式
            sb.Append(img.Wrap switch
            {
                "square" => "<wp:wrapSquare wrapText=\"bothSides\"/>",
                "tight" => "<wp:wrapTight wrapText=\"bothSides\"/>",
                "through" => "<wp:wrapThrough wrapText=\"bothSides\"/>",
                "topAndBottom" => "<wp:wrapTopAndBottom/>",
                "none" => "<wp:wrapNone/>",
                _ => "<wp:wrapSquare wrapText=\"bothSides\"/>",
            });
        }
        else
        {
            sb.Append("<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        }

        sb.Append($"<wp:extent cx=\"{img.WidthEmu}\" cy=\"{img.HeightEmu}\"/>");
        sb.Append($"<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>");
        var descr = String.IsNullOrEmpty(img.AltText) ? "" : $" descr=\"{Esc(img.AltText)}\"";
        sb.Append($"<wp:docPr id=\"{id}\" name=\"Image{id}\"{descr}/>");
        sb.Append("<wp:cNvGraphicFramePr/>");
        sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        sb.Append("<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"\"/><pic:cNvPicPr/></pic:nvPicPr>");
        sb.Append($"<pic:blipFill><a:blip r:embed=\"{img.RelId}\"");
        // SVG 矢量图片：在 a:blip 内嵌 asvg:svgBlip（Word 2016+ 渲染矢量）
        if (img.IsSvg)
            sb.Append($"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"{img.RelId}\"/></a:blip>");
        else
            sb.Append("/>");
        sb.Append("<a:stretch><a:fillRect/></a:stretch></pic:blipFill>");
        sb.Append($"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{img.WidthEmu}\" cy=\"{img.HeightEmu}\"/></a:xfrm>");
        sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>");
        sb.Append("</pic:pic></a:graphicData></a:graphic>");
        sb.Append(img.AnchorType == "anchor" ? "</wp:anchor>" : "</wp:inline>");
        sb.Append("</w:drawing></w:r></w:p>");
    }

    private static void BuildShapeXml(StringBuilder sb, WordShape shape)
    {
        // 避免 Math.Abs(int.MinValue) 溢出，用掩码取非负
        var id = (shape.GetHashCode() & 0x7FFFFFFF) + 2;
        var cx = (Int64)(shape.WidthCm * 360000);
        var cy = (Int64)(shape.HeightCm * 360000);
        sb.Append("<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        sb.Append($"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>");
        sb.Append($"<wp:docPr id=\"{id}\" name=\"Shape{id}\"/>");
        sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">");
        sb.Append("<wps:wsp><wps:cNvSpPr/><wps:spPr>");
        sb.Append($"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>");
        sb.Append($"<a:prstGeom prst=\"{shape.Type}\"><a:avLst/></a:prstGeom>");
        if (!String.IsNullOrEmpty(shape.FillColor))
            sb.Append($"<a:solidFill><a:srgbClr val=\"{shape.FillColor.TrimStart('#')}\"/></a:solidFill>");
        else
            sb.Append("<a:noFill/>");
        if (!String.IsNullOrEmpty(shape.LineColor) || shape.LineWidth > 0)
        {
            var w = (Int32)(shape.LineWidth * 12700);
            sb.Append($"<a:ln w=\"{w}\">");
            if (!String.IsNullOrEmpty(shape.LineColor))
                sb.Append($"<a:solidFill><a:srgbClr val=\"{shape.LineColor.TrimStart('#')}\"/></a:solidFill>");
            sb.Append("</a:ln>");
        }
        sb.Append("</wps:spPr>");
        if (!String.IsNullOrEmpty(shape.Text))
        {
            sb.Append("<wps:txbx><w:txbxContent><w:p><w:r><w:t>");
            sb.Append(Esc(shape.Text));
            sb.Append("</w:t></w:r></w:p></w:txbxContent></wps:txbx>");
        }
        sb.Append("<wps:bodyPr/>");
        sb.Append("</wps:wsp></a:graphicData></a:graphic>");
        sb.Append("</wp:inline></w:drawing></w:r></w:p>");
    }

    private static void BuildChartXml(StringBuilder sb, WordChart chart, Int32 chartId)
    {
        var cx = (Int64)(chart.WidthCm * 360000);
        var cy = (Int64)(chart.HeightCm * 360000);
        var id = 100 + chartId;
        sb.Append("<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        sb.Append($"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>");
        sb.Append($"<wp:docPr id=\"{id}\" name=\"Chart{chartId}\"/>");
        sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
        sb.Append($"<c:chart r:id=\"rChart{chartId}\"/>");
        sb.Append("</a:graphicData></a:graphic>");
        sb.Append("</wp:inline></w:drawing></w:r></w:p>");
    }

    private void WriteCoreProperties(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
        sb.Append("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        if (DocumentProperties.Title != null) sb.Append($"<dc:title>{Esc(DocumentProperties.Title)}</dc:title>");
        if (DocumentProperties.Author != null) sb.Append($"<dc:creator>{Esc(DocumentProperties.Author)}</dc:creator>");
        if (DocumentProperties.Subject != null) sb.Append($"<dc:subject>{Esc(DocumentProperties.Subject)}</dc:subject>");
        if (DocumentProperties.Description != null) sb.Append($"<dc:description>{Esc(DocumentProperties.Description)}</dc:description>");
        sb.Append($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:created>");
        sb.Append("</cp:coreProperties>");
        WriteEntry(za, "docProps/core.xml", sb.ToString());
    }

    private void WriteCustomProperties(ZipArchive za)
    {
        if (DocumentProperties.CustomProperties.Count == 0) return;

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" ");
        sb.Append("xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");

        var pid = 2; // PID 从 2 开始（1 保留给系统属性）
        foreach (var kv in DocumentProperties.CustomProperties)
        {
            var name = Esc(kv.Key);
            var value = Esc(kv.Value.Value);
            sb.Append($"<property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"{pid}\" name=\"{name}\">");
            switch (kv.Value.Type)
            {
                case "i4":
                    sb.Append($"<vt:i4>{value}</vt:i4>");
                    break;
                case "r8":
                    sb.Append($"<vt:r8>{value}</vt:r8>");
                    break;
                case "bool":
                    sb.Append($"<vt:bool>{(value == "true" || value == "1" ? "true" : "false")}</vt:bool>");
                    break;
                case "date":
                    sb.Append($"<vt:filetime>{value}</vt:filetime>");
                    break;
                default:
                    sb.Append($"<vt:lpwstr>{value}</vt:lpwstr>");
                    break;
            }
            sb.Append("</property>");
            pid++;
        }

        sb.Append("</Properties>");
        WriteEntry(za, "docProps/custom.xml", sb.ToString());
    }
    #endregion

    #region 辅助
    /// <summary>属性路径条目</summary>
    private struct PropPathEntry
    {
        public PropertyInfo Property;
        public String[] Path;
    }

    /// <summary>递归收集属性路径，支持嵌套对象展开</summary>
    private static void CollectProperties(Type type, String[] prefixPath, List<String> headers, List<PropPathEntry> entries, Int32 maxDepth)
    {
        foreach (var prop in type.GetProperties())
        {
            if (prop.GetIndexParameters().Length > 0) continue;

            var currentPath = new String[prefixPath.Length + 1];
            Array.Copy(prefixPath, 0, currentPath, 0, prefixPath.Length);
            currentPath[prefixPath.Length] = prop.Name;

            var headerName = GetHeaderName(prop, currentPath);

            if (maxDepth > 0 && IsExpandableType(prop.PropertyType))
            {
                CollectProperties(prop.PropertyType, currentPath, headers, entries, maxDepth - 1);
            }
            else
            {
                headers.Add(headerName);
                entries.Add(new PropPathEntry { Property = prop, Path = currentPath });
            }
        }
    }

    /// <summary>获取属性的列名（优先 DisplayName，其次是属性名）</summary>
    private static String GetHeaderName(PropertyInfo prop, String[] path)
    {
        var dn = prop.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                      .OfType<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault()?.DisplayName;
        var desc = prop.GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), false)
                        .OfType<System.ComponentModel.DescriptionAttribute>().FirstOrDefault()?.Description;
        var leafName = dn ?? desc ?? prop.Name;

        // 扁平属性直接返回
        if (path.Length <= 1) return leafName;

        // 嵌套属性：ParentProp.LeafDisplayName
        var parentName = path[0];
        return parentName + "." + leafName;
    }

    /// <summary>判断类型是否可展开（嵌套对象，非基础类型）</summary>
    private static Boolean IsExpandableType(Type type)
    {
        if (type.IsPrimitive || type == typeof(String) || type == typeof(Decimal)
            || type == typeof(DateTime) || type == typeof(Guid) || type == typeof(TimeSpan))
            return false;

        if (type.IsEnum) return false;
        if (type.IsGenericType && typeof(IEnumerable).IsAssignableFrom(type)) return false;

        // 跳过 Nullable<T>
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            return false;

        return true;
    }

    /// <summary>根据属性路径解析嵌套对象的值</summary>
    /// <param name="root">根对象</param>
    /// <param name="leafProp">叶子属性</param>
    /// <param name="path">完整属性路径</param>
    /// <returns>属性值</returns>
    private static Object? ResolveNestedValue(Object root, PropertyInfo leafProp, String[] path)
    {
        Object? current = root;
        // 遍历路径中的每个中间对象属性（不含最后的叶子属性）
        for (var i = 0; i < path.Length - 1; i++)
        {
            if (current == null) return null;
            var intermediateProp = current.GetType().GetProperty(path[i]);
            if (intermediateProp == null) return null;
            current = intermediateProp.GetValue(current);
        }

        if (current == null) return null;
        return leafProp.GetValue(current);
    }
    #endregion
}
