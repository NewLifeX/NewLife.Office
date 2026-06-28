using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace NewLife.Office;

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
    public WordPageSettings PageSettings { get; set; } = new();

    /// <summary>文档属性</summary>
    public WordDocumentProperties DocumentProperties { get; set; } = new();
    #endregion

    #region 私有字段
    private readonly List<WordElement> _elements = [];
    private readonly List<(String RelId, String Url)> _hyperlinkRels = [];
    private readonly List<(String RelId, String Ext, Byte[] Data)> _imageRels = [];
    private Int32 _relCounter = 1;
    private Int32 _imgCounter = 1;
    private Int32 _bookmarkId = 1;
    private readonly Dictionary<Int32, Int32> _orderedStartOverrides = []; // level → startValue

    // 原始 XML 透传（非空时覆盖生成默认）
    private String? _stylesXml;
    private String? _numberingXml;
    private String? _settingsXml;
    private String? _sectPrXml;           // sectPr 原始 XML
    private String? _documentXmlNsDecls;  // document.xml 根元素命名空间声明
    private String? _documentXml;         // word/document.xml 全文（非空时直接写入，跳过重建）

    // 原样透传的所有 ZIP 部件（除 word/document.xml 外）
    private Dictionary<String, Byte[]> _otherParts = [];

    /// <summary>是否启用只读保护</summary>
    public Boolean ProtectionReadOnly { get; set; }

    private Dictionary<String, String> _documentVariables = [];
    #endregion

    #region 构造
    /// <summary>实例化写入器</summary>
    public WordWriter() { }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 段落方法
    /// <summary>追加段落（WordParagraph 对象）</summary>
    /// <param name="para">段落对象</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendParagraph(WordParagraph para)
    {
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加普通段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="style">段落样式</param>
    /// <returns>段落对象（可进一步设置间距/缩进等属性）</returns>
    public WordParagraph AppendParagraph(String text, WordParagraphStyle style = WordParagraphStyle.Normal)
    {
        var para = new WordParagraph { Style = style };
        para.Runs.Add(new WordRun { Text = text });
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加带格式的段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="style">段落样式</param>
    /// <param name="runProps">文字格式</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendParagraph(String text, WordParagraphStyle style, WordRunProperties runProps)
    {
        var para = new WordParagraph { Style = style };
        para.Runs.Add(new WordRun { Text = text, Properties = runProps });
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加标题</summary>
    /// <param name="text">标题文本</param>
    /// <param name="level">标题级别（1-6）</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendHeading(String text, Int32 level = 1)
    {
        if (level < 1) level = 1;
        if (level > 6) level = 6;
        return AppendParagraph(text, (WordParagraphStyle)level);
    }

    /// <summary>追加多格式 Run 的段落</summary>
    /// <param name="runs">Run 集合</param>
    /// <param name="style">段落样式</param>
    /// <param name="alignment">对齐（left/center/right/both）</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendFormattedParagraph(IEnumerable<WordRun> runs, WordParagraphStyle style = WordParagraphStyle.Normal, String? alignment = null)
    {
        var para = new WordParagraph { Style = style, Alignment = alignment };
        para.Runs.AddRange(runs);
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加超链接段落</summary>
    /// <param name="displayText">显示文本</param>
    /// <param name="url">目标 URL</param>
    /// <param name="runProps">可选文字格式</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendHyperlink(String displayText, String url, WordRunProperties? runProps = null)
    {
        var relId = $"rHyp{_relCounter++}";
        _hyperlinkRels.Add((relId, url));
        var para = new WordParagraph();
        para.Runs.Add(new WordRun { Text = displayText, Properties = runProps, HyperlinkRelId = relId });
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        return para;
    }

    /// <summary>追加带书签的段落</summary>
    /// <param name="text">文本内容</param>
    /// <param name="bookmarkName">书签名称</param>
    /// <param name="style">段落样式</param>
    /// <returns>段落对象</returns>
    public WordParagraph AppendBookmarkedParagraph(String text, String bookmarkName, WordParagraphStyle style = WordParagraphStyle.Normal)
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
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, RawXml = xml });
    }

    /// <summary>追加邮件合并域（MERGEFIELD）</summary>
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
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, RawXml = xml });
    }

    /// <summary>追加分页符</summary>
    public void AppendPageBreak()
    {
        var para = new WordParagraph { IsPageBreak = true };
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
    }

    /// <summary>追加内容控件（SDT）</summary>
    /// <param name="sdt">SDT 内容控件元素</param>
    public void AppendSdt(WordSdtElement sdt)
    {
        _elements.Add(new WordElement { Type = WordElementType.Sdt, Sdt = sdt });
    }

    /// <summary>追加纯文本内容控件</summary>
    /// <param name="content">控件内容文本</param>
    /// <param name="tag">标签（可选，用于标识控件）</param>
    /// <param name="alias">别名（可选，用于展示名称）</param>
    public void AppendPlainTextSdt(String content, String? tag = null, String? alias = null)
    {
        AppendSdt(new WordSdtElement
        {
            SdtType = WordSdtType.PlainText,
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
        AppendSdt(new WordSdtElement
        {
            SdtType = WordSdtType.Date,
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
        AppendSdt(new WordSdtElement
        {
            SdtType = WordSdtType.DropDownList,
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
        AppendSdt(new WordSdtElement
        {
            SdtType = WordSdtType.RichText,
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
        AppendSdt(new WordSdtElement
        {
            SdtType = WordSdtType.ComboBox,
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
            var para = new WordParagraph { IsBullet = true };
            para.Runs.Add(new WordRun { Text = item });
            _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        }
    }

    /// <summary>追加多级嵌套无序列表</summary>
    /// <param name="items">列表项，每项为 (文本, 级别) 元组，级别 0=一级, 1=二级...</param>
    public void AppendMultiLevelBulletList(IEnumerable<(String Text, Int32 Level)> items)
    {
        foreach (var (text, level) in items)
        {
            var para = new WordParagraph { IsBullet = true, ListLevel = level };
            para.Runs.Add(new WordRun { Text = text });
            _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        }
    }

    /// <summary>追加有序列表</summary>
    /// <param name="items">列表项</param>
    public void AppendOrderedList(IEnumerable<String> items)
    {
        foreach (var item in items)
        {
            var para = new WordParagraph { IsOrderedList = true };
            para.Runs.Add(new WordRun { Text = item });
            _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
        }
    }
    #endregion

    #region 表格方法
    /// <summary>追加表格（字符串二维数组）</summary>
    /// <param name="rows">行集合，每行为列字符串集合</param>
    /// <param name="firstRowHeader">首行是否表头</param>
    /// <param name="style">表格样式，null=默认黑色边框</param>
    public void AppendTable(IEnumerable<IEnumerable<String>> rows, Boolean firstRowHeader = false, WordTableStyle? style = null)
    {
        var tableRows = rows.Select(row => row.Select(cellText =>
        {
            var cell = new WordCell();
            var para = new WordParagraph();
            para.Runs.Add(new WordRun { Text = cellText });
            cell.Paragraphs.Add(para);
            return cell;
        }).ToList()).ToList();

        _elements.Add(new WordElement
        {
            Type = WordElementType.Table,
            TableRows = tableRows,
            TableFirstRowHeader = firstRowHeader,
            TableStyle = style,
        });
    }

    /// <summary>追加对象集合为表格</summary>
    /// <param name="data">对象集合</param>
    /// <param name="firstRowHeader">首行表头</param>
    /// <param name="style">表格样式</param>

    /// <summary>追加水平分隔线（段落底部边框）</summary>
    /// <param name="colorHex">线条颜色（16进制RGB，默认 000000）</param>
    /// <param name="width">线宽（1/8pt，默认 6 = 0.75pt）</param>
    public void AppendHorizontalRule(String? colorHex = null, Int32 width = 6)
    {
        var para = new WordParagraph
        {
            Borders = new WordParagraphBorders
            {
                Bottom = new WordBorder
                {
                    Style = WordBorderStyle.Single,
                    Width = width,
                    Color = colorHex ?? "000000"
                }
            }
        };
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
    }
    public void WriteObjects<T>(IEnumerable<T> data, Boolean firstRowHeader = true, WordTableStyle? style = null) where T : class
    {
        var props = typeof(T).GetProperties();
        var headers = props.Select(p =>
        {
            var dn = p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                      .OfType<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault()?.DisplayName;
            var desc = p.GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), false)
                        .OfType<System.ComponentModel.DescriptionAttribute>().FirstOrDefault()?.Description;
            return dn ?? desc ?? p.Name;
        }).ToArray();

        var allRows = new List<IEnumerable<String>> { headers };
        foreach (var item in data)
        {
            allRows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());
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
        var img = new WordImage
        {
            ImageData = imageData,
            Extension = ext,
            RelId = relId,
            WidthEmu = (Int64)(widthCm * 360000),
            HeightEmu = (Int64)(heightCm * 360000),
        };
        _elements.Add(new WordElement { Type = WordElementType.Image, Image = img });
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
        var psave = PageSettings;
        var hdrInOther = _otherParts.ContainsKey("word/header1.xml");
        var ftrInOther = _otherParts.ContainsKey("word/footer1.xml");
        if ((psave.HeaderText != null || psave.WatermarkText != null) && !hdrInOther)
            WriteHeaderXml(za);
        if (psave.FooterText != null && !ftrInOther)
            WriteFooterXml(za);
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            WriteCoreProperties(za);
            WriteCustomProperties(za);
        WriteOtherParts(za);
        foreach (var (_, ext, data) in _imageRels)
        {
            var relId = _imageRels.First(r => r.Data == data).RelId;
            using var entry = za.CreateEntry($"word/media/{relId}.{ext}").Open();
            entry.Write(data, 0, data.Length);
        }
    }

    /// <summary>保存文档模型到文件</summary>
    public void Save(String path, WordDocument document)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
        Save(fs, document);
    }

    /// <summary>保存文档模型到流</summary>
    public void Save(Stream stream, WordDocument document)
    {
        _elements.Clear(); _imageRels.Clear(); _hyperlinkRels.Clear();
        _relCounter = 1; _imgCounter = 1; _bookmarkId = 1;
        _orderedStartOverrides.Clear();
        _stylesXml = document.StylesXml;
        _numberingXml = document.NumberingXml;
        _settingsXml = document.SettingsXml;
        _sectPrXml = document.SectPrXml;
        _documentXmlNsDecls = document.DocumentXmlNsDecls;
        _documentXml = document.DocumentXml;
        _otherParts = document.OtherParts.Count > 0 ? new Dictionary<String, Byte[]>(document.OtherParts) : [];
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
        Save(stream);
    }

    /// <summary>追加文档元素</summary>
    public void AppendDocument(WordDocument document)
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
        if (_numberingXml != null || _elements.Any(e => e.Type == WordElementType.Paragraph && e.Paragraph?.IsBullet == true))
            sb.Append("<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>");
        var ps = PageSettings;
        if (ps.HeaderText != null || ps.WatermarkText != null)
            sb.Append("<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
        if (ps.FooterText != null)
            sb.Append("<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        if (DocumentProperties.CustomProperties.Count > 0)
            sb.Append("<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>");
        // image content types
        var addedPng = false; var addedJpeg = false;
        foreach (var (_, ext, _) in _imageRels)
        {
            if ((ext == "png") && !addedPng)
            {
                sb.Append("<Default Extension=\"png\" ContentType=\"image/png\"/>");
                addedPng = true;
            }
            else if ((ext is "jpg" or "jpeg") && !addedJpeg)
            {
                sb.Append("<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>");
                addedJpeg = true;
            }
        }
        sb.Append("</Types>");
        WriteEntry(za, "[Content_Types].xml", sb.ToString());
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
        if (psRels.HeaderText != null || psRels.WatermarkText != null)
            sb.Append("<Relationship Id=\"rHdr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>");
        if (psRels.FooterText != null)
            sb.Append("<Relationship Id=\"rFtr1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>");
        foreach (var (relId, url) in _hyperlinkRels)
        {
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{Esc(url)}\" TargetMode=\"External\"/>");
        }
        foreach (var (relId, ext, _) in _imageRels)
        {
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{relId}.{ext}\"/>");
        }
        sb.Append("</Relationships>");
        WriteEntry(za, "word/_rels/document.xml.rels", sb.ToString());
    }

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
        var hasBullets = _elements.Any(e => e.Type == WordElementType.Paragraph && e.Paragraph?.IsBullet == true);
        var hasOrdered = _elements.Any(e => e.Type == WordElementType.Paragraph && e.Paragraph?.IsOrderedList == true);
        if (!hasBullets && !hasOrdered) return;

        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var hasMultiLevel = _elements.Any(e => e.Type == WordElementType.Paragraph && e.Paragraph?.ListLevel > 0);
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

        foreach (var el in _elements)
        {
            if (el.RawXml != null)
            {
                // 有原始 XML：直接写入，100% 保留所有格式
                sb.Append(el.RawXml);
            }
            else
            {
                switch (el.Type)
                {
                    case WordElementType.Paragraph when el.Paragraph != null:
                        BuildParagraphXml(sb, el.Paragraph);
                        break;
                    case WordElementType.Table when el.TableRows != null:
                        BuildTableXml(sb, el.TableRows, el.TableFirstRowHeader, el.TableStyle);
                        break;
                    case WordElementType.Image when el.Image != null:
                        BuildImageXml(sb, el.Image);
                        break;
                    case WordElementType.Sdt:
                        if (el.RawXml != null)
                            sb.Append(el.RawXml);
                        else if (el.Sdt != null)
                            BuildSdtXml(sb, el.Sdt);
                        break;
                }
            }
        }

        // 节属性（页面尺寸/页眉页脚引用）
        if (_sectPrXml != null)
        {
            sb.Append(_sectPrXml);
        }
        else
        {
            var ps = PageSettings;
            var pgW = ps.Landscape ? ps.PageHeight : ps.PageWidth;
            var pgH = ps.Landscape ? ps.PageWidth : ps.PageHeight;
            sb.Append("<w:sectPr>");
            if (ps.TitlePage) sb.Append("<w:titlePg/>");
            if (ps.EvenAndOddHeaders) sb.Append("<w:evenAndOddHeaders/>");
            if (ps.HeaderText != null || ps.WatermarkText != null)
                sb.Append("<w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>");
            if (ps.FooterText != null)
                sb.Append("<w:footerReference w:type=\"default\" r:id=\"rFtr1\"/>");
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
        }

        sb.Append("</w:body></w:document>");
        WriteEntry(za, "word/document.xml", sb.ToString());
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

    private void BuildParagraphXml(StringBuilder sb, WordParagraph para)
    {
        // 书签开始标记放在 <w:p> 之前（包住整段）
        if (para.BookmarkName != null)
        {
            var bmId = _bookmarkId++;
            sb.Append($"<w:bookmarkStart w:id=\"{bmId}\" w:name=\"{Esc(para.BookmarkName)}\"/>");
            sb.Append("<w:p>");
            sb.Append($"<w:bookmarkEnd w:id=\"{bmId}\"/>");
        }
        else
        {
            sb.Append("<w:p>");
        }

        // paragraph properties
        var hasPPr = para.StyleId != null || para.Style != WordParagraphStyle.Normal || para.Alignment != null
            || para.IndentLeft.HasValue || para.IndentRight.HasValue || para.FirstLineIndent.HasValue
            || para.SpaceBefore.HasValue || para.SpaceAfter.HasValue || para.LineSpacingPct.HasValue
            || para.IsBullet || para.IsOrderedList || para.BackgroundColor != null
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
            else if (para.Style != WordParagraphStyle.Normal)
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
            if (para.IsBullet)
                sb.Append($"<w:numPr><w:ilvl w:val=\"{para.ListLevel}\"/><w:numId w:val=\"1\"/></w:numPr>");
            else if (para.IsOrderedList)
            {
                var numId = para.ListStartOverride.HasValue ? 3 : 2; // numId=3 用于有 startOverride 的列表
                sb.Append($"<w:numPr><w:ilvl w:val=\"{para.ListLevel}\"/><w:numId w:val=\"{numId}\"/></w:numPr>");
                if (para.ListStartOverride.HasValue)
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
        sb.Append("</w:p>");
    }

    private static void BuildSdtXml(StringBuilder sb, WordSdtElement sdt)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
            WordSdtType.PlainText => "单击或点击输入文字",
            WordSdtType.RichText => "单击或点击输入文字",
            WordSdtType.Date => "单击或点击输入日期",
            WordSdtType.DropDownList => "选择一项",
            WordSdtType.ComboBox => "选择或输入一项",
            _ => "单击或点击输入文字"
        };
        sb.Append($"<w:placeholder><w:docPart w:val=\"{Esc(placeholder)}\"/></w:placeholder>");

        // 类型特定属性
        switch (sdt.SdtType)
        {
            case WordSdtType.PlainText:
                sb.Append("<w:text/>");
                break;
            case WordSdtType.RichText:
                sb.Append("<w:richText/>");
                break;
            case WordSdtType.Date:
                sb.Append("<w:date>");
                var fmt = sdt.DateFormat ?? "yyyy-MM-dd";
                sb.Append($"<w:dateFormat w:val=\"{Esc(fmt)}\"/>");
                sb.Append("<w:lid w:val=\"zh-CN\"/>");
                sb.Append("<w:storeMappedDataAs w:val=\"dateTime\"/>");
                sb.Append("<w:calendar w:val=\"gregorian\"/>");
                sb.Append("</w:date>");
                break;
            case WordSdtType.DropDownList:
                sb.Append("<w:dropDownList>");
                if (sdt.ListItems != null)
                {
                    foreach (var item in sdt.ListItems)
                        sb.Append($"<w:listItem w:displayText=\"{Esc(item)}\" w:value=\"{Esc(item)}\"/>");
                }
                sb.Append("</w:dropDownList>");
                break;
            case WordSdtType.ComboBox:
                sb.Append("<w:comboBox>");
                if (sdt.ListItems != null)
                {
                    foreach (var item in sdt.ListItems)
                        sb.Append($"<w:listItem w:displayText=\"{Esc(item)}\" w:value=\"{Esc(item)}\"/>");
                }
                sb.Append("</w:comboBox>");
                break;
            case WordSdtType.CheckBox:
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

    private static void BuildRunXml(StringBuilder sb, WordRun run)
    {
        if (run.HyperlinkRelId != null)
            sb.Append($"<w:hyperlink r:id=\"{run.HyperlinkRelId}\" w:history=\"1\">");

        sb.Append("<w:r>");
        var p = run.Properties;
        if (p != null)
        {
            sb.Append("<w:rPr>");
            if (p.Bold) sb.Append("<w:b/>");
            if (p.Italic) sb.Append("<w:i/>");
            if (p.Strikethrough) sb.Append("<w:strike/>");
            if (p.Superscript) sb.Append("<w:vertAlign w:val=\"superscript\"/>");
            else if (p.Subscript) sb.Append("<w:vertAlign w:val=\"subscript\"/>");
            // 下划线（支持样式）
            if (p.Underline || p.UnderlineStyle != null)
            {
                var uVal = p.UnderlineStyle ?? "single";
                sb.Append($"<w:u w:val=\"{uVal}\"/>");
            }
            if (p.ForeColor != null) sb.Append($"<w:color w:val=\"{p.ForeColor.TrimStart('#')}\"/>");
            if (p.FontSize.HasValue) sb.Append($"<w:sz w:val=\"{(Int32)(p.FontSize.Value * 2)}\"/>");
            if (p.CharacterSpacing.HasValue) sb.Append($"<w:spacing w:val=\"{p.CharacterSpacing.Value}\"/>");
            if (p.CharacterScaling.HasValue) sb.Append($"<w:w w:val=\"{p.CharacterScaling.Value}\"/>");
            if (p.FontName != null) sb.Append($"<w:rFonts w:ascii=\"{Esc(p.FontName)}\" w:hAnsi=\"{Esc(p.FontName)}\" w:eastAsia=\"{Esc(p.FontName)}\"/>");
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

    /// <summary>生成单边段落边框 XML（w:top / w:left / w:bottom / w:right）</summary>
    private static void AppendBorderXml(StringBuilder sb, String edge, WordBorder? border)
    {
        if (border == null || border.Style == WordBorderStyle.None) return;
        var val = border.Style switch
        {
            WordBorderStyle.Single     => "single",
            WordBorderStyle.Thick      => "thick",
            WordBorderStyle.Double     => "double",
            WordBorderStyle.Dotted     => "dotted",
            WordBorderStyle.Dashed     => "dashed",
            WordBorderStyle.DotDash    => "dotDash",
            WordBorderStyle.DotDotDash => "dotDotDash",
            _                          => "single",
        };
        sb.Append($"<w:{edge} w:val=\"{val}\" w:sz=\"{border.Width}\" w:space=\"1\"");
        if (!String.IsNullOrEmpty(border.Color)) sb.Append($" w:color=\"{border.Color!.TrimStart('#')}\"");
        if (!String.IsNullOrEmpty(border.ThemeColor)) sb.Append($" w:themeColor=\"{border.ThemeColor}\"");
        if (border.Shadow) sb.Append(" w:shadow=\"1\"");
        sb.Append("/>");
    }

    private static void AppendPgBorderXml(StringBuilder sb, String edge, String? style, WordPageBorder pb)
    {
        if (style == null || style == "none") return;
        sb.Append($"<w:{edge} w:val=\"{style}\" w:sz=\"{pb.Size}\" w:space=\"{pb.Space}\"");
        if (!String.IsNullOrEmpty(pb.Color)) sb.Append($" w:color=\"{pb.Color!.TrimStart('#')}\"");
        sb.Append("/>");
    }

    private void BuildTableXml(StringBuilder sb, List<List<WordCell>> tableRows, Boolean firstRowHeader, WordTableStyle? style = null)
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

        for (var ri = 0; ri < tableRows.Count; ri++)
        {
            var row = tableRows[ri];
            sb.Append("<w:tr>");
            if (ri == 0 && firstRowHeader)
                sb.Append("<w:trPr><w:tblHeader/></w:trPr>");

            var colCount = row.Count;
            var availW = ps.PageWidth - ps.MarginLeft - ps.MarginRight;

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
                sb.Append("</w:tcPr>");

                foreach (var para in cell.Paragraphs)
                {
                    // 表头行加粗
                    if (ri == 0 && firstRowHeader && style is { HeaderBold: true })
                    {
                        foreach (var run in para.Runs)
                        {
                            run.Properties ??= new WordRunProperties();
                            run.Properties.Bold = true;
                        }
                    }
                    BuildParagraphXml(sb, para);
                }
                sb.Append("</w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
    }

    private static void BuildImageXml(StringBuilder sb, WordImage img)
    {
        var id = Math.Abs(img.RelId.GetHashCode());
        sb.Append("<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        sb.Append($"<wp:extent cx=\"{img.WidthEmu}\" cy=\"{img.HeightEmu}\"/>");
        sb.Append($"<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>");
        sb.Append($"<wp:docPr id=\"{id}\" name=\"Image{id}\"/>");
        sb.Append("<wp:cNvGraphicFramePr/>");
        sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        sb.Append("<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"\"/><pic:cNvPicPr/></pic:nvPicPr>");
        sb.Append($"<pic:blipFill><a:blip r:embed=\"{img.RelId}\"/>");
        sb.Append("<a:stretch><a:fillRect/></a:stretch></pic:blipFill>");
        sb.Append($"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{img.WidthEmu}\" cy=\"{img.HeightEmu}\"/></a:xfrm>");
        sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>");
        sb.Append("</pic:pic></a:graphicData></a:graphic>");
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
}
