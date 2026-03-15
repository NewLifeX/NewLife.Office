using System.IO.Compression;
using System.Security;
using System.Text;

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

    /// <summary>是否启用只读保护</summary>
    public Boolean ProtectionReadOnly { get; set; }
    #endregion

    #region 构造
    /// <summary>实例化写入器</summary>
    public WordWriter() { }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 段落方法
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

    /// <summary>追加分页符</summary>
    public void AppendPageBreak()
    {
        var para = new WordParagraph { IsPageBreak = true };
        _elements.Add(new WordElement { Type = WordElementType.Paragraph, Paragraph = para });
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

    /// <summary>追加有序列表</summary>
    /// <param name="items">列表项</param>
    public void AppendOrderedList(IEnumerable<String> items)
    {
        var idx = 1;
        foreach (var item in items)
        {
            var para = new WordParagraph();
            para.Runs.Add(new WordRun { Text = $"{idx++}. {item}" });
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
            allRows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());

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
        var img = new WordImageElement
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
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        using var za = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: Encoding);
        WriteContentTypes(za);
        WriteRels(za);
        WriteStyles(za);
        WriteSettings(za);
        WriteDocumentRels(za);
        WriteDocument(za);
        var psave = PageSettings;
        if (psave.HeaderText != null || psave.WatermarkText != null)
            WriteHeaderXml(za);
        if (psave.FooterText != null)
            WriteFooterXml(za);
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            WriteCoreProperties(za);
        foreach (var (_, ext, data) in _imageRels)
        {
            var relId = _imageRels.First(r => r.Data == data).RelId;
            using var entry = za.CreateEntry($"word/media/{relId}.{ext}").Open();
            entry.Write(data, 0, data.Length);
        }
    }
    #endregion

    #region 私有方法
    private void WriteEntry(ZipArchive za, String path, String content)
    {
        using var sw = new StreamWriter(za.CreateEntry(path).Open(), Encoding);
        sw.Write(content);
    }

    private static String Esc(String? s) => s == null ? String.Empty : (SecurityElement.Escape(s) ?? s);

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
        var ps = PageSettings;
        if (ps.HeaderText != null || ps.WatermarkText != null)
            sb.Append("<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
        if (ps.FooterText != null)
            sb.Append("<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");
        if (DocumentProperties.Title != null || DocumentProperties.Author != null)
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
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
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{Esc(url)}\" TargetMode=\"External\"/>");
        foreach (var (relId, ext, _) in _imageRels)
            sb.Append($"<Relationship Id=\"{relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{relId}.{ext}\"/>");
        sb.Append("</Relationships>");
        WriteEntry(za, "word/_rels/document.xml.rels", sb.ToString());
    }

    private void WriteStyles(ZipArchive za)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:styles xmlns:w=\"{W}\">");
        sb.Append("<w:docDefaults><w:rPrDefault><w:rPr>");
        sb.Append("<w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:eastAsia=\"SimSun\"/>");
        sb.Append("<w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults>");
        // Normal
        sb.Append("<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>");
        // Headings
        int[] headSizes = [40, 32, 28, 26, 24, 22];
        for (var i = 1; i <= 6; i++)
            sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"Heading{i}\"><w:name w:val=\"heading {i}\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"{i - 1}\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"{headSizes[i - 1]}\"/></w:rPr></w:style>");
        // Table Grid
        sb.Append("<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/>");
        sb.Append("<w:tblPr><w:tblBorders>");
        foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
            sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        sb.Append("</w:tblBorders></w:tblPr></w:style>");
        sb.Append("</w:styles>");
        WriteEntry(za, "word/styles.xml", sb.ToString());
    }

    private void WriteSettings(ZipArchive za)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        sb.Append("<w:defaultTabStop w:val=\"720\"/>");
        if (ProtectionReadOnly)
            sb.Append("<w:documentProtection w:edit=\"readOnly\" w:enforcement=\"1\"/>");
        sb.Append("</w:settings>");
        WriteEntry(za, "word/settings.xml", sb.ToString());
    }

    private void WriteDocument(ZipArchive za)
    {
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const String WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const String A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const String PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:wp=\"{WP}\" xmlns:a=\"{A}\" xmlns:pic=\"{PIC}\">");
        sb.Append("<w:body>");

        foreach (var el in _elements)
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
            }
        }

        var ps = PageSettings;
        var pgW = ps.Landscape ? ps.PageHeight : ps.PageWidth;
        var pgH = ps.Landscape ? ps.PageWidth : ps.PageHeight;
        sb.Append("<w:sectPr>");
        if (ps.HeaderText != null || ps.WatermarkText != null)
            sb.Append("<w:headerReference w:type=\"default\" r:id=\"rHdr1\"/>");
        if (ps.FooterText != null)
            sb.Append("<w:footerReference w:type=\"default\" r:id=\"rFtr1\"/>");
        var orientAttr = ps.Landscape ? " w:orient=\"landscape\"" : String.Empty;
        sb.Append($"<w:pgSz w:w=\"{pgW}\" w:h=\"{pgH}\"{orientAttr}/>");
        sb.Append($"<w:pgMar w:top=\"{ps.MarginTop}\" w:right=\"{ps.MarginRight}\" w:bottom=\"{ps.MarginBottom}\" w:left=\"{ps.MarginLeft}\" w:header=\"720\" w:footer=\"720\"/>");
        sb.Append("</w:sectPr>");
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
        var hasPPr = para.Style != WordParagraphStyle.Normal || para.Alignment != null
            || para.IndentLeft.HasValue || para.IndentRight.HasValue || para.FirstLineIndent.HasValue
            || para.SpaceBefore.HasValue || para.SpaceAfter.HasValue || para.LineSpacingPct.HasValue
            || para.IsBullet;
        if (hasPPr)
        {
            sb.Append("<w:pPr>");
            if (para.Style != WordParagraphStyle.Normal)
                sb.Append($"<w:pStyle w:val=\"Heading{(Int32)para.Style}\"/>");
            if (para.Alignment != null)
                sb.Append($"<w:jc w:val=\"{para.Alignment}\"/>");
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
                sb.Append("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>");
            sb.Append("</w:pPr>");
        }
        if (para.IsPageBreak)
        {
            sb.Append("<w:r><w:br w:type=\"page\"/></w:r>");
        }
        else
        {
            foreach (var run in para.Runs)
                BuildRunXml(sb, run);
        }
        sb.Append("</w:p>");
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
            if (p.Underline) sb.Append("<w:u w:val=\"single\"/>");
            if (p.ForeColor != null) sb.Append($"<w:color w:val=\"{p.ForeColor.TrimStart('#')}\"/>");
            if (p.FontSize.HasValue) sb.Append($"<w:sz w:val=\"{(Int32)(p.FontSize.Value * 2)}\"/>");
            if (p.FontName != null) sb.Append($"<w:rFonts w:ascii=\"{Esc(p.FontName)}\" w:hAnsi=\"{Esc(p.FontName)}\" w:eastAsia=\"{Esc(p.FontName)}\"/>");
            if (run.HyperlinkRelId != null) sb.Append("<w:rStyle w:val=\"Hyperlink\"/><w:color w:val=\"0563C1\"/><w:u w:val=\"single\"/>");
            sb.Append("</w:rPr>");
        }
        var spaceAttr = (run.Text.Length > 0 && (run.Text[0] == ' ' || run.Text[^1] == ' '))
            ? " xml:space=\"preserve\"" : "";
        sb.Append($"<w:t{spaceAttr}>{Esc(run.Text)}</w:t>");
        sb.Append("</w:r>");

        if (run.HyperlinkRelId != null)
            sb.Append("</w:hyperlink>");
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
                sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"{borderSize}\" w:space=\"0\" w:color=\"{borderColor}\"/>");
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
                        sb.Append($"<w:{edge} w:val=\"single\" w:sz=\"{borderSize}\" w:space=\"0\" w:color=\"{borderColor}\"/>");
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
        sb.Append("</w:tbl><w:p/>");
    }

    private static void BuildImageXml(StringBuilder sb, WordImageElement img)
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
    #endregion
}