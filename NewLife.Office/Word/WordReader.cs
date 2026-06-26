using System.IO.Compression;
using System.Text;
using System.Xml;

namespace NewLife.Office;

/// <summary>Word docx 读取器</summary>
/// <remarks>
/// 直接解析 Open XML（ZIP+XML）提取文本、表格、图片等内容。
/// </remarks>
public class WordReader : IDisposable, ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>源文件路径（从文件构造时有效）</summary>
    public String? FilePath { get; private set; }
    #endregion

    #region 私有字段
    private readonly ZipArchive _zip;
    private Boolean _disposed;
    #endregion

    #region 构造
    /// <summary>从文件路径打开</summary>
    /// <param name="path">docx 文件路径</param>
    public WordReader(String path)
    {
        FilePath = path.GetFullPath();
        _zip = ZipFile.OpenRead(FilePath);
    }

    /// <summary>从流打开</summary>
    /// <param name="stream">包含 docx 内容的流</param>
    public WordReader(Stream stream)
    {
        _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
    }

    /// <summary>释放资源</summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _zip.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
    #endregion

    #region 读取方法
    /// <summary>读取所有段落文本</summary>
    /// <returns>段落字符串序列</returns>
    public IEnumerable<String> ReadParagraphs()
    {
        var doc = LoadDocumentXml();
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        foreach (XmlElement para in doc.SelectNodes("//w:p", ns)!)
        {
            var sb = new StringBuilder();
            foreach (XmlElement t in para.SelectNodes(".//w:t", ns)!)
            {
                sb.Append(t.InnerText);
            }
            var text = sb.ToString();
            if (text.Length > 0)
                yield return text;
        }
    }

    /// <summary>读取全文（段落间用换行分隔）</summary>
    /// <returns>完整文本</returns>
    public String ReadFullText() => String.Join(Environment.NewLine, ReadParagraphs());

    /// <summary>读取所有表格数据</summary>
    /// <returns>每个表格是 string[][] 的序列</returns>
    public IEnumerable<String[][]> ReadTables()
    {
        var doc = LoadDocumentXml();
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        foreach (XmlElement tbl in doc.SelectNodes("//w:tbl", ns)!)
        {
            var rows = new List<String[]>();
            foreach (XmlElement tr in tbl.SelectNodes("w:tr", ns)!)
            {
                var cells = new List<String>();
                foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
                {
                    var sb = new StringBuilder();
                    foreach (XmlElement t in tc.SelectNodes(".//w:t", ns)!)
                    {
                        sb.Append(t.InnerText);
                    }
                    cells.Add(sb.ToString());
                }
                if (cells.Count > 0)
                    rows.Add(cells.ToArray());
            }
            if (rows.Count > 0)
                yield return rows.ToArray();
        }
    }

    /// <summary>提取所有图片数据</summary>
    /// <returns>（扩展名, 字节数据）序列</returns>
    public IEnumerable<(String Extension, Byte[] Data)> ExtractImages()
    {
        foreach (var entry in _zip.Entries)
        {
            if (!entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                continue;
            var ext = Path.GetExtension(entry.Name).TrimStart('.').ToLowerInvariant();
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            yield return (ext, ms.ToArray());
        }
    }

    /// <summary>获取文档属性</summary>
    /// <returns>属性对象</returns>
    public WordProperties GetProperties()
    {
        var props = new WordProperties();
        var entry = _zip.GetEntry("docProps/core.xml");
        if (entry == null) return props;

        var doc = new XmlDocument();
        using (var s = entry.Open())
            doc.Load(s);

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
        ns.AddNamespace("dcterms", "http://purl.org/dc/terms/");
        ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");

        props.Title = doc.SelectSingleNode("//dc:title", ns)?.InnerText;
        props.Author = doc.SelectSingleNode("//dc:creator", ns)?.InnerText;
        props.Subject = doc.SelectSingleNode("//dc:subject", ns)?.InnerText;
        props.Description = doc.SelectSingleNode("//dc:description", ns)?.InnerText;
        var createdText = doc.SelectSingleNode("//dcterms:created", ns)?.InnerText;
        if (DateTime.TryParse(createdText, out var dt))
            props.Created = dt;

        return props;
    }

    /// <summary>读取对象集合（将第一行表格映射到属性）</summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <returns>对象序列</returns>
    public IEnumerable<T> ReadObjects<T>() where T : class, new()
    {
        var props = typeof(T).GetProperties();
        foreach (var tbl in ReadTables())
        {
            if (tbl.Length < 2) continue;
            var headers = tbl[0];
            for (var ri = 1; ri < tbl.Length; ri++)
            {
                var row = tbl[ri];
                var obj = new T();
                for (var ci = 0; ci < Math.Min(headers.Length, row.Length); ci++)
                {
                    var hdr = headers[ci].Trim();
                    var prop = props.FirstOrDefault(p =>
                        p.Name.Equals(hdr, StringComparison.OrdinalIgnoreCase) ||
                        p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                         .OfType<System.ComponentModel.DisplayNameAttribute>().Any(a => a.DisplayName == hdr));
                    if (prop == null) continue;
                    try
                    {
                        var value = row[ci];
                        if (prop.PropertyType == typeof(String))
                            prop.SetValue(obj, value);
                        else
                            prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
                    }
                    catch { /* skip conversion errors */ }
                }
                yield return obj;
            }
        }
    }
    #endregion

    #region 文档模型读取
    /// <summary>读取完整文档模型（含格式、图片、表格、页面设置等）</summary>
    public WordDocument ReadDocument()
    {
        var doc = new WordDocument();
        var xml = LoadDocumentXml();
        var ns = WmlNs(xml);
        var rels = LoadRels();
        var styleMap = LoadStyles();

        // 保存 document.xml 根元素的命名空间声明，用于 Writer 重建时保展扩命名空间
        if (xml.DocumentElement != null)
        {
            var nsSb = new StringBuilder();
            foreach (XmlAttribute attr in xml.DocumentElement.Attributes)
            {
                if (attr.Name.StartsWith("xmlns", StringComparison.Ordinal))
                    nsSb.Append($" {attr.Name}=\"{attr.Value}\"");
            }
            doc.DocumentXmlNsDecls = nsSb.ToString();
        }

        var body = xml.SelectSingleNode("//w:body", ns);
        if (body != null)
        {
            foreach (XmlNode child in body.ChildNodes)
            {
                if (child is not XmlElement el) continue;

                if (el.LocalName == "p")
                {
                    var drawing = el.SelectSingleNode(".//w:drawing", ns) as XmlElement;
                    var hasText = el.SelectSingleNode(".//w:t", ns) != null;

                    if (!hasText && drawing != null)
                    {
                        // 纯图片段落：用 Image 元素表示（不保留空段落）
                        var ie = ParseDrawing(drawing, ns, rels, doc);
                        if (ie != null) doc.Elements.Add(ie);
                    }
                    else
                    {
                        // 文字段落（可能含内嵌图片）——保存完整 RawXml，图片已包含在其中
                        var pe = ParsePara(el, ns, rels, styleMap);
                        if (pe.Paragraph != null && !IsEmptyPara(pe.Paragraph))
                        {
                            pe.RawXml = el.OuterXml;
                            doc.Elements.Add(pe);
                        }
                        // 加载图片数据到 doc.Images（需要，即使 RawXml 已包含 XML）
                        if (drawing != null)
                            ParseDrawing(drawing, ns, rels, doc); // 将图片加入 Images，不追加到 Elements
                    }
                }
                else if (el.LocalName == "tbl")
                {
                    var te = ParseTable(el, ns, rels, styleMap);
                    te.RawXml = el.OuterXml; // 保存表格原始 XML
                    doc.Elements.Add(te);
                }
                else if (el.LocalName == "sectPr")
                {
                    ParseSectPr(el, ns, doc);
                    doc.SectPrXml = el.OuterXml; // 保存页面设置原始 XML
                }
            }
        }

        foreach (var kv in rels)
        {
            var target = kv.Value;
            if (!target.StartsWith("media/", StringComparison.OrdinalIgnoreCase) || doc.Images.ContainsKey(kv.Key))
                continue;
            var entry = _zip.GetEntry($"word/{target}");
            if (entry == null) continue;
            var ext = Path.GetExtension(target).TrimStart('.').ToLowerInvariant();
            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            doc.Images[kv.Key] = (ext, ms.ToArray());
        }

        foreach (var kv in rels)
        {
            if (kv.Value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                kv.Value.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                doc.Hyperlinks.Add((kv.Key, kv.Value));
        }

        var props = GetProperties();
        doc.DocumentProperties.Title = props.Title;
        doc.DocumentProperties.Author = props.Author;
        doc.DocumentProperties.Subject = props.Subject;
        doc.DocumentProperties.Description = props.Description;

        // 保存原始 XML 部件，用于 Writer 完美还原视觉效果
        doc.StylesXml = ReadZipEntryText("word/styles.xml");
        doc.NumberingXml = ReadZipEntryText("word/numbering.xml");
        doc.SettingsXml = ReadZipEntryText("word/settings.xml");
        doc.DocumentXml = ReadZipEntryText("word/document.xml");

        LoadHdrFtr(rels, doc);

        // 收集所有 ZIP 部件（除 document.xml 外全部透传）
        CollectOtherParts(doc);

        return doc;
    }

    /// <summary>读取 ZIP 入口的文本内容</summary>
    private String? ReadZipEntryText(String entryPath)
    {
        var entry = _zip.GetEntry(entryPath);
        if (entry == null) return null;
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    /// <summary>读取 ZIP 入口的原始字节</summary>
    private Byte[]? ReadZipEntryBytes(String entryPath)
    {
        var entry = _zip.GetEntry(entryPath);
        if (entry == null) return null;
        using var ms = new MemoryStream();
        using var es = entry.Open();
        es.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>收集所有 ZIP 部件（除 word/document.xml 外）到 OtherParts，用于透传模式保真</summary>
    private void CollectOtherParts(WordDocument doc)
    {
        foreach (var entry in _zip.Entries)
        {
            var name = entry.FullName;
            if (name.EndsWith("/")) continue; // 目录条目
            // 仅排除 word/document.xml —— 它是唯一需要重新生成的部件
            if (name.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase)) continue;

            using var ms = new MemoryStream();
            using var es = entry.Open();
            es.CopyTo(ms);
            doc.OtherParts[name] = ms.ToArray();
        }
    }
    #endregion

    #region 解析辅助
    private static readonly String Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly String Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly String WPns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly String Ans = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static XmlNamespaceManager WmlNs(XmlDocument doc)
    {
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", Wns);
        ns.AddNamespace("r", Rns);
        ns.AddNamespace("wp", WPns);
        ns.AddNamespace("a", Ans);
        return ns;
    }

    private XmlDocument LoadDocumentXml()
    {
        var entry = _zip.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("无效的 docx 文件：缺少 word/document.xml");
        var doc = new XmlDocument();
        using var s = entry.Open();
        doc.Load(s);
        return doc;
    }

    private Dictionary<String, String> LoadRels()
    {
        var map = new Dictionary<String, String>();
        var entry = _zip.GetEntry("word/_rels/document.xml.rels");
        if (entry == null) return map;
        var doc = new XmlDocument();
        using (var s = entry.Open()) doc.Load(s);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (XmlElement rel in doc.SelectNodes("//rel:Relationship", ns)!)
        {
            var id = rel.GetAttribute("Id");
            var target = rel.GetAttribute("Target");
            if (id != null && target != null) map[id] = target;
        }
        return map;
    }

    private Dictionary<String, WordRunProperties> LoadStyles()
    {
        var map = new Dictionary<String, WordRunProperties>(StringComparer.OrdinalIgnoreCase);
        var entry = _zip.GetEntry("word/styles.xml");
        if (entry == null) return map;
        var doc = new XmlDocument();
        using (var s = entry.Open()) doc.Load(s);
        var ns = WmlNs(doc);
        foreach (XmlElement st in doc.SelectNodes("//w:style", ns)!)
        {
            var styleId = st.GetAttribute("w:styleId");
            if (String.IsNullOrEmpty(styleId)) continue;
            var rPr = st.SelectSingleNode("w:rPr", ns) as XmlElement;
            if (rPr == null)
            {
                var pPr = st.SelectSingleNode("w:pPr", ns) as XmlElement;
                if (pPr != null) rPr = pPr.SelectSingleNode("w:rPr", ns) as XmlElement;
            }
            if (rPr != null) map[styleId] = ParseRunPr(rPr, ns, false);
        }
        return map;
    }
    #endregion

    #region 段落解析
    private WordElement ParsePara(XmlElement pEl, XmlNamespaceManager ns, Dictionary<String, String> rels, Dictionary<String, WordRunProperties> styleMap)
    {
        var para = new WordParagraph();
        var isPageBreak = false;
        var isBullet = false;

        var bmStart = pEl.SelectSingleNode("w:bookmarkStart", ns);
        if (bmStart is XmlElement bmEl)
        {
            para.BookmarkName = bmEl.GetAttribute("w:name");
            if (para.BookmarkName == null) para.BookmarkName = bmEl.GetAttribute("name");
        }

        var pPr = pEl.SelectSingleNode("w:pPr", ns);
        WordRunProperties? styleDefaults = null;
        if (pPr is XmlElement pPrEl)
        {
            var pStyle = pPrEl.SelectSingleNode("w:pStyle", ns) as XmlElement;
            if (pStyle != null)
            {
                var styleId = pStyle.GetAttribute("w:val");
                if (styleId == null) styleId = pStyle.GetAttribute("val");
                para.StyleId = styleId;
                para.Style = ParseStyleId(styleId);
                if (styleId != null) styleMap.TryGetValue(styleId, out styleDefaults);
            }

            var jc = pPrEl.SelectSingleNode("w:jc", ns) as XmlElement;
            if (jc != null)
            {
                para.Alignment = jc.GetAttribute("w:val");
                if (para.Alignment == null) para.Alignment = jc.GetAttribute("val");
            }

            var shd = pPrEl.SelectSingleNode("w:shd", ns) as XmlElement;
            if (shd != null)
                para.BackgroundColor = shd.GetAttribute("w:fill") ?? shd.GetAttribute("fill");

            var ind = pPrEl.SelectSingleNode("w:ind", ns) as XmlElement;
            if (ind != null)
            {
                var left = ind.GetAttribute("w:left") ?? ind.GetAttribute("left");
                if (Int32.TryParse(left, out var lv)) para.IndentLeft = lv;
                var right = ind.GetAttribute("w:right") ?? ind.GetAttribute("right");
                if (Int32.TryParse(right, out var rv)) para.IndentRight = rv;
                var firstLine = ind.GetAttribute("w:firstLine") ?? ind.GetAttribute("firstLine");
                if (Int32.TryParse(firstLine, out var fl)) para.FirstLineIndent = fl;
                else
                {
                    var hanging = ind.GetAttribute("w:hanging") ?? ind.GetAttribute("hanging");
                    if (Int32.TryParse(hanging, out var hg)) para.FirstLineIndent = -hg;
                }
            }

            var spacing = pPrEl.SelectSingleNode("w:spacing", ns) as XmlElement;
            if (spacing != null)
            {
                var before = spacing.GetAttribute("w:before") ?? spacing.GetAttribute("before");
                if (Int32.TryParse(before, out var bv)) para.SpaceBefore = bv;
                var after = spacing.GetAttribute("w:after") ?? spacing.GetAttribute("after");
                if (Int32.TryParse(after, out var av)) para.SpaceAfter = av;
                var line = spacing.GetAttribute("w:line") ?? spacing.GetAttribute("line");
                var lineRule = spacing.GetAttribute("w:lineRule") ?? spacing.GetAttribute("lineRule");
                if (Int32.TryParse(line, out var lv) && lineRule == "auto")
                    para.LineSpacingPct = lv * 100 / 240;
            }

            if (pPrEl.SelectSingleNode("w:numPr", ns) != null)
                isBullet = true;
        }

        var br = pEl.SelectSingleNode("w:r/w:br", ns) as XmlElement;
        if (br != null)
        {
            var brType = br.GetAttribute("w:type") ?? br.GetAttribute("type");
            if (brType == "page") isPageBreak = true;
        }

        para.IsBullet = isBullet;
        para.IsPageBreak = isPageBreak;

        if (!isPageBreak)
        {
            foreach (XmlNode child in pEl.ChildNodes)
            {
                if (child is not XmlElement el) continue;
                if (el.LocalName == "r")
                {
                    var run = ParseRun(el, ns, null);
                    ApplyStyleDefaults(run, styleDefaults);
                    para.Runs.Add(run);
                }
                else if (el.LocalName == "hyperlink")
                {
                    var hlRelId = el.GetAttribute("r:id");
                    if (hlRelId == null) hlRelId = el.GetAttribute("id");
                    foreach (XmlNode hlChild in el.ChildNodes)
                    {
                        if (hlChild is XmlElement hlRun && hlRun.LocalName == "r")
                        {
                            var run = ParseRun(hlRun, ns, hlRelId);
                            ApplyStyleDefaults(run, styleDefaults);
                            para.Runs.Add(run);
                        }
                    }
                }
            }
        }

        return new WordElement { Type = WordElementType.Paragraph, Paragraph = para };
    }

    private static WordRun ParseRun(XmlElement rEl, XmlNamespaceManager ns, String? hyperlinkRelId)
    {
        var run = new WordRun { HyperlinkRelId = hyperlinkRelId };
        var rPr = rEl.SelectSingleNode("w:rPr", ns) as XmlElement;
        if (rPr != null)
            run.Properties = ParseRunPr(rPr, ns, hyperlinkRelId != null);
        var tEl = rEl.SelectSingleNode("w:t", ns);
        if (tEl != null) run.Text = tEl.InnerText;
        return run;
    }

    private static WordRunProperties ParseRunPr(XmlElement rPrEl, XmlNamespaceManager ns, Boolean isHyperlink)
    {
        var p = new WordRunProperties();
        if (rPrEl.SelectSingleNode("w:b", ns) != null) p.Bold = true;
        if (rPrEl.SelectSingleNode("w:i", ns) != null) p.Italic = true;
        if (rPrEl.SelectSingleNode("w:u", ns) != null) p.Underline = true;
        var color = rPrEl.SelectSingleNode("w:color", ns) as XmlElement;
        if (color != null)
        {
            var val = color.GetAttribute("w:val") ?? color.GetAttribute("val");
            if (val != null && val != "auto") p.ForeColor = val;
        }
        var sz = rPrEl.SelectSingleNode("w:sz", ns) as XmlElement;
        if (sz != null)
        {
            var val = sz.GetAttribute("w:val") ?? sz.GetAttribute("val");
            if (Single.TryParse(val, out var sv)) p.FontSize = sv / 2f;
        }
        var rFonts = rPrEl.SelectSingleNode("w:rFonts", ns) as XmlElement;
        if (rFonts != null)
            p.FontName = rFonts.GetAttribute("w:ascii") ?? rFonts.GetAttribute("ascii")
                ?? rFonts.GetAttribute("w:hAnsi") ?? rFonts.GetAttribute("hAnsi");
        if (isHyperlink && p.ForeColor == null)
        {
            p.ForeColor = "0563C1";
            p.Underline = true;
        }
        return p;
    }

    private static void ApplyStyleDefaults(WordRun run, WordRunProperties? defaults)
    {
        if (defaults == null) return;
        var rp = run.Properties;
        if (rp == null)
        {
            var np = new WordRunProperties();
            if (defaults.Bold) np.Bold = true;
            if (defaults.Italic) np.Italic = true;
            if (defaults.Underline) np.Underline = true;
            if (defaults.ForeColor != null) np.ForeColor = defaults.ForeColor;
            if (defaults.FontSize.HasValue) np.FontSize = defaults.FontSize;
            if (defaults.FontName != null) np.FontName = defaults.FontName;
            if (np.Bold || np.Italic || np.Underline || np.ForeColor != null || np.FontSize.HasValue || np.FontName != null)
                run.Properties = np;
        }
        else
        {
            if (!rp.Bold && defaults.Bold) rp.Bold = true;
            if (!rp.Italic && defaults.Italic) rp.Italic = true;
            if (!rp.Underline && defaults.Underline) rp.Underline = true;
            if (rp.ForeColor == null) rp.ForeColor = defaults.ForeColor;
            if (rp.FontSize == null) rp.FontSize = defaults.FontSize;
            if (rp.FontName == null) rp.FontName = defaults.FontName;
        }
    }
    #endregion

    #region 表格/图片/节属性
    private WordElement ParseTable(XmlElement tblEl, XmlNamespaceManager ns, Dictionary<String, String> rels, Dictionary<String, WordRunProperties> styleMap)
    {
        var rows = new List<List<WordCell>>();
        var firstRowHeader = false;
        var style = new WordTableStyle();

        var tblPr = tblEl.SelectSingleNode("w:tblPr", ns) as XmlElement;
        if (tblPr != null)
        {
            var borders = tblPr.SelectSingleNode("w:tblBorders", ns) as XmlElement;
            if (borders != null)
            {
                var topB = borders.SelectSingleNode("w:top", ns) as XmlElement;
                if (topB != null)
                {
                    style.BorderColor = topB.GetAttribute("w:color") ?? topB.GetAttribute("color") ?? "000000";
                    var szVal = topB.GetAttribute("w:sz") ?? topB.GetAttribute("sz");
                    if (Int32.TryParse(szVal, out var bsz)) style.BorderSize = bsz;
                }
            }
        }

        var tblGrid = tblEl.SelectSingleNode("w:tblGrid", ns) as XmlElement;
        var colWidths = new List<Int32>();
        if (tblGrid != null)
        {
            foreach (XmlElement gc in tblGrid.SelectNodes("w:gridCol", ns)!)
            {
                var w = gc.GetAttribute("w:w") ?? gc.GetAttribute("w");
                if (Int32.TryParse(w, out var cw)) colWidths.Add(cw);
            }
            if (colWidths.Count > 0) style.ColumnWidths = colWidths.ToArray();
        }

        foreach (XmlElement tr in tblEl.SelectNodes("w:tr", ns)!)
        {
            var cells = new List<WordCell>();
            var trPr = tr.SelectSingleNode("w:trPr", ns);
            if (trPr != null && trPr.SelectSingleNode("w:tblHeader", ns) != null)
                firstRowHeader = true;
            foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
                cells.Add(ParseCell(tc, ns));
            if (cells.Count > 0) rows.Add(cells);
        }

        return new WordElement
        {
            Type = WordElementType.Table,
            TableRows = rows,
            TableFirstRowHeader = firstRowHeader,
            TableStyle = style,
        };
    }

    private static WordCell ParseCell(XmlElement tcEl, XmlNamespaceManager ns)
    {
        var cell = new WordCell();
        var tcPr = tcEl.SelectSingleNode("w:tcPr", ns) as XmlElement;
        if (tcPr != null)
        {
            var shd = tcPr.SelectSingleNode("w:shd", ns) as XmlElement;
            if (shd != null)
                cell.BackgroundColor = shd.GetAttribute("w:fill") ?? shd.GetAttribute("fill");
            var gridSpan = tcPr.SelectSingleNode("w:gridSpan", ns) as XmlElement;
            if (gridSpan != null)
            {
                var val = gridSpan.GetAttribute("w:val") ?? gridSpan.GetAttribute("val");
                if (Int32.TryParse(val, out var gs)) cell.ColSpan = gs;
            }
            if (tcPr.SelectSingleNode("w:vMerge", ns) != null) cell.RowSpan = 0;
        }

        foreach (XmlElement pEl in tcEl.SelectNodes("w:p", ns)!)
        {
            var para = new WordParagraph();
            foreach (XmlNode child in pEl.ChildNodes)
            {
                if (child is XmlElement el && el.LocalName == "r")
                    para.Runs.Add(ParseRun(el, ns, null));
            }
            var pPr = pEl.SelectSingleNode("w:pPr", ns) as XmlElement;
            if (pPr != null)
            {
                var jc = pPr.SelectSingleNode("w:jc", ns) as XmlElement;
                if (jc != null) para.Alignment = jc.GetAttribute("w:val") ?? jc.GetAttribute("val");
            }
            cell.Paragraphs.Add(para);
        }
        return cell;
    }

    private WordElement? ParseDrawing(XmlElement drawing, XmlNamespaceManager ns, Dictionary<String, String> rels, WordDocument doc)
    {
        var inline = drawing.SelectSingleNode("wp:inline", ns) as XmlElement
            ?? drawing.SelectSingleNode("wp:anchor", ns) as XmlElement;
        if (inline == null) return null;

        var extent = inline.SelectSingleNode("wp:extent", ns) as XmlElement;
        var cx = 0L;
        var cy = 0L;
        if (extent != null)
        {
            Int64.TryParse(extent.GetAttribute("cx"), out cx);
            Int64.TryParse(extent.GetAttribute("cy"), out cy);
        }

        var blip = drawing.SelectSingleNode(".//a:blip", ns) as XmlElement;
        var rId = blip?.GetAttribute("r:embed");
        if (rId == null) return null;

        if (!doc.Images.ContainsKey(rId) && rels.TryGetValue(rId, out var target))
        {
            var entry = _zip.GetEntry($"word/{target}");
            if (entry != null)
            {
                var ext = Path.GetExtension(target).TrimStart('.').ToLowerInvariant();
                using var ms = new MemoryStream();
                using var es = entry.Open();
                es.CopyTo(ms);
                doc.Images[rId] = (ext, ms.ToArray());
            }
        }

        return new WordElement
        {
            Type = WordElementType.Image,
            Image = new WordImageElement
            {
                RelId = rId,
                WidthEmu = cx > 0 ? cx : 3600000,
                HeightEmu = cy > 0 ? cy : 2700000,
                Extension = doc.Images.TryGetValue(rId, out var img) ? img.Extension : "png",
            },
        };
    }

    private static void ParseSectPr(XmlElement sectPr, XmlNamespaceManager ns, WordDocument doc)
    {
        var ps = doc.PageSettings;
        var pgSz = sectPr.SelectSingleNode("w:pgSz", ns) as XmlElement;
        if (pgSz != null)
        {
            if (Int32.TryParse(pgSz.GetAttribute("w:w") ?? pgSz.GetAttribute("w"), out var pw)) ps.PageWidth = pw;
            if (Int32.TryParse(pgSz.GetAttribute("w:h") ?? pgSz.GetAttribute("h"), out var ph)) ps.PageHeight = ph;
            if ((pgSz.GetAttribute("w:orient") ?? pgSz.GetAttribute("orient")) == "landscape") ps.Landscape = true;
        }

        var pgMar = sectPr.SelectSingleNode("w:pgMar", ns) as XmlElement;
        if (pgMar != null)
        {
            if (Int32.TryParse(pgMar.GetAttribute("w:top") ?? pgMar.GetAttribute("top"), out var tv)) ps.MarginTop = tv;
            if (Int32.TryParse(pgMar.GetAttribute("w:right") ?? pgMar.GetAttribute("right"), out var rv)) ps.MarginRight = rv;
            if (Int32.TryParse(pgMar.GetAttribute("w:bottom") ?? pgMar.GetAttribute("bottom"), out var bv)) ps.MarginBottom = bv;
            if (Int32.TryParse(pgMar.GetAttribute("w:left") ?? pgMar.GetAttribute("left"), out var lv)) ps.MarginLeft = lv;
        }

        var hdrRef = sectPr.SelectSingleNode("w:headerReference", ns) as XmlElement;
        if (hdrRef != null) { var rId = hdrRef.GetAttribute("r:id"); if (rId != null) ps.HeaderText = rId; }

        var ftrRef = sectPr.SelectSingleNode("w:footerReference", ns) as XmlElement;
        if (ftrRef != null) { var rId = ftrRef.GetAttribute("r:id"); if (rId != null) ps.FooterText = rId; }
    }

    private void LoadHdrFtr(Dictionary<String, String> rels, WordDocument doc)
    {
        var headerRId = doc.PageSettings.HeaderText;
        if (headerRId != null && rels.TryGetValue(headerRId, out var headerTarget))
        {
            var text = LoadHdrFtrText(headerTarget, doc);
            if (text != null) doc.HeaderText = text;
        }

        var footerRId = doc.PageSettings.FooterText;
        if (footerRId != null && rels.TryGetValue(footerRId, out var footerTarget))
        {
            var text = LoadHdrFtrText(footerTarget, doc);
            if (text != null) doc.FooterText = text;
        }

        doc.PageSettings.HeaderText = doc.HeaderText;
        doc.PageSettings.FooterText = doc.FooterText;
    }

    private String? LoadHdrFtrText(String target, WordDocument doc)
    {
        var entry = _zip.GetEntry($"word/{target}");
        if (entry == null) return null;

        var xml = new XmlDocument();
        using (var s = entry.Open()) xml.Load(s);
        var ns = WmlNs(xml);

        var sb = new StringBuilder();
        foreach (XmlElement t in xml.SelectNodes("//w:t", ns)!) sb.Append(t.InnerText);

        var relsName = Path.GetFileName(target);
        var relsEntry = _zip.GetEntry($"word/_rels/{relsName}.rels");
        var hfRels = new Dictionary<String, String>();
        if (relsEntry != null)
        {
            var relsDoc = new XmlDocument();
            using (var rs = relsEntry.Open()) relsDoc.Load(rs);
            var relsNs = new XmlNamespaceManager(relsDoc.NameTable);
            relsNs.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
            foreach (XmlElement rel in relsDoc.SelectNodes("//rel:Relationship", relsNs)!)
            {
                var id = rel.GetAttribute("Id");
                var tgt = rel.GetAttribute("Target");
                if (id != null && tgt != null) hfRels[id] = tgt;
            }
        }

        foreach (XmlElement drawing in xml.SelectNodes("//w:drawing", ns)!)
        {
            var blip = drawing.SelectSingleNode(".//a:blip", ns) as XmlElement;
            var rId = blip?.GetAttribute("r:embed");
            if (rId == null || !hfRels.TryGetValue(rId, out var imgTarget)) continue;
            if (doc.Images.ContainsKey(rId)) continue;
            var imgEntry = _zip.GetEntry($"word/{imgTarget}");
            if (imgEntry == null) continue;
            var ext = Path.GetExtension(imgTarget).TrimStart('.').ToLowerInvariant();
            using var ms = new MemoryStream();
            using var es = imgEntry.Open();
            es.CopyTo(ms);
            doc.Images[rId] = (ext, ms.ToArray());
        }

        return sb.Length > 0 ? sb.ToString() : null;
    }

    private static Boolean IsEmptyPara(WordParagraph p)
    {
        return !p.IsPageBreak && !p.IsBullet && p.BookmarkName == null && p.Runs.Count == 0;
    }

    private static WordParagraphStyle ParseStyleId(String? styleId)
    {
        if (String.IsNullOrEmpty(styleId)) return WordParagraphStyle.Normal;
        return styleId.ToLowerInvariant() switch
        {
            "heading1" or "1" or "heading 1" => WordParagraphStyle.Heading1,
            "heading2" or "2" or "heading 2" => WordParagraphStyle.Heading2,
            "heading3" or "3" or "heading 3" => WordParagraphStyle.Heading3,
            "heading4" or "4" or "heading 4" => WordParagraphStyle.Heading4,
            "heading5" or "5" or "heading 5" => WordParagraphStyle.Heading5,
            "heading6" or "6" or "heading 6" => WordParagraphStyle.Heading6,
            _ => WordParagraphStyle.Normal,
        };
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（段落间换行分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText() => ReadFullText();

    /// <summary>提取 Markdown 格式（段落+表格）</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown()
    {
        var sb = new StringBuilder();

        // 输出段落
        foreach (var para in ReadParagraphs())
        {
            sb.AppendLine(para);
            sb.AppendLine();
        }

        // 输出表格
        foreach (var table in ReadTables())
        {
            if (table.Length == 0) continue;

            // 第一行作为表头
            var header = table[0];
            sb.Append('|');
            foreach (var cell in header)
            {
                sb.Append(' ').Append(MdEscape(cell)).Append(" |");
            }
            sb.AppendLine();

            sb.Append('|');
            for (var i = 0; i < header.Length; i++)
            {
                sb.Append(" --- |");
            }
            sb.AppendLine();

            for (var ri = 1; ri < table.Length; ri++)
            {
                var row = table[ri];
                sb.Append('|');
                for (var i = 0; i < header.Length; i++)
                {
                    var val = i < row.Length ? row[i] : "";
                    sb.Append(' ').Append(MdEscape(val)).Append(" |");
                }
                sb.AppendLine();
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static String MdEscape(String? value)
    {
        if (String.IsNullOrEmpty(value)) return "";
        return value.Replace("|", "\\|").Replace("\n", " ").Replace("\r", "");
    }
    #endregion
}
