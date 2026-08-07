using System.IO.Compression;
using System.Text;
using System.Xml;
using NewLife.Office.Markdown;

namespace NewLife.Office.Word;

/// <summary>Word docx 读取器</summary>
/// <remarks>
/// 直接解析 Open XML（ZIP+XML）提取文本、表格、图片等内容。
/// </remarks>
public class WordReader : IDisposable, ITextExtractable, IMarkdownExtractable, IMarkdownDocumentExtractable
{
    #region 属性
    /// <summary>源文件路径（从文件构造时有效）</summary>
    public String? FilePath { get; private set; }
    #endregion

    #region 私有字段
    private readonly ZipArchive _zip;
    private Boolean _disposed;

    // Markdown AST 转换状态（列表聚合缓冲 + 图片计数）
    private readonly List<MarkdownBlock> _bulletItems = [];
    private readonly List<(Int32 Num, List<MarkdownInline> Inlines)> _orderedItems = [];
    private Int32 _orderedNum;
    private Int32 _imageIdx;
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
    /// <remarks>
    /// 排除文本框（w:txbxContent）内层段落——它们由外层形状段落的文本提取一并覆盖，
    /// 避免文本框文本重复出现。w:tab→\t、w:br→换行、w:cr→换行，域指令文本不泄漏。
    /// </remarks>
    public IEnumerable<String> ReadParagraphs()
    {
        var doc = LoadDocumentXml();
        const String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", W);

        foreach (XmlElement para in doc.SelectNodes("//w:p[not(ancestor::w:txbxContent)]", ns)!)
        {
            var sb = new StringBuilder();
            AppendParaText(para, ns, sb);
            var text = sb.ToString();
            if (text.Length > 0)
                yield return text;
        }
    }

    /// <summary>按文档顺序拼接段落文本，处理 w:t/w:tab/w:br/w:cr 语义；递归下钻到超链接/文本框等容器</summary>
    private static void AppendParaText(XmlElement el, XmlNamespaceManager ns, StringBuilder sb)
    {
        foreach (XmlNode child in el.ChildNodes)
        {
            if (child is not XmlElement c) continue;
            switch (c.LocalName)
            {
                case "t":
                    sb.Append(c.InnerText);
                    break;
                case "tab":
                    sb.Append('\t');
                    break;
                case "br":
                    var bt = c.GetAttribute("w:type") ?? c.GetAttribute("type");
                    sb.Append(bt == "page" ? '\f' : bt == "column" ? '\v' : '\n');
                    break;
                case "cr":
                    sb.Append('\n');
                    break;
                case "noBreakHyphen":
                    sb.Append('-');
                    break;
                case "softHyphen":
                    sb.Append('\u00AD');
                    break;
                case "instrText":   // 域指令（如 MERGEFIELD/PAGE）不泄漏到纯文本
                case "delText":     // 修订删除文本不提取
                case "fldChar":
                case "footnoteReference":
                case "endnoteReference":
                case "commentReference":
                    break;
                default:
                    AppendParaText(c, ns, sb);
                    break;
            }
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

        // 仅提取顶层表格（排除嵌套在 w:tc 内的表格，避免重复）
        foreach (XmlElement tbl in doc.SelectNodes("//w:tbl[not(ancestor::w:tc)]", ns)!)
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
    public Properties GetProperties()
    {
        var props = new Properties();
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

    /// <summary>读取自定义文档属性（docProps/custom.xml）</summary>
    private void ReadCustomProperties(DocumentProperties props)
    {
        var entry = _zip.GetEntry("docProps/custom.xml");
        if (entry == null) return;

        try
        {
            var doc = new XmlDocument();
            using var s = entry.Open();
            doc.Load(s);

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            var propNodes = doc.SelectNodes("//*[local-name()='property']");
            if (propNodes == null) return;

            foreach (XmlElement propEl in propNodes)
            {
                var name = propEl.GetAttribute("name");
                if (String.IsNullOrEmpty(name)) continue;

                String? type = "lpwstr";
                String? value = null;

                // 检查所有可能的 vt 子元素类型
                var vtNodes = propEl.SelectNodes("*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes']");
                if (vtNodes == null || vtNodes.Count == 0) continue;

                var vtEl = vtNodes[0] as XmlElement;
                if (vtEl == null) continue;

                type = vtEl.LocalName;
                value = vtEl.InnerText;

                if (value != null)
                    props.CustomProperties[name] = (value, type);
            }
        }
        catch { /* custom.xml 损坏时静默跳过 */ }
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
    public Document ReadDocument()
    {
        var doc = new Document();
        var xml = LoadDocumentXml();
        var ns = WmlNs(xml);
        var rels = LoadRels();
        var styleMap = LoadStyles();
        var numberingXml = ReadZipEntryText("word/numbering.xml");
        var numIdFormats = ParseNumIdFormats(numberingXml);

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
            // 按 w:sectPr 切分节：段内嵌 sectPr 结束当前节（该段是本节最后一段），body 级 sectPr 属于最后一节
            var currentSection = new Section();
            foreach (XmlNode child in body.ChildNodes)
            {
                if (child is not XmlElement el) continue;

                var embeddedSectPr = el.LocalName == "p"
                    ? (el.SelectSingleNode("w:pPr/w:sectPr", ns) as XmlElement)
                    : null;

                var before = doc.Elements.Count;
                ProcessBodyElement(el, ns, rels, styleMap, numIdFormats, doc);

                // 将新增元素同步到当前节（doc.Elements 保持全文档平铺视图）
                if (doc.Elements.Count > before)
                    currentSection.Elements.AddRange(doc.Elements.GetRange(before, doc.Elements.Count - before));

                // 段内嵌 sectPr：解析本节属性并提交
                if (embeddedSectPr != null)
                {
                    ParseSectPr(embeddedSectPr, ns, currentSection.PageSettings);
                    currentSection.SectPrXml = embeddedSectPr.OuterXml;
                    doc.Sections.Add(currentSection);
                    currentSection = new Section();
                }
            }

            // body 级 sectPr（最后一节属性，ProcessBodyElement 已设置 doc.SectPrXml/doc.PageSettings）
            var finalSectPr = body.SelectSingleNode("w:sectPr", ns) as XmlElement;
            if (finalSectPr != null)
            {
                ParseSectPr(finalSectPr, ns, currentSection.PageSettings);
                currentSection.SectPrXml = finalSectPr.OuterXml;
            }
            if (currentSection.Elements.Count > 0 || doc.Sections.Count == 0)
                doc.Sections.Add(currentSection);
        }

        // 修订追踪（W11）：收集全文 w:ins/w:del 记录
        ParseTrackChanges(doc, xml, ns);

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

        // 读取自定义属性
        ReadCustomProperties(doc.DocumentProperties);

        // 保存原始 XML 部件，用于 Writer 完美还原视觉效果
        doc.StylesXml = ReadZipEntryText("word/styles.xml");
        doc.NumberingXml = numberingXml;
        // 解析编号定义到模型（无损往返时 NumberingXml 已足够，模型解析便于程序化修改）
        if (doc.NumberingXml != null)
            doc.Numbering = ParseNumbering(doc.NumberingXml);
        doc.SettingsXml = ReadZipEntryText("word/settings.xml");
        // 解析文档变量
        if (doc.SettingsXml != null)
            ParseDocumentVariables(doc.SettingsXml, doc.DocumentVariables);
        doc.DocumentXml = ReadZipEntryText("word/document.xml");

        LoadHdrFtr(rels, styleMap, numIdFormats, doc);

        // 脚注/尾注（W22）：解析 footnotes.xml/endnotes.xml 到模型（不含内置分隔符标记）
        ParseNotes(ReadZipEntryText("word/footnotes.xml"), doc.Footnotes);
        ParseNotes(ReadZipEntryText("word/endnotes.xml"), doc.Endnotes);

        // 收集所有 ZIP 部件（除 document.xml 外全部透传）
        CollectOtherParts(doc);

        return doc;
    }

    /// <summary>解析脚注/尾注 XML 到模型（w:footnote/w:endnote，跳过内置分隔符/续接标记）</summary>
    /// <param name="xml">footnotes.xml 或 endnotes.xml 内容</param>
    /// <param name="target">目标列表（Footnotes 或 Endnotes）</param>
    private static void ParseNotes(String? xml, List<Footnote> target)
    {
        if (String.IsNullOrEmpty(xml)) return;
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", Wns);
            foreach (XmlElement fn in doc.SelectNodes("//w:footnote | //w:endnote", ns)!)
            {
                var type = fn.GetAttribute("w:type");
                if (String.IsNullOrEmpty(type)) type = fn.GetAttribute("type");
                // 内置标记脚注（分隔符/续接分隔符/续接提示）不建模
                if (type is "separator" or "continuationSeparator" or "continuationNotice") continue;
                if (!Int32.TryParse(fn.GetAttribute("w:id") ?? fn.GetAttribute("id"), out var id)) continue;

                var note = new Footnote { Id = id, Type = String.IsNullOrEmpty(type) ? null : type };
                var sb = new StringBuilder();
                foreach (XmlElement p in fn.SelectNodes("w:p", ns)!)
                {
                    var para = new Paragraph();
                    foreach (XmlElement r in p.SelectNodes("w:r", ns)!)
                        para.Runs.Add(ParseRun(r, ns, null));
                    if (para.Runs.Count > 0)
                    {
                        var paraText = new StringBuilder();
                        foreach (var r in para.Runs) paraText.Append(r.Text);
                        if (paraText.Length > 0)
                        {
                            if (sb.Length > 0) sb.Append('\n');
                            sb.Append(paraText);
                        }
                        note.Paragraphs.Add(para);
                    }
                }
                note.Text = sb.ToString();
                target.Add(note);
            }
        }
        catch { /* 脚注 XML 损坏时静默跳过 */ }
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
    private void CollectOtherParts(Document doc)
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
            var data = ms.ToArray();
            doc.OtherParts[name] = data;

            // 收集自定义 XML 部件到专用集合
            if (name.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                var partName = name.Substring("customXml/".Length);
                doc.CustomXmlParts[partName] = data;
            }
        }
    }
    #endregion

    #region 解析辅助
    private static readonly String Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly String Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly String WPns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly String Ans = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>处理 body 下的单个块级元素（段落/表格/内容控件/修订包裹块/节设置）</summary>
    private void ProcessBodyElement(XmlElement el, XmlNamespaceManager ns, Dictionary<String, String> rels,
        Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats, Document doc)
    {
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
                // 保留空段落（含书签/分页符的段落），保证"读入→写出"结构完整（空行/占位段落不丢失）
                var pe = ParsePara(el, ns, rels, styleMap, numIdFormats);
                pe.RawXml = el.OuterXml;
                doc.Elements.Add(pe);
                // 加载图片数据到 doc.Images（需要，即使 RawXml 已包含 XML）
                if (drawing != null)
                    ParseDrawing(drawing, ns, rels, doc); // 将图片加入 Images，不追加到 Elements
            }
        }
        else if (el.LocalName == "tbl")
        {
            var te = ParseTable(el, ns, rels, styleMap, numIdFormats, doc);
            te.RawXml = el.OuterXml; // 保存表格原始 XML
            doc.Elements.Add(te);
        }
        else if (el.LocalName == "sdt")
        {
            var se = ParseSdt(el, ns);
            if (se != null)
            {
                se.RawXml = el.OuterXml;
                doc.SdtElements.Add(se);
                doc.Elements.Add(new Element { Type = ElementType.Sdt, Sdt = se, RawXml = el.OuterXml });
            }
        }
        else if (el.LocalName == "sectPr")
        {
            ParseSectPr(el, ns, doc.PageSettings);
            doc.SectPrXml = el.OuterXml; // 保存页面设置原始 XML
        }
        else if (el.LocalName is "ins" or "del")
        {
            // 段落级修订（w:ins/w:del 包裹整个段落）：递归处理子块，修订记录由 ParseTrackChanges 统一收集
            foreach (XmlNode inner in el.ChildNodes)
            {
                if (inner is XmlElement innerEl)
                    ProcessBodyElement(innerEl, ns, rels, styleMap, numIdFormats, doc);
            }
        }
    }

    /// <summary>收集全文修订追踪记录（w:ins 插入 / w:del 删除，W11）</summary>
    /// <param name="doc">文档对象（填充 TrackChanges）</param>
    /// <param name="xml">document.xml</param>
    /// <param name="ns">命名空间管理器</param>
    private void ParseTrackChanges(Document doc, XmlDocument xml, XmlNamespaceManager ns)
    {
        var changes = new List<TrackChange>();
        foreach (XmlElement el in xml.SelectNodes("//w:ins | //w:del", ns)!)
        {
            var change = new TrackChange
            {
                Type = el.LocalName == "ins" ? TrackChangeType.Insert : TrackChangeType.Delete,
                Author = el.GetAttribute("w:author") ?? "",
                Date = el.GetAttribute("w:date"),
                Id = el.GetAttribute("w:id"),
            };

            // 修订文本：插入取 w:t，删除取 w:delText
            var sb = new StringBuilder();
            foreach (XmlElement t in el.SelectNodes(".//w:t | .//w:delText", ns)!)
                sb.Append(t.InnerText);
            change.Text = sb.ToString();

            // 所在段落纯文本上下文（含 w:delText，反映修订前后全貌）
            var para = el.SelectSingleNode("ancestor::w:p[1]", ns) as XmlElement;
            if (para != null)
            {
                var ctx = new StringBuilder();
                foreach (XmlElement t in para.SelectNodes(".//w:t | .//w:delText", ns)!)
                    ctx.Append(t.InnerText);
                change.ParagraphText = ctx.ToString();
            }

            changes.Add(change);
        }
        doc.TrackChanges = changes;
    }

    private static XmlNamespaceManager WmlNs(XmlDocument doc)
    {
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("w", Wns);
        ns.AddNamespace("r", Rns);
        ns.AddNamespace("wp", WPns);
        ns.AddNamespace("a", Ans);
        ns.AddNamespace("v", "urn:schemas-microsoft-com:vml");
        ns.AddNamespace("asvg", "http://schemas.microsoft.com/office/drawing/2016/SVG/main");
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

    /// <summary>样式信息（含字符格式 RPr 与段落默认格式 PPr），用于把样式默认应用到段落/文字模型</summary>
    private sealed class WordStyle
    {
        /// <summary>字符格式（展开继承后）</summary>
        public RunProperties? RPr { get; set; }

        /// <summary>基础样式 ID（基于关系）</summary>
        public String? BasedOn { get; set; }

        // ── 段落默认（来自样式 w:pPr，段落未显式设置时应用）──
        public String? Alignment { get; set; }
        public Int32? IndentLeft { get; set; }
        public Int32? IndentRight { get; set; }
        public Int32? FirstLineIndent { get; set; }
        public Int32? SpaceBefore { get; set; }
        public Int32? SpaceAfter { get; set; }
        public Int32? LineSpacingPct { get; set; }
        public String? BackgroundColor { get; set; }
    }

    /// <summary>加载样式表并展开继承链（docDefaults + basedOn），返回 styleId → 合并后的样式信息</summary>
    private Dictionary<String, WordStyle> LoadStyles()
    {
        var map = new Dictionary<String, WordStyle>(StringComparer.OrdinalIgnoreCase);
        var entry = _zip.GetEntry("word/styles.xml");
        if (entry == null) return map;
        var doc = new XmlDocument();
        using (var s = entry.Open()) doc.Load(s);
        var ns = WmlNs(doc);

        // 文档默认格式（docDefaults/rPrDefault），所有样式的基础
        var defaults = new RunProperties();
        var rPrDefault = doc.SelectSingleNode("//w:docDefaults/w:rPrDefault/w:rPr", ns) as XmlElement;
        if (rPrDefault != null)
            defaults = ParseRunPr(rPrDefault, ns, false);

        // 收集每个样式的直接 rPr / 段落默认 / basedOn 引用
        var raw = new Dictionary<String, WordStyle>(StringComparer.OrdinalIgnoreCase);
        foreach (XmlElement st in doc.SelectNodes("//w:style", ns)!)
        {
            var styleId = st.GetAttribute("w:styleId");
            if (String.IsNullOrEmpty(styleId)) continue;

            var style = new WordStyle();
            var rPrEl = st.SelectSingleNode("w:rPr", ns) as XmlElement;
            if (rPrEl == null)
            {
                var pPr0 = st.SelectSingleNode("w:pPr", ns) as XmlElement;
                if (pPr0 != null) rPrEl = pPr0.SelectSingleNode("w:rPr", ns) as XmlElement;
            }
            if (rPrEl != null) style.RPr = ParseRunPr(rPrEl, ns, false);
            style.BasedOn = (st.SelectSingleNode("w:basedOn", ns) as XmlElement)?.GetAttribute("w:val");

            // 段落默认格式（w:pPr）
            var pPr = st.SelectSingleNode("w:pPr", ns) as XmlElement;
            if (pPr != null)
            {
                var jc = pPr.SelectSingleNode("w:jc", ns) as XmlElement;
                if (jc != null) style.Alignment = jc.GetAttribute("w:val") ?? jc.GetAttribute("val");
                var ind = pPr.SelectSingleNode("w:ind", ns) as XmlElement;
                if (ind != null)
                {
                    var left = ind.GetAttribute("w:left") ?? ind.GetAttribute("left");
                    if (Int32.TryParse(left, out var lv)) style.IndentLeft = lv;
                    var right = ind.GetAttribute("w:right") ?? ind.GetAttribute("right");
                    if (Int32.TryParse(right, out var rv)) style.IndentRight = rv;
                    var firstLine = ind.GetAttribute("w:firstLine") ?? ind.GetAttribute("firstLine");
                    if (Int32.TryParse(firstLine, out var fl)) style.FirstLineIndent = fl;
                    else
                    {
                        var hanging = ind.GetAttribute("w:hanging") ?? ind.GetAttribute("hanging");
                        if (Int32.TryParse(hanging, out var hg)) style.FirstLineIndent = -hg;
                    }
                }
                var spacing = pPr.SelectSingleNode("w:spacing", ns) as XmlElement;
                if (spacing != null)
                {
                    var before = spacing.GetAttribute("w:before") ?? spacing.GetAttribute("before");
                    if (Int32.TryParse(before, out var bv)) style.SpaceBefore = bv;
                    var after = spacing.GetAttribute("w:after") ?? spacing.GetAttribute("after");
                    if (Int32.TryParse(after, out var av)) style.SpaceAfter = av;
                    var line = spacing.GetAttribute("w:line") ?? spacing.GetAttribute("line");
                    var lineRule = spacing.GetAttribute("w:lineRule") ?? spacing.GetAttribute("lineRule");
                    if (Int32.TryParse(line, out var lv) && lineRule == "auto")
                        style.LineSpacingPct = lv * 100 / 240;
                }
                var shd = pPr.SelectSingleNode("w:shd", ns) as XmlElement;
                if (shd != null) style.BackgroundColor = shd.GetAttribute("w:fill") ?? shd.GetAttribute("fill");
            }
            raw[styleId] = style;
        }

        // 展开继承链（基于 Normal → 目标样式，带环保护）
        foreach (var kv in raw)
            map[kv.Key] = ExpandStyle(kv.Key, raw, defaults, []);

        return map;
    }

    /// <summary>递归展开样式继承链：docDefaults → basedOn 链 → 样式自身（字符 + 段落默认）</summary>
    /// <param name="styleId">样式 ID</param>
    /// <param name="raw">样式原始信息表</param>
    /// <param name="defaults">文档默认格式</param>
    /// <param name="visited">已访问样式集合（环保护）</param>
    private static WordStyle ExpandStyle(String styleId,
        Dictionary<String, WordStyle> raw,
        RunProperties defaults, HashSet<String> visited)
    {
        var merged = new WordStyle();
        // 环保护：重复引用同一样式时返回空，避免无限递归
        if (!visited.Add(styleId)) return merged;

        // 1. 应用文档默认字符格式
        merged.RPr = new RunProperties();
        merged.RPr = MergeRunProps(merged.RPr, defaults);

        // 2. 递归应用 basedOn 链
        if (raw.TryGetValue(styleId, out var self) && self.BasedOn != null && raw.ContainsKey(self.BasedOn))
        {
            var baseStyle = ExpandStyle(self.BasedOn!, raw, defaults, visited);
            merged.RPr = MergeRunProps(merged.RPr, baseStyle.RPr);
            merged = MergeStyleDefaults(merged, baseStyle);
        }

        // 3. 应用样式自身格式（最高优先级）
        if (self.RPr != null)
            merged.RPr = MergeRunProps(merged.RPr, self.RPr);
        merged = MergeStyleDefaults(merged, self);

        visited.Remove(styleId);
        return merged;
    }

    /// <summary>合并段落默认格式（源非空即覆盖）</summary>
    private static WordStyle MergeStyleDefaults(WordStyle target, WordStyle src)
    {
        if (src.Alignment != null) target.Alignment = src.Alignment;
        if (src.IndentLeft.HasValue) target.IndentLeft = src.IndentLeft;
        if (src.IndentRight.HasValue) target.IndentRight = src.IndentRight;
        if (src.FirstLineIndent.HasValue) target.FirstLineIndent = src.FirstLineIndent;
        if (src.SpaceBefore.HasValue) target.SpaceBefore = src.SpaceBefore;
        if (src.SpaceAfter.HasValue) target.SpaceAfter = src.SpaceAfter;
        if (src.LineSpacingPct.HasValue) target.LineSpacingPct = src.LineSpacingPct;
        if (src.BackgroundColor != null) target.BackgroundColor = src.BackgroundColor;
        return target;
    }

    /// <summary>合并源格式到目标格式（源属性非空即覆盖）</summary>
    private static RunProperties MergeRunProps(RunProperties target, RunProperties src)
    {
        if (src.Bold.HasValue) target.Bold = src.Bold;
        if (src.Italic.HasValue) target.Italic = src.Italic;
        if (src.Underline.HasValue) target.Underline = src.Underline;
        if (src.Strikethrough.HasValue) target.Strikethrough = src.Strikethrough;
        if (src.Superscript.HasValue) target.Superscript = src.Superscript;
        if (src.Subscript.HasValue) target.Subscript = src.Subscript;
        if (src.ForeColor != null) target.ForeColor = src.ForeColor;
        if (src.FontSize.HasValue) target.FontSize = src.FontSize;
        if (src.FontName != null) target.FontName = src.FontName;
        if (src.EastAsiaFontName != null) target.EastAsiaFontName = src.EastAsiaFontName;
        if (src.HighlightColor != null) target.HighlightColor = src.HighlightColor;
        if (src.SmallCaps.HasValue) target.SmallCaps = src.SmallCaps;
        if (src.AllCaps.HasValue) target.AllCaps = src.AllCaps;
        if (src.Hidden.HasValue) target.Hidden = src.Hidden;
        if (src.Language != null) target.Language = src.Language;
        return target;
    }

    /// <summary>从 numbering.xml 解析 numId → 各层 numFmt 映射，用于段落列表类型判定</summary>
    /// <param name="numberingXml">numbering.xml 内容</param>
    /// <returns>numId → (ilvl → numFmt) 字典</returns>
    private static Dictionary<Int32, Dictionary<Int32, String>> ParseNumIdFormats(String? numberingXml)
    {
        var map = new Dictionary<Int32, Dictionary<Int32, String>>();
        if (String.IsNullOrEmpty(numberingXml)) return map;
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(numberingXml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            // numId → abstractNumId
            var numToAbs = new Dictionary<Int32, Int32>();
            foreach (XmlElement numEl in doc.SelectNodes("//w:num", ns)!)
            {
                if (!Int32.TryParse(numEl.GetAttribute("w:numId"), out var numId)) continue;
                var absEl = numEl.SelectSingleNode("w:abstractNumId", ns) as XmlElement;
                if (absEl != null && Int32.TryParse(absEl.GetAttribute("w:val"), out var absId))
                    numToAbs[numId] = absId;
            }

            foreach (var kv in numToAbs)
            {
                var absEl = doc.SelectSingleNode($"//w:abstractNum[@w:abstractNumId='{kv.Value}']", ns) as XmlElement;
                if (absEl == null) continue;
                var levels = new Dictionary<Int32, String>();
                foreach (XmlElement lvl in absEl.SelectNodes("w:lvl", ns)!)
                {
                    if (!Int32.TryParse(lvl.GetAttribute("w:ilvl"), out var ilvl)) continue;
                    var fmt = (lvl.SelectSingleNode("w:numFmt", ns) as XmlElement)?.GetAttribute("w:val");
                    if (fmt != null) levels[ilvl] = fmt;
                }
                map[kv.Key] = levels;
            }
        }
        catch { /* 解析失败回退为按 numId==2 的旧逻辑 */ }
        return map;
    }
    #endregion

    #region 段落解析
    private Element ParsePara(XmlElement pEl, XmlNamespaceManager ns, Dictionary<String, String> rels, Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats)
    {
        var para = new Paragraph();
        var isPageBreak = false;
        var isBullet = false;
        var isOrderedList = false;

        var bmStart = pEl.SelectSingleNode("w:bookmarkStart", ns);
        if (bmStart is XmlElement bmEl)
        {
            para.BookmarkName = bmEl.GetAttribute("w:name");
            if (para.BookmarkName == null) para.BookmarkName = bmEl.GetAttribute("name");
        }

        var pPr = pEl.SelectSingleNode("w:pPr", ns);
        RunProperties? styleDefaults = null;
        WordStyle? style = null;
        if (pPr is XmlElement pPrEl)
        {
            var pStyle = pPrEl.SelectSingleNode("w:pStyle", ns) as XmlElement;
            if (pStyle != null)
            {
                var styleId = pStyle.GetAttribute("w:val");
                if (styleId == null) styleId = pStyle.GetAttribute("val");
                para.StyleId = styleId;
                para.Style = ParseStyleId(styleId);
                if (styleId != null && styleMap.TryGetValue(styleId, out style))
                    styleDefaults = style.RPr;
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
            {
                var numIdEl = pPrEl.SelectSingleNode("w:numPr/w:numId", ns) as XmlElement;
                var numIdStr = numIdEl?.GetAttribute("w:val") ?? numIdEl?.GetAttribute("val");

                var ilvlEl = pPrEl.SelectSingleNode("w:numPr/w:ilvl", ns) as XmlElement;
                var ilvlStr = ilvlEl?.GetAttribute("w:val") ?? ilvlEl?.GetAttribute("val");
                if (Int32.TryParse(ilvlStr, out var ilvlVal)) para.ListLevel = ilvlVal;

                // 通过 numbering.xml 的 numId → numFmt 映射判定列表类型（而非硬编码 numId==2）
                if (Int32.TryParse(numIdStr, out var numIdVal))
                {
                    para.NumId = numIdVal;
                    if (numIdVal == 0)
                    {
                        // numId=0 表示无编号
                        isBullet = false;
                        isOrderedList = false;
                    }
                    else if (numIdFormats.TryGetValue(numIdVal, out var levels) &&
                             levels.TryGetValue(para.ListLevel, out var fmt))
                    {
                        para.ListFormat = fmt;
                        var isBulletFmt = fmt == "bullet";
                        isBullet = isBulletFmt;
                        isOrderedList = !isBulletFmt;
                    }
                    else
                    {
                        // 映射缺失时回退为无序列表（保持原行为）
                        isBullet = true;
                    }
                }
                else
                {
                    isBullet = true;
                }
            }

            // 段落边框：<w:pBdr>
            var pBdr = pPrEl.SelectSingleNode("w:pBdr", ns) as XmlElement;
            if (pBdr != null)
            {
                var borders = new ParagraphBorders();
                borders.Top    = ParseSingleBorder(pBdr.SelectSingleNode("w:top", ns) as XmlElement);
                borders.Bottom = ParseSingleBorder(pBdr.SelectSingleNode("w:bottom", ns) as XmlElement);
                borders.Left   = ParseSingleBorder(pBdr.SelectSingleNode("w:left", ns) as XmlElement);
                borders.Right  = ParseSingleBorder(pBdr.SelectSingleNode("w:right", ns) as XmlElement);
                if (borders.Top != null || borders.Bottom != null || borders.Left != null || borders.Right != null)
                    para.Borders = borders;
            }

            // 制表位：<w:tabs>
            var tabsEl = pPrEl.SelectSingleNode("w:tabs", ns) as XmlElement;
            if (tabsEl != null)
            {
                var list = new List<TabStop>();
                foreach (XmlElement tabEl in tabsEl.SelectNodes("w:tab", ns)!)
                {
                    var pos = tabEl.GetAttribute("w:pos") ?? tabEl.GetAttribute("pos");
                    if (!Int32.TryParse(pos, out var pv)) continue;
                    var ts = new TabStop { Position = pv };
                    ts.Alignment = tabEl.GetAttribute("w:val") ?? tabEl.GetAttribute("val") ?? "left";
                    var leader = tabEl.GetAttribute("w:leader") ?? tabEl.GetAttribute("leader");
                    if (!String.IsNullOrEmpty(leader) && leader != "none") ts.Leader = leader;
                    list.Add(ts);
                }
                if (list.Count > 0) para.TabStops = list;
            }

            // 首字下沉：<w:framePr w:dropCap="drop" w:lines="3"/>
            var framePr = pPrEl.SelectSingleNode("w:framePr", ns) as XmlElement;
            if (framePr != null)
            {
                var dc = framePr.GetAttribute("w:dropCap") ?? framePr.GetAttribute("dropCap");
                if (dc == "drop")
                {
                    var lines = framePr.GetAttribute("w:lines") ?? framePr.GetAttribute("lines");
                    if (Int32.TryParse(lines, out var ln) && ln > 0) para.DropCapLines = ln;
                    else para.DropCapLines = 3;
                }
            }

            // 分页控制：<w:keepNext/> <w:keepLines/> <w:widowControl>
            para.KeepNext = pPrEl.SelectSingleNode("w:keepNext", ns) != null;
            para.KeepLines = pPrEl.SelectSingleNode("w:keepLines", ns) != null;
            var widowEl = pPrEl.SelectSingleNode("w:widowControl", ns) as XmlElement;
            if (widowEl != null)
            {
                var val = widowEl.GetAttribute("w:val") ?? widowEl.GetAttribute("val");
                para.WidowControl = val != "0" && val != "false";
            }
        }

        // 应用段落样式默认（段落未显式设置时填充；优先级：内联显式 > 样式默认）
        if (style != null)
        {
            if (para.Alignment == null) para.Alignment = style.Alignment;
            if (para.IndentLeft == null && style.IndentLeft.HasValue) para.IndentLeft = style.IndentLeft;
            if (para.IndentRight == null && style.IndentRight.HasValue) para.IndentRight = style.IndentRight;
            if (para.FirstLineIndent == null && style.FirstLineIndent.HasValue) para.FirstLineIndent = style.FirstLineIndent;
            if (para.SpaceBefore == null && style.SpaceBefore.HasValue) para.SpaceBefore = style.SpaceBefore;
            if (para.SpaceAfter == null && style.SpaceAfter.HasValue) para.SpaceAfter = style.SpaceAfter;
            if (para.LineSpacingPct == null && style.LineSpacingPct.HasValue) para.LineSpacingPct = style.LineSpacingPct;
            if (para.BackgroundColor == null) para.BackgroundColor = style.BackgroundColor;
        }

        var br = pEl.SelectSingleNode("w:r/w:br", ns) as XmlElement;
        if (br != null)
        {
            var brType = br.GetAttribute("w:type") ?? br.GetAttribute("type");
            if (brType == "page") isPageBreak = true;
        }

        para.IsBullet = isBullet;
        para.IsOrderedList = isOrderedList;
        para.IsPageBreak = isPageBreak;

        // 始终解析 Run（含分页符段落中的文本，如"正文后加分页符"），分页符在 Run 中以 \f 表示
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
            else if (el.LocalName == "sdt")
            {
                // 内联内容控件（W22）：把 sdtContent 内的 Run 并入段落 Runs，Find/Replace 可达
                var sdtContent = el.SelectSingleNode("w:sdtContent", ns) as XmlElement;
                if (sdtContent != null)
                {
                    foreach (XmlNode sc in sdtContent.ChildNodes)
                    {
                        if (sc is not XmlElement sdtChild) continue;
                        if (sdtChild.LocalName == "r")
                        {
                            var run = ParseRun(sdtChild, ns, null);
                            ApplyStyleDefaults(run, styleDefaults);
                            para.Runs.Add(run);
                        }
                        else if (sdtChild.LocalName == "hyperlink")
                        {
                            var hlRelId = sdtChild.GetAttribute("r:id") ?? sdtChild.GetAttribute("id");
                            foreach (XmlNode hlChild in sdtChild.ChildNodes)
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
            }
        }

        // 文本框（W22）：解析 w:drawing 内 wps:txbx/w:txbxContent 的段落到 TextBoxes，
        // FindText/ReplaceText 可覆盖文本框文本；递归支持文本框嵌套。
        // 只取最外层 txbxContent，内层由 ParsePara 递归解析，避免嵌套文本框被双重解析
        foreach (XmlElement txbxContent in pEl.SelectNodes(".//w:txbxContent[not(ancestor::w:txbxContent)]", ns)!)
        {
            foreach (XmlNode txbxChild in txbxContent.ChildNodes)
            {
                if (txbxChild is XmlElement txbxPara && txbxPara.LocalName == "p")
                {
                    var tp = ParsePara(txbxPara, ns, rels, styleMap, numIdFormats).Paragraph!;
                    tp.RawXml = txbxPara.OuterXml; // 保留内层原 XML，替换时优先使用
                    para.TextBoxes.Add(tp);
                }
            }
        }

        return new Element { Type = ElementType.Paragraph, Paragraph = para };
    }

    private static Run ParseRun(XmlElement rEl, XmlNamespaceManager ns, String? hyperlinkRelId)
    {
        var run = new Run { HyperlinkRelId = hyperlinkRelId };
        var rPr = rEl.SelectSingleNode("w:rPr", ns) as XmlElement;
        if (rPr != null)
            run.Properties = ParseRunPr(rPr, ns, hyperlinkRelId != null);

        // 一个 Run 可能包含多个 w:t（被域代码/分隔符分割），也可能含 w:tab / w:br / w:cr 等特殊字符元素
        var sb = new StringBuilder();
        foreach (XmlNode child in rEl.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            switch (el.LocalName)
            {
                case "t":
                    sb.Append(el.InnerText);
                    break;
                case "tab":
                    sb.Append('\t');
                    break;
                case "br":
                    var brType = el.GetAttribute("w:type") ?? el.GetAttribute("type");
                    if (brType == "page")
                        sb.Append('\f');   // 分页符
                    else if (brType == "column")
                        sb.Append('\v');   // 分栏符
                    else
                        sb.Append('\n');   // 普通换行
                    break;
                case "cr":
                    sb.Append('\n');
                    break;
                case "noBreakHyphen":
                    sb.Append('-');
                    break;
                case "softHyphen":
                    sb.Append('\u00AD');
                    break;
            }
        }
        run.Text = sb.ToString();
        return run;
    }

    /// <summary>解析布尔开关元素（w:b / w:i / w:strike 等），支持 w:val="0/1/true/false" 显式关闭</summary>
    /// <param name="el">开关元素，null 表示未出现</param>
    /// <returns>null=未设置（继承），true=开启，false=显式关闭</returns>
    private static Boolean? ParseSwitch(XmlElement? el)
    {
        if (el == null) return null;
        var val = el.GetAttribute("w:val");
        if (String.IsNullOrEmpty(val)) val = el.GetAttribute("val");
        if (String.IsNullOrEmpty(val)) return true; // 无 w:val = 开启
        return val switch
        {
            "0" or "false" or "off" or "none" => false,
            _ => true,
        };
    }

    private static RunProperties ParseRunPr(XmlElement rPrEl, XmlNamespaceManager ns, Boolean isHyperlink)
    {
        var p = new RunProperties();
        // 单次遍历子节点解析全部格式（避免逐属性 XPath 查询，显著提升大文档读取性能）
        foreach (XmlNode child in rPrEl.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            switch (el.LocalName)
            {
                case "b": p.Bold = ParseSwitch(el); break;
                case "i": p.Italic = ParseSwitch(el); break;
                case "strike":
                    p.Strikethrough = ParseSwitch(el);
                    break;
                case "dstrike": // 双删除线同归 Strikethrough（首个非空生效）
                    if (p.Strikethrough == null) p.Strikethrough = ParseSwitch(el);
                    break;
                case "smallCaps": p.SmallCaps = ParseSwitch(el); break;
                case "caps": p.AllCaps = ParseSwitch(el); break;
                case "vanish": p.Hidden = ParseSwitch(el); break;
                case "color":
                    var colorVal = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (colorVal != null && colorVal != "auto") p.ForeColor = colorVal;
                    break;
                case "sz":
                    var szVal = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (Single.TryParse(szVal, out var sv)) p.FontSize = sv / 2f;
                    break;
                case "rFonts":
                    var ascii = el.GetAttribute("w:ascii") ?? el.GetAttribute("ascii")
                        ?? el.GetAttribute("w:hAnsi") ?? el.GetAttribute("hAnsi");
                    if (ascii != null) p.FontName = ascii;
                    var eastAsia = el.GetAttribute("w:eastAsia") ?? el.GetAttribute("eastAsia");
                    if (eastAsia != null) p.EastAsiaFontName = eastAsia;
                    break;
                case "vertAlign": // 上标/下标/baseline
                    var va = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (va == "superscript") { p.Superscript = true; p.Subscript = false; }
                    else if (va == "subscript") { p.Superscript = false; p.Subscript = true; }
                    else if (va == "baseline") { p.Superscript = false; p.Subscript = false; }
                    break;
                case "highlight":
                    var hv = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (!String.IsNullOrEmpty(hv) && hv != "none") p.HighlightColor = hv;
                    break;
                case "lang":
                    var lv = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (!String.IsNullOrEmpty(lv)) p.Language = lv;
                    break;
                case "u": // 下划线样式（val=none/nil 表示显式取消）
                    var uVal = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (uVal == null || uVal == "single" || uVal == "1" || uVal == "true")
                        p.Underline = true;
                    else if (uVal == "none" || uVal == "nil" || uVal == "0" || uVal == "false")
                        p.Underline = false;
                    else
                    {
                        p.Underline = true;
                        p.UnderlineStyle = uVal;
                    }
                    break;
                case "spacing": // 字符间距
                    var spVal = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (Single.TryParse(spVal, out var spv)) p.CharacterSpacing = spv;
                    break;
                case "w": // 字符缩放 <w:w w:val="..."/>
                    var wVal = el.GetAttribute("w:val") ?? el.GetAttribute("val");
                    if (Int32.TryParse(wVal, out var wv)) p.CharacterScaling = wv;
                    break;
            }
        }
        if (isHyperlink && p.ForeColor == null)
        {
            p.ForeColor = "0563C1";
            p.Underline = true;
        }
        return p;
    }

    private static void ApplyStyleDefaults(Run run, RunProperties? defaults)
    {
        if (defaults == null) return;
        var rp = run.Properties;
        if (rp == null)
        {
            rp = new RunProperties();
            run.Properties = rp;
        }
        // 仅填充 run 上未显式设置的属性（样式优先级低于 run 内联格式）
        if (rp.Bold == null) rp.Bold = defaults.Bold;
        if (rp.Italic == null) rp.Italic = defaults.Italic;
        if (rp.Underline == null) rp.Underline = defaults.Underline;
        if (rp.Strikethrough == null) rp.Strikethrough = defaults.Strikethrough;
        if (rp.Superscript == null) rp.Superscript = defaults.Superscript;
        if (rp.Subscript == null) rp.Subscript = defaults.Subscript;
        if (rp.ForeColor == null) rp.ForeColor = defaults.ForeColor;
        if (rp.FontSize == null) rp.FontSize = defaults.FontSize;
        if (rp.FontName == null) rp.FontName = defaults.FontName;
        if (rp.EastAsiaFontName == null) rp.EastAsiaFontName = defaults.EastAsiaFontName;
        if (rp.HighlightColor == null) rp.HighlightColor = defaults.HighlightColor;
        if (rp.SmallCaps == null) rp.SmallCaps = defaults.SmallCaps;
        if (rp.AllCaps == null) rp.AllCaps = defaults.AllCaps;
        if (rp.Hidden == null) rp.Hidden = defaults.Hidden;
        if (rp.Language == null) rp.Language = defaults.Language;
    }

    /// <summary>解析单边边框 XML 元素（w:top / w:bottom / w:left / w:right）</summary>
    private static Border? ParseSingleBorder(XmlElement? el)
    {
        if (el == null) return null;
        var val = el.GetAttribute("w:val") ?? el.GetAttribute("val");
        if (String.IsNullOrEmpty(val) || val == "none" || val == "nil") return null;
        var border = new Border
        {
            Style = val switch
            {
                "single"     => BorderStyle.Single,
                "thick"      => BorderStyle.Thick,
                "double"     => BorderStyle.Double,
                "dotted"     => BorderStyle.Dotted,
                "dashed"     => BorderStyle.Dashed,
                "dotDash"    => BorderStyle.DotDash,
                "dotDotDash" => BorderStyle.DotDotDash,
                _            => BorderStyle.Single,
            }
        };
        border.Color = el.GetAttribute("w:color") ?? el.GetAttribute("color");
        var sz = el.GetAttribute("w:sz") ?? el.GetAttribute("sz");
        if (Int32.TryParse(sz, out var s)) border.Width = s;
        var shadow = el.GetAttribute("w:shadow") ?? el.GetAttribute("shadow");
        if (shadow == "1" || shadow == "true") border.Shadow = true;
        return border;
    }
    #endregion

    #region 表格/图片/节属性
    private Element ParseTable(XmlElement tblEl, XmlNamespaceManager ns, Dictionary<String, String> rels, Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats, Document doc)
    {
        var rows = new List<List<Cell>>();
        var firstRowHeader = false;
        var style = new TableStyle();
        String? tblStyleId = null;
        String? tblAlignment = null;
        Int32? tblWidth = null;

        var tblPr = tblEl.SelectSingleNode("w:tblPr", ns) as XmlElement;
        if (tblPr != null)
        {
            // 表格样式引用
            var tblStyle = tblPr.SelectSingleNode("w:tblStyle", ns) as XmlElement;
            tblStyleId = tblStyle?.GetAttribute("w:val") ?? tblStyle?.GetAttribute("val");
            // 表格对齐
            var jc = tblPr.SelectSingleNode("w:jc", ns) as XmlElement;
            tblAlignment = jc?.GetAttribute("w:val") ?? jc?.GetAttribute("val");
            // 表格宽度
            var tblW = tblPr.SelectSingleNode("w:tblW", ns) as XmlElement;
            if (tblW != null)
            {
                var wv = tblW.GetAttribute("w:w") ?? tblW.GetAttribute("w");
                if (Int32.TryParse(wv, out var tw)) tblWidth = tw;
            }

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

        // 同时构建 Table 富模型（行高/行级表头/行背景/禁止跨页），与 TableRows 并存
        var tableModel = new Table { FirstRowHeader = firstRowHeader };
        if (style.ColumnWidths != null) tableModel.ColumnWidths = style.ColumnWidths;
        var modelRows = new List<TableRow>();

        foreach (XmlElement tr in tblEl.SelectNodes("w:tr", ns)!)
        {
            var cells = new List<Cell>();
            var trPr = tr.SelectSingleNode("w:trPr", ns) as XmlElement;
            var isHeader = false;
            Int32? trHeight = null;
            var cantSplit = false;
            String? rowBg = null;
            if (trPr != null)
            {
                if (trPr.SelectSingleNode("w:tblHeader", ns) != null) { firstRowHeader = true; isHeader = true; }
                var trHeightEl = trPr.SelectSingleNode("w:trHeight", ns) as XmlElement;
                if (trHeightEl != null)
                {
                    var hv = trHeightEl.GetAttribute("w:val") ?? trHeightEl.GetAttribute("val");
                    if (Int32.TryParse(hv, out var th)) trHeight = th;
                }
                cantSplit = trPr.SelectSingleNode("w:cantSplit", ns) != null;
                var trShd = trPr.SelectSingleNode("w:shd", ns) as XmlElement;
                if (trShd != null) rowBg = trShd.GetAttribute("w:fill") ?? trShd.GetAttribute("fill");
            }
            foreach (XmlElement tc in tr.SelectNodes("w:tc", ns)!)
                cells.Add(ParseCell(tc, ns, rels, styleMap, numIdFormats, doc));
            if (cells.Count > 0)
            {
                rows.Add(cells);
                modelRows.Add(new TableRow
                {
                    Cells = cells,
                    IsHeader = isHeader,
                    Height = trHeight,
                    CantSplit = cantSplit,
                    BackgroundColor = rowBg,
                });
            }
        }
        tableModel.Rows = modelRows;
        tableModel.FirstRowHeader = firstRowHeader;
        tableModel.StyleId = tblStyleId;
        tableModel.Alignment = tblAlignment;
        tableModel.Width = tblWidth;

        return new Element
        {
            Type = ElementType.Table,
            TableRows = rows,
            TableFirstRowHeader = firstRowHeader,
            TableStyle = style,
            Table = tableModel,
        };
    }

    private Cell ParseCell(XmlElement tcEl, XmlNamespaceManager ns, Dictionary<String, String> rels,
        Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats, Document doc)
    {
        var cell = new Cell();
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
            // 垂直合并：w:vMerge 无 val = 继续合并；w:val="restart" = 合并起点
            var vMerge = tcPr.SelectSingleNode("w:vMerge", ns) as XmlElement;
            if (vMerge != null)
            {
                var val = vMerge.GetAttribute("w:val") ?? vMerge.GetAttribute("val");
                cell.RowSpan = val == "restart" ? -1 : 0;
            }
            // 单元格宽度 w:tcW
            var tcW = tcPr.SelectSingleNode("w:tcW", ns) as XmlElement;
            if (tcW != null)
            {
                var wVal = tcW.GetAttribute("w:w") ?? tcW.GetAttribute("w");
                if (Int32.TryParse(wVal, out var tw)) cell.Width = tw;
            }
            // 垂直对齐 w:vAlign
            var vAlign = tcPr.SelectSingleNode("w:vAlign", ns) as XmlElement;
            if (vAlign != null)
                cell.VerticalAlignment = vAlign.GetAttribute("w:val") ?? vAlign.GetAttribute("val");
            // 单元格边框 w:tcBorders
            var tcBorders = tcPr.SelectSingleNode("w:tcBorders", ns) as XmlElement;
            if (tcBorders != null)
            {
                var borders = new TableBorders
                {
                    Top = ParseSingleBorder(tcBorders.SelectSingleNode("w:top", ns) as XmlElement),
                    Bottom = ParseSingleBorder(tcBorders.SelectSingleNode("w:bottom", ns) as XmlElement),
                    Left = ParseSingleBorder(tcBorders.SelectSingleNode("w:left", ns) as XmlElement),
                    Right = ParseSingleBorder(tcBorders.SelectSingleNode("w:right", ns) as XmlElement),
                };
                if (borders.Top != null || borders.Bottom != null || borders.Left != null || borders.Right != null)
                    cell.Borders = borders;
            }
        }

        foreach (XmlElement pEl in tcEl.SelectNodes("w:p", ns)!)
        {
            var pe = ParsePara(pEl, ns, rels, styleMap, numIdFormats);
            if (pe.Paragraph != null) cell.Paragraphs.Add(pe.Paragraph);
        }

        // 嵌套表格（W46）：<w:tc><w:tbl>...</w:tbl>
        foreach (XmlElement tblEl in tcEl.SelectNodes("w:tbl", ns)!)
        {
            var te = ParseTable(tblEl, ns, rels, styleMap, numIdFormats, doc);
            if (te.Table != null) cell.NestedTables.Add(te.Table);
        }

        // 单元格内嵌图片：提取图片数据到 doc.Images（RawXml 已保留原 XML）
        foreach (XmlElement drawing in tcEl.SelectNodes(".//w:drawing", ns)!)
            ParseDrawing(drawing, ns, rels, doc);

        return cell;
    }

    private Element? ParseDrawing(XmlElement drawing, XmlNamespaceManager ns, Dictionary<String, String> rels, Document doc)
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

        // SVG 矢量图片：<asvg:svgBlip r:embed="..."/>
        var isSvg = false;
        var svgBlip = drawing.SelectSingleNode(".//asvg:svgBlip", ns) as XmlElement;
        var rId = svgBlip?.GetAttribute("r:embed");
        if (rId == null)
        {
            var blip = drawing.SelectSingleNode(".//a:blip", ns) as XmlElement;
            rId = blip?.GetAttribute("r:embed");
        }
        else
        {
            isSvg = true;
        }

        // VML 图片：<w:pict><v:shape><v:imagedata r:id="..."/></v:shape></w:pict>
        if (rId == null)
        {
            var vmlData = drawing.SelectSingleNode(".//v:imagedata", ns) as XmlElement
                ?? drawing.ParentNode?.SelectSingleNode(".//v:imagedata", ns) as XmlElement;
            rId = vmlData?.GetAttribute("r:id") ?? vmlData?.GetAttribute("id");
        }

        if (rId == null) return null;

        if (!doc.Images.ContainsKey(rId) && rels.TryGetValue(rId, out var target))
        {
            var entry = _zip.GetEntry($"word/{target}");
            if (entry != null)
            {
                var ext = Path.GetExtension(target).TrimStart('.').ToLowerInvariant();
                if (isSvg) ext = "svg";
                using var ms = new MemoryStream();
                using var es = entry.Open();
                es.CopyTo(ms);
                doc.Images[rId] = (ext, ms.ToArray());
            }
        }

        var img = new Image
        {
            RelId = rId,
            WidthEmu = cx > 0 ? cx : 3600000,
            HeightEmu = cy > 0 ? cy : 2700000,
            IsSvg = isSvg,
            Extension = doc.Images.TryGetValue(rId, out var imgData) ? imgData.Extension : (isSvg ? "svg" : "png"),
        };

        // wp:anchor 浮动定位与环绕
        if (inline.LocalName == "anchor")
        {
            img.AnchorType = "anchor";
            var posH = inline.SelectSingleNode("wp:positionH", ns) as XmlElement;
            if (posH != null)
            {
                var align = posH.SelectSingleNode("wp:align", ns);
                var offset = posH.SelectSingleNode("wp:posOffset", ns);
                if (align != null) img.AnchorPosH = align.InnerText;
                if (offset != null && Int64.TryParse(offset.InnerText, out var ox)) img.AnchorOffsetX = ox;
            }
            var posV = inline.SelectSingleNode("wp:positionV", ns) as XmlElement;
            if (posV != null)
            {
                var align = posV.SelectSingleNode("wp:align", ns);
                var offset = posV.SelectSingleNode("wp:posOffset", ns);
                if (align != null) img.AnchorPosV = align.InnerText;
                if (offset != null && Int64.TryParse(offset.InnerText, out var oy)) img.AnchorOffsetY = oy;
            }
            // 环绕方式：wrapSquare/wrapTight/wrapThrough/wrapTopAndBottom/wrapNone
            var wrap = inline.SelectSingleNode("wp:wrapSquare", ns) ?? inline.SelectSingleNode("wp:wrapTight", ns)
                ?? inline.SelectSingleNode("wp:wrapThrough", ns) ?? inline.SelectSingleNode("wp:wrapTopAndBottom", ns)
                ?? inline.SelectSingleNode("wp:wrapNone", ns);
            if (wrap != null)
                img.Wrap = wrap.LocalName switch
                {
                    "wrapSquare" => "square",
                    "wrapTight" => "tight",
                    "wrapThrough" => "through",
                    "wrapTopAndBottom" => "topAndBottom",
                    "wrapNone" => "none",
                    _ => null,
                };
        }

        // 替代文本（wp:docPr descr）
        var docPr = inline.SelectSingleNode("wp:docPr", ns) as XmlElement;
        if (docPr != null)
        {
            var descr = docPr.GetAttribute("descr");
            if (!String.IsNullOrEmpty(descr)) img.AltText = descr;
        }

        return new Element { Type = ElementType.Image, Image = img };
    }

    private static void ParseSectPr(XmlElement sectPr, XmlNamespaceManager ns, PageSettings ps)
    {
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

        // 分栏设置
        var cols = sectPr.SelectSingleNode("w:cols", ns) as XmlElement;
        if (cols != null)
        {
            if (Int32.TryParse(cols.GetAttribute("w:num") ?? cols.GetAttribute("num"), out var cn)) ps.ColumnCount = cn;
            if (Int32.TryParse(cols.GetAttribute("w:space") ?? cols.GetAttribute("space"), out var cs)) ps.ColumnSpacing = cs;
        }

        // 页面边框
        var pgBorders = sectPr.SelectSingleNode("w:pgBorders", ns) as XmlElement;
        if (pgBorders != null)
        {
            var pb = new PageBorder();
            var offset = pgBorders.GetAttribute("w:offsetFrom") ?? pgBorders.GetAttribute("offsetFrom");
            if (offset == "text") pb.OffsetFrom = 0;

            var top = pgBorders.SelectSingleNode("w:top", ns) as XmlElement;
            if (top != null) { pb.Top = top.GetAttribute("w:val") ?? top.GetAttribute("val"); ParsePgBorderAttrs(top, pb); }

            var bottom = pgBorders.SelectSingleNode("w:bottom", ns) as XmlElement;
            if (bottom != null) { pb.Bottom = bottom.GetAttribute("w:val") ?? bottom.GetAttribute("val"); ParsePgBorderAttrs(bottom, pb); }

            var left = pgBorders.SelectSingleNode("w:left", ns) as XmlElement;
            if (left != null) { pb.Left = left.GetAttribute("w:val") ?? left.GetAttribute("val"); ParsePgBorderAttrs(left, pb); }

            var right = pgBorders.SelectSingleNode("w:right", ns) as XmlElement;
            if (right != null) { pb.Right = right.GetAttribute("w:val") ?? right.GetAttribute("val"); ParsePgBorderAttrs(right, pb); }

            ps.PageBorder = pb;
        }

        // 行号
        var lnNumType = sectPr.SelectSingleNode("w:lnNumType", ns) as XmlElement;
        if (lnNumType != null)
        {
            var ln = new LineNumberSettings();
            if (Int32.TryParse(lnNumType.GetAttribute("w:start") ?? lnNumType.GetAttribute("start"), out var st)) ln.Start = st;
            if (Int32.TryParse(lnNumType.GetAttribute("w:countBy") ?? lnNumType.GetAttribute("countBy"), out var cb)) ln.CountBy = cb;
            if (Int32.TryParse(lnNumType.GetAttribute("w:distance") ?? lnNumType.GetAttribute("distance"), out var dist)) ln.Distance = dist;
            var restart = lnNumType.GetAttribute("w:restart") ?? lnNumType.GetAttribute("restart");
            if (!restart.IsNullOrEmpty()) ln.Restart = restart!;
            ps.LineNumber = ln;
        }

        // 页眉/页脚引用（default/first/even 三种类型全部捕获）
        ps.HeaderRefs.Clear();
        foreach (XmlElement hdrRef in sectPr.SelectNodes("w:headerReference", ns)!)
        {
            var type = hdrRef.GetAttribute("w:type") ?? hdrRef.GetAttribute("type") ?? "default";
            var rId = hdrRef.GetAttribute("r:id");
            if (rId != null)
            {
                ps.HeaderRefs[type] = rId;
                if (type == "default") ps.HeaderText = rId;
            }
        }
        ps.FooterRefs.Clear();
        foreach (XmlElement ftrRef in sectPr.SelectNodes("w:footerReference", ns)!)
        {
            var type = ftrRef.GetAttribute("w:type") ?? ftrRef.GetAttribute("type") ?? "default";
            var rId = ftrRef.GetAttribute("r:id");
            if (rId != null)
            {
                ps.FooterRefs[type] = rId;
                if (type == "default") ps.FooterText = rId;
            }
        }
    }

    private static void ParsePgBorderAttrs(XmlElement el, PageBorder pb)
    {
        if (Int32.TryParse(el.GetAttribute("w:sz") ?? el.GetAttribute("sz"), out var sz)) pb.Size = sz;
        if (Int32.TryParse(el.GetAttribute("w:space") ?? el.GetAttribute("space"), out var sp)) pb.Space = sp;
        var color = el.GetAttribute("w:color") ?? el.GetAttribute("color");
        if (!String.IsNullOrEmpty(color)) pb.Color = color;
    }

    /// <summary>从 settings.xml 解析 w:docVars 文档变量</summary>
    private static void ParseDocumentVariables(String settingsXml, Dictionary<String, String> vars)
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
        catch { /* 解析失败不影响整体读取 */ }
    }

    /// <summary>解析 numbering.xml 到 Numbering 模型</summary>
    private static Numbering? ParseNumbering(String numberingXml)
    {
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(numberingXml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            // 收集所有 num/abstractNumId
            var absNumIds = new HashSet<String>();
            var numNodes = doc.SelectNodes("//w:num", ns);
            if (numNodes != null)
            {
                foreach (XmlElement numEl in numNodes)
                {
                    var absIdEl = numEl.SelectSingleNode("w:abstractNumId", ns) as XmlElement;
                    var absId = absIdEl?.GetAttribute("w:val");
                    if (absId != null) absNumIds.Add(absId);
                }
            }

            if (absNumIds.Count == 0) return null;

            var numbering = new Numbering();
            if (numNodes!.Count > 0)
            {
                var firstNum = (XmlElement)numNodes[0]!;
                if (Int32.TryParse(firstNum.GetAttribute("w:numId"), out var nid))
                    numbering.NumberingId = nid;
            }

            // 合并所有 abstractNum 的级别定义
            foreach (var absId in absNumIds)
            {
                var absNumEl = doc.SelectSingleNode($"//w:abstractNum[@w:abstractNumId='{absId}']", ns) as XmlElement;
                if (absNumEl == null) continue;

                var lvlNodes = absNumEl.SelectNodes("w:lvl", ns);
                if (lvlNodes == null) continue;

                foreach (XmlElement lvlEl in lvlNodes)
                {
                    var ilvlStr = lvlEl.GetAttribute("w:ilvl");
                    var fmt = (lvlEl.SelectSingleNode("w:numFmt", ns) as XmlElement)?.GetAttribute("w:val") ?? "decimal";
                    var text = (lvlEl.SelectSingleNode("w:lvlText", ns) as XmlElement)?.GetAttribute("w:val");
                    var startStr = (lvlEl.SelectSingleNode("w:start", ns) as XmlElement)?.GetAttribute("w:val");

                    var level = new NumberingLevel { Format = fmt, Text = text };
                    if (Int32.TryParse(ilvlStr, out var ilvl)) level.Level = ilvl;
                    if (Int32.TryParse(startStr, out var startVal)) level.StartAt = startVal;

                    var pPr = lvlEl.SelectSingleNode("w:pPr", ns) as XmlElement;
                    if (pPr != null)
                    {
                        var ind = pPr.SelectSingleNode("w:ind", ns) as XmlElement;
                        if (ind != null)
                        {
                            var leftStr = ind.GetAttribute("w:left");
                            var hangStr = ind.GetAttribute("w:hanging");
                            if (Int32.TryParse(leftStr, out var leftVal)) level.Indent = leftVal;
                            if (Int32.TryParse(hangStr, out var hangVal)) level.HangingIndent = hangVal;
                        }
                    }

                    // bullet 格式额外属性
                    if (fmt == "bullet")
                    {
                        var rPr = lvlEl.SelectSingleNode("w:rPr", ns) as XmlElement;
                        if (rPr != null)
                        {
                            var rFonts = rPr.SelectSingleNode("w:rFonts", ns) as XmlElement;
                            if (rFonts != null)
                                level.BulletFontName = rFonts.GetAttribute("w:ascii");
                        }
                        level.BulletChar = text;
                    }

                    numbering.Levels.Add(level);
                }
            }

            // 快捷字段
            if (numbering.Levels.Count > 0)
            {
                numbering.Format = numbering.Levels[0].Format;
                numbering.BulletChar = numbering.Levels[0].BulletChar;
            }

            return numbering.Levels.Count > 0 ? numbering : null;
        }
        catch { return null; }
    }

    private static SdtElement? ParseSdt(XmlElement sdtEl, XmlNamespaceManager ns)
    {
        var sdt = new SdtElement();

        // 解析 sdtPr 获取控件类型和标签
        var sdtPr = sdtEl.SelectSingleNode("w:sdtPr", ns) as XmlElement;
        if (sdtPr != null)
        {
            sdt.Tag = (sdtPr.SelectSingleNode("w:tag", ns) as XmlElement)?.GetAttribute("w:val")
                ?? (sdtPr.SelectSingleNode("w:tag", ns) as XmlElement)?.GetAttribute("val");
            sdt.Alias = (sdtPr.SelectSingleNode("w:alias", ns) as XmlElement)?.GetAttribute("w:val")
                ?? (sdtPr.SelectSingleNode("w:alias", ns) as XmlElement)?.GetAttribute("val");

            if (sdtPr.SelectSingleNode("w:date", ns) != null)
                sdt.SdtType = SdtType.Date;
            else if (sdtPr.SelectSingleNode("w:dropDownList", ns) != null)
                sdt.SdtType = SdtType.DropDownList;
            else if (sdtPr.SelectSingleNode("w:comboBox", ns) != null)
                sdt.SdtType = SdtType.ComboBox;
            else if (sdtPr.SelectSingleNode("w:checkBox", ns) != null)
                sdt.SdtType = SdtType.CheckBox;
            else if (sdtPr.SelectSingleNode("w:picture", ns) != null)
                sdt.SdtType = SdtType.Picture;
            else if (sdtPr.SelectSingleNode("w:repeatingSection", ns) != null)
                sdt.SdtType = SdtType.RepeatingSection;
            else if (sdtPr.SelectSingleNode("w:richText", ns) != null)
                sdt.SdtType = SdtType.RichText;
            else
                sdt.SdtType = SdtType.PlainText; // 默认纯文本
        }

        // 提取内容文本（sdtContent 内的 w:t 文本）
        var sdtContent = sdtEl.SelectSingleNode("w:sdtContent", ns) as XmlElement;
        if (sdtContent != null)
        {
            var sb = new StringBuilder();
            foreach (XmlElement t in sdtContent.SelectNodes(".//w:t", ns)!)
                sb.Append(t.InnerText);
            sdt.Content = sb.ToString();
        }

        return sdt;
    }

    private void LoadHdrFtr(Dictionary<String, String> rels, Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats, Document doc)
    {
        var ps = doc.PageSettings;

        // 富文本页眉（default/first/even 三种类型）
        doc.Headers.Clear();
        foreach (var kv in ps.HeaderRefs)
        {
            if (!rels.TryGetValue(kv.Value, out var target)) continue;
            var elements = LoadHdrFtrElements(target, rels, styleMap, numIdFormats, doc);
            if (elements.Count == 0) continue;
            doc.Headers.Add(new Header { Type = kv.Key, Elements = elements });
            if (kv.Key == "default")
            {
                var text = LoadHdrFtrText(target, doc);
                if (text != null) doc.HeaderText = text;
            }
        }

        // 富文本页脚（default/first/even 三种类型）
        doc.Footers.Clear();
        foreach (var kv in ps.FooterRefs)
        {
            if (!rels.TryGetValue(kv.Value, out var target)) continue;
            var elements = LoadHdrFtrElements(target, rels, styleMap, numIdFormats, doc);
            if (elements.Count == 0) continue;
            doc.Footers.Add(new Footer { Type = kv.Key, Elements = elements });
            if (kv.Key == "default")
            {
                var text = LoadHdrFtrText(target, doc);
                if (text != null) doc.FooterText = text;
            }
        }

        // 无富文本引用时回退到纯文本（HeaderText 此时仍是 relId）
        if (doc.Headers.Count == 0 && ps.HeaderText != null && rels.TryGetValue(ps.HeaderText, out var hTarget))
        {
            var text = LoadHdrFtrText(hTarget, doc);
            if (text != null) doc.HeaderText = text;
        }
        if (doc.Footers.Count == 0 && ps.FooterText != null && rels.TryGetValue(ps.FooterText, out var fTarget))
        {
            var text = LoadHdrFtrText(fTarget, doc);
            if (text != null) doc.FooterText = text;
        }

        ps.HeaderText = doc.HeaderText;
        ps.FooterText = doc.FooterText;
    }

    /// <summary>解析页眉/页脚部件内容为元素列表（段落/表格/内容控件 + 图片提取）</summary>
    private List<Element> LoadHdrFtrElements(String target, Dictionary<String, String> rels,
        Dictionary<String, WordStyle> styleMap, Dictionary<Int32, Dictionary<Int32, String>> numIdFormats, Document doc)
    {
        var elements = new List<Element>();
        var entry = _zip.GetEntry($"word/{target}");
        if (entry == null) return elements;

        var xml = new XmlDocument();
        using (var s = entry.Open()) xml.Load(s);
        var ns = WmlNs(xml);
        var root = xml.SelectSingleNode("//*[local-name()='hdr' or local-name()='ftr']", ns) as XmlElement;
        if (root == null) return elements;

        foreach (XmlNode child in root.ChildNodes)
        {
            if (child is not XmlElement el) continue;
            if (el.LocalName == "p")
            {
                var pe = ParsePara(el, ns, rels, styleMap, numIdFormats);
                pe.RawXml = el.OuterXml;
                elements.Add(pe);
            }
            else if (el.LocalName == "tbl")
            {
                var te = ParseTable(el, ns, rels, styleMap, numIdFormats, doc);
                te.RawXml = el.OuterXml;
                elements.Add(te);
            }
            else if (el.LocalName == "sdt")
            {
                var se = ParseSdt(el, ns);
                if (se != null) elements.Add(new Element { Type = ElementType.Sdt, Sdt = se, RawXml = el.OuterXml });
            }
        }

        // 页眉/页脚中的图片提取到 doc.Images
        foreach (XmlElement drawing in root.SelectNodes(".//w:drawing", ns)!)
            ParseDrawing(drawing, ns, rels, doc);

        return elements;
    }

    private String? LoadHdrFtrText(String target, Document doc)
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

    private static ParagraphStyle ParseStyleId(String? styleId)
    {
        if (String.IsNullOrEmpty(styleId)) return ParagraphStyle.Normal;
        return styleId!.ToLowerInvariant() switch
        {
            "heading1" or "1" or "heading 1" => ParagraphStyle.Heading1,
            "heading2" or "2" or "heading 2" => ParagraphStyle.Heading2,
            "heading3" or "3" or "heading 3" => ParagraphStyle.Heading3,
            "heading4" or "4" or "heading 4" => ParagraphStyle.Heading4,
            "heading5" or "5" or "heading 5" => ParagraphStyle.Heading5,
            "heading6" or "6" or "heading 6" => ParagraphStyle.Heading6,
            _ => ParagraphStyle.Normal,
        };
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本（段落间换行分隔）</summary>
    /// <returns>纯文本字符串</returns>
    public String? ExtractText() => ReadFullText();

    /// <summary>提取 Markdown 格式（段落格式+表格+图片），结构化 AST 序列化</summary>
    /// <returns>Markdown 字符串</returns>
    public String? ExtractMarkdown() => ToMarkdownDocument()?.ToMarkdown();

    /// <summary>提取 Markdown 文档对象（AST）</summary>
    /// <remarks>
    /// 结构化输出：标题样式 → Heading、项目符号/编号 → 列表、表格 → GFM 表格、
    /// 图片 → 引用形式、SDT 内容 → 段落。解析失败时回退纯文本+表格。
    /// </remarks>
    /// <returns>Markdown 文档对象，无内容返回 null</returns>
    public MarkdownDocument? ToMarkdownDocument()
    {
        try
        {
            var doc = ReadDocument();
            var md = new MarkdownDocument();

            if (doc.Elements.Count == 0)
            {
                FallbackToDocument(md);
                return md.Blocks.Count > 0 ? md : null;
            }

            foreach (var el in doc.Elements)
            {
                if (el.Type == ElementType.Paragraph && el.Paragraph != null)
                {
                    ParagraphToBlocks(md, el.Paragraph, doc);
                }
                else
                {
                    FlushLists(md);
                    switch (el.Type)
                    {
                        case ElementType.Table:
                            var table = TableToBlock(el);
                            if (table != null) md.Blocks.Add(table);
                            break;
                        case ElementType.Image:
                            if (el.Image != null)
                            {
                                var img = ImageToBlock(el.Image);
                                if (img != null) md.Blocks.Add(img);
                            }
                            break;
                        case ElementType.Sdt:
                            if (el.Sdt != null) SdtToBlocks(md, el.Sdt);
                            break;
                    }
                }
            }
            FlushLists(md);

            // 脚注/尾注（W22：AI 知识库场景包含脚注内容）
            AppendNotesToMarkdown(md, doc.Footnotes, "脚注");
            AppendNotesToMarkdown(md, doc.Endnotes, "尾注");

            return md.Blocks.Count > 0 ? md : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>将脚注/尾注追加为 Markdown 引用块（AI 知识库场景）</summary>
    private static void AppendNotesToMarkdown(MarkdownDocument md, List<Footnote> notes, String label)
    {
        if (notes.Count == 0) return;
        var sb = new StringBuilder();
        var added = false;
        foreach (var fn in notes)
        {
            if (String.IsNullOrWhiteSpace(fn.Text)) continue;
            if (sb.Length > 0) sb.Append('\n');
            sb.Append($"{label} {fn.Id}: {fn.Text}");
            added = true;
        }
        if (added)
            md.Blocks.Add(MarkdownBlock.CreateBlockQuote([MarkdownBlock.CreateParagraph([MarkdownInline.CreateText(sb.ToString())])]));
    }

    /// <summary>回退构建：纯段落 + 表格</summary>
    /// <param name="md">目标文档</param>
    private void FallbackToDocument(MarkdownDocument md)
    {
        foreach (var para in ReadParagraphs())
        {
            var text = MarkdownTextCleaner.Clean(para);
            if (!String.IsNullOrWhiteSpace(text))
                md.Blocks.Add(MarkdownBlock.CreateParagraph([MarkdownInline.CreateText(text)]));
        }

        foreach (var table in ReadTables())
        {
            if (table.Length == 0) continue;
            md.Blocks.Add(MarkdownBlockBuilder.Table(table, true));
        }
    }

    /// <summary>段落 → 块（标题/分页/列表/段落，列表项聚合到缓冲）</summary>
    /// <param name="md">目标文档</param>
    /// <param name="para">段落</param>
    /// <param name="doc">文档模型</param>
    private void ParagraphToBlocks(MarkdownDocument md, Paragraph para, Document doc)
    {
        // 标题
        if (para.Style >= ParagraphStyle.Heading1 && para.Style <= ParagraphStyle.Heading6)
        {
            FlushLists(md);
            var level = (Int32)para.Style - (Int32)ParagraphStyle.Heading1 + 1;
            md.Blocks.Add(MarkdownBlock.CreateHeading(level, BuildRunInlines(para.Runs, doc)));
            return;
        }

        // 分页符
        if (para.IsPageBreak)
        {
            FlushLists(md);
            md.Blocks.Add(MarkdownBlock.CreateThematicBreak());
            return;
        }

        // 无序列表
        if (para.IsBullet)
        {
            FlushOrdered(md);
            _bulletItems.Add(MarkdownBlock.CreateListItem(BuildRunInlines(para.Runs, doc)));
            return;
        }

        // 有序列表
        if (para.IsOrderedList)
        {
            FlushBullet(md);
            var num = para.ListStartOverride ?? _orderedNum + 1;
            _orderedItems.Add((num, BuildRunInlines(para.Runs, doc)));
            _orderedNum = num;
            return;
        }

        // 普通段落
        FlushLists(md);
        var inlines = BuildRunInlines(para.Runs, doc);
        // 文本框内容并入（Word 原生导出/搜索均包含文本框文本，W22）
        if (para.TextBoxes.Count > 0)
        {
            foreach (var tp in EnumerateTextBoxesForMd(para))
            {
                var tbInlines = BuildRunInlines(tp.Runs, doc);
                if (tbInlines.Count > 0)
                {
                    if (inlines.Count > 0) inlines.Add(MarkdownInline.CreateText(" "));
                    inlines.AddRange(tbInlines);
                }
            }
        }
        if (inlines.Count > 0)
            md.Blocks.Add(MarkdownBlock.CreateParagraph(inlines));
    }

    /// <summary>递归枚举段落内文本框内容段落（W22，Markdown 转换用）</summary>
    private static IEnumerable<Paragraph> EnumerateTextBoxesForMd(Paragraph para)
    {
        foreach (var tp in para.TextBoxes)
        {
            yield return tp;
            foreach (var nested in EnumerateTextBoxesForMd(tp))
                yield return nested;
        }
    }

    /// <summary>提交所有列表缓冲</summary>
    /// <param name="md">目标文档</param>
    private void FlushLists(MarkdownDocument md)
    {
        FlushBullet(md);
        FlushOrdered(md);
    }

    /// <summary>提交无序列表缓冲</summary>
    /// <param name="md">目标文档</param>
    private void FlushBullet(MarkdownDocument md)
    {
        if (_bulletItems.Count > 0)
        {
            md.Blocks.Add(MarkdownBlock.CreateBulletList(_bulletItems));
            _bulletItems.Clear();
        }
    }

    /// <summary>提交有序列表缓冲</summary>
    /// <param name="md">目标文档</param>
    private void FlushOrdered(MarkdownDocument md)
    {
        if (_orderedItems.Count > 0)
        {
            var start = _orderedItems[0].Item1;
            var items = _orderedItems.Select(x => MarkdownBlock.CreateListItem(x.Item2)).ToArray();
            md.Blocks.Add(MarkdownBlock.CreateOrderedList(items, start));
            _orderedItems.Clear();
        }
        _orderedNum = 0;
    }

    private List<MarkdownInline> BuildRunInlines(List<Run> runs, Document doc)
    {
        var list = new List<MarkdownInline>();
        foreach (var run in runs)
        {
            var text = run.Text;
            if (String.IsNullOrEmpty(text)) continue;

            var props = run.Properties;

            // 超链接
            if (!run.HyperlinkRelId.IsNullOrEmpty() && doc.Hyperlinks.Count > 0)
            {
                var url = doc.Hyperlinks.FirstOrDefault(h => h.RelId == run.HyperlinkRelId).Url;
                if (!url.IsNullOrEmpty())
                {
                    list.Add(MarkdownInline.CreateLink(url!, "", AppendFormattedInlines(text, props)));
                    continue;
                }
            }

            list.AddRange(AppendFormattedInlines(text, props));
        }
        return list;
    }

    private static List<MarkdownInline> AppendFormattedInlines(String text, RunProperties? props)
    {
        if (props == null) return [MarkdownInline.CreateText(MarkdownTextCleaner.Clean(text))];

        List<MarkdownInline> inlines = [MarkdownInline.CreateText(MarkdownTextCleaner.Clean(text))];
        if (props.Bold == true) inlines = [MarkdownInline.CreateStrong(inlines)];
        if (props.Italic == true) inlines = [MarkdownInline.CreateEmphasis(inlines)];
        if (props.Strikethrough == true) inlines = [MarkdownInline.CreateStrikethrough(inlines)];
        return inlines;
    }

    private static MarkdownBlock? TableToBlock(Element el)
    {
        // 优先使用 Table 模型
        if (el.Table != null && el.Table.Rows.Count > 0)
        {
            var table = el.Table;
            var colCount = table.Rows.Max(r => r.Cells.Count);
            if (colCount == 0) return null;

            var rows = new List<String[]>();
            foreach (var row in table.Rows)
            {
                var cells = new String[colCount];
                for (var i = 0; i < colCount; i++)
                    cells[i] = i < row.Cells.Count ? GetCellText(row.Cells[i]) : "";
                rows.Add(cells);
            }
            return MarkdownBlockBuilder.Table(rows.ToArray(), table.FirstRowHeader);
        }

        // 回退到 TableRows（嵌套列表）
        if (el.TableRows != null && el.TableRows.Count > 0)
        {
            var rows = el.TableRows;
            var colCount = rows.Max(r => r.Count);
            if (colCount == 0) return null;

            var arr = new String[rows.Count][];
            for (var r = 0; r < rows.Count; r++)
            {
                arr[r] = new String[colCount];
                for (var i = 0; i < colCount; i++)
                    arr[r][i] = i < rows[r].Count ? GetCellText(rows[r][i]) : "";
            }
            return MarkdownBlockBuilder.Table(arr, el.TableFirstRowHeader);
        }
        return null;
    }

    private static String GetCellText(Cell cell)
    {
        if (cell.Paragraphs.Count == 0 && cell.NestedTables.Count == 0) return String.Empty;
        var sb = new StringBuilder();
        // 嵌套表格文本（W46）
        foreach (var nested in cell.NestedTables)
        {
            foreach (var row in nested.Rows)
            {
                foreach (var nc in row.Cells)
                {
                    sb.Append(GetCellText(nc));
                    sb.Append('\t');
                }
            }
        }
        for (var i = 0; i < cell.Paragraphs.Count; i++)
        {
            var para = cell.Paragraphs[i];
            foreach (var run in para.Runs)
            {
                sb.Append(run.Text);
            }
            if (i < cell.Paragraphs.Count - 1)
                sb.Append(' ');
        }
        return sb.ToString();
    }

    private MarkdownBlock? ImageToBlock(Image image)
    {
        _imageIdx++;
        var ext = image.Extension ?? "png";
        if (ext.IsNullOrEmpty()) ext = "png";

        // 图片以引用形式输出（对标 anydoc：图片渲染为 alt 文本，字节留在模型层）
        var inline = MarkdownInline.CreateImage($"image{_imageIdx}.{ext}", $"Image {_imageIdx}", "");
        return MarkdownBlock.CreateParagraph([inline]);
    }

    private static void SdtToBlocks(MarkdownDocument md, SdtElement sdt)
    {
        if (!sdt.Content.IsNullOrEmpty())
            md.Blocks.Add(MarkdownBlockBuilder.Paragraph(sdt.Content));
    }
    #endregion
}
