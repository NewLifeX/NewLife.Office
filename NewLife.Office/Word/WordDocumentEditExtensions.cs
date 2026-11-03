namespace NewLife.Office.Word;

/// <summary>Word 文档语义化编辑扩展：按语义增删改段落/内容，自动维护多节结构与修改计数</summary>
/// <remarks>
/// 与 <see cref="WordDocumentExtensions"/>（查找替换）、<see cref="WordDocumentQueryExtensions"/>（定位）互补：
/// 本类负责「增删改」，解决只能按 <c>Elements[i]</c> 下标操作、文档结构一变就错位的问题。
/// <para>
/// 所有修改方法都会递增 <see cref="Document.ModelRevision"/> 并标记正文已改，
/// 因此配合 <see cref="Document.BeginEdit"/> 使用时，忘记关闭透传会在保存阶段抛异常，
/// 而不是静默丢弃修改。
/// </para>
/// <example>
/// <code>
/// using var reader = new WordReader("source.docx");
/// var doc = reader.ReadDocument();
/// doc.BeginEdit();
///
/// // 定位（Query 扩展）+ 修改（本扩展）
/// var para = doc.FindParagraph(p => p.Runs.Count > 0 &amp;&amp; p.Runs[0].Text.StartsWith("结论"));
/// if (para != null)
/// {
///     para.SetText("结论：同意上线。");
///     doc.InsertParagraphAfter(para, "补充说明：……", ParagraphStyle.Normal);
/// }
///
/// using var writer = new WordWriter();
/// writer.Save("output.docx", doc);
/// </code>
/// </example>
/// </remarks>
public static class WordDocumentEditExtensions
{
    #region 增
    /// <summary>在正文末尾追加段落</summary>
    /// <param name="doc">文档</param>
    /// <param name="text">文本（含 \n 时按换行拆成多个段落）</param>
    /// <param name="style">段落样式</param>
    /// <returns>首个新增段落的顶层元素</returns>
    public static Element AddParagraph(this Document doc, String text, ParagraphStyle style = ParagraphStyle.Normal)
    {
        var paras = CreateParagraphs(text, style);
        var sec = doc.Sections.Count > 0 ? doc.Sections[^1] : null;
        var first = new Element { Type = ElementType.Paragraph, Paragraph = paras[0] };
        foreach (var para in paras)
        {
            var el = para == paras[0] ? first : new Element { Type = ElementType.Paragraph, Paragraph = para };
            doc.Elements.Add(el);
            sec?.Elements.Add(el);
        }
        doc.MarkModified();
        return first;
    }

    /// <summary>在正文末尾追加标题段落</summary>
    /// <param name="doc">文档</param>
    /// <param name="text">标题文本</param>
    /// <param name="level">级别（1~6，对应 Heading1~Heading6；超界自动收敛）</param>
    /// <returns>新增段落的顶层元素</returns>
    public static Element AddHeading(this Document doc, String text, Int32 level = 1)
    {
        if (level < 1) level = 1;
        if (level > 6) level = 6;
        return doc.AddParagraph(text, (ParagraphStyle)level);
    }

    /// <summary>在正文末尾追加空表格</summary>
    /// <param name="doc">文档</param>
    /// <param name="rows">行数（至少 1）</param>
    /// <param name="cols">列数（至少 1）</param>
    /// <param name="styleId">表格样式 ID（如 TableGrid），null 表示不指定</param>
    /// <returns>新增表格的顶层元素</returns>
    /// <exception cref="ArgumentOutOfRangeException">行数或列数小于 1</exception>
    public static Element AddTable(this Document doc, Int32 rows, Int32 cols, String? styleId = "TableGrid")
    {
        if (rows < 1) throw new ArgumentOutOfRangeException(nameof(rows));
        if (cols < 1) throw new ArgumentOutOfRangeException(nameof(cols));

        var table = new Table { StyleId = styleId, FirstRowHeader = false };
        for (var r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (var c = 0; c < cols; c++)
                row.Cells.Add(new Cell { Paragraphs = { new Paragraph() } });
            table.Rows.Add(row);
        }
        return doc.AppendTopLevel(new Element { Type = ElementType.Table, Table = table });
    }

    /// <summary>用二维数组在正文末尾追加表格（首行默认作表头）</summary>
    /// <param name="doc">文档</param>
    /// <param name="rows">行数据（每行一个字符串数组，列数以最长行为准，缺列补空单元格）</param>
    /// <param name="firstRowHeader">首行是否作表头（底色加粗 + 跨页重复）</param>
    /// <param name="styleId">表格样式 ID</param>
    /// <returns>新增表格的顶层元素</returns>
    /// <exception cref="ArgumentNullException">rows 为空</exception>
    /// <exception cref="ArgumentException">rows 无有效行</exception>
    public static Element AddTable(this Document doc, IEnumerable<String[]> rows, Boolean firstRowHeader = true, String? styleId = "TableGrid")
    {
        if (rows == null) throw new ArgumentNullException(nameof(rows));

        var data = rows.ToList();
        if (data.Count == 0) throw new ArgumentException(L.T("Table rows are empty."), nameof(rows));

        var cols = data.Max(r => r?.Length ?? 0);
        var table = new Table { StyleId = styleId, FirstRowHeader = firstRowHeader };
        for (var r = 0; r < data.Count; r++)
        {
            var row = new TableRow { IsHeader = firstRowHeader && r == 0 };
            for (var c = 0; c < cols; c++)
            {
                var text = c < (data[r]?.Length ?? 0) ? data[r][c] : null;
                var para = new Paragraph();
                if (!String.IsNullOrEmpty(text)) para.Runs.Add(new Run { Text = text });
                row.Cells.Add(new Cell { Paragraphs = { para } });
            }
            table.Rows.Add(row);
        }
        return doc.AppendTopLevel(new Element { Type = ElementType.Table, Table = table });
    }

    /// <summary>在正文末尾追加分页符</summary>
    /// <param name="doc">文档</param>
    /// <returns>分页符所在段落的顶层元素</returns>
    public static Element AddPageBreak(this Document doc)
        => doc.AppendTopLevel(new Element { Type = ElementType.Paragraph, Paragraph = new Paragraph { IsPageBreak = true } });

    /// <summary>在正文末尾追加图片：按像素与 DPI 计算自然尺寸，只给一个方向时按比例缩放</summary>
    /// <param name="doc">文档</param>
    /// <param name="data">图片字节（PNG/JPEG/GIF/BMP/SVG）</param>
    /// <param name="extension">扩展名（不含点），null 时按文件头识别，无法识别按 png</param>
    /// <param name="width">显示宽度，null 表示用自然宽度</param>
    /// <param name="height">显示高度，null 表示用自然高度或按宽高等比换算</param>
    /// <param name="altText">替代文本（无障碍/检索）</param>
    /// <returns>新增的图片模型（<see cref="Image.RelId"/> 已分配到 <see cref="Document.Images"/></returns>
    /// <example>
    /// <code>
    /// doc.BeginEdit();
    /// doc.AddPicture("logo.png", width: Length.FromCm(4));   // 只给宽度，高度按比例
    /// doc.AddPictureBase64("data:image/png;base64,iVBORw0...");
    /// </code>
    /// </example>
    public static Image AddPicture(this Document doc, Byte[] data, String? extension = null, Length? width = null, Length? height = null, String? altText = null)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0) throw new ArgumentException(L.T("Image data is empty."), nameof(data));

        var ext = extension?.TrimStart('.').ToLowerInvariant();
        if (String.IsNullOrEmpty(ext)) ext = ImageInfo.DetectExtension(data) ?? "png";

        // 自然尺寸：像素按 DPI 折算为物理尺寸
        Length? natW = null;
        Length? natH = null;
        if (ImageInfo.TryGetPixelSize(data, out var px, out var py, out var dpi) && px > 0 && py > 0)
        {
            natW = Length.FromInches(px / dpi);
            natH = Length.FromInches(py / dpi);
        }

        Length w, h;
        if (width == null && height == null)
        {
            w = natW ?? Length.FromCm(10);
            h = natH ?? Length.FromCm(7.5);
        }
        else if (width != null && height == null)
        {
            // 只给宽度：按自然宽高比换算高度，无法探测时按 4:3
            w = width.Value;
            h = natW is { } nw && natH is { } nh && nw.Emu > 0 ? Length.FromEmu(w.Emu * nh.Emu / nw.Emu) : Length.FromEmu(w.Emu * 0.75);
        }
        else if (width == null && height != null)
        {
            // 只给高度：按自然宽高比换算宽度，无法探测时按 4:3
            h = height.Value;
            w = natW is { } nw2 && natH is { } nh2 && nh2.Emu > 0 ? Length.FromEmu(h.Emu * nw2.Emu / nh2.Emu) : Length.FromEmu(h.Emu * 4 / 3);
        }
        else
        {
            w = width!.Value;
            h = height!.Value;
        }

        // 关系 ID 必须在文档内唯一
        var index = 1;
        String relId;
        do { relId = $"rPic{index++}"; } while (doc.Images.ContainsKey(relId));

        doc.Images[relId] = (ext, data);
        var img = new Image
        {
            ImageData = data,
            Extension = ext,
            RelId = relId,
            WidthEmu = w.ToEmuInt64(),
            HeightEmu = h.ToEmuInt64(),
            IsSvg = ext == "svg",
            AnchorType = "inline",
            AltText = altText,
        };
        doc.AppendTopLevel(new Element { Type = ElementType.Image, Image = img });
        return img;
    }

    /// <summary>从文件追加图片</summary>
    /// <param name="doc">文档</param>
    /// <param name="path">图片文件路径</param>
    /// <param name="width">显示宽度，null 表示用自然宽度</param>
    /// <param name="height">显示高度，null 表示用自然高度或等比换算</param>
    /// <param name="altText">替代文本</param>
    /// <returns>新增的图片模型</returns>
    public static Image AddPicture(this Document doc, String path, Length? width = null, Length? height = null, String? altText = null)
    {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return doc.AddPicture(File.ReadAllBytes(path.GetFullPath()), Path.GetExtension(path), width, height, altText);
    }

    /// <summary>从流追加图片（流不会被关闭）</summary>
    /// <param name="doc">文档</param>
    /// <param name="stream">图片流</param>
    /// <param name="extension">扩展名，null 时按文件头识别</param>
    /// <param name="width">显示宽度，null 表示用自然宽度</param>
    /// <param name="height">显示高度，null 表示用自然高度或等比换算</param>
    /// <param name="altText">替代文本</param>
    /// <returns>新增的图片模型</returns>
    public static Image AddPicture(this Document doc, Stream stream, String? extension = null, Length? width = null, Length? height = null, String? altText = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return doc.AddPicture(ms.ToArray(), extension, width, height, altText);
    }

    /// <summary>从 Base64（含 data URI）追加图片</summary>
    /// <param name="doc">文档</param>
    /// <param name="base64">Base64 文本，或 <c>data:image/png;base64,...</c> 形式</param>
    /// <param name="extension">扩展名，null 时从 data URI 或文件头识别</param>
    /// <param name="width">显示宽度，null 表示用自然宽度</param>
    /// <param name="height">显示高度，null 表示用自然高度或等比换算</param>
    /// <param name="altText">替代文本</param>
    /// <returns>新增的图片模型</returns>
    /// <exception cref="ArgumentNullException">base64 为空</exception>
    /// <exception cref="FormatException">data URI 缺少 base64 标记，或 Base64 内容非法</exception>
    public static Image AddPictureBase64(this Document doc, String base64, String? extension = null, Length? width = null, Length? height = null, String? altText = null)
    {
        if (String.IsNullOrEmpty(base64)) throw new ArgumentNullException(nameof(base64));

        var text = base64.Trim();
        if (text.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            var idx = text.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) throw new FormatException(L.T("Unsupported data URI (missing 'base64,' marker)."));

            var mime = text[5..idx].TrimEnd(';');
            if (extension == null && mime.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                extension = mime[6..];
            text = text[(idx + 7)..];
        }
        return doc.AddPicture(Convert.FromBase64String(text), extension, width, height, altText);
    }

    /// <summary>追加新节：此后追加的内容归属新节，返回新节</summary>
    /// <param name="doc">文档</param>
    /// <param name="startType">节起始类型：nextPage（默认）/ continuous / evenPage / oddPage（传 newPage 会自动纠正为 nextPage）</param>
    /// <returns>新节（其 <see cref="Section.PageSettings"/> 为当前页设置的副本，可单独调整）</returns>
    /// <remarks>调用后文档进入多节输出模式，保存前需 <see cref="Document.BeginEdit"/>（与其它模型修改一致）。</remarks>
    public static Section AddSection(this Document doc, String startType = "nextPage")
    {
        // 保证已有内容归属第一节：程序化创建时 Sections 为空，需要先把当前元素收进首节（复制列表，避免后续追加同时落入两节）
        if (doc.Sections.Count == 0)
            doc.Sections.Add(new Section { Elements = new List<Element>(doc.Elements), PageSettings = doc.PageSettings.Clone() });
        else if (doc.Sections.Count == 1 && doc.Sections[0].Elements.Count != doc.Elements.Count)
            doc.Sections[0].Elements = new List<Element>(doc.Elements);

        var sec = new Section { Type = startType, PageSettings = doc.PageSettings.Clone() };
        doc.Sections.Add(sec);
        doc.MarkModified();
        return sec;
    }

    /// <summary>在正文指定位置插入元素（自动同步所属节，多节文档输出以节为准）</summary>
    /// <param name="doc">文档</param>
    /// <param name="index">插入位置（0=最前，超出则追加到末尾）</param>
    /// <param name="element">待插入元素</param>
    /// <returns>同一个元素（便于链式使用）</returns>
    /// <exception cref="ArgumentNullException">element 为空</exception>
    public static Element InsertElementAt(this Document doc, Int32 index, Element element)
    {
        if (element == null) throw new ArgumentNullException(nameof(element));

        index = Math.Max(0, Math.Min(index, doc.Elements.Count));
        var anchorEl = index > 0 ? doc.Elements[index - 1] : null;
        doc.Elements.Insert(index, element);

        if (doc.Sections.Count > 0)
        {
            var inserted = false;
            if (anchorEl != null)
            {
                foreach (var sec in doc.Sections)
                {
                    var i = sec.Elements.IndexOf(anchorEl);
                    if (i < 0) continue;

                    sec.Elements.Insert(i + 1, element);
                    inserted = true;
                    break;
                }
            }
            if (!inserted) doc.Sections[0].Elements.Insert(0, element);
        }

        doc.MarkModified();
        return element;
    }

    /// <summary>把元素追加到正文末尾，并同步到最后一节（多节文档的输出来源）</summary>
    /// <param name="doc">文档</param>
    /// <param name="element">待追加元素</param>
    /// <returns>同一个元素（便于链式使用）</returns>
    private static Element AppendTopLevel(this Document doc, Element element)
    {
        doc.Elements.Add(element);
        if (doc.Sections.Count > 0) doc.Sections[^1].Elements.Add(element);
        doc.MarkModified();
        return element;
    }

    /// <summary>在指定段落之后插入新段落（自动适配正文/单元格/页眉页脚/脚注容器）</summary>
    /// <param name="doc">文档</param>
    /// <param name="anchor">锚点段落</param>
    /// <param name="text">文本（含 \n 时按换行拆成多个段落）</param>
    /// <param name="style">段落样式</param>
    /// <returns>首个新增段落，锚点不在文档中返回 null</returns>
    public static Paragraph? InsertParagraphAfter(this Document doc, Paragraph anchor, String text, ParagraphStyle style = ParagraphStyle.Normal)
        => doc.InsertParagraphCore(anchor, text, style, after: true);

    /// <summary>在指定段落之前插入新段落（自动适配正文/单元格/页眉页脚/脚注容器）</summary>
    /// <param name="doc">文档</param>
    /// <param name="anchor">锚点段落</param>
    /// <param name="text">文本（含 \n 时按换行拆成多个段落）</param>
    /// <param name="style">段落样式</param>
    /// <returns>首个新增段落，锚点不在文档中返回 null</returns>
    public static Paragraph? InsertParagraphBefore(this Document doc, Paragraph anchor, String text, ParagraphStyle style = ParagraphStyle.Normal)
        => doc.InsertParagraphCore(anchor, text, style, after: false);

    private static Paragraph? InsertParagraphCore(this Document doc, Paragraph anchor, String text, ParagraphStyle style, Boolean after)
    {
        if (anchor == null) throw new ArgumentNullException(nameof(anchor));
        var paras = CreateParagraphs(text, style);

        // 1) 正文顶层：插入为 Element，并同步所属节的元素列表（多节输出以节为准）
        for (var i = 0; i < doc.Elements.Count; i++)
        {
            if (!ReferenceEquals(doc.Elements[i].Paragraph, anchor)) continue;

            var anchorEl = doc.Elements[i];
            var at = after ? i + 1 : i;
            var elements = new List<Element>();
            foreach (var para in paras)
            {
                var el = new Element { Type = ElementType.Paragraph, Paragraph = para };
                elements.Add(el);
                doc.Elements.Insert(at++, el);
            }
            SyncSectionInsert(doc, anchorEl, elements, after);
            doc.MarkModified();
            return paras[0];
        }

        // 2) 其他容器：单元格段落、页眉页脚、脚注尾注
        foreach (var list in EnumerateParagraphLists(doc))
        {
            var idx = list.IndexOf(anchor);
            if (idx < 0) continue;
            var at = after ? idx + 1 : idx;
            foreach (var para in paras) list.Insert(at++, para);
            doc.MarkModified();
            return paras[0];
        }

        return null;
    }

    /// <summary>把新元素同步插入所属节的元素列表（多节文档的输出来源）</summary>
    private static void SyncSectionInsert(Document doc, Element anchor, List<Element> inserted, Boolean after)
    {
        foreach (var sec in doc.Sections)
        {
            var idx = sec.Elements.IndexOf(anchor);
            if (idx < 0) continue;
            var at = after ? idx + 1 : idx;
            foreach (var el in inserted) sec.Elements.Insert(at++, el);
            return;
        }
    }
    #endregion

    #region 删
    /// <summary>移除顶层元素（同时从所属节移除）</summary>
    /// <param name="doc">文档</param>
    /// <param name="element">元素</param>
    /// <returns>成功移除返回 true</returns>
    public static Boolean RemoveElement(this Document doc, Element element)
    {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (!doc.Elements.Remove(element)) return false;
        foreach (var sec in doc.Sections)
            sec.Elements.Remove(element);
        doc.MarkModified();
        return true;
    }

    /// <summary>移除段落（正文段落移除整个顶层元素；单元格/页眉页脚/脚注段落只移除该段）</summary>
    /// <param name="doc">文档</param>
    /// <param name="paragraph">段落</param>
    /// <returns>成功移除返回 true</returns>
    public static Boolean RemoveParagraph(this Document doc, Paragraph paragraph)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));

        foreach (var el in doc.Elements)
        {
            if (!ReferenceEquals(el.Paragraph, paragraph)) continue;
            return doc.RemoveElement(el);
        }

        foreach (var list in EnumerateParagraphLists(doc))
        {
            if (!list.Remove(paragraph)) continue;
            doc.MarkModified();
            return true;
        }

        return false;
    }
    #endregion

    #region 改
    /// <summary>设置段落文本，保留首个 Run 的格式并删除其余 Run</summary>
    /// <remarks>
    /// 若原段落是“纯分页符段”（<see cref="Paragraph.IsPageBreak"/>），会被清掉分页标记——
    /// 否则 Writer 会把它整体当分页符输出，新文本被静默丢弃。</remarks>
    /// <param name="paragraph">段落</param>
    /// <param name="text">新文本（不支持换行，需要分段请用 <c>InsertParagraphAfter</c>）</param>
    public static void SetText(this Paragraph paragraph, String text)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        text ??= String.Empty;

        // 清掉段落级 RawXml，避免元素级兜底把改动盖回去
        paragraph.RawXml = null;
        // 替换文本后本段落不再是纯分页段（否则 Writer 只输出 w:br，文本丢失）
        paragraph.IsPageBreak = false;

        if (paragraph.Runs.Count == 0)
        {
            paragraph.Runs.Add(new Run { Text = text });
            return;
        }

        paragraph.Runs[0].Text = text;
        for (var i = paragraph.Runs.Count - 1; i >= 1; i--)
            paragraph.Runs.RemoveAt(i);
    }

    /// <summary>设置单元格文本：保留首段首 Run 格式，清空其余内容构成“整格替换”</summary>
    /// <param name="cell">单元格</param>
    /// <param name="text">新文本</param>
    public static void SetText(this Cell cell, String text)
    {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (cell.Paragraphs.Count == 0) cell.Paragraphs.Add(new Paragraph());

        cell.Paragraphs[0].SetText(text);
        for (var i = cell.Paragraphs.Count - 1; i >= 1; i--)
            cell.Paragraphs.RemoveAt(i);
        cell.NestedTables.Clear();
    }
    #endregion

    #region 批注
    /// <summary>给段落加批注（AI 审阅场景）：批注锚定该段全部 Run</summary>
    /// <param name="doc">文档</param>
    /// <param name="paragraph">目标段落</param>
    /// <param name="text">批注内容</param>
    /// <param name="author">作者，默认 NewLife.Office</param>
    /// <param name="initials">作者缩写（可选）</param>
    /// <param name="date">时间，默认当前时间</param>
    /// <returns>新增的批注模型</returns>
    /// <exception cref="ArgumentNullException">doc 或 paragraph 为空</exception>
    /// <example>
    /// <code>
    /// doc.BeginEdit();
    /// var para = doc.FindParagraph(p => p.Runs[0].Text.Contains("待确认"));
    /// if (para != null) doc.AddComment(para, "此结论缺少数据支撑，请补充。", "AI 审阅");
    /// </code>
    /// </example>
    public static Comment AddComment(this Document doc, Paragraph paragraph, String text, String? author = null, String? initials = null, DateTime? date = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));

        // 文档内唯一 ID
        var id = 0;
        foreach (var c in doc.Comments)
            if (c.Id >= id) id = c.Id + 1;

        var comment = new Comment
        {
            Id = id,
            Text = text,
            Author = author ?? "NewLife.Office",
            Initials = initials,
            Date = date ?? DateTime.Now,
        };
        doc.Comments.Add(comment);

        // 锚定整段（空段落也输出锚点，由 Writer 特殊处理）
        paragraph.RawXml = null;
        paragraph.CommentRanges.Add(new CommentRange
        {
            Id = id,
            StartRunIndex = 0,
            EndRunIndex = Math.Max(0, paragraph.Runs.Count - 1),
        });

        doc.MarkModified();
        return comment;
    }
    #endregion

    #region 修订
    /// <summary>开启修订建议模式：设置后续修订的作者与时间</summary>
    /// <param name="doc">文档</param>
    /// <param name="author">修订作者（写入 w:ins/w:del 的 w:author）</param>
    /// <param name="date">修订时间；null 表示保存时刻</param>
    public static void StartRevisions(this Document doc, String author, DateTime? date = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (String.IsNullOrEmpty(author)) throw new ArgumentException(L.T("The author is empty."), nameof(author));

        doc.RevisionsAuthor = author;
        doc.RevisionsDate = date;
    }

    /// <summary>把段落改成"删除原文 + 插入新文"的修订建议（Word 中显示为建议修改，可接受/拒绝）</summary>
    /// <param name="doc">文档</param>
    /// <param name="paragraph">目标段落</param>
    /// <param name="text">新文本</param>
    /// <param name="author">修订作者；null 表示沿用 <see cref="Document.RevisionsAuthor"/></param>
    /// <param name="date">修订时间；null 表示沿用 <see cref="Document.RevisionsDate"/></param>
    /// <returns>同一个段落（便于链式使用）</returns>
    /// <exception cref="ArgumentNullException">doc 或 paragraph 为空</exception>
    /// <example>
    /// <code>
    /// doc.BeginEdit();
    /// doc.StartRevisions("AI 审阅");
    /// var para = doc.FindParagraph(p => p.Runs[0].Text.Contains("可能"));
    /// if (para != null) doc.SetTextAsRevision(para, "建议改为：必然。");
    /// </code>
    /// </example>
    public static Paragraph SetTextAsRevision(this Document doc, Paragraph paragraph, String text, String? author = null, DateTime? date = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));

        if (author != null) doc.RevisionsAuthor = author;
        if (date.HasValue) doc.RevisionsDate = date;

        // 原文整段标为删除（保留各自格式），新文标为插入追加在后
        var style = paragraph.Runs.Count > 0 ? paragraph.Runs[0].Properties : null;
        foreach (var run in paragraph.Runs)
        {
            if (!String.IsNullOrEmpty(run.Text)) run.Revision = TrackChangeType.Delete;
        }

        paragraph.RawXml = null;
        paragraph.Runs.Add(new Run { Text = text ?? String.Empty, Properties = style, Revision = TrackChangeType.Insert });

        doc.MarkModified();
        return paragraph;
    }
    #endregion

    #region 页眉页脚
    /// <summary>在页脚追加页码（PAGE / NUMPAGES 域），并开启打开时更新域</summary>
    /// <param name="doc">文档</param>
    /// <param name="prefix">页码前缀文本（如"第 "）</param>
    /// <param name="middle">页码与总页数之间的分隔文本（如" 页 / 共 "）</param>
    /// <param name="suffix">总页数后缀文本（如" 页"）</param>
    /// <param name="type">页脚类型：default（默认）/ first（首页）/ even（偶数页）</param>
    /// <returns>新增的页脚段落</returns>
    /// <remarks>
    /// 页码是域而非静态文本：Word/WPS 打开时按 <see cref="Document.UpdateFields"/> 刷新。
    /// 已有 <paramref name="type"/> 页脚时追加到其末尾，否则新建该类型页脚。
    /// </remarks>
    public static Paragraph AddPageNumber(this Document doc, String prefix = "第 ", String middle = " 页 / 共 ", String suffix = " 页", String type = "default")
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        var footer = doc.Footers.Find(f => String.Equals(f.Type, type, StringComparison.OrdinalIgnoreCase));
        if (footer == null)
        {
            footer = new Footer { Type = type };
            doc.Footers.Add(footer);
        }

        var para = new Paragraph { Alignment = "center" };
        if (!String.IsNullOrEmpty(prefix)) para.Runs.Add(new Run { Text = prefix });
        para.AddField("PAGE", "1");
        if (!String.IsNullOrEmpty(middle)) para.Runs.Add(new Run { Text = middle });
        para.AddField("NUMPAGES", "1");
        if (!String.IsNullOrEmpty(suffix)) para.Runs.Add(new Run { Text = suffix });

        footer.Elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });

        doc.UpdateFields = true;
        doc.MarkModified();
        return para;
    }
    #endregion

    #region 域
    /// <summary>在当前段落末尾追加域（复杂域：begin/instrText/separate/结果/end）</summary>
    /// <param name="paragraph">段落</param>
    /// <param name="code">域指令，如 <c>PAGE \* MERGEFORMAT</c>、<c>NUMPAGES</c>、<c>DATE \@ yyyy-MM-dd</c></param>
    /// <param name="result">缓存显示文本；null 时留空，Word 打开后按 <see cref="Document.UpdateFields"/> 刷新</param>
    /// <param name="properties">结果的字符格式</param>
    /// <returns>新增的域 Run</returns>
    /// <exception cref="ArgumentNullException">paragraph 为空</exception>
    /// <exception cref="ArgumentException">code 为空</exception>
    /// <example>
    /// <code>
    /// doc.BeginEdit();
    /// var para = doc.AddParagraph("第 ");
    /// para.Paragraph!.AddField("PAGE", "1");
    /// para.Paragraph!.Runs.Add(new Run { Text = " 页" });
    /// doc.UpdateFields = true;
    /// </code>
    /// </example>
    public static Run AddField(this Paragraph paragraph, String code, String? result = null, RunProperties? properties = null)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (String.IsNullOrEmpty(code)) throw new ArgumentException(L.T("The field code is empty."), nameof(code));

        paragraph.RawXml = null;
        var run = new Run { FieldCode = code, FieldResult = result, Properties = properties };
        paragraph.Runs.Add(run);
        return run;
    }

    /// <summary>在正文末尾追加只含一个域的段落</summary>
    /// <param name="doc">文档</param>
    /// <param name="code">域指令</param>
    /// <param name="result">缓存显示文本</param>
    /// <param name="style">段落样式</param>
    /// <returns>新增段落的顶层元素</returns>
    public static Element AddField(this Document doc, String code, String? result = null, ParagraphStyle style = ParagraphStyle.Normal)
    {
        if (String.IsNullOrEmpty(code)) throw new ArgumentException(L.T("The field code is empty."), nameof(code));

        var para = new Paragraph { Style = style };
        para.AddField(code, result);
        return doc.AppendTopLevel(new Element { Type = ElementType.Paragraph, Paragraph = para });
    }

    /// <summary>在正文末尾追加目录（TOC 域），并开启打开时更新域</summary>
    /// <param name="doc">文档</param>
    /// <param name="levels">收录的标题级别（1~9），默认 3 级</param>
    /// <param name="title">目录标题文本（如"目录"）；null 表示不输出标题段落</param>
    /// <returns>目录域所在段落的顶层元素</returns>
    /// <remarks>
    /// 目录条目由 Word/WPS 打开时按标题样式生成（<see cref="Document.UpdateFields"/> 已置 true）。
    /// 需要离线即可见的目录，请自行按标题生成静态条目。
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">levels 不在 1~9</exception>
    public static Element AddTableOfContents(this Document doc, Int32 levels = 3, String? title = null)
    {
        if (levels < 1) levels = 1;
        if (levels > 9) levels = 9;

        if (!String.IsNullOrEmpty(title))
            doc.AddParagraph(title!, ParagraphStyle.Heading1);

        doc.UpdateFields = true;
        return doc.AddField($"TOC \\o \"1-{levels}\" \\h \\z \\u");
    }
    #endregion

    #region 目录
    /// <summary>插入静态目录：按标题段落生成缩进条目，离线可见、不依赖 Word 更新域</summary>
    /// <param name="doc">文档</param>
    /// <param name="levels">收录的标题级别（1~6），默认 3 级</param>
    /// <param name="title">目录标题文本（null 表示不输出标题行）</param>
    /// <param name="insertAt">插入位置（0=文档最前，默认）</param>
    /// <returns>插入的标题元素（无标题可收录时返回 null）</returns>
    /// <remarks>
    /// 条目为纯文本段落（按级别缩进），不含超链接跳转；需要可点击/自动更新的目录请用 <see cref="AddTableOfContents"/>。
    /// </remarks>
    /// <example>
    /// <code>
    /// doc.BeginEdit();
    /// doc.AddStaticTableOfContents(levels: 3, title: "目录");
    /// </code>
    /// </example>
    public static Element? AddStaticTableOfContents(this Document doc, Int32 levels = 3, String? title = "目录", Int32 insertAt = 0)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (levels < 1) levels = 1;
        if (levels > 6) levels = 6;

        // 只收录正文顶层的标题段落
        var headings = new List<(Int32 Level, String Text)>();
        foreach (var el in doc.Elements)
        {
            var para = el.Paragraph;
            if (para == null) continue;
            if (para.Style < ParagraphStyle.Heading1 || para.Style > ParagraphStyle.Heading6) continue;

            var level = (Int32)para.Style;
            if (level > levels) continue;
            headings.Add((level, BuildParagraphText(para)));
        }
        if (headings.Count == 0) return null;

        var elements = new List<Element>();

        // 标题行：加粗普通段落，避免它自己被算作目录条目
        if (!String.IsNullOrEmpty(title))
        {
            var titlePara = new Paragraph();
            titlePara.Runs.Add(new Run { Text = title!, Properties = new RunProperties { Bold = true, FontSize = 14 } });
            elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = titlePara });
        }

        foreach (var (level, text) in headings)
        {
            var para = new Paragraph();
            if (level > 1) para.IndentLeft = (level - 1) * 360;
            if (text.Length > 0) para.Runs.Add(new Run { Text = text });
            elements.Add(new Element { Type = ElementType.Paragraph, Paragraph = para });
        }

        for (var i = 0; i < elements.Count; i++)
            doc.InsertElementAt(insertAt + i, elements[i]);

        return elements.Count > 0 ? elements[0] : null;
    }

    /// <summary>拼接段落纯文本（域 Run 取缓存结果）</summary>
    private static String BuildParagraphText(Paragraph para)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var run in para.Runs)
        {
            if (!String.IsNullOrEmpty(run.FieldCode))
            {
                if (!String.IsNullOrEmpty(run.FieldResult)) sb.Append(run.FieldResult);
                continue;
            }
            sb.Append(run.Text);
        }
        return sb.ToString();
    }
    #endregion

    #region 辅助
    /// <summary>按换行拆分为多个段落（OOXML 的 Run 不承载换行，只能分段表达）</summary>
    private static List<Paragraph> CreateParagraphs(String? text, ParagraphStyle style)
    {
        var list = new List<Paragraph>();
        var lines = (text ?? String.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        foreach (var line in lines)
        {
            var para = new Paragraph { Style = style };
            if (line.Length > 0) para.Runs.Add(new Run { Text = line });
            list.Add(para);
        }
        return list;
    }

    /// <summary>枚举可承载段落的列表（单元格、嵌套表格单元格、页眉页脚、脚注尾注）</summary>
    private static List<List<Paragraph>> EnumerateParagraphLists(Document doc)
    {
        var lists = new List<List<Paragraph>>();

        void AddCell(Cell cell)
        {
            lists.Add(cell.Paragraphs);
            foreach (var nested in cell.NestedTables)
                foreach (var row in nested.Rows)
                    foreach (var c in row.Cells) AddCell(c);
        }

        void AddTable(Table table)
        {
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells) AddCell(cell);
        }

        void AddElements(List<Element> elements)
        {
            foreach (var el in elements)
            {
                if (el.Table != null) AddTable(el.Table);
                else if (el.TableRows != null)
                {
                    foreach (var row in el.TableRows)
                        foreach (var cell in row) AddCell(cell);
                }
            }
        }

        AddElements(doc.Elements);
        foreach (var h in doc.Headers) AddElements(h.Elements);
        foreach (var f in doc.Footers) AddElements(f.Elements);
        foreach (var fn in doc.Footnotes) lists.Add(fn.Paragraphs);
        foreach (var en in doc.Endnotes) lists.Add(en.Paragraphs);

        return lists;
    }
    #endregion
}
