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
    /// <param name="paragraph">段落</param>
    /// <param name="text">新文本（不支持换行，需要分段请用 <c>InsertParagraphAfter</c>）</param>
    public static void SetText(this Paragraph paragraph, String text)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        text ??= String.Empty;

        // 清掉段落级 RawXml，避免元素级兜底把改动盖回去
        paragraph.RawXml = null;

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
