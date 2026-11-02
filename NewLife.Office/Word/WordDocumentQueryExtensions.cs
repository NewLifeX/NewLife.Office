namespace NewLife.Office.Word;

/// <summary>Word 文档语义化定位扩展：按语义查找段落/表格/元素，替代按 <c>Elements[i]</c> 下标硬编码</summary>
/// <remarks>
/// 与 <see cref="WordDocumentEditExtensions"/>（增删改）成对使用：
/// 本类只读，不修改文档；增删改在编辑扩展中，且会自动维护 <see cref="Document.Sections"/> 与修改计数。
/// <example>
/// <code>
/// using var reader = new WordReader("source.docx");
/// var doc = reader.ReadDocument();
/// doc.BeginEdit();
///
/// // 定位：找到第一个以"结论"开头的段落
/// var para = doc.FindParagraph(p => p.Runs.Count > 0 &amp;&amp; p.Runs[0].Text.StartsWith("结论"));
/// if (para != null) para.SetText("结论：同意上线。");
///
/// using var writer = new WordWriter();
/// writer.Save("output.docx", doc);
/// </code>
/// </example>
/// </remarks>
public static class WordDocumentQueryExtensions
{
    #region 定位
    /// <summary>枚举全部段落（正文、表格单元格含嵌套、页眉页脚、脚注尾注）</summary>
    /// <param name="doc">文档</param>
    /// <returns>按文档顺序排列的段落列表</returns>
    public static List<Paragraph> AllParagraphs(this Document doc)
    {
        var list = new List<Paragraph>();

        void AddCell(Cell cell)
        {
            list.AddRange(cell.Paragraphs);
            foreach (var nested in cell.NestedTables) AddTable(nested);
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
                if (el.Paragraph != null) list.Add(el.Paragraph);
                if (el.Table != null)
                    AddTable(el.Table);
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
        foreach (var fn in doc.Footnotes) list.AddRange(fn.Paragraphs);
        foreach (var en in doc.Endnotes) list.AddRange(en.Paragraphs);

        return list;
    }

    /// <summary>枚举全部表格（顶层表格 + 单元格内嵌套表格）</summary>
    /// <param name="doc">文档</param>
    /// <returns>表格列表</returns>
    public static List<Table> AllTables(this Document doc)
    {
        var list = new List<Table>();

        void AddCell(Cell cell)
        {
            foreach (var nested in cell.NestedTables)
            {
                list.Add(nested);
                foreach (var row in nested.Rows)
                    foreach (var c in row.Cells) AddCell(c);
            }
        }

        foreach (var el in doc.Elements)
        {
            if (el.Table == null) continue;
            list.Add(el.Table);
            foreach (var row in el.Table.Rows)
                foreach (var cell in row.Cells) AddCell(cell);
        }

        return list;
    }

    /// <summary>按条件查找顶层元素</summary>
    /// <param name="doc">文档</param>
    /// <param name="predicate">匹配条件</param>
    /// <returns>匹配的元素列表（无匹配为空列表）</returns>
    public static List<Element> FindElements(this Document doc, Func<Element, Boolean> predicate)
        => doc.Elements.FindAll(new Predicate<Element>(predicate));

    /// <summary>按条件查找第一个匹配段落（范围同 <see cref="AllParagraphs"/>）</summary>
    /// <param name="doc">文档</param>
    /// <param name="predicate">匹配条件</param>
    /// <returns>匹配段落，无匹配返回 null</returns>
    public static Paragraph? FindParagraph(this Document doc, Func<Paragraph, Boolean> predicate)
        => doc.AllParagraphs().Find(p => predicate(p));

    /// <summary>按条件查找第一个匹配表格（范围同 <see cref="AllTables"/>）</summary>
    /// <param name="doc">文档</param>
    /// <param name="predicate">匹配条件</param>
    /// <returns>匹配表格，无匹配返回 null</returns>
    public static Table? FindTable(this Document doc, Func<Table, Boolean> predicate)
        => doc.AllTables().Find(t => predicate(t));

    /// <summary>获取元素在正文中的序号（-1 表示不在正文顶层）</summary>
    /// <param name="doc">文档</param>
    /// <param name="element">元素</param>
    /// <returns>序号</returns>
    public static Int32 IndexOfElement(this Document doc, Element element) => doc.Elements.IndexOf(element);

    /// <summary>获取元素所属的节（找不到返回 null，表示单节文档或元素不在节列表中）</summary>
    /// <param name="doc">文档</param>
    /// <param name="element">元素</param>
    /// <returns>所属节</returns>
    public static Section? GetSectionOf(this Document doc, Element element)
    {
        foreach (var sec in doc.Sections)
            if (sec.Elements.Contains(element)) return sec;
        return null;
    }

    /// <summary>获取段落所在的顶层元素（表格内段落返回其所属表格元素；找不到返回 null）</summary>
    /// <param name="doc">文档</param>
    /// <param name="paragraph">段落</param>
    /// <returns>顶层元素</returns>
    public static Element? GetElementOf(this Document doc, Paragraph paragraph)
    {
        foreach (var el in doc.Elements)
        {
            if (ReferenceEquals(el.Paragraph, paragraph)) return el;
            if (el.Table != null && ContainsParagraph(el.Table, paragraph)) return el;
            if (el.TableRows != null && el.TableRows.Exists(row => row.Exists(cell => ContainsParagraph(cell, paragraph)))) return el;
        }
        return null;
    }

    private static Boolean ContainsParagraph(Table table, Paragraph paragraph)
    {
        foreach (var row in table.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.Paragraphs.Contains(paragraph)) return true;
                foreach (var nested in cell.NestedTables)
                    if (ContainsParagraph(nested, paragraph)) return true;
            }
        return false;
    }

    private static Boolean ContainsParagraph(Cell cell, Paragraph paragraph)
    {
        if (cell.Paragraphs.Contains(paragraph)) return true;
        foreach (var nested in cell.NestedTables)
            if (ContainsParagraph(nested, paragraph)) return true;
        return false;
    }
    #endregion
}
