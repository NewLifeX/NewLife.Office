using System.Text;

namespace NewLife.Office.Word;

/// <summary>文本匹配结果（查找替换用）</summary>
/// <param name="Paragraph">所在段落</param>
/// <param name="StartOffset">匹配起点（段落扁平文本中的偏移）</param>
/// <param name="Length">匹配长度</param>
/// <param name="Text">匹配文本</param>
public readonly record struct TextMatch(Paragraph Paragraph, Int32 StartOffset, Int32 Length, String Text);

/// <summary>
/// Word 文档文本查找/替换扩展方法。
/// 支持跨 Run 匹配（Word 常把一句话拆到多个 <c>w:r</c>），替换时保留首 Run 格式。
/// </summary>
/// <remarks>
/// <example>
/// <code>
/// using var reader = new WordReader("source.docx");
/// var doc = reader.ReadDocument();
///
/// // 查找
/// var matches = doc.FindText("旧公司名");
///
/// // 替换（保留原格式，忽略大小写）
/// var count = doc.ReplaceText("旧公司名", "新公司名", ignoreCase: true);
///
/// // 重要：真实文档读入后 DocumentXml/RawXml 会兜底输出原 XML，覆盖模型修改。
/// // 保存前必须关闭两者，否则替换结果被静默丢弃。
/// doc.DocumentXml = null;   // 关闭 document.xml 原样透传
/// doc.ClearRawXml();        // 关闭元素级 RawXml 兜底
///
/// using var writer = new WordWriter();
/// writer.Save("output.docx", doc);
/// </code>
/// </example>
/// </remarks>
public static class WordDocumentExtensions
{
    #region 查找
    /// <summary>在文档中查找文本（跨 Run 匹配），返回所有匹配位置</summary>
    /// <param name="doc">文档</param>
    /// <param name="find">查找文本</param>
    /// <param name="ignoreCase">忽略大小写</param>
    /// <returns>匹配结果列表（段落 + 偏移 + 长度 + 文本）</returns>
    public static List<TextMatch> FindText(this Document doc, String find, Boolean ignoreCase = false)
    {
        var matches = new List<TextMatch>();
        if (String.IsNullOrEmpty(find)) return matches;
        var comparer = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        foreach (var para in EnumerateParagraphs(doc))
        {
            var text = FlattenRuns(para, out _);
            var idx = 0;
            while ((idx = text.IndexOf(find, idx, comparer)) >= 0)
            {
                matches.Add(new TextMatch(para, idx, find.Length, text.Substring(idx, find.Length)));
                idx += find.Length;
            }
        }
        return matches;
    }
    #endregion

    #region 替换
    /// <summary>在文档中替换文本（保留首 Run 格式），返回替换次数</summary>
    /// <param name="doc">文档</param>
    /// <param name="find">查找文本</param>
    /// <param name="replace">替换文本</param>
    /// <param name="ignoreCase">忽略大小写</param>
    /// <param name="wholeWord">整词匹配（前后不能是字母/数字/下划线）</param>
    /// <returns>替换次数</returns>
    public static Int32 ReplaceText(this Document doc, String find, String replace, Boolean ignoreCase = false, Boolean wholeWord = false)
    {
        if (String.IsNullOrEmpty(find)) return 0;
        var comparer = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var count = 0;
        foreach (var para in EnumerateParagraphs(doc))
            count += ReplaceInParagraph(para, find, replace, comparer, wholeWord);
        // 文本框/脚注/尾注替换后同步对应原始 XML 部件，保证保存时修改生效
        if (count > 0)
        {
            SyncTextBoxesRawXml(doc);
            SyncNotesXml(doc);
        }
        return count;
    }

    /// <summary>在单个段落中替换（支持跨 Run），返回替换次数</summary>
    private static Int32 ReplaceInParagraph(Paragraph para, String find, String replace, StringComparison comparer, Boolean wholeWord)
    {
        if (para.Runs.Count == 0) return 0;

        // 扁平文本 + 每个 Run 在扁平文本中的 [start, end) 边界
        var text = FlattenRuns(para, out var boundaries);
        var matches = new List<(Int32 Start, Int32 End)>();
        var idx = 0;
        while (idx <= text.Length - find.Length)
        {
            var found = text.IndexOf(find, idx, comparer);
            if (found < 0) break;
            if (!wholeWord || IsWordBoundary(text, found, find.Length))
                matches.Add((found, found + find.Length));
            idx = found + find.Length;
        }
        if (matches.Count == 0) return 0;

        // 从后往前替换，保持偏移有效
        for (var m = matches.Count - 1; m >= 0; m--)
        {
            var (start, end) = matches[m];
            ReplaceRange(para, boundaries, start, end, replace);
        }
        return matches.Count;
    }

    /// <summary>将扁平文本 [start,end) 替换为 replacement，跨 Run 时保留首个 Run 格式</summary>
    private static void ReplaceRange(Paragraph para, Int32[] boundaries, Int32 start, Int32 end, String replacement)
    {
        // 定位 start/end 所在 Run 索引
        var startRun = FindRunIndex(boundaries, start);
        var endRun = FindRunIndex(boundaries, end - 1);
        if (startRun < 0 || endRun < 0) return;

        var startOff = start - boundaries[startRun];
        var endOff = end - boundaries[endRun];

        if (startRun == endRun)
        {
            // 单 Run 内替换
            var t = para.Runs[startRun].Text;
            para.Runs[startRun].Text = t[..startOff] + replacement + t[endOff..];
        }
        else
        {
            // 跨 Run：首 Run 保留前缀+替换文本，中间 Run 清空，末 Run 保留后缀
            var first = para.Runs[startRun].Text;
            var last = para.Runs[endRun].Text;
            para.Runs[startRun].Text = first[..startOff] + replacement;
            for (var i = startRun + 1; i < endRun; i++)
                para.Runs[i].Text = String.Empty;
            para.Runs[endRun].Text = last[endOff..];
        }
    }

    /// <summary>整词边界判断：匹配前后不能是字母/数字/下划线</summary>
    private static Boolean IsWordBoundary(String text, Int32 start, Int32 length)
    {
        if (start > 0 && IsWordChar(text[start - 1])) return false;
        var after = start + length;
        if (after < text.Length && IsWordChar(text[after])) return false;
        return true;
    }

    private static Boolean IsWordChar(Char c) => Char.IsLetterOrDigit(c) || c == '_';
    #endregion

    #region 辅助
    /// <summary>扁平化段落 Run 文本，输出每个 Run 在扁平文本中的起始偏移</summary>
    private static String FlattenRuns(Paragraph para, out Int32[] boundaries)
    {
        var sb = new StringBuilder();
        boundaries = new Int32[para.Runs.Count];
        for (var i = 0; i < para.Runs.Count; i++)
        {
            boundaries[i] = sb.Length;
            sb.Append(para.Runs[i].Text);
        }
        return sb.ToString();
    }

    /// <summary>定位扁平偏移所在的 Run 索引</summary>
    private static Int32 FindRunIndex(Int32[] boundaries, Int32 offset)
    {
        for (var i = boundaries.Length - 1; i >= 0; i--)
        {
            if (offset >= boundaries[i]) return i;
        }
        return -1;
    }

    /// <summary>枚举文档中所有可替换的段落（正文/表格单元格/页眉页脚/文本框/脚注/尾注）</summary>
    private static IEnumerable<Paragraph> EnumerateParagraphs(Document doc)
    {
        foreach (var el in doc.Elements)
        {
            if (el.Paragraph != null)
            {
                yield return el.Paragraph;
                foreach (var tp in EnumerateTextBoxes(el.Paragraph))
                    yield return tp;
            }
            else if (el.Table != null)
            {
                foreach (var row in el.Table.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var p in EnumerateCellParagraphs(cell))
                            yield return p;
            }
            else if (el.TableRows != null)
            {
                foreach (var row in el.TableRows)
                    foreach (var cell in row)
                        foreach (var p in EnumerateCellParagraphs(cell))
                            yield return p;
            }
        }
        foreach (var hdr in doc.Headers)
        {
            foreach (var el in hdr.Elements)
            {
                if (el.Paragraph != null)
                {
                    yield return el.Paragraph;
                    foreach (var tp in EnumerateTextBoxes(el.Paragraph))
                        yield return tp;
                }
            }
        }
        foreach (var ftr in doc.Footers)
        {
            foreach (var el in ftr.Elements)
            {
                if (el.Paragraph != null)
                {
                    yield return el.Paragraph;
                    foreach (var tp in EnumerateTextBoxes(el.Paragraph))
                        yield return tp;
                }
            }
        }
        // 脚注/尾注（与 Word 原生查找行为一致）
        foreach (var fn in doc.Footnotes)
            foreach (var p in fn.Paragraphs)
                yield return p;
        foreach (var en in doc.Endnotes)
            foreach (var p in en.Paragraphs)
                yield return p;
    }

    /// <summary>递归枚举段落内文本框内容段落（W22）</summary>
    private static IEnumerable<Paragraph> EnumerateTextBoxes(Paragraph para)
    {
        foreach (var tp in para.TextBoxes)
        {
            yield return tp;
            foreach (var nested in EnumerateTextBoxes(tp))
                yield return nested;
        }
    }

    /// <summary>枚举单元格内段落（含嵌套表格递归，W46）</summary>
    private static IEnumerable<Paragraph> EnumerateCellParagraphs(Cell cell)
    {
        foreach (var p in cell.Paragraphs)
        {
            yield return p;
            foreach (var tp in EnumerateTextBoxes(p))
                yield return tp;
        }
        foreach (var nested in cell.NestedTables)
        {
            foreach (var row in nested.Rows)
                foreach (var nc in row.Cells)
                    foreach (var p in EnumerateCellParagraphs(nc))
                        yield return p;
        }
    }

    /// <summary>
    /// 清空所有元素的 RawXml（含嵌套表格与页眉页脚元素），
    /// 配合 <c>doc.DocumentXml = null</c> 使用，强制 Writer 以模型输出——
    /// 解决"读入→通过模型修改→写回"时 RawXml 兜底覆盖模型修改的问题。
    /// </summary>
    /// <param name="doc">文档</param>
    /// <example>
    /// <code>
    /// var doc = reader.ReadDocument();
    /// doc.DocumentXml = null;      // 关闭 document.xml 原样透传
    /// doc.ClearRawXml();           // 关闭元素级 RawXml 兜底
    /// doc.Elements[0].Paragraph.Runs[0].Text = "新标题";
    /// writer.Save("output.docx", doc);
    /// </code>
    /// </example>
    public static void ClearRawXml(this Document doc)
    {
        void Clear(List<Element> list)
        {
            foreach (var el in list)
            {
                el.RawXml = null;
                if (el.Paragraph != null)
                {
                    foreach (var tp in EnumerateTextBoxes(el.Paragraph))
                        tp.RawXml = null;
                }
                if (el.Table != null)
                {
                    el.Table.RawXml = null;
                    foreach (var row in el.Table.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var nested in cell.NestedTables)
                                nested.RawXml = null;
                }
                if (el.Sdt != null) el.Sdt.RawXml = null;
            }
        }
        Clear(doc.Elements);
        foreach (var h in doc.Headers) Clear(h.Elements);
        foreach (var f in doc.Footers) Clear(f.Elements);
    }

    /// <summary>文本框替换后同步外层元素/表格元素 RawXml（重建 txbxContent 部分），保证保存时修改生效</summary>
    private static void SyncTextBoxesRawXml(Document doc)
    {
        void Sync(List<Element> list)
        {
            foreach (var el in list)
            {
                if (el.Paragraph != null && el.Paragraph.TextBoxes.Count > 0)
                    SyncTextBoxesRawXml(el);
                if (el.Type == ElementType.Table)
                    SyncTableTextBoxes(el);
            }
        }

        Sync(doc.Elements);
        foreach (var h in doc.Headers) Sync(h.Elements);
        foreach (var f in doc.Footers) Sync(f.Elements);
    }

    /// <summary>同步表格元素 RawXml 内所有 txbxContent（单元格内文本框，W22）</summary>
    private static void SyncTableTextBoxes(Element el)
    {
        if (el.RawXml == null) return;

        // 按文档顺序收集表格内所有文本框段落
        var boxes = new List<Paragraph>();
        void CollectCell(Cell cell)
        {
            foreach (var p in cell.Paragraphs)
            {
                foreach (var tb in p.TextBoxes)
                    boxes.Add(tb);
            }
            foreach (var nested in cell.NestedTables)
                foreach (var row in nested.Rows)
                    foreach (var c in row.Cells)
                        CollectCell(c);
        }

        if (el.Table != null)
        {
            foreach (var row in el.Table.Rows)
                foreach (var cell in row.Cells)
                    CollectCell(cell);
        }
        else if (el.TableRows != null)
        {
            foreach (var row in el.TableRows)
                foreach (var cell in row)
                    CollectCell(cell);
        }
        if (boxes.Count == 0) return;

        var idx = 0;
        el.RawXml = System.Text.RegularExpressions.Regex.Replace(
            el.RawXml,
            "<w:txbxContent>.*?</w:txbxContent>",
            m =>
            {
                var tp = idx < boxes.Count ? boxes[idx] : null;
                idx++;
                return tp == null ? m.Value : $"<w:txbxContent>{BuildTextboxParagraphXml(tp)}</w:txbxContent>";
            },
            System.Text.RegularExpressions.RegexOptions.Singleline);
    }

    /// <summary>脚注/尾注替换后同步 footnotes.xml/endnotes.xml 部件，保证保存时修改生效</summary>
    private static void SyncNotesXml(Document doc)
    {
        SyncNotePart(doc, "word/footnotes.xml", "w:footnote", doc.Footnotes);
        SyncNotePart(doc, "word/endnotes.xml", "w:endnote", doc.Endnotes);
    }

    /// <summary>同步单个脚注/尾注部件中文本已变化的条目（保留未修改条目的原始 XML）</summary>
    private static void SyncNotePart(Document doc, String partPath, String elemName, List<Footnote> notes)
    {
        if (notes.Count == 0) return;
        if (!doc.OtherParts.TryGetValue(partPath, out var bytes)) return;
        var xml = Encoding.UTF8.GetString(bytes);
        var changed = false;

        foreach (var note in notes)
        {
            if (note.Paragraphs.Count == 0) continue;

            // 定位 <elem w:id="N" ...>...</elem>
            var openMatch = System.Text.RegularExpressions.Regex.Match(xml,
                $"<{elemName}[^>]*w:id=\"{note.Id}\"[^>]*>");
            if (!openMatch.Success) continue;
            var contentStart = openMatch.Index + openMatch.Length;
            var closeTag = $"</{elemName}>";
            var closeIdx = xml.IndexOf(closeTag, contentStart, StringComparison.Ordinal);
            if (closeIdx < 0) continue;

            // 文本未变化则保留原 XML（保真）；变化则从模型重建
            var oldText = ExtractWtText(xml.Substring(contentStart, closeIdx - contentStart));
            var modelText = new StringBuilder();
            foreach (var p in note.Paragraphs)
                foreach (var r in p.Runs)
                    modelText.Append(r.Text);
            if (oldText == modelText.ToString()) continue;

            var sb = new StringBuilder();
            foreach (var p in note.Paragraphs)
                sb.Append(BuildTextboxParagraphXml(p));
            xml = xml[..contentStart] + sb + xml[closeIdx..];
            changed = true;
        }

        if (changed)
            doc.OtherParts[partPath] = Encoding.UTF8.GetBytes(xml);
    }

    /// <summary>提取 XML 片段内全部文本（w:t 文本 + w:tab→\t + w:br/w:cr→\n，与 Run.Text 语义一致）</summary>
    private static String ExtractWtText(String xml)
    {
        var sb = new StringBuilder();
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
            xml, "<w:t[^>]*>(?<t>.*?)</w:t>|<w:tab[^>]*/>|<w:br[^>]*/>|<w:cr[^>]*/>"))
        {
            if (m.Groups["t"].Success)
                sb.Append(DecodeXml(m.Groups["t"].Value));
            else if (m.Value.StartsWith("<w:tab", StringComparison.Ordinal))
                sb.Append('\t');
            else
                sb.Append('\n'); // w:br / w:cr
        }
        return sb.ToString();
    }

    /// <summary>解码 XML 实体</summary>
    private static String DecodeXml(String s) => s
        .Replace("&lt;", "<").Replace("&gt;", ">")
        .Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&");

    /// <summary>重建单个段落元素 RawXml 的 txbxContent 部分（从 TextBoxes 模型，文本框按序对应）</summary>
    private static void SyncTextBoxesRawXml(Element el)
    {
        if (el.RawXml == null || el.Paragraph == null || el.Paragraph.TextBoxes.Count == 0) return;

        var idx = 0;
        el.RawXml = System.Text.RegularExpressions.Regex.Replace(
            el.RawXml,
            "<w:txbxContent>.*?</w:txbxContent>",
            m =>
            {
                var tp = idx < el.Paragraph.TextBoxes.Count ? el.Paragraph.TextBoxes[idx] : null;
                idx++;
                return tp == null ? m.Value : $"<w:txbxContent>{BuildTextboxParagraphXml(tp)}</w:txbxContent>";
            },
            System.Text.RegularExpressions.RegexOptions.Singleline);
    }

    /// <summary>从 Run 模型构建文本框内段落 XML（替换后重建用，保留常见字符格式）</summary>
    private static String BuildTextboxParagraphXml(Paragraph p)
    {
        static String Esc(String? s) => s == null ? String.Empty : (System.Security.SecurityElement.Escape(s) ?? s);

        var sb = new StringBuilder();
        sb.Append("<w:p>");
        foreach (var run in p.Runs)
        {
            sb.Append("<w:r>");
            var rp = run.Properties;
            if (rp != null)
            {
                sb.Append("<w:rPr>");
                if (rp.Bold == true) sb.Append("<w:b/>");
                if (rp.Italic == true) sb.Append("<w:i/>");
                if (rp.Strikethrough == true) sb.Append("<w:strike/>");
                if (rp.Underline == true) sb.Append("<w:u/>");
                if (rp.Superscript == true) sb.Append("<w:vertAlign w:val=\"superscript\"/>");
                else if (rp.Subscript == true) sb.Append("<w:vertAlign w:val=\"subscript\"/>");
                if (rp.ForeColor != null) sb.Append($"<w:color w:val=\"{rp.ForeColor.TrimStart('#')}\"/>");
                if (rp.FontSize.HasValue) sb.Append($"<w:sz w:val=\"{(Int32)(rp.FontSize.Value * 2)}\"/>");
                if (rp.FontName != null) sb.Append($"<w:rFonts w:ascii=\"{Esc(rp.FontName)}\" w:hAnsi=\"{Esc(rp.FontName)}\"/>");
                sb.Append("</w:rPr>");
            }
            var spaceAttr = (run.Text.Length > 0 && (run.Text[0] == ' ' || run.Text[^1] == ' '))
                ? " xml:space=\"preserve\"" : "";
            sb.Append($"<w:t{spaceAttr}>{Esc(run.Text)}</w:t>");
            sb.Append("</w:r>");
        }
        sb.Append("</w:p>");
        return sb.ToString();
    }
    #endregion
}
