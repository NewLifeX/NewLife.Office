namespace NewLife.Office.Word;

/// <summary>文档补丁：按锚点批量修改文档，供 AI/Agent 以 JSON 驱动（对应结构快照里的 <c>p#3</c>/<c>tbl#5</c>）</summary>
/// <remarks>
/// 工作流：<c>read</c> 取 <see cref="DocxOutline"/> 看清结构与锚点 → 生成补丁 JSON → <see cref="Apply"/> 批量执行。
/// 所有操作在内存模型上进行，由调用方根据 <see cref="DocxPatchResult.Success"/> 决定是否落盘；
/// <b>注意：失败时不会回滚</b>——已执行成功的操作仍留在内存模型上，因此失败后必须丢弃该 <see cref="Document"/> 实例（重新读取源文件），否则会得到“部分修改”的文档。
/// <para>锚点指的是 <b>read 时刻</b> 的 <see cref="Document.Elements"/> 下标。补丁中存在 <c>insert_paragraph</c> / <c>insert_table</c> / <c>delete</c>
/// 时，其后的元素下标会整体位移，后续操作的锚点会对不上。请把结构性操作排在补丁最后，或按锚点从大到小安排顺序。</para>
/// <para>支持的操作：<c>replace_text</c> / <c>set_text</c> / <c>insert_paragraph</c> / <c>delete</c> /
/// <c>set_style</c> / <c>set_cell_text</c> / <c>insert_table</c> / <c>add_comment</c>。</para>
/// <example>
/// <code>
/// var patch = DocxPatch.ParseJson(File.ReadAllText("patch.json"));
/// using var reader = new WordReader("in.docx");
/// var doc = reader.ReadDocument();
/// doc.BeginEdit();
///
/// var result = DocxPatch.Apply(doc, patch);
/// if (!result.Success) throw new InvalidOperationException(String.Join("; ", result.Failed));
///
/// using var writer = new WordWriter();
/// writer.Save("out.docx", doc);
/// </code>
/// </example>
/// </remarks>
public class DocxPatch
{
    #region 属性
    /// <summary>操作列表（按顺序执行）</summary>
    public List<DocxPatchOp> Ops { get; } = [];
    #endregion

    #region 解析
    /// <summary>从 JSON 文本解析补丁</summary>
    /// <param name="json">补丁 JSON，形如 <c>{"ops":[{"op":"set_text","anchor":"p#3","text":"新内容"}]}</c></param>
    /// <returns>补丁对象</returns>
    /// <exception cref="ArgumentNullException">json 为空</exception>
    /// <exception cref="FormatException">JSON 非法或缺少 ops 数组</exception>
    public static DocxPatch ParseJson(String json)
    {
        if (json == null) throw new ArgumentNullException(nameof(json));

        var root = JsonLite.Parse(json) as Dictionary<String, Object?>
            ?? throw new FormatException("Patch root must be a JSON object.");

        // 顶层必须是对象：{"ops":[...]}；只给数组会解析失败
        List<Object?>? ops = null;
        if (root.TryGetValue("ops", out var opsValue)) ops = opsValue as List<Object?>;
        else if (root.Count == 0) ops = [];

        if (ops == null) throw new FormatException("Patch must contain an 'ops' array.");

        var patch = new DocxPatch();
        foreach (var item in ops)
        {
            if (item is not Dictionary<String, Object?> obj)
                throw new FormatException("Each op must be a JSON object.");

            patch.Ops.Add(new DocxPatchOp
            {
                Op = JsonLite.GetString(obj, "op") ?? String.Empty,
                Anchor = JsonLite.GetString(obj, "anchor"),
                Position = JsonLite.GetString(obj, "position"),
                Text = JsonLite.GetString(obj, "text"),
                Style = JsonLite.GetString(obj, "style"),
                Find = JsonLite.GetString(obj, "find"),
                Replace = JsonLite.GetString(obj, "replace"),
                Regex = JsonLite.GetBoolean(obj, "regex") ?? false,
                Author = JsonLite.GetString(obj, "author"),
                Row = JsonLite.GetInt32(obj, "row"),
                Column = JsonLite.GetInt32(obj, "column"),
                Rows = JsonLite.GetInt32(obj, "rows"),
                Columns = JsonLite.GetInt32(obj, "columns"),
                Cells = ReadCells(obj),
            });
        }
        return patch;
    }

    private static List<List<String>>? ReadCells(Dictionary<String, Object?> obj)
    {
        if (!obj.TryGetValue("cells", out var value) || value is not List<Object?> rows) return null;

        var matrix = new List<List<String>>();
        foreach (var rowValue in rows)
        {
            var row = new List<String>();
            if (rowValue is List<Object?> cells)
            {
                foreach (var cell in cells) row.Add(cell as String ?? Convert.ToString(cell) ?? String.Empty);
            }
            matrix.Add(row);
        }
        return matrix;
    }
    #endregion

    #region 执行
    /// <summary>在文档模型上执行补丁（调用前请先 <see cref="Document.BeginEdit"/>）</summary>
    /// <param name="doc">文档</param>
    /// <param name="patch">补丁</param>
    /// <returns>逐条执行结果；<see cref="DocxPatchResult.Success"/> 为 false 时不应保存，且应丢弃当前文档重新读取（无回滚）</returns>
    /// <exception cref="ArgumentNullException">doc 或 patch 为空</exception>
    public static DocxPatchResult Apply(Document doc, DocxPatch patch)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (patch == null) throw new ArgumentNullException(nameof(patch));

        var result = new DocxPatchResult();
        var warnStart = doc.EditWarnings.Count;
        for (var i = 0; i < patch.Ops.Count; i++)
        {
            var op = patch.Ops[i];
            try
            {
                ApplyOne(doc, op);
                result.Applied.Add($"#{i} {op.Op}: ok");
            }
            catch (Exception ex)
            {
                result.Failed.Add($"#{i} {op.Op}: {ex.Message}");
            }
        }

        // 把本次执行新增的告警带回调用方（BeginEdit 已有的告警不回放）
        for (var i = warnStart; i < doc.EditWarnings.Count; i++)
            result.Warnings.Add(doc.EditWarnings[i]);

        if (result.Applied.Count > 0) doc.MarkModified();
        return result;
    }

    private static void ApplyOne(Document doc, DocxPatchOp op)
    {
        switch (op.Op)
        {
            case "replace_text":
                if (String.IsNullOrEmpty(op.Find)) throw new ArgumentException("replace_text 需要 find");
                if (String.IsNullOrEmpty(op.Replace)) throw new ArgumentException("replace_text 需要 replace");
                var count = op.Regex
                    ? doc.ReplaceRegex(op.Find!, op.Replace!)
                    : doc.ReplaceText(op.Find!, op.Replace!);
                if (count == 0) throw new InvalidOperationException("未匹配到 find 文本");
                return;

            case "set_text":
            {
                var (el, para) = RequireParagraphPair(doc, op.Anchor, "set_text");
                DropRawXmlFallback(doc, el, op.Anchor!, "set_text");
                para.SetText(op.Text ?? String.Empty);
                return;
            }

            case "insert_paragraph":
            {
                var style = ParseStyle(op.Style);
                if (String.IsNullOrEmpty(op.Anchor))
                {
                    doc.AddParagraph(op.Text ?? String.Empty, style);
                    return;
                }

                var para = FindParagraphByAnchor(doc, op.Anchor!) ?? throw new ArgumentException($"锚点不存在或不是段落: {op.Anchor}");
                if (String.Equals(op.Position, "before", StringComparison.OrdinalIgnoreCase))
                    doc.InsertParagraphBefore(para, op.Text ?? String.Empty, style);
                else
                    doc.InsertParagraphAfter(para, op.Text ?? String.Empty, style);
                return;
            }

            case "delete":
            {
                var el = ResolveElement(doc, op.Anchor, "delete");
                if (!doc.RemoveElement(el)) throw new InvalidOperationException($"删除失败: {op.Anchor}");
                return;
            }

            case "set_style":
            {
                var (el, para) = RequireParagraphPair(doc, op.Anchor, "set_style");
                if (String.IsNullOrEmpty(op.Style)) throw new ArgumentException("set_style 需要 style");
                DropRawXmlFallback(doc, el, op.Anchor!, "set_style");
                ApplyStyle(para, op.Style!);
                para.RawXml = null;
                return;
            }

            case "set_cell_text":
            {
                var cellEl = ResolveElement(doc, op.Anchor, "set_cell_text");
                var table = cellEl.Table ?? throw new ArgumentException($"锚点不是表格: {op.Anchor}");
                if (!op.Row.HasValue || !op.Column.HasValue) throw new ArgumentException("set_cell_text 需要 row 与 column");
                DropRawXmlFallback(doc, cellEl, op.Anchor!, "set_cell_text");
                table.RawXml = null;
                table.SetCellText(op.Row.Value, op.Column.Value, op.Text ?? String.Empty);
                return;
            }

            case "insert_table":
            {
                Element el;
                if (op.Cells != null)
                {
                    if (String.IsNullOrEmpty(op.Anchor)) { el = doc.AddTable(op.Cells.Select(r => r.ToArray())); return; }
                    var temp = new Document();
                    temp.AddTable(op.Cells.Select(r => r.ToArray()));
                    el = temp.Elements[0];
                }
                else
                {
                    var rows = op.Rows ?? 1;
                    var cols = op.Columns ?? 1;
                    if (String.IsNullOrEmpty(op.Anchor)) { doc.AddTable(rows, cols); return; }
                    var temp = new Document();
                    temp.AddTable(rows, cols);
                    el = temp.Elements[0];
                }

                var index = ResolveIndex(doc, op.Anchor)
                    ?? throw new ArgumentException($"锚点不存在: {op.Anchor}（insert_table 不再静默追加到末尾）");
                if (String.Equals(op.Position, "before", StringComparison.OrdinalIgnoreCase)) index--;
                doc.InsertElementAt(index, el);
                return;
            }

            case "add_comment":
            {
                var (el, para) = RequireParagraphPair(doc, op.Anchor, "add_comment");
                if (String.IsNullOrEmpty(op.Text)) throw new ArgumentException("add_comment 需要 text");
                DropRawXmlFallback(doc, el, op.Anchor!, "add_comment");
                doc.AddComment(para, op.Text!, op.Author);
                return;
            }

            default:
                throw new ArgumentException($"不支持的操作: {op.Op}");
        }
    }
    #endregion

    #region 辅助
    /// <summary>解析锚点序号（如 <c>p#3</c> → 3）；格式非法返回 null</summary>
    private static Int32? ResolveIndex(Document doc, String? anchor)
    {
        if (String.IsNullOrEmpty(anchor)) return null;
        var idx = anchor!.LastIndexOf('#');
        if (idx < 0 || !Int32.TryParse(anchor[(idx + 1)..], out var index)) return null;
        if (index < 0 || index >= doc.Elements.Count) return null;
        return index;
    }

    private static Element ResolveElement(Document doc, String? anchor, String opName)
    {
        var index = ResolveIndex(doc, anchor) ?? throw new ArgumentException($"{opName} 需要有效锚点");
        return doc.Elements[index];
    }

    private static Paragraph? FindParagraphByAnchor(Document doc, String anchor)
    {
        var index = ResolveIndex(doc, anchor);
        return index.HasValue ? doc.Elements[index.Value].Paragraph : null;
    }

    private static Paragraph RequireParagraph(Document doc, String? anchor, String opName)
        => (String.IsNullOrEmpty(anchor) ? null : FindParagraphByAnchor(doc, anchor!))
            ?? throw new ArgumentException($"{opName} 需要指向段落的锚点");

    /// <summary>按锚点取段落及其所属元素（修改前后需要定位元素级 RawXml）</summary>
    /// <param name="doc">文档</param>
    /// <param name="anchor">锚点（如 p#3）</param>
    /// <param name="opName">操作名（用于异常提示）</param>
    /// <returns>元素与段落</returns>
    private static (Element Element, Paragraph Paragraph) RequireParagraphPair(Document doc, String? anchor, String opName)
    {
        var index = ResolveIndex(doc, anchor) ?? throw new ArgumentException($"{opName} 需要指向段落的锚点");
        var el = doc.Elements[index];
        var para = el.Paragraph ?? throw new ArgumentException($"{opName} 需要指向段落的锚点");
        return (el, para);
    }

    /// <summary>清除元素级 RawXml 兜底，保证模型修改真正落盘</summary>
    /// <remarks>
    /// 真实 Word 文档中部分段落含域、批注锚点、对象等未建模内容，Reader 会同时保留原始 XML（RawXml）并建模。
    /// 保存时 Writer 优先输出 RawXml，导致模型修改被静默丢弃。这里在修改前清掉兜底，改为按模型重建，
    /// 并把"未建模内容会丢失"记入 <see cref="Document.EditWarnings"/>，由 <see cref="DocxPatchResult.Warnings"/> 带回调用方。
    /// </remarks>
    /// <param name="doc">文档</param>
    /// <param name="el">被修改的元素</param>
    /// <param name="anchor">锚点（用于告警文本）</param>
    /// <param name="opName">操作名（用于告警文本）</param>
    private static void DropRawXmlFallback(Document doc, Element el, String anchor, String opName)
    {
        if (el.RawXml == null) return;
        el.RawXml = null;
        if (el.Paragraph is { HasUnmodeledContent: true })
            doc.EditWarnings.Add($"{opName} {anchor}: 该段落含域/批注锚点等未建模内容，按模型重建后这些内容会丢失");
    }

    /// <summary>解析样式：Heading1~6 映射枚举，其余作为自定义样式 ID</summary>
    private static ParagraphStyle ParseStyle(String? style)
    {
        if (String.IsNullOrEmpty(style)) return ParagraphStyle.Normal;
        for (var i = 1; i <= 6; i++)
        {
            if (String.Equals(style, $"Heading{i}", StringComparison.OrdinalIgnoreCase)) return (ParagraphStyle)i;
        }
        return ParagraphStyle.Normal;
    }

    /// <summary>应用样式：同步设置 Style 枚举与 StyleId（自定义样式名靠 StyleId 生效）</summary>
    private static void ApplyStyle(Paragraph para, String style)
    {
        para.Style = ParseStyle(style);
        para.StyleId = para.Style == ParagraphStyle.Normal && !String.Equals(style, "Normal", StringComparison.OrdinalIgnoreCase)
            ? style
            : null;
    }
    #endregion
}

/// <summary>补丁操作</summary>
public class DocxPatchOp
{
    #region 属性
    /// <summary>操作名：replace_text / set_text / insert_paragraph / delete / set_style / set_cell_text / insert_table / add_comment</summary>
    public String Op { get; set; } = String.Empty;

    /// <summary>目标锚点（来自结构快照，如 <c>p#3</c>、<c>tbl#5</c>）；部分操作可省略</summary>
    public String? Anchor { get; set; }

    /// <summary>插入位置：after（默认）/ before</summary>
    public String? Position { get; set; }

    /// <summary>文本内容（set_text / insert_paragraph / set_cell_text / add_comment）</summary>
    public String? Text { get; set; }

    /// <summary>样式 ID（heading 用 Heading1~6，其余按自定义样式名）</summary>
    public String? Style { get; set; }

    /// <summary>查找文本（replace_text）</summary>
    public String? Find { get; set; }

    /// <summary>替换文本（replace_text，支持 $1 分组）</summary>
    public String? Replace { get; set; }

    /// <summary>find 是否按正则处理</summary>
    public Boolean Regex { get; set; }

    /// <summary>批注作者（add_comment）</summary>
    public String? Author { get; set; }

    /// <summary>行号（set_cell_text）</summary>
    public Int32? Row { get; set; }

    /// <summary>列号（set_cell_text）</summary>
    public Int32? Column { get; set; }

    /// <summary>表格行数（insert_table 空网格）</summary>
    public Int32? Rows { get; set; }

    /// <summary>表格列数（insert_table 空网格）</summary>
    public Int32? Columns { get; set; }

    /// <summary>表格单元格矩阵（insert_table，优先于 rows/columns）</summary>
    public List<List<String>>? Cells { get; set; }
    #endregion
}

/// <summary>补丁执行结果</summary>
public class DocxPatchResult
{
    #region 属性
    /// <summary>成功的操作（形如 <c>#0 set_text: ok</c>）</summary>
    public List<String> Applied { get; } = [];

    /// <summary>失败的操作与原因</summary>
    public List<String> Failed { get; } = [];

    /// <summary>执行过程中的告警（如修改段落时丢弃了原本未建模的域/批注锚点等内容）</summary>
    /// <remarks>告警不影响 <see cref="Success"/>，但调用方应向用户展示，避免“内容被静默丢弃”。</remarks>
    public List<String> Warnings { get; } = [];

    /// <summary>是否全部成功（为 false 时不应保存文档）</summary>
    public Boolean Success => Failed.Count == 0;
    #endregion
}
