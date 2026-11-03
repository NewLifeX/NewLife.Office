using System.Text;

namespace NewLife.Office.Pdf;

/// <summary>PDF 结构快照：面向 AI/Agent 的页/书签/表单/注释/附件视图</summary>
/// <remarks>
/// 锚点以页码（1 起始）与矩形坐标为定位单位，配合 <see cref="PdfTextSearch.SearchText"/> 可形成"搜索 → 定位 → 标注"链路。
/// <example>
/// <code>
/// using var reader = new PdfReader("report.pdf");
/// var snapshot = PdfSnapshot.From(reader);
/// Console.WriteLine(snapshot.ToJson());
/// </code>
/// </example>
/// </remarks>
public class PdfSnapshot
{
    #region 属性
    /// <summary>页数</summary>
    public Int32 PageCount { get; set; }

    /// <summary>是否加密（需要密码才能解析内容）</summary>
    public Boolean Encrypted { get; set; }

    /// <summary>是否含签名字段</summary>
    public Boolean Signed { get; set; }

    /// <summary>元数据</summary>
    public PdfDocumentInfo? Metadata { get; set; }

    /// <summary>页面快照</summary>
    public List<PdfPageSnapshot> Pages { get; } = [];

    /// <summary>书签树</summary>
    public List<PdfBookmarkSnapshot> Bookmarks { get; } = [];

    /// <summary>表单字段</summary>
    public List<PdfFormFieldSnapshot> FormFields { get; } = [];

    /// <summary>注释总数</summary>
    public Int32 AnnotationCount { get; set; }

    /// <summary>字体数量</summary>
    public Int32 FontCount { get; set; }

    /// <summary>附件文件名集合</summary>
    public List<String> Attachments { get; } = [];
    #endregion

    #region 构造
    /// <summary>从读取器构建快照</summary>
    /// <param name="reader">PDF 读取器</param>
    /// <param name="firstLineLength">每页首行预览的最大字符数</param>
    /// <returns>结构快照</returns>
    public static PdfSnapshot From(PdfReader reader, Int32 firstLineLength = 80)
    {
        if (reader == null) throw new ArgumentNullException(nameof(reader));

        var snapshot = new PdfSnapshot
        {
            PageCount = reader.GetPageCount(),
            Encrypted = reader.XRefTable?.Trailer.ContainsKey("Encrypt") == true,
        };

        snapshot.Metadata = reader.ReadMetadata();

        var doc = reader.ReadDocument();
        foreach (var page in doc.Pages)
        {
            var item = new PdfPageSnapshot
            {
                Index = snapshot.Pages.Count,
                Width = page.Width,
                Height = page.Height,
                Rotation = page.Rotation,
                TextLength = page.TextBlocks.Sum(e => e.Text.Length),
            };

            var first = page.TextBlocks
                .Where(e => !e.Text.IsNullOrEmpty())
                .OrderByDescending(e => e.Y)
                .ThenBy(e => e.X)
                .FirstOrDefault();
            if (first != null)
            {
                var text = first.Text.Trim();
                item.FirstLine = text.Length > firstLineLength ? text[..firstLineLength] : text;
            }

            try
            {
                var annotations = reader.ReadAnnotations(item.Index);
                item.AnnotationCount = annotations.Count;
                snapshot.AnnotationCount += annotations.Count;
            }
            catch
            {
                item.AnnotationCount = 0;
            }

            snapshot.Pages.Add(item);
        }

        foreach (var bookmark in doc.Bookmarks)
            snapshot.Bookmarks.Add(Convert(bookmark));

        var form = reader.ReadFormFields();
        if (form != null)
        {
            foreach (var field in form.Fields)
            {
                snapshot.FormFields.Add(new PdfFormFieldSnapshot
                {
                    Name = field.FullName,
                    Type = field.FieldType.ToString(),
                    Value = field.Value,
                    PageIndex = field.PageIndex,
                    ReadOnly = (field.Flags & FormFieldFlags.ReadOnly) != 0,
                });
                if (field.FieldType == FormFieldType.Sig) snapshot.Signed = true;
            }
        }

        snapshot.FontCount = reader.ReadFonts().Count;

        foreach (var attachment in reader.GetAttachments())
            snapshot.Attachments.Add(attachment.FileName);

        return snapshot;
    }

    private static PdfBookmarkSnapshot Convert(PdfOutline outline)
    {
        var item = new PdfBookmarkSnapshot
        {
            Title = outline.Title,
            PageIndex = outline.PageIndex,
        };
        foreach (var child in outline.Children)
            item.Children.Add(Convert(child));
        return item;
    }
    #endregion

    #region JSON
    /// <summary>输出 JSON 快照</summary>
    /// <param name="pretty">是否缩进美化</param>
    /// <returns>JSON 文本</returns>
    public String ToJson(Boolean pretty = true)
    {
        var sb = new StringBuilder();
        var nl = pretty ? "\n" : String.Empty;
        var pad = pretty ? "  " : String.Empty;

        sb.Append('{').Append(nl);
        sb.Append(pad).Append("\"pageCount\": ").Append(PageCount).Append(',').Append(nl);
        sb.Append(pad).Append("\"encrypted\": ").Append(Encrypted ? "true" : "false").Append(',');
        sb.Append(" \"signed\": ").Append(Signed ? "true" : "false").Append(',').Append(nl);
        sb.Append(pad).Append("\"fontCount\": ").Append(FontCount).Append(',');
        sb.Append(" \"annotationCount\": ").Append(AnnotationCount).Append(',').Append(nl);

        sb.Append(pad).Append("\"metadata\": {");
        if (Metadata != null)
        {
            sb.Append("\"title\": ").Append(OfficeJson.Escape(Metadata.Title));
            sb.Append(", \"author\": ").Append(OfficeJson.Escape(Metadata.Author));
            sb.Append(", \"subject\": ").Append(OfficeJson.Escape(Metadata.Subject));
            sb.Append(", \"creationDate\": ").Append(OfficeJson.Escape(Metadata.CreationDate));
            sb.Append(", \"pdfVersion\": ").Append(OfficeJson.Escape(Metadata.PdfVersion));
        }
        sb.Append("},").Append(nl);

        sb.Append(pad).Append("\"attachments\": ").Append(OfficeJson.Array(Attachments)).Append(',').Append(nl);

        sb.Append(pad).Append("\"bookmarks\": ");
        AppendBookmarks(sb, Bookmarks);
        sb.Append(',').Append(nl);

        sb.Append(pad).Append("\"forms\": [");
        for (var i = 0; i < FormFields.Count; i++)
        {
            var field = FormFields[i];
            if (i > 0) sb.Append(", ");
            sb.Append('{');
            sb.Append("\"name\": ").Append(OfficeJson.Escape(field.Name));
            sb.Append(", \"type\": ").Append(OfficeJson.Escape(field.Type));
            if (field.Value != null) sb.Append(", \"value\": ").Append(OfficeJson.Escape(field.Value));
            sb.Append(", \"page\": ").Append(field.PageIndex + 1);
            if (field.ReadOnly) sb.Append(", \"readOnly\": true");
            sb.Append('}');
        }
        sb.Append("],").Append(nl);

        sb.Append(pad).Append("\"pages\": [");
        for (var i = 0; i < Pages.Count; i++)
        {
            var page = Pages[i];
            sb.Append(nl).Append(pad).Append(pad).Append('{');
            sb.Append("\"anchor\": ").Append(OfficeJson.Escape($"p#{page.Index}"));
            sb.Append(", \"page\": ").Append(page.Index + 1);
            sb.Append(", \"width\": ").Append(OfficeJson.Value(page.Width));
            sb.Append(", \"height\": ").Append(OfficeJson.Value(page.Height));
            if (page.Rotation != 0) sb.Append(", \"rotation\": ").Append(page.Rotation);
            sb.Append(", \"textLength\": ").Append(page.TextLength);
            if (!page.FirstLine.IsNullOrEmpty()) sb.Append(", \"firstLine\": ").Append(OfficeJson.Escape(page.FirstLine));
            if (page.AnnotationCount > 0) sb.Append(", \"annotations\": ").Append(page.AnnotationCount);
            sb.Append('}');
            if (i < Pages.Count - 1) sb.Append(',');
        }
        sb.Append(nl).Append(pad).Append(']').Append(nl).Append('}');
        return sb.ToString();
    }

    private static void AppendBookmarks(StringBuilder sb, List<PdfBookmarkSnapshot> bookmarks)
    {
        sb.Append('[');
        for (var i = 0; i < bookmarks.Count; i++)
        {
            var item = bookmarks[i];
            if (i > 0) sb.Append(", ");
            sb.Append('{');
            sb.Append("\"title\": ").Append(OfficeJson.Escape(item.Title));
            sb.Append(", \"page\": ").Append(item.PageIndex + 1);
            if (item.Children.Count > 0)
            {
                sb.Append(", \"children\": ");
                AppendBookmarks(sb, item.Children);
            }
            sb.Append('}');
        }
        sb.Append(']');
    }
    #endregion
}

/// <summary>页面快照</summary>
public class PdfPageSnapshot
{
    /// <summary>页索引（0 基）</summary>
    public Int32 Index { get; set; }

    /// <summary>页宽（点）</summary>
    public Single Width { get; set; }

    /// <summary>页高（点）</summary>
    public Single Height { get; set; }

    /// <summary>旋转角度</summary>
    public Int32 Rotation { get; set; }

    /// <summary>文本总长度</summary>
    public Int32 TextLength { get; set; }

    /// <summary>首行预览</summary>
    public String? FirstLine { get; set; }

    /// <summary>注释数量</summary>
    public Int32 AnnotationCount { get; set; }
}

/// <summary>书签快照</summary>
public class PdfBookmarkSnapshot
{
    /// <summary>标题</summary>
    public String Title { get; set; } = String.Empty;

    /// <summary>目标页索引（0 基）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>子书签</summary>
    public List<PdfBookmarkSnapshot> Children { get; } = [];
}

/// <summary>表单字段快照</summary>
public class PdfFormFieldSnapshot
{
    /// <summary>字段名</summary>
    public String Name { get; set; } = String.Empty;

    /// <summary>字段类型（Btn/Tx/Ch/Sig）</summary>
    public String Type { get; set; } = String.Empty;

    /// <summary>当前值</summary>
    public String? Value { get; set; }

    /// <summary>所在页索引（0 基）</summary>
    public Int32 PageIndex { get; set; }

    /// <summary>是否只读</summary>
    public Boolean ReadOnly { get; set; }
}
