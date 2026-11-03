using System.Globalization;
using System.Text;

namespace NewLife.Office.Pdf;

/// <summary>PDF 对象级文档：按对象号读写整份 PDF，支持页面重排、标注、元数据与表单填充</summary>
/// <remarks>
/// 与 <see cref="PdfReader"/>（只读提取）+ <see cref="PdfWriter"/>（新建）互补：本类保留原始对象与内容流，
/// 适合对已有 PDF 做"局部修改后整份写回"。
/// <para>加密 PDF 不支持修改（写入时抛 <see cref="NotSupportedException"/>）。</para>
/// </remarks>
public class PdfObjectDocument
{
    #region 属性
    /// <summary>对象号 → 对象</summary>
    public Dictionary<Int32, PdfObject> Objects { get; } = [];

    /// <summary>对象号 → 代数（增量更新文件的间接引用需保持代数一致）</summary>
    public Dictionary<Int32, Int32> Generations { get; } = [];

    /// <summary>trailer 字典</summary>
    public PdfDict Trailer { get; } = [];

    /// <summary>是否加密</summary>
    public Boolean IsEncrypted => Trailer.ContainsKey("Encrypt");

    private Int32 _textFontRef;
    #endregion

    #region 加载
    /// <summary>从 PDF 字节加载</summary>
    /// <param name="data">PDF 字节</param>
    /// <returns>对象级文档</returns>
    public static PdfObjectDocument Load(Byte[] data)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));

        var xref = new PdfXRefTable(data);
        if (xref.Entries.Count == 0) throw new InvalidDataException("无法解析 PDF 交叉引用表");

        var doc = new PdfObjectDocument();
        foreach (var kv in xref.Trailer)
            doc.Trailer[kv.Key] = kv.Value;

        foreach (var kv in xref.Entries)
        {
            if (!kv.Value.InUse) continue;
            var obj = PdfObjectParser.ReadObject(data, xref, kv.Key);
            if (obj != null)
            {
                doc.Objects[kv.Key] = obj;
                doc.Generations[kv.Key] = kv.Value.GenNum;
            }
        }
        return doc;
    }

    /// <summary>从文件加载</summary>
    /// <param name="path">PDF 文件路径</param>
    /// <returns>对象级文档</returns>
    public static PdfObjectDocument Load(String path) => Load(File.ReadAllBytes(path.GetFullPath()));
    #endregion

    #region 保存
    /// <summary>序列化为 PDF 字节</summary>
    /// <returns>PDF 字节</returns>
    /// <exception cref="NotSupportedException">文档已加密</exception>
    public Byte[] ToArray()
    {
        if (IsEncrypted) throw new NotSupportedException("加密 PDF 不支持对象级修改，请先解密");

        var root = Trailer.TryGetValue("Root", out var rootValue) && rootValue is Ref rootRef
            ? rootRef
            : throw new InvalidDataException("PDF 缺少 /Root");

        var size = 0;
        foreach (var objNum in Objects.Keys)
            if (objNum >= size) size = objNum + 1;

        using var ms = new MemoryStream();
        WriteAscii(ms, "%PDF-1.7\n");
        ms.WriteByte(0x25); ms.WriteByte(0xE2); ms.WriteByte(0xE3); ms.WriteByte(0xCF); ms.WriteByte(0xD3); ms.WriteByte(0x0A);

        var offsets = new Int64[size];
        for (var i = 0; i < size; i++) offsets[i] = -1;

        for (var objNum = 1; objNum < size; objNum++)
        {
            if (!Objects.TryGetValue(objNum, out var obj)) continue;

            offsets[objNum] = ms.Position;
            var gen = GetGeneration(objNum);
            WriteAscii(ms, $"{objNum} {gen} obj\n");
            WriteObject(ms, obj);
            WriteAscii(ms, "\nendobj\n");
        }

        var xrefOffset = ms.Position;
        WriteAscii(ms, $"xref\n0 {size}\n");
        WriteAscii(ms, "0000000000 65535 f \n");
        for (var i = 1; i < size; i++)
        {
            if (offsets[i] >= 0) WriteAscii(ms, $"{offsets[i]:D10} {GetGeneration(i):D5} n \n");
            else WriteAscii(ms, "0000000000 00000 f \n");
        }

        var trailer = new PdfDict { ["Size"] = new PdfNumber(size), ["Root"] = root };
        if (Trailer.TryGetValue("Info", out var info)) trailer["Info"] = info;
        if (Trailer.TryGetValue("ID", out var id)) trailer["ID"] = id;

        WriteAscii(ms, "trailer\n");
        WriteDict(ms, trailer);
        WriteAscii(ms, $"\nstartxref\n{xrefOffset}\n%%EOF\n");
        return ms.ToArray();
    }

    /// <summary>保存为 PDF 文件</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        var file = path.GetFullPath();
        file.EnsureDirectory(true);
        File.WriteAllBytes(file, ToArray());
    }

    private static void WriteAscii(Stream stream, String text)
    {
        foreach (var ch in text) stream.WriteByte((Byte)ch);
    }
    #endregion

    #region 对象序列化
    private static void WriteObject(Stream stream, PdfObject obj)
    {
        switch (obj)
        {
            case PdfNumber number:
                WriteAscii(stream, number.Value.ToString("0.###", CultureInfo.InvariantCulture));
                break;
            case PdfName name:
                WriteAscii(stream, "/" + EscapeName(name.Value));
                break;
            case PdfString text:
                WriteTextString(stream, text.Value);
                break;
            case HexString hex:
                WriteAscii(stream, "<" + hex.Value + ">");
                break;
            case PdfBoolean boolean:
                WriteAscii(stream, boolean.Value ? "true" : "false");
                break;
            case PdfNull:
                WriteAscii(stream, "null");
                break;
            case Ref reference:
                WriteAscii(stream, $"{reference.ObjNum} {reference.GenNum} R");
                break;
            case PdfArray array:
                WriteAscii(stream, "[");
                for (var i = 0; i < array.Items.Count; i++)
                {
                    if (i > 0) stream.WriteByte(0x20);
                    WriteObject(stream, array.Items[i]);
                }
                WriteAscii(stream, "]");
                break;
            case DictObj dict:
                WriteDict(stream, dict.Value);
                break;
            case PdfStream pdfStream:
            {
                var dict = CopyDict(pdfStream.Dict);
                dict["Length"] = new PdfNumber(pdfStream.Data.Length);
                WriteDict(stream, dict);
                WriteAscii(stream, "\nstream\n");
                stream.Write(pdfStream.Data, 0, pdfStream.Data.Length);
                WriteAscii(stream, "\nendstream");
                break;
            }
            default:
                throw new NotSupportedException($"不支持序列化的 PDF 对象: {obj.GetType().Name}");
        }
    }

    private static void WriteDict(Stream stream, PdfDict dict)
    {
        WriteAscii(stream, "<<");
        foreach (var kv in dict)
        {
            WriteAscii(stream, " /" + EscapeName(kv.Key) + " ");
            WriteObject(stream, kv.Value);
        }
        WriteAscii(stream, " >>");
    }

    private static void WriteTextString(Stream stream, String? value)
    {
        if (value == null)
        {
            WriteAscii(stream, "()");
            return;
        }

        // 超出 Latin-1 范围时改用 UTF-16BE 十六进制串（带 BOM）
        foreach (var ch in value)
        {
            if (ch > 255)
            {
                WriteAscii(stream, "<FEFF");
                foreach (var c in value) WriteAscii(stream, ((Int32)c).ToString("X4", CultureInfo.InvariantCulture));
                WriteAscii(stream, ">");
                return;
            }
        }

        stream.WriteByte(0x28);
        foreach (var ch in value)
        {
            if (ch == '(' || ch == ')' || ch == '\\')
            {
                stream.WriteByte(0x5C);
                stream.WriteByte((Byte)ch);
            }
            else if (ch < 32 || ch > 126)
            {
                WriteAscii(stream, "\\" + ((Int32)ch).ToString("000", CultureInfo.InvariantCulture));
            }
            else
                stream.WriteByte((Byte)ch);
        }
        stream.WriteByte(0x29);
    }

    private static String EscapeName(String name)
    {
        var sb = new StringBuilder(name.Length);
        foreach (var ch in name)
        {
            if (ch <= 32 || ch > 126 || ch == '#' || ch == '/' || ch == '(' || ch == ')' || ch == '<' || ch == '>' || ch == '[' || ch == ']' || ch == '%')
                sb.Append('#').Append(((Int32)ch).ToString("X2", CultureInfo.InvariantCulture));
            else
                sb.Append(ch);
        }
        return sb.ToString();
    }
    #endregion

    #region 对象访问
    /// <summary>新增对象并返回对象号</summary>
    /// <param name="obj">PDF 对象</param>
    /// <returns>对象号</returns>
    public Int32 AddObject(PdfObject obj)
    {
        if (obj == null) throw new ArgumentNullException(nameof(obj));

        var max = 0;
        foreach (var key in Objects.Keys)
            if (key > max) max = key;
        var objNum = max + 1;
        Objects[objNum] = obj;
        Generations[objNum] = 0;
        return objNum;
    }

    /// <summary>获取对象代数</summary>
    /// <param name="objNum">对象号</param>
    /// <returns>代数（默认 0）</returns>
    public Int32 GetGeneration(Int32 objNum) => Generations.TryGetValue(objNum, out var gen) ? gen : 0;

    /// <summary>解析间接引用（非引用原样返回）</summary>
    /// <param name="obj">对象或引用</param>
    /// <returns>实际对象，找不到返回 null</returns>
    public PdfObject? Resolve(PdfObject? obj)
    {
        if (obj is Ref reference) return Objects.TryGetValue(reference.ObjNum, out var target) ? target : null;
        return obj;
    }

    /// <summary>解析字典（引用或直接字典）</summary>
    /// <param name="obj">对象或引用</param>
    /// <returns>字典，非字典返回 null</returns>
    public DictObj? GetDict(PdfObject? obj) => Resolve(obj) as DictObj;

    /// <summary>解析数组（引用或直接数组）</summary>
    /// <param name="obj">对象或引用</param>
    /// <returns>数组，非数组返回 null</returns>
    public PdfArray? GetArray(PdfObject? obj) => Resolve(obj) as PdfArray;

    private DictObj? GetCatalog() => Trailer.TryGetValue("Root", out var root) ? GetDict(root) : null;

    private (Int32 ObjNum, DictObj Dict)? GetPagesRoot()
    {
        var catalog = GetCatalog();
        if (catalog == null || !catalog.Value.TryGetValue("Pages", out var pages)) return null;
        if (pages is Ref reference && Objects.TryGetValue(reference.ObjNum, out var obj) && obj is DictObj dict)
            return (reference.ObjNum, dict);
        if (pages is DictObj direct) return (0, direct);
        return null;
    }
    #endregion

    #region 页面
    /// <summary>按页面树顺序返回页面对象号</summary>
    /// <returns>页面对象号集合</returns>
    public List<Int32> GetPageObjectNumbers()
    {
        var result = new List<Int32>();
        var pages = GetPagesRoot();
        if (pages == null) return result;

        var kids = pages.Value.Dict.Value.TryGetValue("Kids", out var kidsValue) ? GetArray(kidsValue) : null;
        if (kids != null)
        {
            foreach (var child in kids.Items) CollectPages(child, result, 0);
        }
        return result;
    }

    private void CollectPages(PdfObject? node, List<Int32> result, Int32 depth)
    {
        if (depth > 64 || node == null) return;

        if (node is Ref reference)
        {
            if (!Objects.TryGetValue(reference.ObjNum, out var obj)) return;
            if (obj is DictObj dict && dict.Value.TryGetValue("Kids", out var kids))
            {
                var array = GetArray(kids);
                if (array != null)
                {
                    foreach (var child in array.Items) CollectPages(child, result, depth + 1);
                }
                return;
            }
            result.Add(reference.ObjNum);
            return;
        }

        if (node is DictObj direct && direct.Value.TryGetValue("Kids", out var directKids))
        {
            var array = GetArray(directKids);
            if (array != null)
            {
                foreach (var child in array.Items) CollectPages(child, result, depth + 1);
            }
        }
    }

    /// <summary>重排页面（传入完整页面对象号列表）</summary>
    /// <param name="pageObjectNumbers">按新顺序排列的页面对象号</param>
    public void SetPageOrder(IList<Int32> pageObjectNumbers)
    {
        if (pageObjectNumbers == null) throw new ArgumentNullException(nameof(pageObjectNumbers));

        var pages = GetPagesRoot() ?? throw new InvalidDataException("PDF 缺少页面树");
        var array = new PdfArray();
        foreach (var objNum in pageObjectNumbers)
            array.Items.Add(new Ref(objNum, GetGeneration(objNum)));

        pages.Dict.Value["Kids"] = array;
        pages.Dict.Value["Count"] = new PdfNumber(pageObjectNumbers.Count);

        var parentRef = pages.ObjNum > 0 ? new Ref(pages.ObjNum, GetGeneration(pages.ObjNum)) : null;
        if (parentRef != null)
        {
            foreach (var objNum in pageObjectNumbers)
            {
                var page = GetDict(new Ref(objNum, 0));
                if (page != null) page.Value["Parent"] = parentRef;
            }
        }
    }

    /// <summary>旋转页面</summary>
    /// <param name="pageIndex">页索引（0 基）</param>
    /// <param name="rotation">旋转角度（0/90/180/270）</param>
    public void RotatePage(Int32 pageIndex, Int32 rotation)
    {
        var pages = GetPageObjectNumbers();
        if (pageIndex < 0 || pageIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), $"页码越界: {pageIndex}");

        var page = GetDict(new Ref(pages[pageIndex], 0)) ?? throw new InvalidDataException("页面对象无效");
        page.Value["Rotate"] = new PdfNumber(rotation / 90 * 90);
    }

    /// <summary>只保留指定页（按给定顺序）</summary>
    /// <param name="pageIndices">保留的页索引（0 基）</param>
    public void KeepPages(IList<Int32> pageIndices)
    {
        var pages = GetPageObjectNumbers();
        var kept = new List<Int32>();
        foreach (var index in pageIndices)
        {
            if (index < 0 || index >= pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndices), $"页码越界: {index}");
            kept.Add(pages[index]);
        }
        if (kept.Count == 0) throw new ArgumentException("至少保留一页");
        SetPageOrder(kept);
    }
    #endregion

    #region 元数据
    /// <summary>设置文档信息字典（Title/Author/Subject/Keywords/Creator/Producer）</summary>
    /// <param name="values">键值集合，null 表示删除该键</param>
    public void SetMetadata(IDictionary<String, String?> values)
    {
        if (values == null) throw new ArgumentNullException(nameof(values));

        DictObj info;
        if (Trailer.TryGetValue("Info", out var infoValue) && Resolve(infoValue) is DictObj existing)
            info = existing;
        else
        {
            info = new DictObj(new PdfDict());
            Trailer["Info"] = new Ref(AddObject(info), 0);
        }

        foreach (var kv in values)
        {
            if (kv.Value == null) info.Value.Remove(kv.Key);
            else info.Value[kv.Key] = NewTextString(kv.Value);
        }
    }

    /// <summary>创建 UTF-16BE 文本字符串对象</summary>
    /// <param name="text">文本</param>
    /// <returns>十六进制字符串对象（带 BOM）</returns>
    public static HexString NewTextString(String text)
    {
        var sb = new StringBuilder("FEFF");
        foreach (var ch in text)
        {
            var bytes = Encoding.BigEndianUnicode.GetBytes([ch]);
            sb.Append(bytes[0].ToString("X2", CultureInfo.InvariantCulture));
            sb.Append(bytes[1].ToString("X2", CultureInfo.InvariantCulture));
        }
        return new HexString(sb.ToString());
    }
    #endregion

    #region 注释与表单
    /// <summary>向页面追加注释对象</summary>
    /// <param name="pageIndex">页索引（0 基）</param>
    /// <param name="annotation">注释字典</param>
    public void AddAnnotation(Int32 pageIndex, PdfObject annotation)
    {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));

        var pages = GetPageObjectNumbers();
        if (pageIndex < 0 || pageIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), $"页码越界: {pageIndex}");

        var pageObjNum = pages[pageIndex];
        var page = GetDict(new Ref(pageObjNum, 0)) ?? throw new InvalidDataException("页面对象无效");
        if (annotation is DictObj dict) dict.Value["P"] = new Ref(pageObjNum, GetGeneration(pageObjNum));

        var annotationRef = new Ref(AddObject(annotation), 0);
        var annots = GetArray(page.Value.TryGetValue("Annots", out var annotsValue) ? annotsValue : null);
        if (annots == null)
        {
            annots = new PdfArray();
            page.Value["Annots"] = annots;
        }
        annots.Items.Add(annotationRef);
    }

    /// <summary>填充表单字段</summary>
    /// <param name="name">字段名（支持点号全名）</param>
    /// <param name="value">值；"true"/"false" 用于复选框</param>
    /// <returns>是否找到并设置</returns>
    public Boolean FillForm(String name, String value)
    {
        var catalog = GetCatalog();
        if (catalog == null || !catalog.Value.TryGetValue("AcroForm", out var formValue)) return false;

        var form = GetDict(formValue);
        if (form == null) return false;

        var fields = GetArray(form.Value.TryGetValue("Fields", out var fieldsValue) ? fieldsValue : null);
        if (fields == null) return false;

        var applied = false;
        foreach (var field in fields.Items)
        {
            if (SetFieldValue(field, null, name, value)) applied = true;
        }

        if (applied) form.Value["NeedAppearances"] = new PdfBoolean(true);
        return applied;
    }

    private Boolean SetFieldValue(PdfObject node, String? prefix, String name, String value)
    {
        var field = GetDict(node);
        if (field == null) return false;

        var fieldName = ReadText(field.Value.TryGetValue("T", out var title) ? title : null);
        var fullName = prefix.IsNullOrEmpty() ? fieldName : (fieldName.IsNullOrEmpty() ? prefix : $"{prefix}.{fieldName}");
        var applied = false;

        if (!fullName.IsNullOrEmpty() && String.Equals(fullName, name, StringComparison.OrdinalIgnoreCase))
        {
            var type = field.Value.TryGetValue("FT", out var ft) ? Resolve(ft) as PdfName : null;
            var isButton = type?.Value == "Btn";

            if (isButton)
            {
                var on = String.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                    || String.Equals(value, "yes", StringComparison.OrdinalIgnoreCase)
                    || String.Equals(value, "1", StringComparison.OrdinalIgnoreCase);
                field.Value["V"] = new PdfName(on ? "Yes" : "Off");
                field.Value["AS"] = new PdfName(on ? "Yes" : "Off");
            }
            else
                field.Value["V"] = NewTextString(value);

            applied = true;
        }

        if (field.Value.TryGetValue("Kids", out var kids))
        {
            var array = GetArray(kids);
            if (array != null)
            {
                foreach (var child in array.Items)
                    if (SetFieldValue(child, fullName, name, value)) applied = true;
            }
        }

        return applied;
    }

    private String? ReadText(PdfObject? obj) => Resolve(obj) switch
    {
        PdfString text => text.Value,
        HexString hex => DecodeUtf16(hex.Value),
        _ => null,
    };

    private static String DecodeUtf16(String hex)
    {
        if (hex.Length < 4) return String.Empty;
        var bytes = new Byte[hex.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
            bytes[i] = Byte.Parse(hex.Substring(i * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        return Encoding.BigEndianUnicode.GetString(bytes);
    }
    #endregion

    #region 文字叠加
    /// <summary>在页面上叠加文字（拉丁字符，Helvetica，无嵌入字体）</summary>
    /// <param name="pageIndex">页索引（0 基）</param>
    /// <param name="text">文本（仅支持 Latin-1 字符）</param>
    /// <param name="x">左下角 X（点）</param>
    /// <param name="y">左下角 Y（点）</param>
    /// <param name="fontSize">字号（点）</param>
    /// <param name="colorHex">颜色（RRGGBB），null 为黑色</param>
    /// <param name="rotation">旋转角度（度，逆时针）</param>
    public void AddText(Int32 pageIndex, String text, Single x, Single y, Single fontSize = 12f, String? colorHex = null, Single rotation = 0)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        foreach (var ch in text)
            if (ch > 255) throw new NotSupportedException("叠加文字仅支持 Latin-1 字符；中文请改用表格/模板方式生成");

        var pages = GetPageObjectNumbers();
        if (pageIndex < 0 || pageIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), $"页码越界: {pageIndex}");

        var pageObjNum = pages[pageIndex];
        var page = GetDict(new Ref(pageObjNum, 0)) ?? throw new InvalidDataException("页面对象无效");
        EnsureFont(page);

        var radians = rotation * Math.PI / 180;
        var cos = Math.Cos(radians).ToString("0.0000", CultureInfo.InvariantCulture);
        var sin = Math.Sin(radians).ToString("0.0000", CultureInfo.InvariantCulture);
        var negSin = (-Math.Sin(radians)).ToString("0.0000", CultureInfo.InvariantCulture);
        var color = ParseColor(colorHex);

        var ops = new StringBuilder();
        ops.Append("q\n");
        ops.Append("BT\n");
        ops.Append($"/NLOfficeF1 {fontSize.ToString("0.##", CultureInfo.InvariantCulture)} Tf\n");
        ops.Append($"{color}");
        ops.Append($"{cos} {sin} {negSin} {cos} {x.ToString("0.##", CultureInfo.InvariantCulture)} {y.ToString("0.##", CultureInfo.InvariantCulture)} Tm\n");
        ops.Append('(').Append(EscapeContent(text)).Append(") Tj\n");
        ops.Append("ET\n");
        ops.Append("Q\n");

        AppendContent(page, Encoding.ASCII.GetBytes(ops.ToString()));
    }

    /// <summary>给所有页叠加水印（45° 灰色居中）</summary>
    /// <param name="text">水印文本（Latin-1）</param>
    /// <param name="fontSize">字号</param>
    /// <param name="colorHex">颜色（RRGGBB），默认浅灰</param>
    public void AddWatermark(String text, Single fontSize = 48f, String? colorHex = null)
    {
        var pages = GetPageObjectNumbers();
        for (var i = 0; i < pages.Count; i++)
        {
            var page = GetDict(new Ref(pages[i], 0));
            if (page == null) continue;

            var width = GetNumber(page.Value.TryGetValue("MediaBox", out var box) ? GetArray(box) : null, 2, 595f);
            var height = GetNumber(page.Value.TryGetValue("MediaBox", out var box2) ? GetArray(box2) : null, 3, 842f);
            var textWidth = PdfTextSearch.Measure(text, fontSize);
            var x = (width - textWidth * 0.7071f) / 2;
            var y = (height - textWidth * 0.7071f) / 2;
            AddText(i, text, x, y, fontSize, colorHex ?? "C0C0C0", 45f);
        }
    }

    private void EnsureFont(DictObj page)
    {
        if (_textFontRef == 0)
        {
            var font = new DictObj(new PdfDict
            {
                ["Type"] = new PdfName("Font"),
                ["Subtype"] = new PdfName("Type1"),
                ["BaseFont"] = new PdfName("Helvetica"),
                ["Encoding"] = new PdfName("WinAnsiEncoding"),
            });
            _textFontRef = AddObject(font);
        }

        DictObj? resources;
        if (page.Value.TryGetValue("Resources", out var resourcesValue) && GetDict(resourcesValue) is DictObj existing)
            resources = existing;
        else
        {
            resources = new DictObj(new PdfDict());
            page.Value["Resources"] = resources;
        }

        DictObj? fonts;
        if (resources.Value.TryGetValue("Font", out var fontValue) && GetDict(fontValue) is DictObj existingFonts)
            fonts = existingFonts;
        else
        {
            fonts = new DictObj(new PdfDict());
            resources.Value["Font"] = fonts;
        }

        fonts.Value["NLOfficeF1"] = new Ref(_textFontRef, 0);
    }

    private void AppendContent(DictObj page, Byte[] data)
    {
        var stream = new PdfStream { Dict = new PdfDict { ["Length"] = new PdfNumber(data.Length) }, Data = data };
        var streamRef = new Ref(AddObject(stream), 0);

        if (!page.Value.TryGetValue("Contents", out var contents) || contents == null)
        {
            page.Value["Contents"] = streamRef;
            return;
        }

        if (contents is Ref)
        {
            var array = new PdfArray { Items = { contents, streamRef } };
            page.Value["Contents"] = array;
            return;
        }

        if (GetArray(contents) is PdfArray existing)
        {
            existing.Items.Add(streamRef);
            page.Value["Contents"] = existing;
            return;
        }

        page.Value["Contents"] = streamRef;
    }

    private static String ParseColor(String? colorHex)
    {
        if (colorHex.IsNullOrEmpty() || colorHex!.Length < 6) return "0 0 0 rg\n";
        var r = Convert.ToInt32(colorHex.Substring(0, 2), 16) / 255d;
        var g = Convert.ToInt32(colorHex.Substring(2, 2), 16) / 255d;
        var b = Convert.ToInt32(colorHex.Substring(4, 2), 16) / 255d;
        return $"{r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b.ToString("0.###", CultureInfo.InvariantCulture)} rg\n";
    }

    private static String EscapeContent(String text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (ch == '(' || ch == ')' || ch == '\\') sb.Append('\\').Append(ch);
            else if (ch < 32 || ch > 126) sb.Append('\\').Append(((Int32)ch).ToString("000", CultureInfo.InvariantCulture));
            else sb.Append(ch);
        }
        return sb.ToString();
    }

    private static Single GetNumber(PdfArray? array, Int32 index, Single fallback)
    {
        if (array == null || index >= array.Items.Count) return fallback;
        return array.Items[index] is PdfNumber number ? number.Value : fallback;
    }
    #endregion

    #region 合并
    /// <summary>把另一份 PDF 的全部页面追加到当前文档</summary>
    /// <param name="other">另一份 PDF</param>
    public void Merge(PdfObjectDocument other)
    {
        if (other == null) throw new ArgumentNullException(nameof(other));

        var offset = 0;
        foreach (var key in Objects.Keys)
            if (key >= offset) offset = key + 1;
        if (offset <= 0) offset = 1;

        var map = new Dictionary<Int32, Int32>();
        foreach (var key in other.Objects.Keys) map[key] = key + offset;

        foreach (var kv in other.Objects)
        {
            Objects[map[kv.Key]] = Remap(kv.Value, map);
            Generations[map[kv.Key]] = other.GetGeneration(kv.Key);
        }

        var pages = GetPageObjectNumbers();
        foreach (var page in other.GetPageObjectNumbers())
            pages.Add(map[page]);

        SetPageOrder(pages);
    }

    private static PdfObject Remap(PdfObject obj, Dictionary<Int32, Int32> map) => obj switch
    {
        Ref reference => new Ref(map.TryGetValue(reference.ObjNum, out var mapped) ? mapped : reference.ObjNum, reference.GenNum),
        DictObj dict => new DictObj(CopyDict(dict.Value.Select(kv => new KeyValuePair<String, PdfObject>(kv.Key, Remap(kv.Value, map))))),
        PdfArray array => new PdfArray { Items = array.Items.Select(e => Remap(e, map)).ToList() },
        PdfStream stream => new PdfStream
        {
            Dict = CopyDict(stream.Dict.Select(kv => new KeyValuePair<String, PdfObject>(kv.Key, Remap(kv.Value, map)))),
            Data = stream.Data,
        },
        _ => obj,
    };

    private static PdfDict CopyDict(IEnumerable<KeyValuePair<String, PdfObject>> source)
    {
        var dict = new PdfDict();
        foreach (var kv in source) dict[kv.Key] = kv.Value;
        return dict;
    }

    /// <summary>复制当前文档（可只保留指定页）</summary>
    /// <param name="pageIndices">保留的页索引（0 基），null 表示全部</param>
    /// <returns>新文档</returns>
    public PdfObjectDocument Clone(IList<Int32>? pageIndices = null)
    {
        var clone = new PdfObjectDocument();
        var map = new Dictionary<Int32, Int32>();
        foreach (var key in Objects.Keys) map[key] = key;
        foreach (var kv in Objects) clone.Objects[kv.Key] = Remap(kv.Value, map);
        foreach (var kv in Generations) clone.Generations[kv.Key] = kv.Value;
        foreach (var kv in Trailer) clone.Trailer[kv.Key] = Remap(kv.Value, map);

        if (pageIndices != null) clone.KeepPages(pageIndices);
        return clone;
    }
    #endregion
}
