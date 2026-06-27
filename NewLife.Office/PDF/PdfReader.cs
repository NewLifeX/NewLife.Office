using System.Globalization;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 读取器</summary>
/// <remarks>
/// 基于交叉引用表（xref）精确定位 PDF 对象，支持 FlateDecode 解压缩内容流。
/// 支持 PDF 1.0-1.7，正确提取文本、图片、元数据和书签结构。
/// 对加密 PDF 的文本提取有限（需先解密）。
/// </remarks>
public class PdfReader : IDisposable, ITextExtractable, IMarkdownExtractable
{
    #region 属性
    /// <summary>源文件路径</summary>
    public String? FilePath { get; private set; }

    /// <summary>交叉引用表</summary>
    public PdfXRefTable? XRefTable { get; private set; }
    #endregion

    #region 私有字段
    private readonly Byte[] _data;
    private readonly Encoding _latin1 = Encoding.GetEncoding(28591);
    #endregion

    #region 构造
    /// <summary>从文件路径打开</summary>
    /// <param name="path">PDF 文件路径</param>
    public PdfReader(String path)
    {
        FilePath = path.GetFullPath();
        _data = File.ReadAllBytes(FilePath);
        XRefTable = new PdfXRefTable(_data);
    }

    /// <summary>从流打开</summary>
    /// <param name="stream">包含 PDF 内容的流</param>
    public PdfReader(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        _data = ms.ToArray();
        XRefTable = new PdfXRefTable(_data);
    }

    /// <summary>释放资源</summary>
    public void Dispose() => GC.SuppressFinalize(this);
    #endregion

    #region FlateDecode 解压（公开给 XRefTable 使用）
    /// <summary>解压 FlateDecode（zlib/Deflate）压缩的数据</summary>
    /// <param name="data">压缩数据（含 zlib 头 + deflate + Adler-32 尾）</param>
    /// <returns>解压后的字节数组</returns>
    public static Byte[] DecompressFlate(Byte[] data)
    {
        // zlib 格式：2字节头 + deflate 数据 + 4字节 Adler-32
        // CMF=0x78, FLG 常见值：0x01/0x5E/0x9C/0xDA
        if (data.Length < 6) return data;

        var headerOffset = 0;
        // 跳过 2 字节 zlib 头
        if (data.Length >= 2 && data[0] == 0x78)
            headerOffset = 2;

        using var output = new MemoryStream();
        using var input = new MemoryStream(data, headerOffset, data.Length - headerOffset - (headerOffset > 0 ? 4 : 0));
        using var deflate = new System.IO.Compression.DeflateStream(input, System.IO.Compression.CompressionMode.Decompress);
        deflate.CopyTo(output);
        return output.ToArray();
    }

    /// <summary>根据流字典解压缩内容流</summary>
    /// <param name="streamData">原始流数据</param>
    /// <param name="dict">流字典</param>
    /// <returns>解压后的字节数组</returns>
    internal static Byte[] DecompressStreamData(Byte[] streamData, PdfDict dict)
    {
        if (!dict.TryGetValue("Filter", out var filterVal)) return streamData;

        // 支持单一过滤器或过滤器数组
        var filters = new List<String>();
        if (filterVal is PdfName fn)
            filters.Add(fn.Value);
        else if (filterVal is PdfArray fa)
            filters.AddRange(fa.Items.OfType<PdfName>().Select(n => n.Value));

        var result = streamData;
        foreach (var filter in filters)
        {
            switch (filter)
            {
                case "FlateDecode":
                    try { result = DecompressFlate(result); } catch { }
                    break;
                case "ASCIIHexDecode":
                    result = DecodeAsciiHex(result);
                    break;
                case "ASCII85Decode":
                    result = DecodeAscii85(result);
                    break;
                case "LZWDecode":
                    // 从 dict 中读取 EarlyChange 参数（默认 1）
                    var earlyChange = 1;
                    if (dict.TryGetValue("EarlyChange", out var ecVal) && ecVal is PdfNumber ecn)
                        earlyChange = (Int32)ecn.Value;
                    result = DecodeLzw(result, earlyChange);
                    break;
                case "RunLengthDecode":
                    result = DecodeRunLength(result);
                    break;
            }
        }
        return result;
    }
    #endregion

    #region 读取方法
    /// <summary>获取总页数（通过 /Count 字段）</summary>
    /// <returns>页数</returns>
    public Int32 GetPageCount()
    {
        var pdf = _latin1.GetString(_data);
        var countIdx = FindToken(pdf, "/Count");
        if (countIdx < 0) return 0;
        var numStr = ExtractNextToken(pdf, countIdx + 6);
        return Int32.TryParse(numStr.Trim(), out var count) ? count : 0;
    }

    /// <summary>读取 PDF 中所有嵌入/引用的字体信息</summary>
    /// <returns>字体信息列表</returns>
    public List<PdfFontInfo> ReadFonts()
    {
        var fonts = new List<PdfFontInfo>();
        var pdf = _latin1.GetString(_data);
        var visited = new HashSet<String>();

        // 扫描所有字体字典：查找 /BaseFont 或 /FontName
        var pos = 0;
        while (true)
        {
            var found = FindToken(pdf.Substring(pos), "/BaseFont");
            if (found < 0) break;
            pos += found + 9;
            var name = ExtractNameValue(pdf, ref pos);
            if (name == null || !visited.Add(name)) continue;

            var info = new PdfFontInfo { Name = name };

            // 回溯查找字体类型和编码
            var searchBack = Math.Max(0, pos - 200);
            var context = pdf.Substring(searchBack, pos - searchBack);

            // 检测类型：/Type /Font 附近或 /Subtype
            var typeIdx = context.LastIndexOf("/Subtype");
            if (typeIdx >= 0)
            {
                var subtype = ExtractNextToken(context, typeIdx + 8).Trim();
                if (subtype.Length > 0 && subtype[0] == '/') subtype = subtype.Substring(1);
                info.Type = subtype;
            }

            // 检测编码
            var encIdx = context.LastIndexOf("/Encoding");
            if (encIdx >= 0)
            {
                var encoding = ExtractNextToken(context, encIdx + 9).Trim();
                if (encoding.Length > 0 && encoding[0] == '/') encoding = encoding.Substring(1);
                info.Encoding = encoding;
            }

            // 排除重复的字体描述符（字体名过短且以 F 开头的通常不是真实字体）
            if (info.Name != null && (info.Name.Length > 3 || !info.Name.StartsWith("F")))
                fonts.Add(info);
        }

        // 如果找到太多（包括 FontDescriptor 等），去重过滤
        var result = fonts.Where(f => f.Name?.Length > 1).GroupBy(f => f.Name).Select(g => g.First()).ToList();
        return result;
    }

    /// <summary>提取 PDF name 值（/xxx 格式）</summary>
    private static String? ExtractNameValue(String pdf, ref Int32 pos)
    {
        while (pos < pdf.Length && pdf[pos] == ' ') pos++;
        if (pos >= pdf.Length) return null;

        // 跳过 #XX 十六进制转义
        if (pos + 1 < pdf.Length && pdf[pos] == '#') { pos += 3; }

        var start = pos;
        if (start < pdf.Length && pdf[start] == '/') start++;

        while (pos < pdf.Length && !IsPdfDelimiter(pdf[pos]))
            pos++;

        if (start >= pos) return null;

        var len = pos - start;
        if (start + len > pdf.Length) len = pdf.Length - start;
        if (len <= 0) return null;

        var name = pdf.Substring(start, len);
        name = name.Trim();
        if (name.StartsWith("/")) name = name.Substring(1);
        return name.Length > 0 ? name : null;
    }

    private static Boolean IsPdfDelimiter(Char c)
    {
        return c == ' ' || c == '\r' || c == '\n' || c == '\t' || c == '/' || c == '<' || c == '>' || c == '[' || c == ']' || c == '(' || c == ')';
    }

    /// <summary>提取全部文本（基于 xref + 解压缩内容流）</summary>
    /// <returns>合并后的文本</returns>
    public String ExtractText()
    {
        var sb = new StringBuilder();
        ExtractFromStreams(_data, sb);
        return sb.ToString();
    }

    /// <summary>基于 xref 表提取全部文本</summary>
    private void ExtractTextViaXref(StringBuilder sb)
    {
        var pageObjNums = GetPageObjectNumbers();
        foreach (var pageObjNum in pageObjNums)
        {
            var pageObj = PdfObjectParser.ReadObject(_data, XRefTable!, pageObjNum);
            if (pageObj is not PdfDictObj pageDictObj) continue;

            var pageDict = pageDictObj.Value;
            if (!pageDict.TryGetValue("Contents", out var contentsVal)) continue;

            // 处理单个内容流或内容流数组
            var contentRefs = new List<PdfRef>();
            if (contentsVal is PdfRef r)
                contentRefs.Add(r);
            else if (contentsVal is PdfArray arr)
                contentRefs.AddRange(arr.Items.OfType<PdfRef>());

            foreach (var cref in contentRefs)
            {
                var contentObj = PdfObjectParser.ReadObject(_data, XRefTable!, cref.ObjNum);
                if (contentObj is not PdfStream contentStream) continue;

                var decompressed = DecompressStreamData(contentStream.Data, contentStream.Dict);
                var contentText = _latin1.GetString(decompressed);
                ExtractTextFromContent(contentText, sb);

                sb.AppendLine();
            }
        }
    }

    /// <summary>获取所有页面对象号</summary>
    private List<Int32> GetPageObjectNumbers()
    {
        var pageObjNums = new List<Int32>();
        if (XRefTable == null) return pageObjNums;

        // 从 Catalog → Pages → Kids 获取页面引用
        if (!XRefTable.Trailer.TryGetValue("Root", out var rootVal) || rootVal is not PdfRef rootRef)
            return pageObjNums;

        var catalogObj = PdfObjectParser.ReadObject(_data, XRefTable, rootRef.ObjNum);
        if (catalogObj is not PdfDictObj catalogDictObj) return pageObjNums;

        if (!catalogDictObj.Value.TryGetValue("Pages", out var pagesVal) || pagesVal is not PdfRef pagesRef)
            return pageObjNums;

        var pagesObj = PdfObjectParser.ReadObject(_data, XRefTable, pagesRef.ObjNum);
        if (pagesObj is not PdfDictObj pagesDictObj) return pageObjNums;

        var pagesDict = pagesDictObj.Value;
        if (!pagesDict.TryGetValue("Kids", out var kidsVal) || kidsVal is not PdfArray kidsArr)
            return pageObjNums;

        foreach (var kid in kidsArr.Items)
        {
            if (kid is PdfRef kidRef)
                pageObjNums.Add(kidRef.ObjNum);
        }

        return pageObjNums;
    }

    /// <summary>读取文档元数据</summary>
    /// <returns>元数据对象</returns>
    public PdfMetadata ReadMetadata()
    {
        var meta = new PdfMetadata { PageCount = GetPageCount() };
        var pdf = _latin1.GetString(_data);

        // 读取 %PDF-x.x 版本
        if (pdf.StartsWith("%PDF-"))
            meta.PdfVersion = pdf.Substring(5, Math.Min(3, pdf.Length - 5)).Trim();

        // 优先从 xref trailer 读取
        if (XRefTable != null && XRefTable.Trailer.TryGetValue("Info", out var infoVal) && infoVal is PdfRef infoRef)
        {
            var infoObj = PdfObjectParser.ReadObject(_data, XRefTable, infoRef.ObjNum);
            if (infoObj is PdfDictObj infoDict)
            {
                meta.Title = GetDictString(infoDict.Value, "Title");
                meta.Author = GetDictString(infoDict.Value, "Author");
                meta.Subject = GetDictString(infoDict.Value, "Subject");
                meta.CreationDate = GetDictString(infoDict.Value, "CreationDate");
                return meta;
            }
        }

        // 回退：扫描 Info 字典
        var infoStart = FindToken(pdf, "/Info");
        if (infoStart >= 0)
        {
            var dictText = ExtractDict(pdf, infoStart);
            meta.Title = GetDictValue(dictText, "Title");
            meta.Author = GetDictValue(dictText, "Author");
            meta.Subject = GetDictValue(dictText, "Subject");
            meta.CreationDate = GetDictValue(dictText, "CreationDate");
        }

        return meta;
    }

    /// <summary>从 PdfDict 中提取字符串值</summary>
    private static String? GetDictString(PdfDict dict, String key)
    {
        if (!dict.TryGetValue(key, out var val)) return null;
        return val switch
        {
            PdfString s => s.Value,
            PdfHexString h => DecodeHexFromString(h.Value),
            _ => val.ToString(),
        };
    }

    /// <summary>提取带坐标位置的文本</summary>
    /// <remarks>基于 xref + 内容流解压，正确追踪文本矩阵</remarks>
    /// <returns>文本项序列</returns>
    public IEnumerable<PdfText> ExtractTextWithPositions()
    {
        var results = new List<PdfText>();
        if (XRefTable == null || !XRefTable.Trailer.TryGetValue("Root", out var rv) || rv is not PdfRef)
        {
            // 回退到字符串扫描
            var pdf = _latin1.GetString(_data);
            var pos = 0;
            while (pos < pdf.Length)
            {
                var streamStart = pdf.IndexOf("stream", pos, StringComparison.Ordinal);
                if (streamStart < 0) break;
                var contentStart = streamStart + 6;
                if (contentStart < pdf.Length && pdf[contentStart] == '\r') contentStart++;
                if (contentStart < pdf.Length && pdf[contentStart] == '\n') contentStart++;
                var streamEnd = pdf.IndexOf("endstream", contentStart, StringComparison.Ordinal);
                if (streamEnd < 0) break;
                var content = pdf[contentStart..streamEnd];
                ExtractPositionedText(content, results);
                pos = streamEnd + 9;
            }
            return results;
        }

        var pageObjNums = GetPageObjectNumbers();
        foreach (var pageObjNum in pageObjNums)
        {
            var pageObj = PdfObjectParser.ReadObject(_data, XRefTable, pageObjNum);
            if (pageObj is not PdfDictObj pageDictObj) continue;

            var pageDict = pageDictObj.Value;
            if (!pageDict.TryGetValue("Contents", out var contentsVal)) continue;

            var contentRefs = new List<PdfRef>();
            if (contentsVal is PdfRef r) contentRefs.Add(r);
            else if (contentsVal is PdfArray arr) contentRefs.AddRange(arr.Items.OfType<PdfRef>());

            foreach (var cref in contentRefs)
            {
                var contentObj = PdfObjectParser.ReadObject(_data, XRefTable, cref.ObjNum);
                if (contentObj is not PdfStream contentStream) continue;

                var decompressed = DecompressStreamData(contentStream.Data, contentStream.Dict);
                var contentText = _latin1.GetString(decompressed);
                ExtractPositionedText(contentText, results);
            }
        }
        return results;
    }

    /// <summary>读取 AcroForm 表单字段及当前值</summary>
    /// <returns>表单对象（含字段列表和值），null 表示无表单</returns>
    public PdfForm? ReadFormFields()
    {
        if (XRefTable == null || !XRefTable.Trailer.TryGetValue("Root", out var rv) || rv is not PdfRef rootRef)
            return null;

        var rootObj = PdfObjectParser.ReadObject(_data, XRefTable, rootRef.ObjNum);
        if (rootObj is not PdfDictObj rootDictObj) return null;
        if (!rootDictObj.Value.TryGetValue("AcroForm", out var acroVal)) return null;

        PdfRef acroRef;
        if (acroVal is PdfRef ar) acroRef = ar;
        else return null;

        var acroObj = PdfObjectParser.ReadObject(_data, XRefTable, acroRef.ObjNum);
        if (acroObj is not PdfDictObj acroDictObj) return null;

        var form = new PdfForm();
        var dict = acroDictObj.Value;

        if (dict.TryGetValue("NeedAppearances", out var naVal) && naVal is PdfBoolean na)
            form.NeedAppearances = na.Value;

        if (dict.TryGetValue("Fields", out var fieldsVal))
        {
            var fieldRefs = new List<PdfRef>();
            if (fieldsVal is PdfRef fr) fieldRefs.Add(fr);
            else if (fieldsVal is PdfArray fa) fieldRefs.AddRange(fa.Items.OfType<PdfRef>());

            foreach (var fref in fieldRefs)
            {
                var field = ReadFormField(fref);
                if (field != null) form.Fields.Add(field);
            }
        }

        return form;
    }

    private PdfFormField? ReadFormField(PdfRef fref)
    {
        var fobj = PdfObjectParser.ReadObject(_data, XRefTable!, fref.ObjNum);
        if (fobj is not PdfDictObj fdictObj) return null;
        var dict = fdictObj.Value;

        var field = new PdfFormField();

        if (dict.TryGetValue("T", out var tVal) && tVal is PdfString ts)
            field.FullName = ts.Value;

        if (dict.TryGetValue("FT", out var ftVal) && ftVal is PdfName ftn)
            field.FieldType = ftn.Value switch { "Tx" => PdfFormFieldType.Tx, "Btn" => PdfFormFieldType.Btn, "Ch" => PdfFormFieldType.Ch, "Sig" => PdfFormFieldType.Sig, _ => PdfFormFieldType.Tx };

        if (dict.TryGetValue("V", out var vVal))
            field.Value = vVal switch { PdfString vs => vs.Value, PdfName vn => vn.Value, _ => vVal.ToString() };

        if (dict.TryGetValue("DV", out var dvVal) && dvVal is PdfString dvs)
            field.DefaultValue = dvs.Value;

        // 解析 Rect（[x1 y1 x2 y2]）
        if (dict.TryGetValue("Rect", out var rectVal) && rectVal is PdfArray rectArr && rectArr.Items.Count >= 4)
        {
            if (rectArr.Items[0] is PdfNumber rx) field.X = (Single)rx.Value;
            if (rectArr.Items[1] is PdfNumber ry) field.Y = (Single)ry.Value;
            if (rectArr.Items[2] is PdfNumber rw) field.Width = (Single)rw.Value - field.X;
            if (rectArr.Items[3] is PdfNumber rh) field.Height = (Single)rh.Value - field.Y;
        }

        // Kids（子字段，递归读取）
        if (dict.TryGetValue("Kids", out var kidsVal))
        {
            var kidRefs = new List<PdfRef>();
            if (kidsVal is PdfRef kr) kidRefs.Add(kr);
            else if (kidsVal is PdfArray ka) kidRefs.AddRange(ka.Items.OfType<PdfRef>());

            foreach (var kref in kidRefs)
            {
                var kid = ReadFormField(kref);
                if (kid != null) field.Kids.Add(kid);
            }
        }

        // Opt（下拉/列表选项）
        if (dict.TryGetValue("Opt", out var optVal) && optVal is PdfArray optArr)
        {
            foreach (var oi in optArr.Items)
            {
                if (oi is PdfString os) field.Options.Add(os.Value);
                else if (oi is PdfArray oa && oa.Items.Count > 0 && oa.Items[0] is PdfString oas) field.Options.Add(oas.Value);
            }
        }

        return field;
    }

    /// <summary>从 PDF 中提取嵌入图片</summary>
    /// <returns>图片流对象序列</returns>
    public IEnumerable<PdfImage> ExtractImageStreams()
    {
        if (XRefTable == null || !XRefTable.Trailer.TryGetValue("Root", out var _))
            return ExtractImageStreamsLegacy();

        var results = new List<PdfImage>();
        var imgIdx = 0;
        foreach (var kv in XRefTable.Entries)
        {
            if (!kv.Value.InUse) continue;
            var obj = PdfObjectParser.ReadObject(_data, XRefTable, kv.Key);
            if (obj is not PdfStream stream) continue;

            var dict = stream.Dict;
            if (!dict.TryGetValue("Subtype", out var subtypeVal) || subtypeVal is not PdfName subName || subName.Value != "Image") continue;
            if (!dict.TryGetValue("Type", out var typeVal) || typeVal is not PdfName tName || tName.Value != "XObject") continue;

            var width = 0; var height = 0;
            if (dict.TryGetValue("Width", out var wVal) && wVal is PdfNumber wn) width = (Int32)wn.Value;
            if (dict.TryGetValue("Height", out var hVal) && hVal is PdfNumber hn) height = (Int32)hn.Value;
            if (width <= 0 || height <= 0) continue;

            var filter = String.Empty;
            if (dict.TryGetValue("Filter", out var fVal))
            {
                if (fVal is PdfName fn) filter = fn.Value;
                else if (fVal is PdfArray fa) filter = String.Join(",", fa.Items.OfType<PdfName>().Select(n => n.Value));
            }

            var rawData = stream.Data;
            var isJpeg = filter.IndexOf("DCTDecode", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!isJpeg && filter.IndexOf("FlateDecode", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                try { rawData = DecompressStreamData(stream.Data, dict); } catch { }
            }

            results.Add(new PdfImage { Index = imgIdx++, Width = width, Height = height, Filter = filter, RawData = rawData });
        }
        return results;
    }

    private IEnumerable<PdfImage> ExtractImageStreamsLegacy()
    {
        var imgIdx = 0;
        var text = _latin1.GetString(_data);
        var pos = 0;
        while (pos < text.Length)
        {
            var dictStart = text.IndexOf("<<", pos, StringComparison.Ordinal);
            if (dictStart < 0) break;
            var dictEnd = text.IndexOf(">>", dictStart + 2, StringComparison.Ordinal);
            if (dictEnd < 0) break;
            var dict = text.Substring(dictStart, dictEnd - dictStart + 2);

            if (dict.IndexOf("/Subtype", StringComparison.Ordinal) >= 0 && dict.IndexOf("/Image", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var widthTok = GetDictIntValue(dict, "Width");
                var heightTok = GetDictIntValue(dict, "Height");
                var filter = GetDictToken(dict, "Filter");
                var strmPos = text.IndexOf("stream", dictEnd, StringComparison.Ordinal);
                if (strmPos >= 0 && strmPos < dictEnd + 100)
                {
                    var dataStart = strmPos + 6;
                    if (dataStart < text.Length && text[dataStart] == '\r') dataStart++;
                    if (dataStart < text.Length && text[dataStart] == '\n') dataStart++;
                    var dataEnd = text.IndexOf("endstream", dataStart, StringComparison.Ordinal);
                    if (dataEnd > dataStart && widthTok > 0 && heightTok > 0)
                    {
                        var rawBytes = _data.AsSpan(dataStart, Math.Min(dataEnd - dataStart, _data.Length - dataStart)).ToArray();
                        yield return new PdfImage { Index = imgIdx++, Width = widthTok, Height = heightTok, Filter = filter ?? String.Empty, RawData = rawBytes };
                        pos = dataEnd + 9;
                        continue;
                    }
                }
            }
            pos = dictEnd + 2;
        }
    }

    /// <summary>读取整个 PDF 文档为 PdfDocument 数据模型</summary>
    /// <returns>包含元数据、页面和书签信息的 PdfDocument</returns>
    public PdfDocument ReadDocument()
    {
        var meta = ReadMetadata();
        var doc = new PdfDocument { Metadata = meta };

        // 读取页面信息
        var pageObjNums = GetPageObjectNumbers();
        foreach (var pageObjNum in pageObjNums)
        {
            var pageObj = PdfObjectParser.ReadObject(_data, XRefTable!, pageObjNum);
            if (pageObj is not PdfDictObj pageDictObj) continue;

            var pageDict = pageDictObj.Value;
            var page = new PdfPage();

            // MediaBox
            if (pageDict.TryGetValue("MediaBox", out var mbVal) && mbVal is PdfArray mb && mb.Items.Count >= 4)
            {
                page.Width = ((PdfNumber)mb.Items[2]).Value;
                page.Height = ((PdfNumber)mb.Items[3]).Value;
            }

            // Rotate
            if (pageDict.TryGetValue("Rotate", out var rotVal) && rotVal is PdfNumber rn)
                page.Rotation = (Int32)rn.Value;

            doc.Pages.Add(page);
        }

        // 读取书签大纲（/Outlines）
        if (XRefTable!.Trailer.TryGetValue("Root", out var rootVal2) && rootVal2 is PdfRef rootRef2)
        {
            var catalogObj2 = PdfObjectParser.ReadObject(_data, XRefTable, rootRef2.ObjNum);
            if (catalogObj2 is PdfDictObj catDict2 &&
                catDict2.Value.TryGetValue("Outlines", out var outVal) && outVal is PdfRef outRef)
            {
                var outlinesObj = PdfObjectParser.ReadObject(_data, XRefTable, outRef.ObjNum);
                if (outlinesObj is PdfDictObj outDict)
                {
                    if (outDict.Value.TryGetValue("First", out var firstVal) && firstVal is PdfRef firstRef)
                        ReadBookmarkTree(firstRef.ObjNum, doc.Bookmarks);
                }
            }
        }

        return doc;
    }

    /// <summary>递归读取书签树</summary>
    private void ReadBookmarkTree(Int32 objNum, List<PdfBookmark> bookmarks)
    {
        var obj = PdfObjectParser.ReadObject(_data, XRefTable!, objNum);
        if (obj is not PdfDictObj dictObj) return;

        var dict = dictObj.Value;
        var bm = new PdfBookmark();

        if (dict.TryGetValue("Title", out var titleVal))
        {
            if (titleVal is PdfString ts) bm.Title = ts.Value;
            else if (titleVal is PdfHexString ths) bm.Title = DecodeHexFromString(ths.Value);
        }

        // 解析 /Dest 获取目标页索引
        if (dict.TryGetValue("Dest", out var destVal) && destVal is PdfArray destArr && destArr.Items.Count >= 1)
        {
            if (destArr.Items[0] is PdfRef pageRef)
            {
                var pageObjNums = GetPageObjectNumbers();
                bm.PageIndex = pageObjNums.IndexOf(pageRef.ObjNum);
                if (bm.PageIndex < 0) bm.PageIndex = 0;
            }
        }

        bookmarks.Add(bm);

        // 递归读取子书签（/First）
        if (dict.TryGetValue("First", out var childVal) && childVal is PdfRef childRef)
            ReadBookmarkTree(childRef.ObjNum, bm.Children);

        // 读取下一个兄弟书签（/Next）
        if (dict.TryGetValue("Next", out var nextVal) && nextVal is PdfRef nextRef)
            ReadBookmarkTree(nextRef.ObjNum, bookmarks);
    }
    #endregion

    #region 私有方法
    /// <summary>从 PDF 内容流中提取文本（解析 Tj/TJ 操作符）</summary>
    private static void ExtractFromStreams(Byte[] pdfData, StringBuilder sb)
    {
        // 扫描所有 stream...endstream 块
        var pdf = Encoding.GetEncoding(1252).GetString(pdfData);
        var pos = 0;
        while (pos < pdf.Length)
        {
            var streamKeyStart = pdf.IndexOf("stream", pos, StringComparison.Ordinal);
            if (streamKeyStart < 0) break;

            // 跳过 "stream\r\n" 或 "stream\n"
            var contentStart = streamKeyStart + 6;
            if (contentStart < pdf.Length && pdf[contentStart] == '\r') contentStart++;
            if (contentStart < pdf.Length && pdf[contentStart] == '\n') contentStart++;

            // ── 解析紧邻 stream 关键字之前的流字典 ──
            // 向前最多搜索 800 字节，找 >> 再找对应 <<
            var lookBack  = Math.Min(streamKeyStart, 800);
            var dictEndPos = pdf.LastIndexOf(">>", streamKeyStart, lookBack, StringComparison.Ordinal);
            var dict = String.Empty;
            if (dictEndPos > 0)
            {
                var dictStartPos = pdf.LastIndexOf("<<", dictEndPos, Math.Min(dictEndPos, 600), StringComparison.Ordinal);
                if (dictStartPos >= 0)
                    dict = pdf.Substring(dictStartPos, dictEndPos - dictStartPos + 2);
            }

            // 从字典解析 /Length，用于精确跳过二进制数据中可能出现的假 "endstream"
            var streamLength = ParseStreamLength(dict);

            // ── 查找真正的 endstream ──
            // 若已知 /Length，从 contentStart+length 附近开始搜，避免二进制内的假命中
            int streamEnd;
            if (streamLength > 0)
            {
                var searchFrom = Math.Min(contentStart + streamLength, pdf.Length - 9);
                streamEnd = pdf.IndexOf("endstream", searchFrom, StringComparison.Ordinal);
                if (streamEnd < 0)
                    streamEnd = pdf.IndexOf("endstream", contentStart, StringComparison.Ordinal);
            }
            else
            {
                streamEnd = pdf.IndexOf("endstream", contentStart, StringComparison.Ordinal);
            }
            if (streamEnd < 0) break;

            pos = streamEnd + 9; // 无论是否提取文本，都正确前进

            // ── 跳过非内容流 ──

            // 1. /Length1 = 嵌入字体二进制流；/ColorSpace = 图片流
            if (dict.IndexOf("/Length1", StringComparison.Ordinal) >= 0 ||
                dict.IndexOf("/ColorSpace", StringComparison.Ordinal) >= 0)
                continue;

            // 2. ToUnicode/CMap 流（以 /CIDInit 或 begincmap 开头）
            var peekLen = Math.Min(40, streamEnd - contentStart);
            if (peekLen > 0)
            {
                var peek = pdf.Substring(contentStart, peekLen).TrimStart();
                if (peek.StartsWith("/CIDInit", StringComparison.Ordinal) ||
                    peek.StartsWith("begincmap", StringComparison.Ordinal))
                    continue;
            }

            // 3. 开头 200 字节中非打印字符比例超过 25%（CIDToGIDMap 等二进制表）
            var checkLen     = Math.Min(streamEnd - contentStart, 200);
            var nonPrintable = 0;
            for (var ci = contentStart; ci < contentStart + checkLen; ci++)
            {
                var b = pdf[ci];
                if (b < 9 || (b > 13 && b < 32)) nonPrintable++;
            }
            if (checkLen > 0 && nonPrintable * 4 > checkLen) continue;

            var streamContent = pdf[contentStart..streamEnd];
            ExtractTextFromContent(streamContent, sb);
        }
    }

    /// <summary>从流字典文本中解析 /Length 值（不含 /Length1/Length2 等衍生键）</summary>
    private static Int32 ParseStreamLength(String dict)
    {
        var idx = 0;
        while (idx < dict.Length)
        {
            var found = dict.IndexOf("/Length", idx, StringComparison.Ordinal);
            if (found < 0) break;
            var afterKey = found + 7; // 跳过 "/Length"
            // 排除 /Length1、/Length2 等
            if (afterKey < dict.Length && (dict[afterKey] >= '0' && dict[afterKey] <= '9' || dict[afterKey] == 'a' || dict[afterKey] == 'A'))
            {
                idx = afterKey;
                continue;
            }
            // 跳过空白，读取数字
            while (afterKey < dict.Length && dict[afterKey] == ' ') afterKey++;
            var numEnd = afterKey;
            while (numEnd < dict.Length && dict[numEnd] >= '0' && dict[numEnd] <= '9') numEnd++;
            if (numEnd > afterKey &&
                Int32.TryParse(dict.Substring(afterKey, numEnd - afterKey), out var len))
                return len;
            idx = afterKey;
        }
        return -1;
    }

    /// <summary>从 PDF 内容流字符串中提取文本操作符</summary>
    private static void ExtractTextFromContent(String content, StringBuilder sb)
    {
        // 解析 (text) Tj 和 [(text)] TJ 操作符
        var i = 0;
        while (i < content.Length)
        {
            if (content[i] == '(')
            {
                // 读取括号字符串
                var str = ReadParenString(content, ref i);
                // 查找后续操作符
                var opPos = i;
                SkipWhitespace(content, ref opPos);
                if (opPos < content.Length - 1)
                {
                    var op = content.Substring(opPos, 2);
                    if (op.StartsWith("Tj") || op.StartsWith("TJ") || op.StartsWith("'") || op.StartsWith("\""))
                    {
                        sb.Append(DecodePdfString(str));
                        i = opPos + (op.StartsWith("Tj") || op.StartsWith("TJ") ? 2 : 1);
                        continue;
                    }
                }
            }
            else if (content[i] == '<' && i + 1 < content.Length && content[i + 1] != '<')
            {
                // 读取 <hex> 字符串（CJK UTF-16BE 编码或 Latin-1 hex）
                var hexEnd = content.IndexOf('>', i + 1);
                if (hexEnd > i)
                {
                    var hexStr = content.Substring(i + 1, hexEnd - i - 1);
                    i = hexEnd + 1;
                    var opPos = i;
                    SkipWhitespace(content, ref opPos);
                    if (opPos + 1 < content.Length)
                    {
                        var op2 = content.Substring(opPos, Math.Min(2, content.Length - opPos));
                        if (op2.StartsWith("Tj") || op2.StartsWith("TJ"))
                        {
                            sb.Append(DecodeHexString(hexStr));
                            i = opPos + 2;
                            continue;
                        }
                    }
                    continue;
                }
            }
            else if (content[i] == '[')
            {
                // TJ array
                var arrEnd = content.IndexOf(']', i);
                if (arrEnd > i)
                {
                    var arr = content.Substring(i + 1, arrEnd - i - 1);
                    ExtractTextFromContent(arr, sb);
                    i = arrEnd + 1;
                    // skip TJ
                    SkipWhitespace(content, ref i);
                    if (i < content.Length - 1 && content.Substring(i, 2) == "TJ")
                        i += 2;
                    continue;
                }
            }
            else if (content[i] == 'T' && i + 1 < content.Length && content[i + 1] == '*')
            {
                sb.AppendLine();
                i += 2;
                continue;
            }
            else if (content[i] == 'B' && i + 1 < content.Length && content[i + 1] == 'T')
            {
                i += 2;
                continue;
            }
            else if (content[i] == 'E' && i + 3 < content.Length && content.Substring(i, 2) == "ET")
            {
                sb.AppendLine();
                i += 2;
                continue;
            }
            i++;
        }
    }

    private static String ReadParenString(String s, ref Int32 pos)
    {
        pos++; // skip '('
        var sb = new StringBuilder();
        var depth = 1;
        while (pos < s.Length && depth > 0)
        {
            var c = s[pos];
            if (c == '\\' && pos + 1 < s.Length)
            {
                sb.Append(s[pos + 1]);
                pos += 2;
                continue;
            }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) { pos++; break; } }
            if (depth > 0) sb.Append(c);
            pos++;
        }
        return sb.ToString();
    }

    private static void SkipWhitespace(String s, ref Int32 pos)
    {
        while (pos < s.Length && (s[pos] == ' ' || s[pos] == '\t' || s[pos] == '\r' || s[pos] == '\n'))
        {
            pos++;
        }
    }

    private static String DecodePdfString(String s)
    {
        // Basic: remove non-printable control chars, keep Latin-1 printables
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (c >= 32 && c < 256) sb.Append(c);
            else if (c == '\n' || c == '\r') sb.Append(' ');
        }
        return sb.ToString();
    }

    /// <summary>解码 PDF hex 字符串（&lt;XXXX...&gt;）为 Unicode 文本</summary>
    private static String DecodeHexString(String hex)
    {
        // 移除空白字符
        var clean = new StringBuilder(hex.Length);
        foreach (var c in hex)
        {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n')
                clean.Append(c);
        }
        var h = clean.ToString();
        if (h.Length == 0 || h.Length % 2 != 0) return String.Empty;
        var byteCount = h.Length / 2;
        var bytes = new Byte[byteCount];
        for (var j = 0; j < byteCount; j++)
        {
            if (!Byte.TryParse(h.Substring(j * 2, 2), NumberStyles.HexNumber, null, out bytes[j]))
                return String.Empty;
        }
        // UTF-16BE（我方 CJK 字体编码）：字节数必须为 2 的倍数
        if (byteCount % 2 == 0)
        {
            try
            {
                var text = Encoding.BigEndianUnicode.GetString(bytes);
                // 确认解码结果有打印字符（避免将 Latin-1 hex 误识为 UTF-16BE）
                var printable = 0;
                foreach (var c in text)
                    if (c >= 32) printable++;
                if (printable > 0 && printable * 2 >= text.Length)
                    return text;
            }
            catch { }
        }
        // 回退：Latin-1 单字节解码
        try { return Encoding.GetEncoding(1252).GetString(bytes); }
        catch { return String.Empty; }
    }

    private static Int32 FindToken(String pdf, String token)
    {
        var idx = pdf.IndexOf(token, StringComparison.Ordinal);
        return idx;
    }

    private static String ExtractNextToken(String pdf, Int32 pos)
    {
        SkipWhitespace(pdf, ref pos);
        var end = pos;
        while (end < pdf.Length && pdf[end] != ' ' && pdf[end] != '\n' && pdf[end] != '\r'
               && pdf[end] != '/' && pdf[end] != '<' && pdf[end] != '>')
            end++;
        return pdf[pos..end];
    }

    private static String ExtractDict(String pdf, Int32 startOffset)
    {
        // find << ... >>
        var start = pdf.IndexOf("<<", startOffset, StringComparison.Ordinal);
        if (start < 0) return String.Empty;
        var end = pdf.IndexOf(">>", start + 2, StringComparison.Ordinal);
        if (end < 0) return String.Empty;
        return pdf.Substring(start, end - start + 2);
    }

    private static String? GetDictValue(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return null;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        if (valStart >= dict.Length) return null;
        if (dict[valStart] == '(')
        {
            var tmp = valStart;
            return ReadParenString(dict, ref tmp);
        }
        return ExtractNextToken(dict, valStart);
    }

    /// <summary>从 PDF 字典字符串中提取整型值</summary>
    /// <param name="dict">字典字符串</param>
    /// <param name="key">键名（不含前导 /）</param>
    /// <returns>整型值，未找到时返回 0</returns>
    private static Int32 GetDictIntValue(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return 0;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        var str = ExtractNextToken(dict, valStart);
        return Int32.TryParse(str.Trim(), out var v) ? v : 0;
    }

    /// <summary>从 PDF 字典字符串中提取 Name 类型的值（不含前导 /）</summary>
    /// <param name="dict">字典字符串</param>
    /// <param name="key">键名（不含前导 /）</param>
    /// <returns>Name 值字符串，未找到时返回 null</returns>
    private static String? GetDictToken(String dict, String key)
    {
        var tag = $"/{key}";
        var idx = dict.IndexOf(tag, StringComparison.Ordinal);
        if (idx < 0) return null;
        var valStart = idx + tag.Length;
        SkipWhitespace(dict, ref valStart);
        if (valStart >= dict.Length) return null;
        if (dict[valStart] == '/')
        {
            var nameEnd = valStart + 1;
            while (nameEnd < dict.Length && dict[nameEnd] != ' ' && dict[nameEnd] != '\t'
                   && dict[nameEnd] != '\r' && dict[nameEnd] != '\n'
                   && dict[nameEnd] != '/' && dict[nameEnd] != '<' && dict[nameEnd] != '>')
                nameEnd++;
            return dict.Substring(valStart + 1, nameEnd - valStart - 1);
        }
        return ExtractNextToken(dict, valStart).Trim();
    }

    /// <summary>从 PDF 内容流中提取带位置信息的文本</summary>
    /// <param name="content">内容流字符串</param>
    /// <param name="results">结果列表</param>
    private static void ExtractPositionedText(String content, List<PdfText> results)
    {
        var curX = 0f;
        var curY = 0f;
        var fontSize = 0f;
        var inText = false;
        var i = 0;
        var numStack = new List<Single>();

        while (i < content.Length)
        {
            SkipWhitespace(content, ref i);
            if (i >= content.Length) break;

            var c = content[i];

            // PDF 注释行
            if (c == '%')
            {
                while (i < content.Length && content[i] != '\n') i++;
                continue;
            }

            // 括号字符串 (text)
            if (c == '(')
            {
                var s = ReadParenString(content, ref i);
                if (inText)
                {
                    var peek = i;
                    SkipWhitespace(content, ref peek);
                    if (peek + 1 < content.Length && content[peek] == 'T'
                        && (content[peek + 1] == 'j' || content[peek + 1] == 'J'
                            || content[peek + 1] == '\'' || content[peek + 1] == '"'))
                    {
                        var decoded = DecodePdfString(s);
                        if (decoded.Length > 0)
                            results.Add(new PdfText { Text = decoded, X = curX, Y = curY, FontSize = fontSize });
                        i = peek + 2;
                        numStack.Clear();
                        continue;
                    }
                }
                numStack.Clear();
                continue;
            }

            // 嵌套字典 << >>
            if (c == '<' && i + 1 < content.Length && content[i + 1] == '<')
            {
                var end = content.IndexOf(">>", i + 2, StringComparison.Ordinal);
                i = end >= 0 ? end + 2 : i + 2;
                numStack.Clear();
                continue;
            }

            // 十六进制字符串 <hex>
            if (c == '<')
            {
                var end = content.IndexOf('>', i + 1);
                if (end > i)
                {
                    if (inText)
                    {
                        var hex = content.Substring(i + 1, end - i - 1);
                        var decoded = HexToString(hex);
                        if (decoded.Length > 0)
                        {
                            var peek = end + 1;
                            SkipWhitespace(content, ref peek);
                            if (peek + 1 < content.Length && content[peek] == 'T' && content[peek + 1] == 'j')
                            {
                                results.Add(new PdfText { Text = decoded, X = curX, Y = curY, FontSize = fontSize });
                                i = peek + 2;
                                numStack.Clear();
                                continue;
                            }
                        }
                    }
                    i = end + 1;
                }
                else
                    i++;
                continue;
            }

            // TJ 数组 [...]
            if (c == '[')
            {
                var arrEnd = content.IndexOf(']', i);
                if (arrEnd > i && inText)
                {
                    var arr = content.Substring(i + 1, arrEnd - i - 1);
                    var arrSb = new StringBuilder();
                    var ap = 0;
                    while (ap < arr.Length)
                    {
                        SkipWhitespace(arr, ref ap);
                        if (ap >= arr.Length) break;
                        if (arr[ap] == '(')
                        {
                            var s = ReadParenString(arr, ref ap);
                            arrSb.Append(DecodePdfString(s));
                        }
                        else
                        {
                            // 数字（字间距调整）或其他 — 跳过到下一个空白或 (
                            while (ap < arr.Length && arr[ap] != '(' && arr[ap] != ' '
                                   && arr[ap] != '\t' && arr[ap] != '\r' && arr[ap] != '\n')
                                ap++;
                        }
                    }
                    var txt = arrSb.ToString();
                    if (txt.Length > 0)
                        results.Add(new PdfText { Text = txt, X = curX, Y = curY, FontSize = fontSize });
                    i = arrEnd + 1;
                    SkipWhitespace(content, ref i);
                    if (i + 1 < content.Length && content[i] == 'T' && content[i + 1] == 'J')
                        i += 2;
                }
                else
                    i = arrEnd >= 0 ? arrEnd + 1 : i + 1;
                numStack.Clear();
                continue;
            }

            // PDF Name（/name），用于 Tf 的字体名等
            if (c == '/')
            {
                var nameEnd = i + 1;
                while (nameEnd < content.Length && content[nameEnd] != ' ' && content[nameEnd] != '\t'
                       && content[nameEnd] != '\r' && content[nameEnd] != '\n'
                       && content[nameEnd] != '/' && content[nameEnd] != '<' && content[nameEnd] != '>')
                    nameEnd++;
                i = nameEnd;
                continue;
            }

            // 数字（操作符参数）
            if (Char.IsDigit(c) || c == '-'
                || (c == '.' && i + 1 < content.Length && Char.IsDigit(content[i + 1])))
            {
                var numEnd = i + 1;
                while (numEnd < content.Length && (Char.IsDigit(content[numEnd]) || content[numEnd] == '.'))
                {
                    numEnd++;
                }
                if (Single.TryParse(content[i..numEnd],
                    NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                    numStack.Add(num);
                i = numEnd;
                continue;
            }

            // 操作符
            if (Char.IsLetter(c) || c == '\'' || c == '"' || c == '*')
            {
                var opEnd = i + 1;
                while (opEnd < content.Length
                       && (Char.IsLetterOrDigit(content[opEnd]) || content[opEnd] == '*'))
                    opEnd++;
                var op = content[i..opEnd];
                i = opEnd;

                switch (op)
                {
                    case "BT":
                        inText = true;
                        curX = 0;
                        curY = 0;
                        break;
                    case "ET":
                        inText = false;
                        break;
                    case "Td":
                    case "TD":
                        if (numStack.Count >= 2)
                        {
                            curX += numStack[numStack.Count - 2];
                            curY += numStack[numStack.Count - 1];
                        }
                        break;
                    case "Tm":
                        if (numStack.Count >= 6)
                        {
                            curX = numStack[numStack.Count - 2];
                            curY = numStack[numStack.Count - 1];
                        }
                        break;
                    case "Tf":
                        if (numStack.Count >= 1)
                            fontSize = numStack[numStack.Count - 1];
                        break;
                    case "T*":
                        curY -= fontSize > 0 ? fontSize * 1.2f : 12f;
                        break;
                }
                numStack.Clear();
                continue;
            }

            i++;
        }
    }

    /// <summary>将 PDF 十六进制字符串转换为可读文本</summary>
    /// <param name="hex">十六进制字符串（可含空白符）</param>
    /// <returns>可打印字符序列</returns>
    private static String HexToString(String hex)
    {
        var sb = new StringBuilder();
        var clean = hex.ToCharArray();
        var ci = 0;
        var cleanBuf = new StringBuilder(hex.Length);
        while (ci < clean.Length)
        {
            var ch = clean[ci++];
            if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n') continue;
            cleanBuf.Append(ch);
        }
        var s = cleanBuf.ToString();
        for (var k = 0; k + 1 < s.Length; k += 2)
        {
            if (Byte.TryParse(s.Substring(k, 2), NumberStyles.HexNumber, null, out var b) && b >= 32)
                sb.Append((Char)b);
        }
        return sb.ToString();
    }
    #endregion

    #region 流解码器
    /// <summary>解码 ASCIIHexDecode 编码的数据</summary>
    private static Byte[] DecodeAsciiHex(Byte[] data)
    {
        var text = Encoding.ASCII.GetString(data);
        var hex = new StringBuilder();
        foreach (var c in text)
        {
            if (c == '>') break; // EOD
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n')
                hex.Append(c);
        }
        var h = hex.ToString();
        if (h.Length % 2 != 0) h += "0";
        var bytes = new Byte[h.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
        {
            if (!Byte.TryParse(h.Substring(i * 2, 2), NumberStyles.HexNumber, null, out bytes[i]))
                bytes[i] = 0;
        }
        return bytes;
    }

    /// <summary>解码 ASCII85Decode 编码的数据</summary>
    private static Byte[] DecodeAscii85(Byte[] data)
    {
        var text = Encoding.ASCII.GetString(data);
        using var output = new MemoryStream();
        var pos = 0;
        while (pos < text.Length)
        {
            var c = text[pos];
            if (c == '~' && pos + 1 < text.Length && text[pos + 1] == '>') break; // EOD

            if (c >= '!' && c <= 'u')
            {
                var group = new List<Byte>(5);
                while (group.Count < 5 && pos < text.Length)
                {
                    var ch = text[pos];
                    if (ch >= '!' && ch <= 'u')
                        group.Add((Byte)(ch - '!'));
                    else if (ch == '~')
                        break;
                    else
                    {
                        pos++;
                        continue;
                    }
                    pos++;
                }

                var n = group.Count;
                if (n == 0) continue;

                // 补到 5 字节
                while (group.Count < 5) group.Add(84); // 'u' value

                UInt32 value = 0;
                for (var i = 0; i < 5; i++)
                    value = value * 85 + group[i];

                for (var i = 0; i < n - 1 && i < 4; i++)
                    output.WriteByte((Byte)(value >> (24 - i * 8)));
            }
            else pos++;
        }
        return output.ToArray();
    }

    /// <summary>LZW 解码（PDF LZWDecode 过滤器）</summary>
    /// <param name="data">LZW 编码数据</param>
    /// <param name="earlyChange">EarlyChange 参数（PDF 默认 1）</param>
    /// <returns>解码后的原始数据</returns>
    /// <remarks>
    /// PDF LZW 解码参数：
    /// - 初始码宽 9 位
    /// - 清除码 256，EOD 码 257
    /// - EarlyChange=1 时在码值 511（而非 512）后增加码宽
    /// </remarks>
    public static Byte[] DecodeLzw(Byte[] data, Int32 earlyChange = 1)
    {
        if (data == null || data.Length == 0) return [];

        using var output = new MemoryStream();
        var bitPos = 0;     // 当前位位置
        var codeSize = 9;   // 当前码宽
        var clearCode = 256;
        var eodCode = 257;
        var nextCode = 258;
        var maxCode = (1 << codeSize) - 1;

        // 字典：code → byte[]（可变长度）
        var dict = new Dictionary<Int32, Byte[]>();
        for (var i = 0; i < 256; i++)
            dict[i] = [(Byte)i];

        Int32 prevCode = -1;
        var firstCode = true;

        while (true)
        {
            var code = ReadLzwBits(data, ref bitPos, codeSize);
            if (code < 0) break; // 数据不足

            if (code == eodCode) break;
            if (code == clearCode)
            {
                // 重置字典
                dict.Clear();
                for (var i = 0; i < 256; i++)
                    dict[i] = [(Byte)i];
                nextCode = 258;
                codeSize = 9;
                maxCode = (1 << codeSize) - 1;
                prevCode = -1;
                firstCode = true;
                continue;
            }

            Byte[] entry;
            if (dict.TryGetValue(code, out var existing))
            {
                entry = existing;
            }
            else if (code == nextCode && prevCode >= 0)
            {
                // KwKwK 特殊情况：prev 序列 + prev 首字节
                var prevEntry = dict[prevCode];
                entry = new Byte[prevEntry.Length + 1];
                Array.Copy(prevEntry, entry, prevEntry.Length);
                entry[prevEntry.Length] = prevEntry[0];
                dict[nextCode] = entry;
                nextCode++;
            }
            else
            {
                // 无效码值，尝试恢复
                break;
            }

            output.Write(entry, 0, entry.Length);

            // 构建新字典条目
            if (!firstCode && code != nextCode - 1)
            {
                var prevEntry = dict[prevCode];
                var newEntry = new Byte[prevEntry.Length + 1];
                Array.Copy(prevEntry, newEntry, prevEntry.Length);
                newEntry[prevEntry.Length] = entry[0];

                if (nextCode <= maxCode - (1 - earlyChange))
                {
                    dict[nextCode] = newEntry;
                    nextCode++;
                }
            }

            // 检查是否需要增加码宽
            var threshold = (1 << codeSize) - earlyChange;
            if (nextCode > threshold && codeSize < 12)
            {
                codeSize++;
                maxCode = (1 << codeSize) - 1;
            }

            prevCode = code;
            firstCode = false;
        }

        return output.ToArray();
    }

    /// <summary>从字节数组中读取指定位数的 LZW 码值</summary>
    private static Int32 ReadLzwBits(Byte[] data, ref Int32 bitPos, Int32 numBits)
    {
        if (numBits <= 0 || numBits > 16) return -1;

        var value = 0;
        for (var i = 0; i < numBits; i++)
        {
            var byteIdx = (bitPos + i) / 8;
            if (byteIdx >= data.Length) return -1;
            var bitIdx = (bitPos + i) % 8;
            if ((data[byteIdx] & (1 << bitIdx)) != 0)
                value |= (1 << i);
        }
        bitPos += numBits;
        return value;
    }

    /// <summary>RunLengthDecode（PDF RLE）解压缩</summary>
    /// <param name="data">RLE 编码数据</param>
    /// <returns>解码后的原始数据</returns>
    /// <remarks>
    /// PDF RunLength 编码：
    /// - 字节 n (0-127): 复制接下来的 n+1 个字节原样输出
    /// - 字节 n (128-255): 将下一个字节重复 257-n 次
    /// - 字节 128: EOD（流结束）
    /// </remarks>
    public static Byte[] DecodeRunLength(Byte[] data)
    {
        if (data == null || data.Length == 0) return [];

        using var output = new MemoryStream();
        var pos = 0;
        while (pos < data.Length)
        {
            var n = data[pos++];
            if (n == 128) break; // EOD

            if (n < 128)
            {
                // 复制接下来的 n+1 个字节
                var count = n + 1;
                if (pos + count > data.Length) count = data.Length - pos;
                output.Write(data, pos, count);
                pos += count;
            }
            else
            {
                // 重复下一个字节 257-n 次
                if (pos >= data.Length) break;
                var count = 257 - n;
                var b = data[pos++];
                for (var i = 0; i < count; i++)
                    output.WriteByte(b);
            }
        }
        return output.ToArray();
    }

    /// <summary>从十六进制字符串解码为文本</summary>
    internal static String DecodeHexFromString(String hex)
    {
        var clean = new StringBuilder(hex.Length);
        foreach (var c in hex)
        {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n')
                clean.Append(c);
        }
        var h = clean.ToString();
        if (h.Length == 0 || h.Length % 2 != 0) return String.Empty;

        var bytes = new Byte[h.Length / 2];
        for (var j = 0; j < bytes.Length; j++)
        {
            if (!Byte.TryParse(h.Substring(j * 2, 2), NumberStyles.HexNumber, null, out bytes[j]))
                return String.Empty;
        }

        // UTF-16BE 猜测
        if (bytes.Length % 2 == 0)
        {
            try
            {
                var text = Encoding.BigEndianUnicode.GetString(bytes);
                var printable = 0;
                foreach (var ch in text) if (ch >= 32) printable++;
                if (printable > 0 && printable * 2 >= text.Length) return text;
            }
            catch { }
        }

        try { return Encoding.GetEncoding(1252).GetString(bytes); }
        catch { return String.Empty; }
    }
    #endregion

    #region 文本提取
    /// <summary>提取纯文本</summary>
    /// <returns>纯文本字符串</returns>
    String? ITextExtractable.ExtractText() => ExtractText();

    /// <summary>提取 Markdown 格式</summary>
    /// <returns>Markdown 字符串</returns>
    String? IMarkdownExtractable.ExtractMarkdown() => ExtractText();
    #endregion
}
