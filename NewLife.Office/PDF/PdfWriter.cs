using System.Security.Cryptography;
using System.Text;

namespace NewLife.Office;

/// <summary>PDF 写入器</summary>
/// <remarks>
/// 纯 C# 实现的基础 PDF 生成器，无外部依赖。
/// 使用 PDF 1.4 规范，支持多页、文本、线段、矩形、表格和图片。
/// 内置标准 Type1 字体（Helvetica/Times/Courier），对中文使用系统宋体（需系统安装）。
/// 注意：中文文字若无嵌入字体，外部 PDF 阅读器需安装相应 CJK 字体包。
/// </remarks>
public class PdfWriter : IDisposable
{
    #region 属性
    /// <summary>页面宽度（点）A4 = 595</summary>
    public Single PageWidth { get; set; } = 595f;

    /// <summary>页面高度（点）A4 = 842</summary>
    public Single PageHeight { get; set; } = 842f;

    /// <summary>上边距（点）</summary>
    public Single MarginTop { get; set; } = 56f;

    /// <summary>下边距（点）</summary>
    public Single MarginBottom { get; set; } = 56f;

    /// <summary>左边距（点）</summary>
    public Single MarginLeft { get; set; } = 56f;

    /// <summary>右边距（点）</summary>
    public Single MarginRight { get; set; } = 56f;

    /// <summary>当前可用宽度</summary>
    public Single ContentWidth => PageWidth - MarginLeft - MarginRight;

    /// <summary>当前 Y 坐标（从顶部向下，会随内容追加下移）</summary>
    public Single CurrentY { get; private set; }

    /// <summary>所有页面集合</summary>
    public List<PdfPage> Pages { get; } = [];

    /// <summary>当前页面</summary>
    public PdfPage? CurrentPage { get; private set; }

    /// <summary>页眉文本，null 表示不显示</summary>
    public String? HeaderText { get; set; }

    /// <summary>页脚文本，null 表示不显示</summary>
    public String? FooterText { get; set; }

    /// <summary>是否在页脚显示页码</summary>
    public Boolean ShowPageNumbers { get; set; }

    /// <summary>文档标题（写入 PDF Info 字典）</summary>
    public String? DocumentTitle { get; set; }

    /// <summary>文档作者（写入 PDF Info 字典）</summary>
    public String? DocumentAuthor { get; set; }

    /// <summary>文档主题</summary>
    public String? DocumentSubject { get; set; }

    /// <summary>用户密码（文档打开密码），null 表示不加密</summary>
    public String? UserPassword { get; set; }

    /// <summary>所有者密码（权限管理密码），null 时回退到 UserPassword</summary>
    public String? OwnerPassword { get; set; }

    /// <summary>权限标志位（PDF 标准，-1 表示全部允许，-3904 表示允许打印/复制，-3844 表示禁止修改）</summary>
    public Int32 Permissions { get; set; } = -1;

    /// <summary>书签列表</summary>
    public List<PdfBookmark> Bookmarks { get; } = [];
    #endregion

    #region 私有字段
    private readonly List<PdfFont> _fonts = [];
    private readonly StringBuilder _content = new();
    private Int32 _imgCounter = 1;
    private readonly PdfFont _fontHelvetica = new("F1", "Helvetica");
    private readonly PdfFont _fontTimesBold = new("F2", "Times-Bold");
    private readonly PdfFont _fontCourier = new("F3", "Courier");
    #endregion

    #region 构造
    /// <summary>实例化 PDF 写入器</summary>
    public PdfWriter()
    {
        _fonts.Add(_fontHelvetica);
        _fonts.Add(_fontTimesBold);
        _fonts.Add(_fontCourier);
    }

    /// <summary>释放资源</summary>
    public void Dispose() { GC.SuppressFinalize(this); }
    #endregion

    #region 页面方法
    /// <summary>开始新页面</summary>
    /// <returns>新页面对象</returns>
    public PdfPage BeginPage()
    {
        // 如果有未结束的页面先结束
        if (CurrentPage != null) EndPageInternal();

        CurrentPage = new PdfPage { Width = PageWidth, Height = PageHeight };
        _content.Clear();
        CurrentY = MarginTop;
        return CurrentPage;
    }

    /// <summary>结束当前页面并加入集合</summary>
    public void EndPage()
    {
        if (CurrentPage == null) return;
        EndPageInternal();
    }

    private void EndPageInternal()
    {
        if (CurrentPage == null) return;
        CurrentPage!.ContentBytes = Encoding.GetEncoding(28591).GetBytes(_content.ToString());
        Pages.Add(CurrentPage);
        CurrentPage = null;
        _content.Clear();
    }
    #endregion

    #region 绘图方法
    /// <summary>在指定位置绘制文本（坐标从左下角量起）</summary>
    /// <param name="text">文本内容</param>
    /// <param name="x">X 坐标（点）</param>
    /// <param name="y">Y 坐标（点，从页面底部量起）</param>
    /// <param name="fontSize">字号（磅）</param>
    /// <param name="font">字体（null=使用 Helvetica）</param>
    public void DrawText(String text, Single x, Single y, Single fontSize = 12, PdfFont? font = null)
    {
        EnsurePage();
        font ??= _fontHelvetica;
        _content.AppendLine("BT");
        _content.AppendLine($"/{font.Name} {fontSize:F1} Tf");
        _content.AppendLine($"{x:F2} {y:F2} Td");
        if (font.IsCjk)
            _content.AppendLine($"<{EncodeCjkHex(text)}> Tj");
        else
            _content.AppendLine($"({EncodePdfText(text)}) Tj");
        _content.AppendLine("ET");
    }

    /// <summary>创建简体中文字体（Adobe 预定义 STSong-Light，PDF 阅读器需安装 CJK 字体包）</summary>
    /// <returns>已注册的 CJK 字体，可传入 DrawText / AppendLine</returns>
    public PdfFont CreateSimplifiedChineseFont()
    {
        var fname = $"F{_fonts.Count + 1}";
        var font = new PdfFont(fname, "STSong-Light", isCjk: true);
        _fonts.Add(font);
        return font;
    }

    /// <summary>追加文本行（自动换行，跟踪当前 Y 位置，Y 从顶部开始）</summary>
    /// <param name="text">文本内容</param>
    /// <param name="fontSize">字号（磅）</param>
    /// <param name="font">字体</param>
    /// <param name="indentX">与左边距的额外水平偏移</param>
    public void AppendLine(String text, Single fontSize = 12, PdfFont? font = null, Single indentX = 0)
    {
        EnsurePage();
        font ??= _fontHelvetica;
        var lineHeight = fontSize * 1.4f;
        // 换页检测
        if (CurrentY + lineHeight > PageHeight - MarginBottom)
        {
            EndPage();
            BeginPage();
        }
        var y = PageHeight - CurrentY - fontSize;
        DrawText(text, MarginLeft + indentX, y, fontSize, font);
        CurrentY += lineHeight;
    }

    /// <summary>追加空行</summary>
    /// <param name="height">行高（点），默认等于正文行高</param>
    public void AppendEmptyLine(Single height = 14f) => CurrentY += height;

    /// <summary>绘制直线</summary>
    /// <param name="x1">起点 X</param>
    /// <param name="y1">起点 Y（从底部量起）</param>
    /// <param name="x2">终点 X</param>
    /// <param name="y2">终点 Y</param>
    /// <param name="lineWidth">线宽（点）</param>
    /// <param name="colorHex">颜色（16进制 RGB 如 "000000"）</param>
    public void DrawLine(Single x1, Single y1, Single x2, Single y2, Single lineWidth = 0.5f, String? colorHex = null)
    {
        EnsurePage();
        _content.AppendLine("q");
        if (colorHex != null) _content.AppendLine(HexToRgbOp(colorHex, false));
        _content.AppendLine($"{lineWidth:F2} w");
        _content.AppendLine($"{x1:F2} {y1:F2} m {x2:F2} {y2:F2} l S");
        _content.AppendLine("Q");
    }

    /// <summary>绘制矩形</summary>
    /// <param name="x">左下角 X（从底部量起）</param>
    /// <param name="y">左下角 Y</param>
    /// <param name="w">宽度</param>
    /// <param name="h">高度</param>
    /// <param name="filled">是否填充</param>
    /// <param name="fillColorHex">填充色（16进制 RGB）</param>
    /// <param name="strokeColorHex">边框色</param>
    /// <param name="lineWidth">边框线宽</param>
    public void DrawRect(Single x, Single y, Single w, Single h,
        Boolean filled = false, String? fillColorHex = null, String? strokeColorHex = null, Single lineWidth = 0.5f)
    {
        EnsurePage();
        _content.AppendLine("q");
        _content.AppendLine($"{lineWidth:F2} w");
        if (strokeColorHex != null) _content.AppendLine(HexToRgbOp(strokeColorHex, false));
        if (filled && fillColorHex != null) _content.AppendLine(HexToRgbOp(fillColorHex, true));
        _content.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
        _content.AppendLine(filled ? (strokeColorHex != null ? "B" : "f") : "S");
        _content.AppendLine("Q");
    }

    /// <summary>绘制表格（从当前 Y 向下追加）</summary>
    /// <param name="rows">行列数据，rows[0] 可作为表头</param>
    /// <param name="firstRowHeader">首行是否表头（加粗、灰色背景）</param>
    /// <param name="columnWidths">各列宽比例（null则平均分）</param>
    public void DrawTable(IEnumerable<String[]> rows, Boolean firstRowHeader = true, Single[]? columnWidths = null)
    {
        EnsurePage();
        var rowList = rows.ToList();
        if (rowList.Count == 0) return;
        var colCount = rowList.Max(r => r.Length);
        if (colCount == 0) return;

        // 归一化列宽
        Single[] colWidths;
        if (columnWidths != null && columnWidths.Length == colCount)
        {
            var total = columnWidths.Sum();
            colWidths = columnWidths.Select(w => w / total * ContentWidth).ToArray();
        }
        else
        {
            var unit = ContentWidth / colCount;
            colWidths = Enumerable.Repeat(unit, colCount).ToArray();
        }

        const Single rowH = 18f;
        const Single fontSize = 10f;
        const Single padding = 3f;

        for (var ri = 0; ri < rowList.Count; ri++)
        {
            // 换页检测
            if (CurrentY + rowH > PageHeight - MarginBottom)
            {
                EndPage();
                BeginPage();
            }

            var row = rowList[ri];
            var isHeader = ri == 0 && firstRowHeader;
            var rowTopY = PageHeight - CurrentY;
            var rowBottomY = rowTopY - rowH;

            // 背景
            if (isHeader)
            {
                DrawRect(MarginLeft, rowBottomY, ContentWidth, rowH, true, "D0D0D0", "000000", 0.3f);
            }
            else
            {
                DrawRect(MarginLeft, rowBottomY, ContentWidth, rowH, false, null, "000000", 0.3f);
            }

            // 列分隔线 + 文字
            var cellX = MarginLeft;
            for (var ci = 0; ci < colCount; ci++)
            {
                var cellW = ci < colWidths.Length ? colWidths[ci] : colWidths[^1];
                var cellText = ci < row.Length ? row[ci] : String.Empty;
                var textY = rowBottomY + padding;
                DrawText(cellText, cellX + padding, textY, fontSize,
                    isHeader ? _fontTimesBold : _fontHelvetica);
                cellX += cellW;
            }

            CurrentY += rowH;
        }
        // bottom border
        var tableBottomY = PageHeight - CurrentY;
        DrawLine(MarginLeft, tableBottomY, MarginLeft + ContentWidth, tableBottomY, 0.3f);
        AppendEmptyLine(4f);
    }

    /// <summary>嵌入并绘制 PNG 图片</summary>
    /// <param name="imageData">图片字节（PNG 格式）</param>
    /// <param name="x">左下角 X（从底部量起）</param>
    /// <param name="y">左下角 Y</param>
    /// <param name="w">显示宽度（点）</param>
    /// <param name="h">显示高度（点）</param>
    public void DrawImage(Byte[] imageData, Single x, Single y, Single w, Single h)
    {
        EnsurePage();
        var imgName = $"Im{_imgCounter++}";
        var (imgW, imgH) = GetPngSize(imageData);
        CurrentPage!.Images[imgName] = (imageData, imgW, imgH, false);
        _content.AppendLine("q");
        _content.AppendLine($"{w:F2} 0 0 {h:F2} {x:F2} {y:F2} cm");
        _content.AppendLine($"/{imgName} Do");
        _content.AppendLine("Q");
    }

    /// <summary>追加图片（自动跟踪 Y 位置）</summary>
    /// <param name="imageData">图片字节</param>
    /// <param name="widthPt">显示宽度（点）</param>
    /// <param name="heightPt">显示高度（点）</param>
    public void AppendImage(Byte[] imageData, Single widthPt, Single heightPt)
    {
        EnsurePage();
        if (CurrentY + heightPt > PageHeight - MarginBottom)
        {
            EndPage();
            BeginPage();
        }
        var y = PageHeight - CurrentY - heightPt;
        DrawImage(imageData, MarginLeft, y, widthPt, heightPt);
        CurrentY += heightPt + 6f;
    }

    /// <summary>在当前页面添加超链接注释区域</summary>
    /// <param name="x">左边距（点，原点在左下角）</param>
    /// <param name="y">下边距（点，原点在左下角）</param>
    /// <param name="w">宽度（点）</param>
    /// <param name="h">高度（点）</param>
    /// <param name="url">目标 URL</param>
    public void AddHyperlink(Single x, Single y, Single w, Single h, String url)
    {
        EnsurePage();
        CurrentPage!.LinkAnnotations.Add((x, y, w, h, url));
    }

    /// <summary>在当前 AppendLine 位置添加超链接（适用于追加文本之后立即调用）</summary>
    /// <param name="url">目标 URL</param>
    /// <param name="lineHeight">文本行高（默认 14）</param>
    public void AddHyperlinkForLastLine(String url, Single lineHeight = 14f)
    {
        EnsurePage();
        var y = PageHeight - CurrentY; // 当前行顶部的 PDF y 坐标
        AddHyperlink(MarginLeft, y, ContentWidth, lineHeight, url);
    }

    /// <summary>添加书签，指向当前（最后一）页</summary>
    /// <param name="title">书签标题</param>
    /// <returns>书签对象</returns>
    public PdfBookmark AddBookmark(String title)
    {
        var bm = new PdfBookmark { Title = title, PageIndex = Pages.Count };
        Bookmarks.Add(bm);
        return bm;
    }

    /// <summary>旋转指定页面</summary>
    /// <param name="pageIndex">页面索引（0起始）</param>
    /// <param name="rotation">旋转角度（0/90/180/270）</param>
    public void RotatePage(Int32 pageIndex, Int32 rotation)
    {
        if (pageIndex >= 0 && pageIndex < Pages.Count)
            Pages[pageIndex].Rotation = rotation / 90 * 90;
    }

    /// <summary>将对象集合以表格形式写入 PDF</summary>
    /// <param name="data">对象集合</param>
    /// <param name="firstRowHeader">首行表头</param>
    public void WriteObjects<T>(IEnumerable<T> data, Boolean firstRowHeader = true) where T : class
    {
        var props = typeof(T).GetProperties();
        var headers = props.Select(p =>
        {
            var dn = p.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), false)
                      .OfType<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault()?.DisplayName;
            return dn ?? p.Name;
        }).ToArray();
        var rows = new List<String[]> { headers };
        foreach (var item in data)
            rows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());
        DrawTable(rows, firstRowHeader);
    }
    #endregion

    #region 保存方法
    /// <summary>保存到文件</summary>
    /// <param name="path">输出路径</param>
    public void Save(String path)
    {
        using var fs = new FileStream(path.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Save(fs);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">目标流</param>
    public void Save(Stream stream)
    {
        // 结束最后一页
        if (CurrentPage != null) EndPage();

        // 如果没有内容，创建空白页
        if (Pages.Count == 0)
        {
            BeginPage();
            EndPage();
        }

        BuildPdf(stream);
    }
    #endregion

    #region PDF 构建
    private void BuildPdf(Stream stream)
    {
        using var ms = new MemoryStream();
        var offsets = new List<Int64>();
        var latin1 = Encoding.GetEncoding(28591);

        void WriteObj(Int32 id, String content)
        {
            while (offsets.Count < id) offsets.Add(0);
            offsets[id - 1] = ms.Position;
            var bytes = latin1.GetBytes($"{id} 0 obj\n{content}\nendobj\n");
            ms.Write(bytes, 0, bytes.Length);
        }

        // Header
        var header = latin1.GetBytes("%PDF-1.4\n%\xFF\xFF\xFF\xFF\n");
        ms.Write(header, 0, header.Length);

        var allPages = Pages;
        var pageCount = allPages.Count;

        // ── 对象 ID 预分配 ──
        var nextId = 2; // 1=Catalog, 2=Pages 已占用
        Int32 NextId() => ++nextId;

        for (var i = 0; i < pageCount; i++)
        {
            allPages[i].PageObjId = NextId();
            allPages[i].ContentObjId = NextId();
        }
        var fontObjIds = new Int32[_fonts.Count];
        var cidFontObjIds = new Int32[_fonts.Count];
        for (var fi = 0; fi < _fonts.Count; fi++)
        {
            fontObjIds[fi] = NextId();
            cidFontObjIds[fi] = _fonts[fi].IsCjk ? NextId() : 0;
        }
        var imgObjMap = new Dictionary<String, Int32>();
        var allImages = new List<(String Name, Byte[] Data, Int32 W, Int32 H, Boolean IsJpeg)>();
        foreach (var page in allPages)
            foreach (var kv in page.Images)
                if (!imgObjMap.ContainsKey(kv.Key))
                {
                    imgObjMap[kv.Key] = NextId();
                    allImages.Add((kv.Key, kv.Value.Data, kv.Value.Width, kv.Value.Height, kv.Value.IsJpeg));
                }

        // 超链接注释对象 ID (每个注释一个对象)
        var pageAnnotObjIds = new Dictionary<Int32, List<Int32>>();
        foreach (var page in allPages)
        {
            if (page.LinkAnnotations.Count > 0)
            {
                var ids = page.LinkAnnotations.Select(_ => NextId()).ToList();
                pageAnnotObjIds[page.PageObjId] = ids;
            }
        }

        // 书签对象 ID
        var outlineObjId = 0;
        var bookmarkObjIds = new List<Int32>();
        if (Bookmarks.Count > 0)
        {
            outlineObjId = NextId();
            bookmarkObjIds.AddRange(Bookmarks.Select(_ => NextId()));
        }

        // 文档属性 Info 对象 ID
        var infoObjId = 0;
        if (DocumentTitle != null || DocumentAuthor != null || DocumentSubject != null)
            infoObjId = NextId();

        // 加密字典对象 ID
        var encryptObjId = 0;
        if (UserPassword != null || OwnerPassword != null)
            encryptObjId = NextId();

        var totalObjs = nextId;
        while (offsets.Count < totalObjs) offsets.Add(0);

        // 创建加密器
        Byte[]? fileIdBytes = null;
        PdfEncryptor? enc = null;
        if (encryptObjId > 0)
        {
            using var encMd5 = MD5.Create();
            fileIdBytes = encMd5.ComputeHash(latin1.GetBytes(DateTime.Now.Ticks.ToString()));
            enc = new PdfEncryptor(UserPassword, OwnerPassword ?? UserPassword ?? String.Empty, Permissions, fileIdBytes);
        }

        String PdfStr(String text, Int32 objId)
        {
            if (enc == null) return $"({EncodePdfText(text)})";
            return enc.EncryptString(text, objId, 0);
        }

        // ── 写入 Catalog (obj 1) ──
        var catalogSb = new StringBuilder();
        catalogSb.Append("<< /Type /Catalog\n/Pages 2 0 R");
        if (outlineObjId > 0) catalogSb.Append($"\n/Outlines {outlineObjId} 0 R\n/PageMode /UseOutlines");
        if (encryptObjId > 0) catalogSb.Append($"\n/Encrypt {encryptObjId} 0 R");
        catalogSb.Append("\n>>");
        WriteObj(1, catalogSb.ToString());

        // ── 写入 Pages (obj 2) ──
        var kidsStr = String.Join(" ", allPages.Select(p => $"{p.PageObjId} 0 R"));
        WriteObj(2, $"<< /Type /Pages\n/Kids [{kidsStr}]\n/Count {pageCount}\n>>");

        // ── 写入加密字典 (encryptObjId) ──
        if (enc != null)
        {
            var oHex = BitConverter.ToString(enc.OEntry).Replace("-", "");
            var uHex = BitConverter.ToString(enc.UEntry).Replace("-", "");
            WriteObj(encryptObjId,
                $"<< /Filter /Standard /V 2 /R 3 /Length 128\n" +
                $"/P {enc.EncPermissions}\n" +
                $"/O <{oHex}>\n" +
                $"/U <{uHex}>\n>>");
        }

        // ── 写入字体对象 ──
        for (var fi = 0; fi < _fonts.Count; fi++)
        {
            var f = _fonts[fi];
            if (f.IsCjk)
            {
                // CIDFont（后代字体），用于描述字形宽度与 CIDSystemInfo
                WriteObj(cidFontObjIds[fi],
                    $"<< /Type /Font /Subtype /CIDFontType0\n" +
                    $"/BaseFont /{f.BaseFont}\n" +
                    $"/CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >>\n" +
                    $"/DW 1000\n>>");
                // Type0（主字体），指定 UniGB-UCS2-H 编码
                WriteObj(fontObjIds[fi],
                    $"<< /Type /Font /Subtype /Type0\n" +
                    $"/BaseFont /{f.BaseFont}\n" +
                    $"/Encoding /UniGB-UCS2-H\n" +
                    $"/DescendantFonts [{cidFontObjIds[fi]} 0 R]\n>>");
            }
            else
            {
                WriteObj(fontObjIds[fi], $"<< /Type /Font\n/Subtype /Type1\n/BaseFont /{f.BaseFont}\n/Encoding /WinAnsiEncoding\n>>");
            }
        }

        // ── 写入图片 XObject ──
        foreach (var (name, data, imgW, imgH, isJpeg) in allImages)
        {
            var rawRgb = ExtractPngRgb(data, imgW, imgH);
            var imgObjId = imgObjMap[name];
            var imgData = enc != null ? enc.EncryptBytes(rawRgb, imgObjId, 0) : rawRgb;
            offsets[imgObjId - 1] = ms.Position;
            var imgHdr = latin1.GetBytes(
                $"{imgObjId} 0 obj\n" +
                $"<< /Type /XObject /Subtype /Image\n/Width {imgW} /Height {imgH}\n" +
                $"/ColorSpace /DeviceRGB\n/BitsPerComponent 8\n/Length {imgData.Length}\n>>\nstream\n");
            ms.Write(imgHdr, 0, imgHdr.Length);
            ms.Write(imgData, 0, imgData.Length);
            ms.Write(latin1.GetBytes("\nendstream\nendobj\n"), 0, "\nendstream\nendobj\n".Length);
        }

        // ── 写入超链接注释对象 ──
        foreach (var page in allPages)
        {
            if (!pageAnnotObjIds.TryGetValue(page.PageObjId, out var annotIds)) continue;
            for (var ai = 0; ai < page.LinkAnnotations.Count; ai++)
            {
                var (ax, ay, aw, ah, url) = page.LinkAnnotations[ai];
                var rect = $"[{ax:F2} {ay:F2} {(ax + aw):F2} {(ay + ah):F2}]";
                WriteObj(annotIds[ai],
                    $"<< /Type /Annot /Subtype /Link\n/Rect {rect}\n/Border [0 0 0]\n" +
                    $"/A << /Type /Action /S /URI /URI {PdfStr(url, annotIds[ai])} >>\n>>");
            }
        }

        // ── 写入书签（Outlines）对象 ──
        if (outlineObjId > 0)
        {
            var firstBmId = bookmarkObjIds[0];
            var lastBmId = bookmarkObjIds[^1];
            WriteObj(outlineObjId,
                $"<< /Type /Outlines /First {firstBmId} 0 R /Last {lastBmId} 0 R /Count {Bookmarks.Count} >>");

            for (var bi = 0; bi < Bookmarks.Count; bi++)
            {
                var bm = Bookmarks[bi];
                var pageRef = (bm.PageIndex < allPages.Count) ? allPages[bm.PageIndex].PageObjId : allPages[0].PageObjId;
                var pageSz = allPages[Math.Min(bm.PageIndex, allPages.Count - 1)];
                var bmSb = new StringBuilder();
                bmSb.Append($"<< /Title {PdfStr(bm.Title, bookmarkObjIds[bi])}\n");
                bmSb.Append($"/Parent {outlineObjId} 0 R\n");
                bmSb.Append($"/Dest [{pageRef} 0 R /XYZ 0 {pageSz.Height} 0]\n");
                if (bi > 0) bmSb.Append($"/Prev {bookmarkObjIds[bi - 1]} 0 R\n");
                if (bi < Bookmarks.Count - 1) bmSb.Append($"/Next {bookmarkObjIds[bi + 1]} 0 R\n");
                bmSb.Append(">>");
                WriteObj(bookmarkObjIds[bi], bmSb.ToString());
            }
        }

        // ── 写入 Info 字典 ──
        if (infoObjId > 0)
        {
            var infoSb = new StringBuilder("<< ");
            if (DocumentTitle != null) infoSb.Append($"/Title {PdfStr(DocumentTitle, infoObjId)} ");
            if (DocumentAuthor != null) infoSb.Append($"/Author {PdfStr(DocumentAuthor, infoObjId)} ");
            if (DocumentSubject != null) infoSb.Append($"/Subject {PdfStr(DocumentSubject, infoObjId)} ");
            infoSb.Append(">>");
            WriteObj(infoObjId, infoSb.ToString());
        }

        // ── 写入页面和内容流 ──
        var needHdrFtr = HeaderText != null || FooterText != null || ShowPageNumbers;
        for (var pi = 0; pi < allPages.Count; pi++)
        {
            var page = allPages[pi];
            var fontRefs = String.Join("\n", _fonts.Select((f, fi) => $"/{f.Name} {fontObjIds[fi]} 0 R"));
            var imgRefs = page.Images.Count > 0
                ? String.Join("\n", page.Images.Keys.Select(n => $"/{n} {imgObjMap[n]} 0 R"))
                : String.Empty;

            var resSb = new StringBuilder("<< /Font << ");
            resSb.Append(fontRefs);
            resSb.Append(" >>");
            if (imgRefs.Length > 0) { resSb.Append("\n/XObject << "); resSb.Append(imgRefs); resSb.Append(" >>"); }
            resSb.Append(" >>");

            // 超链接注释引用
            var annotStr = String.Empty;
            if (pageAnnotObjIds.TryGetValue(page.PageObjId, out var annotIds2))
                annotStr = $"\n/Annots [{String.Join(" ", annotIds2.Select(id => $"{id} 0 R"))}]";

            // 旋转
            var rotateStr = page.Rotation != 0 ? $"\n/Rotate {page.Rotation}" : String.Empty;

            WriteObj(page.PageObjId,
                $"<< /Type /Page\n/Parent 2 0 R\n" +
                $"/MediaBox [0 0 {page.Width:F0} {page.Height:F0}]\n" +
                $"/Resources {resSb}\n" +
                $"/Contents {page.ContentObjId} 0 R{rotateStr}{annotStr}\n>>");

            // 内容流 = 原始内容 + 页眉/页脚
            Byte[] finalContent;
            if (needHdrFtr)
            {
                var hfSb = new StringBuilder();
                var f1Name = _fonts[0].Name;
                // 页眉
                if (HeaderText != null)
                {
                    var hdrY = page.Height - 18f;
                    hfSb.Append($"BT /{f1Name} 9 Tf\n{MarginLeft} {hdrY:F2} Td\n({EncodePdfText(HeaderText)}) Tj\nET\n");
                    // 分隔线
                    hfSb.Append($"{MarginLeft} {hdrY - 3:F2} m {page.Width - MarginRight} {hdrY - 3:F2} l S\n");
                }
                // 页脚
                var ftrY = MarginBottom - 14f;
                if (ftrY < 4f) ftrY = 4f;
                if (FooterText != null)
                    hfSb.Append($"BT /{f1Name} 9 Tf\n{MarginLeft} {ftrY:F2} Td\n({EncodePdfText(FooterText)}) Tj\nET\n");
                if (ShowPageNumbers)
                {
                    var pageNumText = $"- {pi + 1} -";
                    var pgX = (page.Width - pageNumText.Length * 4f) / 2f;
                    hfSb.Append($"BT /{f1Name} 9 Tf\n{pgX:F2} {ftrY:F2} Td\n({pageNumText}) Tj\nET\n");
                }
                var hfBytes = latin1.GetBytes(hfSb.ToString());
                finalContent = new Byte[page.ContentBytes.Length + hfBytes.Length];
                page.ContentBytes.CopyTo(finalContent, 0);
                hfBytes.CopyTo(finalContent, page.ContentBytes.Length);
            }
            else
            {
                finalContent = page.ContentBytes;
            }

            var encContent = enc != null ? enc.EncryptBytes(finalContent, page.ContentObjId, 0) : finalContent;
            offsets[page.ContentObjId - 1] = ms.Position;
            var contentHdr = latin1.GetBytes($"{page.ContentObjId} 0 obj\n<< /Length {encContent.Length} >>\nstream\n");
            ms.Write(contentHdr, 0, contentHdr.Length);
            ms.Write(encContent, 0, encContent.Length);
            ms.Write(latin1.GetBytes("\nendstream\nendobj\n"), 0, "\nendstream\nendobj\n".Length);
        }

        // ── xref 表 ──
        var xrefPos = ms.Position;
        var xrefSb = new StringBuilder();
        xrefSb.AppendLine("xref");
        xrefSb.AppendLine($"0 {totalObjs + 1}");
        xrefSb.AppendLine("0000000000 65535 f ");
        foreach (var off in offsets) xrefSb.AppendLine($"{off:D10} 00000 n ");
        ms.Write(latin1.GetBytes(xrefSb.ToString()), 0, xrefSb.Length);

        // ── trailer ──
        var trailerStr = new StringBuilder("trailer\n<< /Size ");
        trailerStr.Append($"{totalObjs + 1}\n/Root 1 0 R");
        if (infoObjId > 0) trailerStr.Append($"\n/Info {infoObjId} 0 R");
        if (encryptObjId > 0) trailerStr.Append($"\n/Encrypt {encryptObjId} 0 R");
        if (fileIdBytes != null)
        {
            var idHex = BitConverter.ToString(fileIdBytes).Replace("-", "");
            trailerStr.Append($"\n/ID [<{idHex}><{idHex}>]");
        }
        trailerStr.Append($" >>\nstartxref\n{xrefPos}\n%%EOF\n");
        ms.Write(latin1.GetBytes(trailerStr.ToString()), 0, trailerStr.Length);

        ms.Position = 0;
        ms.CopyTo(stream);
    }
    #endregion

    #region 辅助方法
    private void EnsurePage()
    {
        if (CurrentPage == null) BeginPage();
    }

    private static String EncodePdfText(String text)
    {
        // 只保留 Latin-1 可打印字符，其他替换为 ?
        var sb = new StringBuilder(text.Length * 2);
        foreach (var ch in text)
        {
            if (ch == '(' || ch == ')' || ch == '\\')
                sb.Append('\\');
            if (ch < 256 && ch >= 32)
                sb.Append(ch);
            else if (ch >= 32)
                sb.Append('?'); // CJK fallback（请改用 CreateSimplifiedChineseFont 绘制中文）
        }
        return sb.ToString();
    }

    private static String EncodeCjkHex(String text)
    {
        // UTF-16BE 大端编码后转十六进制，作为 PDF 字符串 <hex> 运算符
        var bytes = Encoding.BigEndianUnicode.GetBytes(text);
        var sb = new StringBuilder(bytes.Length * 2);
        foreach (var b in bytes)
            sb.Append(b.ToString("X2"));
        return sb.ToString();
    }

    private static String HexToRgbOp(String hex, Boolean fill)
    {
        hex = hex.TrimStart('#');
        if (hex.Length < 6) hex = "000000";
        var r = Convert.ToInt32(hex.Substring(0, 2), 16) / 255f;
        var g = Convert.ToInt32(hex.Substring(2, 2), 16) / 255f;
        var b = Convert.ToInt32(hex.Substring(4, 2), 16) / 255f;
        return fill
            ? $"{r:F3} {g:F3} {b:F3} rg"
            : $"{r:F3} {g:F3} {b:F3} RG";
    }

    /// <summary>从 PNG 数据读取宽高（从 IHDR chunk）</summary>
    private static (Int32 Width, Int32 Height) GetPngSize(Byte[] png)
    {
        // PNG Signature: 8 bytes
        // IHDR chunk: 4(length) + 4(type) + 4(width) + 4(height) = starts at offset 8
        if (png.Length < 24) return (1, 1);
        var w = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
        var h = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
        return (w > 0 ? w : 1, h > 0 ? h : 1);
    }

    /// <summary>从 PNG 提取原始 RGB 字节（简化：跳过压缩，直接返回后 IDAT 内容占位）</summary>
    /// <remarks>
    /// 完整实现需要解码 zlib 压缩+过滤器。此处返回白色占位矩形（不影响 PDF 结构正确性）。
    /// 实际项目中可替换为 System.Drawing 或 ImageSharp 解码。
    /// </remarks>
    private static Byte[] ExtractPngRgb(Byte[] png, Int32 w, Int32 h)
    {
        // 尝试用简单方案：如果系统有 System.Drawing，用它；否则返回白色占位
        try
        {
#if NET6_0_OR_GREATER
            using var ms = new System.IO.MemoryStream(png);
            using var bmp = System.Drawing.Image.FromStream(ms);
            return ExtractBitmapRgb(bmp, w, h);
#else
            return CreateWhiteRgb(w, h);
#endif
        }
        catch
        {
            return CreateWhiteRgb(w, h);
        }
    }

#if NET6_0_OR_GREATER
    private static Byte[] ExtractBitmapRgb(System.Drawing.Image img, Int32 w, Int32 h)
    {
        using var bmp = new System.Drawing.Bitmap(img);
        var rgb = new Byte[w * h * 3];
        var idx = 0;
        for (var y = 0; y < h; y++)
        {
            for (var x = 0; x < w; x++)
            {
                var c = bmp.GetPixel(x, y);
                rgb[idx++] = c.R;
                rgb[idx++] = c.G;
                rgb[idx++] = c.B;
            }
        }
        return rgb;
    }
#endif

    private static Byte[] CreateWhiteRgb(Int32 w, Int32 h)
    {
        var rgb = new Byte[w * h * 3];
        for (var i = 0; i < rgb.Length; i++) rgb[i] = 255; // white
        return rgb;
    }
    #endregion
}