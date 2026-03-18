using System.IO.Compression;
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
    private PdfFont? _fontCjk;

    // WinAnsiEncoding 在 Latin-1 之外的 CP1252 扩展字符映射（U+0080-U+009F 区段）
    private static readonly Dictionary<Char, Char> _cp1252Map = new Dictionary<Char, Char>
    {
        ['\u20AC'] = (Char)0x80, // €
        ['\u201A'] = (Char)0x82, // ‚
        ['\u0192'] = (Char)0x83, // ƒ
        ['\u201E'] = (Char)0x84, // „
        ['\u2026'] = (Char)0x85, // …
        ['\u2020'] = (Char)0x86, // †
        ['\u2021'] = (Char)0x87, // ‡
        ['\u02C6'] = (Char)0x88, // ˆ
        ['\u2030'] = (Char)0x89, // ‰
        ['\u0160'] = (Char)0x8A, // Š
        ['\u2039'] = (Char)0x8B, // ‹
        ['\u0152'] = (Char)0x8C, // Œ
        ['\u017D'] = (Char)0x8E, // Ž
        ['\u2018'] = (Char)0x91, // '
        ['\u2019'] = (Char)0x92, // '
        ['\u201C'] = (Char)0x93, // "
        ['\u201D'] = (Char)0x94, // "
        ['\u2022'] = (Char)0x95, // •
        ['\u2013'] = (Char)0x96, // –
        ['\u2014'] = (Char)0x97, // —
        ['\u02DC'] = (Char)0x98, // ˜
        ['\u2122'] = (Char)0x99, // ™
        ['\u0161'] = (Char)0x9A, // š
        ['\u203A'] = (Char)0x9B, // ›
        ['\u0153'] = (Char)0x9C, // œ
        ['\u017E'] = (Char)0x9E, // ž
        ['\u0178'] = (Char)0x9F, // Ÿ
    };

    // 常见中文字体名
    private static readonly Dictionary<String, (String FileName, Int32 Index)> _fontFileMap =
        new Dictionary<String, (String, Int32)>(StringComparer.OrdinalIgnoreCase)
        {
            ["宋体"]               = ("simsun.ttc",  0),
            ["SimSun"]            = ("simsun.ttc",  0),
            ["新宋体"]             = ("simsun.ttc",  1),
            ["NSimSun"]           = ("simsun.ttc",  1),
            ["黑体"]               = ("simhei.ttf",  0),
            ["SimHei"]            = ("simhei.ttf",  0),
            ["楷体"]               = ("simkai.ttf",  0),
            ["KaiTi"]             = ("simkai.ttf",  0),
            ["仿宋"]               = ("simfang.ttf", 0),
            ["FangSong"]          = ("simfang.ttf", 0),
            ["微软雅黑"]           = ("msyh.ttc",    0),
            ["Microsoft YaHei"]   = ("msyh.ttc",    0),
            ["MicrosoftYaHei"]    = ("msyh.ttc",    0),
            ["微软雅黑 Light"]     = ("msyhl.ttc",   0),
            ["等线"]               = ("Deng.ttf",    0),
            ["DengXian"]          = ("Deng.ttf",    0),
            ["新細明體"]           = ("mingliu.ttc", 0),
            ["MingLiU"]           = ("mingliu.ttc", 0),
        };

    // 中文字体名 → PostScript 名
    private static readonly Dictionary<String, String> _fontPsNameMap =
        new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase)
        {
            ["宋体"]            = "SimSun",
            ["新宋体"]          = "NSimSun",
            ["黑体"]            = "SimHei",
            ["楷体"]            = "KaiTi",
            ["仿宋"]            = "FangSong",
            ["微软雅黑"]        = "MicrosoftYaHei",
            ["微软雅黑 Light"]  = "MicrosoftYaHei-Light",
            ["等线"]            = "DengXian",
            ["新細明體"]        = "MingLiU",
        };

    // 标准 PDF Type1（14 个内置字体）及常用英文别名 → BaseFont 名
    private static readonly Dictionary<String, String> _latinFontMap =
        new Dictionary<String, String>(StringComparer.OrdinalIgnoreCase)
        {
            ["Helvetica"]             = "Helvetica",
            ["Helvetica-Bold"]        = "Helvetica-Bold",
            ["Helvetica-Oblique"]     = "Helvetica-Oblique",
            ["Helvetica-BoldOblique"] = "Helvetica-BoldOblique",
            ["Arial"]                 = "Helvetica",
            ["Arial Bold"]            = "Helvetica-Bold",
            ["Arial-Bold"]            = "Helvetica-Bold",
            ["Times"]                 = "Times-Roman",
            ["Times-Roman"]           = "Times-Roman",
            ["Times-Bold"]            = "Times-Bold",
            ["Times-Italic"]          = "Times-Italic",
            ["Times-BoldItalic"]      = "Times-BoldItalic",
            ["Courier"]               = "Courier",
            ["Courier-Bold"]          = "Courier-Bold",
            ["Courier-Oblique"]       = "Courier-Oblique",
            ["Courier-BoldOblique"]   = "Courier-BoldOblique",
            ["Symbol"]                = "Symbol",
            ["ZapfDingbats"]          = "ZapfDingbats",
        };
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
        font ??= ContainsCjk(text) ? EnsureCjkFont() : _fontHelvetica;
        // 若调用方显式传入 Latin 字体但文本含 CJK 字符，自动切换 CJK 字体，避免输出问号
        if (!font.IsCjk && ContainsCjk(text))
            font = EnsureCjkFont();
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

    /// <summary>根据字体名称创建字体，支持 PDF 标准 Type1 英文字体和系统 TrueType 中文字体</summary>
    /// <param name="fontName">
    /// 字体名称。标准英文字体（如 "Helvetica-Bold"、"Times-Roman"、"Courier-Bold"、"Arial" 等）无需嵌入；
    /// 中文字体（如 "微软雅黑"、"宋体"、"SimHei" 等）默认嵌入字体文件，可通过 embed=false 禁止嵌入。
    /// </param>
    /// <param name="embed">是否嵌入字体文件（仅对找到字体文件的 TrueType 字体有效，默认 true）</param>
    /// <returns>已注册的字体；中文字体未找到字体文件时回退到 Adobe STSong-Light</returns>
    public PdfFont CreateFont(String fontName, Boolean embed = true)
    {
        var fname = $"F{_fonts.Count + 1}";
        // 优先匹配标准 Type1 英文字体（不需要嵌入字体文件）
        if (_latinFontMap.TryGetValue(fontName, out var type1Base))
        {
            var font = new PdfFont(fname, type1Base, isCjk: false);
            _fonts.Add(font);
            return font;
        }
        // CJK TrueType 字体：搜索系统字体文件
        if (TryFindFontFile(fontName, out var filePath, out var ttcIndex))
        {
            if (!_fontPsNameMap.TryGetValue(fontName, out var psName))
                psName = fontName.Replace(" ", "-");
            var font = new PdfFont(fname, psName, isCjk: true);
            font.FontFilePath = filePath;
            font.TtcFontIndex = ttcIndex;
            font.EmbedFont = embed;
            _fonts.Add(font);
            return font;
        }
        // 未找到字体文件，回退到 Adobe STSong-Light
        return CreateSimplifiedChineseFont();
    }

    /// <summary>追加文本行（自动换行，跟踪当前 Y 位置，Y 从顶部开始）</summary>
    /// <param name="text">文本内容</param>
    /// <param name="fontSize">字号（磅）</param>
    /// <param name="font">字体</param>
    /// <param name="indentX">与左边距的额外水平偏移</param>
    public void AppendLine(String text, Single fontSize = 12, PdfFont? font = null, Single indentX = 0)
    {
        EnsurePage();
        font ??= ContainsCjk(text) ? EnsureCjkFont() : _fontHelvetica;
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
                var cellFont = ContainsCjk(cellText) ? EnsureCjkFont() : (isHeader ? _fontTimesBold : _fontHelvetica);
                DrawText(cellText, cellX + padding, textY, fontSize, cellFont);
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
        {
            rows.Add(props.Select(p => Convert.ToString(p.GetValue(item)) ?? String.Empty).ToArray());
        }
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
        var written = 0L;
        var offsets = new List<Int64>();
        var latin1 = Encoding.GetEncoding(28591);

        void WriteBytes(Byte[] bytes, Int32 offset, Int32 count)
        {
            stream.Write(bytes, offset, count);
            written += count;
        }

        void WriteObj(Int32 id, String content)
        {
            while (offsets.Count < id) offsets.Add(0);
            offsets[id - 1] = written;
            var bytes = latin1.GetBytes($"{id} 0 obj\n{content}\nendobj\n");
            WriteBytes(bytes, 0, bytes.Length);
        }

        // Header
        var header = latin1.GetBytes("%PDF-1.4\n%\xFF\xFF\xFF\xFF\n");
        WriteBytes(header, 0, header.Length);

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
        var fontObjIds      = new Int32[_fonts.Count];
        var cidFontObjIds   = new Int32[_fonts.Count];
        var fontDescObjIds  = new Int32[_fonts.Count];
        var fontFile2ObjIds = new Int32[_fonts.Count];
        var cidToGidObjIds  = new Int32[_fonts.Count];
        var toUnicodeObjIds = new Int32[_fonts.Count];
        for (var fi = 0; fi < _fonts.Count; fi++)
        {
            fontObjIds[fi]    = NextId();
            cidFontObjIds[fi] = _fonts[fi].IsCjk ? NextId() : 0;
            // 有字体文件的 CJK 字体均需 CIDToGIDMap 和 ToUnicode；嵌入时还需 FontFile2 和 FontDescriptor
            if (_fonts[fi].IsCjk && _fonts[fi].FontFilePath != null)
            {
                if (_fonts[fi].EmbedFont)
                {
                    fontFile2ObjIds[fi] = NextId();
                    fontDescObjIds[fi]  = NextId();
                }
                cidToGidObjIds[fi]  = NextId();
                toUnicodeObjIds[fi] = NextId();
            }
        }
        var imgObjMap = new Dictionary<String, Int32>();
        var allImages = new List<(String Name, Byte[] Data, Int32 W, Int32 H, Boolean IsJpeg)>();
        foreach (var page in allPages)
        {
            foreach (var kv in page.Images)
            {
                if (!imgObjMap.ContainsKey(kv.Key))
                {
                    imgObjMap[kv.Key] = NextId();
                    allImages.Add((kv.Key, kv.Value.Data, kv.Value.Width, kv.Value.Height, kv.Value.IsJpeg));
                }
            }
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
        var streamEndBytes = latin1.GetBytes("\nendstream\nendobj\n");
        for (var fi = 0; fi < _fonts.Count; fi++)
        {
            var f = _fonts[fi];
            if (f.IsCjk)
            {
                if (f.FontFilePath != null && File.Exists(f.FontFilePath) && f.EmbedFont)
                {
                    // ── 嵌入 TrueType/TTC 字体 ──
                    var fontData = File.ReadAllBytes(f.FontFilePath);
                    var sfOff = GetSfOffset(fontData, f.TtcFontIndex);
                    var (upm, ascent, descent, xMin, yMin, xMax, yMax) = ReadTtfMetrics(fontData, sfOff);
                    var scale  = upm > 0 ? 1000.0 / upm : 1.0;
                    var a1000  = (Int32)(ascent  * scale);
                    var d1000  = (Int32)(descent * scale);
                    var bb     = $"[{(Int32)(xMin*scale)} {(Int32)(yMin*scale)} {(Int32)(xMax*scale)} {(Int32)(yMax*scale)}]";

                    // FontFile2 流（原始字体字节）
                    offsets[fontFile2ObjIds[fi] - 1] = written;
                    var ff2h = latin1.GetBytes($"{fontFile2ObjIds[fi]} 0 obj\n<< /Length {fontData.Length} /Length1 {fontData.Length} >>\nstream\n");
                    WriteBytes(ff2h, 0, ff2h.Length);
                    WriteBytes(fontData, 0, fontData.Length);
                    WriteBytes(streamEndBytes, 0, streamEndBytes.Length);

                    // FontDescriptor
                    WriteObj(fontDescObjIds[fi],
                        $"<< /Type /FontDescriptor\n/FontName /{f.BaseFont}\n/Flags 32\n" +
                        $"/FontBBox {bb}\n/ItalicAngle 0\n/Ascent {a1000}\n/Descent {d1000}\n" +
                        $"/CapHeight {a1000}\n/StemV 80\n/FontFile2 {fontFile2ObjIds[fi]} 0 R\n>>");

                    // CIDToGIDMap 流（Unicode → GlyphID，压缩以减小体积）
                    var glyphMap = ParseTtfCmap(fontData, sfOff);
                    var ctgData  = ZlibCompress(BuildCidToGidMap(glyphMap));
                    offsets[cidToGidObjIds[fi] - 1] = written;
                    var ctgh = latin1.GetBytes($"{cidToGidObjIds[fi]} 0 obj\n<< /Length {ctgData.Length} /Filter /FlateDecode >>\nstream\n");
                    WriteBytes(ctgh, 0, ctgh.Length);
                    WriteBytes(ctgData, 0, ctgData.Length);
                    WriteBytes(streamEndBytes, 0, streamEndBytes.Length);

                    // ToUnicode 流（Identity 映射，供文本提取）
                    var tuData = BuildIdentityToUnicodeStream();
                    offsets[toUnicodeObjIds[fi] - 1] = written;
                    var tuh = latin1.GetBytes($"{toUnicodeObjIds[fi]} 0 obj\n<< /Length {tuData.Length} >>\nstream\n");
                    WriteBytes(tuh, 0, tuh.Length);
                    WriteBytes(tuData, 0, tuData.Length);
                    WriteBytes(streamEndBytes, 0, streamEndBytes.Length);

                    // CIDFont（CIDFontType2，通过 CIDToGIDMap 流映射 Unicode → GlyphID）
                    WriteObj(cidFontObjIds[fi],
                        $"<< /Type /Font /Subtype /CIDFontType2\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>\n" +
                        $"/DW 1000\n/CIDToGIDMap {cidToGidObjIds[fi]} 0 R\n" +
                        $"/FontDescriptor {fontDescObjIds[fi]} 0 R\n>>");

                    // Type0（主字体，Identity-H 编码）
                    WriteObj(fontObjIds[fi],
                        $"<< /Type /Font /Subtype /Type0\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/Encoding /Identity-H\n" +
                        $"/DescendantFonts [{cidFontObjIds[fi]} 0 R]\n" +
                        $"/ToUnicode {toUnicodeObjIds[fi]} 0 R\n>>");
                }
                else if (f.FontFilePath != null && !f.EmbedFont)
                {
                    // ── TrueType 字体引用（不嵌入字体数据，但写入 CIDToGIDMap 确保正确映射）──
                    var fontData = File.ReadAllBytes(f.FontFilePath);
                    var sfOff = GetSfOffset(fontData, f.TtcFontIndex);

                    // CIDToGIDMap 流（从 cmap 解析 Unicode → GlyphID，压缩以减小体积）
                    var glyphMap = ParseTtfCmap(fontData, sfOff);
                    var ctgData  = ZlibCompress(BuildCidToGidMap(glyphMap));
                    offsets[cidToGidObjIds[fi] - 1] = written;
                    var ctgh = latin1.GetBytes($"{cidToGidObjIds[fi]} 0 obj\n<< /Length {ctgData.Length} /Filter /FlateDecode >>\nstream\n");
                    WriteBytes(ctgh, 0, ctgh.Length);
                    WriteBytes(ctgData, 0, ctgData.Length);
                    WriteBytes(streamEndBytes, 0, streamEndBytes.Length);

                    // ToUnicode 流（供文本提取）
                    var tuData = BuildIdentityToUnicodeStream();
                    offsets[toUnicodeObjIds[fi] - 1] = written;
                    var tuh = latin1.GetBytes($"{toUnicodeObjIds[fi]} 0 obj\n<< /Length {tuData.Length} >>\nstream\n");
                    WriteBytes(tuh, 0, tuh.Length);
                    WriteBytes(tuData, 0, tuData.Length);
                    WriteBytes(streamEndBytes, 0, streamEndBytes.Length);

                    WriteObj(cidFontObjIds[fi],
                        $"<< /Type /Font /Subtype /CIDFontType2\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>\n" +
                        $"/DW 1000\n/CIDToGIDMap {cidToGidObjIds[fi]} 0 R\n>>");
                    WriteObj(fontObjIds[fi],
                        $"<< /Type /Font /Subtype /Type0\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/Encoding /Identity-H\n" +
                        $"/DescendantFonts [{cidFontObjIds[fi]} 0 R]\n" +
                        $"/ToUnicode {toUnicodeObjIds[fi]} 0 R\n>>");
                }
                else
                {
                    // ── Adobe 预定义 CJK 字体（STSong-Light 等，无需嵌入）──
                    WriteObj(cidFontObjIds[fi],
                        $"<< /Type /Font /Subtype /CIDFontType0\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >>\n" +
                        $"/DW 1000\n>>");
                    WriteObj(fontObjIds[fi],
                        $"<< /Type /Font /Subtype /Type0\n" +
                        $"/BaseFont /{f.BaseFont}\n" +
                        $"/Encoding /UniGB-UCS2-H\n" +
                        $"/DescendantFonts [{cidFontObjIds[fi]} 0 R]\n>>");
                }
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
            offsets[imgObjId - 1] = written;
            var imgHdr = latin1.GetBytes(
                $"{imgObjId} 0 obj\n" +
                $"<< /Type /XObject /Subtype /Image\n/Width {imgW} /Height {imgH}\n" +
                $"/ColorSpace /DeviceRGB\n/BitsPerComponent 8\n/Length {imgData.Length}\n>>\nstream\n");
            WriteBytes(imgHdr, 0, imgHdr.Length);
            WriteBytes(imgData, 0, imgData.Length);
            var imgEnd = latin1.GetBytes("\nendstream\nendobj\n");
            WriteBytes(imgEnd, 0, imgEnd.Length);
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
            offsets[page.ContentObjId - 1] = written;
            var contentHdr = latin1.GetBytes($"{page.ContentObjId} 0 obj\n<< /Length {encContent.Length} >>\nstream\n");
            WriteBytes(contentHdr, 0, contentHdr.Length);
            WriteBytes(encContent, 0, encContent.Length);
            var contentEnd = latin1.GetBytes("\nendstream\nendobj\n");
            WriteBytes(contentEnd, 0, contentEnd.Length);
        }

        // ── xref 表 ──
        var xrefPos = written;
        var xrefSb = new StringBuilder();
        xrefSb.AppendLine("xref");
        xrefSb.AppendLine($"0 {totalObjs + 1}");
        xrefSb.AppendLine("0000000000 65535 f ");
        foreach (var off in offsets) xrefSb.AppendLine($"{off:D10} 00000 n ");
        var xrefBytes = latin1.GetBytes(xrefSb.ToString());
        WriteBytes(xrefBytes, 0, xrefBytes.Length);

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
        var trailerBytes = latin1.GetBytes(trailerStr.ToString());
        WriteBytes(trailerBytes, 0, trailerBytes.Length);
    }
    #endregion

    #region 辅助方法
    private void EnsurePage()
    {
        if (CurrentPage == null) BeginPage();
    }

    /// <summary>惰性获取简体中文字体，首次调用时自动注册</summary>
    private PdfFont EnsureCjkFont()
    {
        if (_fontCjk == null)
            _fontCjk = CreateSimplifiedChineseFont();
        return _fontCjk;
    }

    /// <summary>判断文本中是否含有 CJK 字符（中日韩统一表意文字及相关区块）</summary>
    private static Boolean ContainsCjk(String text)
    {
        foreach (var ch in text)
        {
            // CJK 符号与标点、假名、CJK 统一表意文字主区（U+3000–U+9FFF）
            if (ch >= '\u3000' && ch <= '\u9FFF') return true;
            // CJK 兼容表意文字（U+F900–U+FAFF）
            if (ch >= '\uF900' && ch <= '\uFAFF') return true;
            // 全角/半角字符（U+FF00–U+FFEF）
            if (ch >= '\uFF00' && ch <= '\uFFEF') return true;
        }
        return false;
    }

    #region TTF/TTC 字体解析
    /// <summary>从字体文件数据中获取指定 TTC 索引对应的 sfOffset（单个 TTF 则为 0）</summary>
    private static Int32 GetSfOffset(Byte[] data, Int32 ttcIndex)
    {
        // TTC magic: 'ttcf' = 0x74 0x74 0x63 0x66
        if (data.Length > 4 && data[0] == 0x74 && data[1] == 0x74 && data[2] == 0x63 && data[3] == 0x66)
        {
            var numFonts = ReadU32BeAsInt(data, 8);
            if (ttcIndex >= numFonts) ttcIndex = 0;
            return ReadU32BeAsInt(data, 12 + ttcIndex * 4);
        }
        return 0;
    }

    /// <summary>在 TTF 偏移表中查找指定 4 字节 tag 的表偏移量（-1=未找到）</summary>
    private static Int32 FindTtfTable(Byte[] data, Int32 sfOffset, String tag)
    {
        if (sfOffset + 12 > data.Length) return -1;
        var numTables = ReadU16Be(data, sfOffset + 4);
        var tableDir = sfOffset + 12;
        for (var i = 0; i < numTables; i++)
        {
            var pos = tableDir + i * 16;
            if (pos + 16 > data.Length) break;
            if (data[pos] == (Byte)tag[0] && data[pos + 1] == (Byte)tag[1]
                && data[pos + 2] == (Byte)tag[2] && data[pos + 3] == (Byte)tag[3])
                return ReadU32BeAsInt(data, pos + 8);
        }
        return -1;
    }

    /// <summary>从 TTF 字体数据中读取关键字体度量（单位：design units）</summary>
    private static (Int32 UnitsPerEm, Int32 Ascent, Int32 Descent, Int32 XMin, Int32 YMin, Int32 XMax, Int32 YMax) ReadTtfMetrics(Byte[] data, Int32 sfOffset)
    {
        var headOff = FindTtfTable(data, sfOffset, "head");
        var hheaOff = FindTtfTable(data, sfOffset, "hhea");
        var os2Off  = FindTtfTable(data, sfOffset, "OS/2");

        var upm  = (headOff >= 0 && headOff + 20 <= data.Length) ? ReadU16Be(data, headOff + 18) : 1000;
        var xMin = (headOff >= 0 && headOff + 44 <= data.Length) ? ReadS16Be(data, headOff + 36) : 0;
        var yMin = (headOff >= 0 && headOff + 44 <= data.Length) ? ReadS16Be(data, headOff + 38) : -200;
        var xMax = (headOff >= 0 && headOff + 44 <= data.Length) ? ReadS16Be(data, headOff + 40) : 1000;
        var yMax = (headOff >= 0 && headOff + 44 <= data.Length) ? ReadS16Be(data, headOff + 42) : 800;

        Int32 ascent, descent;
        if (os2Off >= 0 && os2Off + 72 <= data.Length)
        {
            // OS/2 表 sTypoAscender/sTypoDescender（offset 68/70）
            ascent  = ReadS16Be(data, os2Off + 68);
            descent = ReadS16Be(data, os2Off + 70);
        }
        else if (hheaOff >= 0 && hheaOff + 8 <= data.Length)
        {
            ascent  = ReadS16Be(data, hheaOff + 4);
            descent = ReadS16Be(data, hheaOff + 6);
        }
        else
        {
            ascent = yMax; descent = yMin;
        }
        return (upm, ascent, descent, xMin, yMin, xMax, yMax);
    }

    /// <summary>解析 TTF cmap 表，返回 Unicode 码点 → GlyphID 的映射</summary>
    private static Dictionary<UInt16, UInt16> ParseTtfCmap(Byte[] data, Int32 sfOffset)
    {
        var cmapOff = FindTtfTable(data, sfOffset, "cmap");
        if (cmapOff < 0 || cmapOff + 4 > data.Length)
            return new Dictionary<UInt16, UInt16>();

        var numSubtables = ReadU16Be(data, cmapOff + 2);
        var format4Offset = -1;
        for (var i = 0; i < numSubtables; i++)
        {
            var entryBase = cmapOff + 4 + i * 8;
            if (entryBase + 8 > data.Length) break;
            var platformId  = ReadU16Be(data, entryBase);
            var encodingId  = ReadU16Be(data, entryBase + 2);
            var subtableOff = cmapOff + ReadU32BeAsInt(data, entryBase + 4);
            if (subtableOff + 2 > data.Length) continue;
            if (ReadU16Be(data, subtableOff) != 4) continue;
            // 优先 Windows Unicode BMP (3,1)，其次 Unicode (0,3/4)
            if ((platformId == 3 && encodingId == 1) || (platformId == 0 && (encodingId == 3 || encodingId == 4)))
            {
                format4Offset = subtableOff;
                if (platformId == 3) break;
            }
        }
        return format4Offset >= 0 ? ParseCmapFormat4(data, format4Offset) : new Dictionary<UInt16, UInt16>();
    }

    /// <summary>解析 cmap Format 4 子表，返回 Unicode → GlyphID</summary>
    private static Dictionary<UInt16, UInt16> ParseCmapFormat4(Byte[] data, Int32 offset)
    {
        if (offset + 14 > data.Length) return new Dictionary<UInt16, UInt16>();
        var segCount       = ReadU16Be(data, offset + 6) / 2;
        var endCodeBase    = offset + 14;
        var startCodeBase  = endCodeBase + segCount * 2 + 2; // +2 = reservedPad
        var idDeltaBase    = startCodeBase + segCount * 2;
        var idRangeOffBase = idDeltaBase + segCount * 2;

        var map = new Dictionary<UInt16, UInt16>(8192);
        for (var i = 0; i < segCount; i++)
        {
            var endCode    = ReadU16Be(data, endCodeBase + i * 2);
            var startCode  = ReadU16Be(data, startCodeBase + i * 2);
            var idDelta    = ReadS16Be(data, idDeltaBase + i * 2);
            var idRangeOff = ReadU16Be(data, idRangeOffBase + i * 2);
            if (startCode == 0xFFFF) break;
            for (var code = (UInt32)startCode; code <= endCode; code++)
            {
                UInt16 glyphId;
                if (idRangeOff == 0)
                {
                    glyphId = (UInt16)((code + (UInt32)(UInt16)idDelta) & 0xFFFF);
                }
                else
                {
                    var idxPos = idRangeOffBase + i * 2 + idRangeOff + (code - startCode) * 2;
                    if (idxPos + 1 >= data.Length) continue;
                    glyphId = ReadU16Be(data, (Int32)idxPos);
                    if (glyphId != 0)
                        glyphId = (UInt16)((glyphId + (UInt32)(UInt16)idDelta) & 0xFFFF);
                }
                if (glyphId != 0) map[(UInt16)code] = glyphId;
            }
        }
        return map;
    }

    /// <summary>构建 CIDToGIDMap 二进制流（65536 × UInt16，索引=Unicode，值=GlyphID）</summary>
    private static Byte[] BuildCidToGidMap(Dictionary<UInt16, UInt16> glyphMap)
    {
        var data = new Byte[131072]; // 65536 * 2
        foreach (var kv in glyphMap)
        {
            var idx = kv.Key * 2;
            data[idx]     = (Byte)(kv.Value >> 8);
            data[idx + 1] = (Byte)(kv.Value & 0xFF);
        }
        return data;
    }

    /// <summary>构建 Identity ToUnicode CMap 流（Unicode 码点映射到自身，用于文本提取）</summary>
    private static Byte[] BuildIdentityToUnicodeStream()
    {
        const String cmap =
            "/CIDInit /ProcSet findresource begin\n" +
            "12 dict begin\nbegincmap\n" +
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS2) /Supplement 0 >> def\n" +
            "/CMapName /Adobe-Identity-UCS2 def\n" +
            "/CMapType 2 def\n" +
            "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n" +
            "1 beginbfrange\n<0000> <FFFF> <0000>\nendbfrange\n" +
            "endcmap\nCMapType currentdict end\nend\n";
        return Encoding.ASCII.GetBytes(cmap);
    }

    /// <summary>搜索系统字体目录，查找指定字体名称对应的文件路径和 TTC 索引</summary>
    private static Boolean TryFindFontFile(String fontName, out String? filePath, out Int32 ttcIndex)
    {
        filePath = null;
        ttcIndex = 0;
        if (_fontFileMap.TryGetValue(fontName, out var entry))
        {
            ttcIndex = entry.Index;
            foreach (var dir in GetFontDirectories())
            {
                var path = Path.Combine(dir, entry.FileName);
                if (File.Exists(path)) { filePath = path; return true; }
            }
        }
        // 直接按扩展名搜索
        foreach (var dir in GetFontDirectories())
        {
            if (!Directory.Exists(dir)) continue;
            foreach (var ext in new[] { ".ttf", ".otf", ".ttc" })
            {
                var path = Path.Combine(dir, fontName + ext);
                if (File.Exists(path)) { filePath = path; return true; }
            }
        }
        return false;
    }

    /// <summary>将数据压缩为 zlib 格式（2字节头 + deflate + 4字节 Adler-32），用于 PDF FlateDecode</summary>
    private static Byte[] ZlibCompress(Byte[] data)
    {
        using var output = new MemoryStream();
        // zlib header: CMF=0x78 (deflate, window=32KB), FLG=0x9C (default level, check bits)
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true))
        {
            deflate.Write(data, 0, data.Length);
        }
        // Adler-32 checksum (big-endian)
        UInt32 a = 1, b = 0;
        for (var i = 0; i < data.Length; i++)
        {
            a = (a + data[i]) % 65521;
            b = (b + a) % 65521;
        }
        var adler = (b << 16) | a;
        output.WriteByte((Byte)(adler >> 24));
        output.WriteByte((Byte)(adler >> 16));
        output.WriteByte((Byte)(adler >> 8));
        output.WriteByte((Byte)adler);
        return output.ToArray();
    }

    /// <summary>返回常见系统字体目录</summary>
    private static IEnumerable<String> GetFontDirectories()
    {
        var dirs = new List<String>();
        // Windows — SpecialFolder.Fonts 可能抛异常，捕获后跳过
        try
        {
            var winFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            if (!String.IsNullOrEmpty(winFonts) && !dirs.Contains(winFonts))
                dirs.Add(winFonts);
        }
        catch { }
        if (Directory.Exists(@"C:\Windows\Fonts") && !dirs.Contains(@"C:\Windows\Fonts"))
            dirs.Add(@"C:\Windows\Fonts");
        // Linux
        if (Directory.Exists("/usr/share/fonts")) dirs.Add("/usr/share/fonts");
        if (Directory.Exists("/usr/local/share/fonts")) dirs.Add("/usr/local/share/fonts");
        // macOS
        if (Directory.Exists("/System/Library/Fonts")) dirs.Add("/System/Library/Fonts");
        if (Directory.Exists("/Library/Fonts")) dirs.Add("/Library/Fonts");
        return dirs;
    }

    private static UInt16 ReadU16Be(Byte[] d, Int32 o) => (UInt16)((d[o] << 8) | d[o + 1]);
    private static Int16 ReadS16Be(Byte[] d, Int32 o) => (Int16)((d[o] << 8) | d[o + 1]);
    private static Int32 ReadU32BeAsInt(Byte[] d, Int32 o)
        => (d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3];
    #endregion

    private static String EncodePdfText(String text)
    {
        // Latin-1 (< 256) 直接输出；CP1252 扩展区通过映射转换；其余替换为 ?
        var sb = new StringBuilder(text.Length * 2);
        foreach (var ch in text)
        {
            if (ch == '(' || ch == ')' || ch == '\\')
                sb.Append('\\');
            if (ch < 256 && ch >= 32)
                sb.Append(ch);
            else if (_cp1252Map.TryGetValue(ch, out var mapped))
                sb.Append(mapped);
            else if (ch >= 32)
                sb.Append('?'); // 非 WinAnsiEncoding 字符（请改用 CJK 字体绘制中文）
        }
        return sb.ToString();
    }

    private static String EncodeCjkHex(String text)
    {
        // UTF-16BE 大端编码后转十六进制，作为 PDF 字符串 <hex> 运算符
        var bytes = Encoding.BigEndianUnicode.GetBytes(text);
        var sb = new StringBuilder(bytes.Length * 2);
        foreach (var b in bytes)
        {
            sb.Append(b.ToString("X2"));
        }
        return sb.ToString();
    }

    private static String HexToRgbOp(String hex, Boolean fill)
    {
        hex = hex.TrimStart('#');
        if (hex.Length < 6) hex = "000000";
        var r = Convert.ToInt32(hex[..2], 16) / 255f;
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
