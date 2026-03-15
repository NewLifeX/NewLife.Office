namespace NewLife.Office;

/// <summary>声明式 Fluent PDF 文档生成器（P07）</summary>
/// <remarks>
/// 以链式调用方式声明文档内容，自动管理页面生命周期（BeginPage/EndPage），
/// 支持文本、表格、图片、线条、矩形等基本元素，以及自动分页、书签、超链接等高级特性。
/// 内置组件复用机制：通过 UseComponent 传入委托，可在多页或多文档中复用布局片段。
/// </remarks>
/// <example>
/// <code>
/// using var doc = new PdfFluentDocument();
/// doc.Title = "报表标题";
/// doc.Header = "公司名称";
/// doc.ShowPageNumbers = true;
/// doc.AddText("一级标题", fontSize: 20)
///    .AddEmptyLine()
///    .AddText("正文段落...")
///    .AddTable(rows, firstRowHeader: true)
///    .AddImage(imageBytes, 200, 100)
///    .PageBreak()
///    .AddText("第二页");
/// doc.Save("output.pdf");
/// </code>
/// </example>
public class PdfFluentDocument : IDisposable
{
    #region 属性
    private readonly PdfWriter _writer;
    private Boolean _pageOpen;

    /// <summary>文档标题（写入 PDF Info 字典）</summary>
    public String? Title { get => _writer.DocumentTitle; set => _writer.DocumentTitle = value; }

    /// <summary>文档作者</summary>
    public String? Author { get => _writer.DocumentAuthor; set => _writer.DocumentAuthor = value; }

    /// <summary>文档主题</summary>
    public String? Subject { get => _writer.DocumentSubject; set => _writer.DocumentSubject = value; }

    /// <summary>每页顶部页眉文本</summary>
    public String? Header { get => _writer.HeaderText; set => _writer.HeaderText = value; }

    /// <summary>每页底部页脚文本</summary>
    public String? Footer { get => _writer.FooterText; set => _writer.FooterText = value; }

    /// <summary>是否在页脚显示页码</summary>
    public Boolean ShowPageNumbers { get => _writer.ShowPageNumbers; set => _writer.ShowPageNumbers = value; }

    /// <summary>页面宽度（点，默认 A4 = 595）</summary>
    public Single PageWidth { get => _writer.PageWidth; set => _writer.PageWidth = value; }

    /// <summary>页面高度（点，默认 A4 = 842）</summary>
    public Single PageHeight { get => _writer.PageHeight; set => _writer.PageHeight = value; }

    /// <summary>上边距（点）</summary>
    public Single MarginTop { get => _writer.MarginTop; set => _writer.MarginTop = value; }

    /// <summary>下边距（点）</summary>
    public Single MarginBottom { get => _writer.MarginBottom; set => _writer.MarginBottom = value; }

    /// <summary>左边距（点）</summary>
    public Single MarginLeft { get => _writer.MarginLeft; set => _writer.MarginLeft = value; }

    /// <summary>右边距（点）</summary>
    public Single MarginRight { get => _writer.MarginRight; set => _writer.MarginRight = value; }

    /// <summary>当前可用内容宽度（点）</summary>
    public Single ContentWidth => _writer.ContentWidth;

    /// <summary>当前 Y 坐标（从顶部向下）</summary>
    public Single CurrentY => _writer.CurrentY;
    #endregion

    #region 构造
    /// <summary>实例化 Fluent PDF 文档（默认 A4，自动打开第一页）</summary>
    public PdfFluentDocument()
    {
        _writer = new PdfWriter();
        EnsurePage();
    }

    /// <summary>释放资源</summary>
    public void Dispose()
    {
        if (_pageOpen) { _writer.EndPage(); _pageOpen = false; }
        _writer.Dispose();
        GC.SuppressFinalize(this);
    }
    #endregion

    #region 文本方法
    /// <summary>追加文本行（自动分页）</summary>
    /// <param name="text">文字内容</param>
    /// <param name="fontSize">字号，默认 12</param>
    /// <param name="font">字体，null 使用默认</param>
    /// <param name="indentX">水平缩进（点）</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddText(String text, Single fontSize = 12f, PdfFont? font = null, Single indentX = 0f)
    {
        EnsurePage();
        _writer.AppendLine(text, fontSize, font, indentX);
        return this;
    }

    /// <summary>追加空行</summary>
    /// <param name="height">行高（点）</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddEmptyLine(Single height = 14f)
    {
        EnsurePage();
        _writer.AppendEmptyLine(height);
        return this;
    }

    /// <summary>在指定坐标绘制文字（绝对定位，不影响当前 Y）</summary>
    /// <param name="text">文字内容</param>
    /// <param name="x">X 坐标（点）</param>
    /// <param name="y">Y 坐标（点，从左下角起算）</param>
    /// <param name="fontSize">字号</param>
    /// <param name="font">字体，null 使用默认</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument DrawText(String text, Single x, Single y, Single fontSize = 12f, PdfFont? font = null)
    {
        EnsurePage();
        _writer.DrawText(text, x, y, fontSize, font);
        return this;
    }
    #endregion

    #region 表格方法
    /// <summary>追加表格（自动分页）</summary>
    /// <param name="rows">行数据（每个 String[] 为一行）</param>
    /// <param name="firstRowHeader">首行是否为表头</param>
    /// <param name="columnWidths">列宽数组（占比），null 表示均分</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddTable(IEnumerable<String[]> rows, Boolean firstRowHeader = true, Single[]? columnWidths = null)
    {
        EnsurePage();
        _writer.DrawTable(rows, firstRowHeader, columnWidths);
        return this;
    }

    /// <summary>追加泛型对象表格（通过反射读取属性，支持 DisplayName 注解）</summary>
    /// <typeparam name="T">数据类型</typeparam>
    /// <param name="data">数据集合</param>
    /// <param name="firstRowHeader">是否输出表头行</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddTable<T>(IEnumerable<T> data, Boolean firstRowHeader = true) where T : class
    {
        EnsurePage();
        _writer.WriteObjects(data, firstRowHeader);
        return this;
    }
    #endregion

    #region 图片方法
    /// <summary>追加图片（流式排版，自动分页）</summary>
    /// <param name="imageData">图片字节（PNG/JPEG）</param>
    /// <param name="width">宽度（点）</param>
    /// <param name="height">高度（点）</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddImage(Byte[] imageData, Single width, Single height)
    {
        EnsurePage();
        _writer.AppendImage(imageData, width, height);
        return this;
    }

    /// <summary>在指定坐标绘制图片（绝对定位）</summary>
    /// <param name="imageData">图片字节</param>
    /// <param name="x">X 坐标（点）</param>
    /// <param name="y">Y 坐标（点，从左下角起算）</param>
    /// <param name="width">宽度（点）</param>
    /// <param name="height">高度（点）</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument DrawImage(Byte[] imageData, Single x, Single y, Single width, Single height)
    {
        EnsurePage();
        _writer.DrawImage(imageData, x, y, width, height);
        return this;
    }
    #endregion

    #region 绘图方法（P07-04 低层绘图 API）
    /// <summary>绘制直线（P07-04）</summary>
    /// <param name="x1">起点 X</param>
    /// <param name="y1">起点 Y</param>
    /// <param name="x2">终点 X</param>
    /// <param name="y2">终点 Y</param>
    /// <param name="lineWidth">线宽（点）</param>
    /// <param name="colorHex">颜色（16进制 RGB），null 表示黑色</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument DrawLine(Single x1, Single y1, Single x2, Single y2, Single lineWidth = 0.5f, String? colorHex = null)
    {
        EnsurePage();
        _writer.DrawLine(x1, y1, x2, y2, lineWidth, colorHex);
        return this;
    }

    /// <summary>绘制矩形（P07-04）</summary>
    /// <param name="x">左下角 X</param>
    /// <param name="y">左下角 Y</param>
    /// <param name="w">宽度</param>
    /// <param name="h">高度</param>
    /// <param name="fill">是否填充</param>
    /// <param name="fillColor">填充色（16进制 RGB）</param>
    /// <param name="borderColor">边框色（16进制 RGB）</param>
    /// <param name="borderWidth">边框线宽</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument DrawRect(Single x, Single y, Single w, Single h,
        Boolean fill = false, String? fillColor = null, String? borderColor = null, Single borderWidth = 0.5f)
    {
        EnsurePage();
        _writer.DrawRect(x, y, w, h, fill, fillColor, borderColor, borderWidth);
        return this;
    }

    /// <summary>绘制水平分隔线</summary>
    /// <param name="lineWidth">线宽</param>
    /// <param name="colorHex">颜色</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddRule(Single lineWidth = 0.5f, String? colorHex = null)
    {
        EnsurePage();
        // CurrentY 是从页顶向下；DrawLine 需要从页底向上的 PDF 坐标
        var y = _writer.PageHeight - _writer.CurrentY;
        _writer.DrawLine(_writer.MarginLeft, y, _writer.PageWidth - _writer.MarginRight, y, lineWidth, colorHex);
        _writer.AppendEmptyLine(6f);
        return this;
    }
    #endregion

    #region 导航方法
    /// <summary>插入分页符（P07-02 手动分页）</summary>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument PageBreak()
    {
        if (_pageOpen) { _writer.EndPage(); _pageOpen = false; }
        EnsurePage();
        return this;
    }

    /// <summary>添加超链接注释区域</summary>
    /// <param name="x">X 坐标（点）</param>
    /// <param name="y">Y 坐标（点，从左下角起算）</param>
    /// <param name="w">宽度</param>
    /// <param name="h">高度</param>
    /// <param name="url">链接 URL</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddHyperlink(Single x, Single y, Single w, Single h, String url)
    {
        EnsurePage();
        _writer.AddHyperlink(x, y, w, h, url);
        return this;
    }

    /// <summary>为最后一行文字追加超链接注释</summary>
    /// <param name="url">链接 URL</param>
    /// <param name="lineHeight">行高（点）</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddHyperlinkForLastLine(String url, Single lineHeight = 14f)
    {
        EnsurePage();
        _writer.AddHyperlinkForLastLine(url, lineHeight);
        return this;
    }

    /// <summary>添加书签（大纲导航）</summary>
    /// <param name="title">书签标题</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument AddBookmark(String title)
    {
        EnsurePage();
        _writer.AddBookmark(title);
        return this;
    }
    #endregion

    #region 组件复用方法（P07-03）
    /// <summary>应用可复用组件（P07-03）</summary>
    /// <remarks>通过传入委托可在多页或多文档间复用布局片段，实现组件化排版。</remarks>
    /// <param name="component">组件委托，接收当前文档实例</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument UseComponent(Action<PdfFluentDocument> component)
    {
        component(this);
        return this;
    }

    /// <summary>设置页边距（便捷方法）</summary>
    /// <param name="top">上边距</param>
    /// <param name="right">右边距</param>
    /// <param name="bottom">下边距</param>
    /// <param name="left">左边距</param>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument SetMargins(Single top, Single right, Single bottom, Single left)
    {
        _writer.MarginTop = top;
        _writer.MarginRight = right;
        _writer.MarginBottom = bottom;
        _writer.MarginLeft = left;
        return this;
    }

    /// <summary>设置页面大小为 A4 横向</summary>
    /// <returns>自身，支持链式调用</returns>
    public PdfFluentDocument SetLandscape()
    {
        _writer.PageWidth = 842f;
        _writer.PageHeight = 595f;
        return this;
    }
    #endregion

    #region 保存方法
    /// <summary>保存到文件</summary>
    /// <param name="outputPath">输出路径</param>
    public void Save(String outputPath)
    {
        ClosePage();
        _writer.Save(outputPath);
    }

    /// <summary>保存到流</summary>
    /// <param name="stream">输出流</param>
    public void Save(Stream stream)
    {
        ClosePage();
        _writer.Save(stream);
    }

    /// <summary>渲染为字节数组</summary>
    /// <returns>PDF 字节数组</returns>
    public Byte[] ToBytes()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }
    #endregion

    #region 辅助
    private void EnsurePage()
    {
        if (!_pageOpen)
        {
            _writer.BeginPage();
            _pageOpen = true;
        }
    }

    private void ClosePage()
    {
        if (_pageOpen)
        {
            _writer.EndPage();
            _pageOpen = false;
        }
    }
    #endregion
}
