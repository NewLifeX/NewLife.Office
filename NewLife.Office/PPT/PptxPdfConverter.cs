namespace NewLife.Office;

/// <summary>PowerPoint pptx 转 PDF 转换器</summary>
/// <remarks>
/// 将 .pptx 文件中的幻灯片文本内容提取并排版到 PDF 页面。
/// 每张幻灯片对应 PDF 中的一页（横向 A4，842×595 pt）。
/// 形状按其在幻灯片中的纵向位置排序后，标题以大字号、正文以小字号输出。
/// <para>注意：当前版本为文本提取转换，不支持原始图形、图片的像素级还原；
/// 如需精确像素渲染，请使用带有图形库（如 SkiaSharp）的渲染扩展，参见 TODO。</para>
/// </remarks>
public class PptxPdfConverter
{
    #region 属性

    /// <summary>PDF 标题（写入 Info 字典）</summary>
    public String? DocumentTitle { get; set; }

    /// <summary>PDF 作者</summary>
    public String? DocumentAuthor { get; set; }

    /// <summary>是否在页脚显示幻灯片编号</summary>
    public Boolean ShowSlideNumbers { get; set; } = true;

    /// <summary>标题字号（点），默认 24</summary>
    public Single TitleFontSize { get; set; } = 24f;

    /// <summary>正文字号（点），默认 14</summary>
    public Single BodyFontSize { get; set; } = 14f;

    /// <summary>PPTX 标准幻灯片宽度（EMU），默认 9144000（10 英寸）</summary>
    public Int64 SlideWidthEmu { get; set; } = 9_144_000L;

    /// <summary>PPTX 标准幻灯片高度（EMU），默认 6858000（7.5 英寸）</summary>
    public Int64 SlideHeightEmu { get; set; } = 6_858_000L;

    #endregion

    #region 转换方法

    /// <summary>将 pptx 文件转换为 PDF 文件</summary>
    /// <param name="pptxPath">源 pptx 文件路径</param>
    /// <param name="pdfPath">目标 PDF 文件路径</param>
    public void Convert(String pptxPath, String pdfPath)
    {
        using var reader = new PptxReader(pptxPath);
        using var doc = CreateDocument();
        RenderSlides(reader, doc);
        doc.Save(pdfPath);
    }

    /// <summary>将 pptx 流转换为 PDF 流</summary>
    /// <param name="input">pptx 输入流</param>
    /// <param name="output">PDF 输出流（须可写）</param>
    public void Convert(Stream input, Stream output)
    {
        using var reader = new PptxReader(input);
        using var doc = CreateDocument();
        RenderSlides(reader, doc);
        doc.Save(output);
    }

    /// <summary>将 pptx 每张幻灯片渲染为图片（PNG/JPEG）</summary>
    /// <remarks>
    /// TODO: 本方法尚未实现。将幻灯片渲染为光栅图片需要引入渲染引擎（如 SkiaSharp + Docnet.Core）
    /// 将 PDF 中间格式解码为帧位图，当前版本不依赖外部库，故此功能暂不支持。
    /// 如需此功能，建议先调用 Convert 获得 PDF 字节，再用 PdfDocument.RenderToImages 处理。
    /// </remarks>
    /// <param name="pptxPath">pptx 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI）</param>
    /// <returns>每张幻灯片图片字节（PNG 格式）</returns>
    /// <exception cref="NotSupportedException">当前版本始终抛出，待引入渲染库后实现</exception>
    public IEnumerable<Byte[]> ConvertToImages(String pptxPath, Int32 dpi = 150)
    {
        throw new NotSupportedException(
            "渲染为图片需要引入 SkiaSharp/Docnet.Core 等渲染库，当前版本不支持。" +
            "建议先调用 Convert 转换为 PDF，再使用 PdfDocument.RenderToImages 处理。");
    }

    #endregion

    #region 私有方法

    /// <summary>创建并初始化 PDF 文档（横向 A4）</summary>
    private PdfFluentDocument CreateDocument()
    {
        var doc = new PdfFluentDocument();
        doc.SetLandscape();   // 842 × 595
        doc.SetMargins(40f, 40f, 40f, 40f);
        if (DocumentTitle != null) doc.Title = DocumentTitle;
        if (DocumentAuthor != null) doc.Author = DocumentAuthor;
        if (ShowSlideNumbers) doc.ShowPageNumbers = true;
        return doc;
    }

    /// <summary>将所有幻灯片渲染到文档</summary>
    /// <param name="reader">pptx 读取器</param>
    /// <param name="doc">PDF 文档</param>
    private void RenderSlides(PptxReader reader, PdfFluentDocument doc)
    {
        var slides = reader.ReadSlides().ToList();
        if (slides.Count == 0) return;

        for (var i = 0; i < slides.Count; i++)
        {
            if (i > 0) doc.PageBreak();
            RenderSlide(reader, doc, slides[i], i + 1);
        }
    }

    /// <summary>将单张幻灯片渲染到当前 PDF 页</summary>
    /// <param name="reader">pptx 读取器</param>
    /// <param name="doc">PDF 文档</param>
    /// <param name="slide">幻灯片摘要</param>
    /// <param name="slideNumber">幻灯片编号（1起始）</param>
    private void RenderSlide(PptxReader reader, PdfFluentDocument doc, PptSlideSummary slide, Int32 slideNumber)
    {
        // 按形状 Y 坐标排序，找标题（最靠上的非空文本）
        var shapes = slide.Shapes
            .Where(s => s.Text.Length > 0)
            .OrderBy(s => s.Top)
            .ToList();

        String? title = null;
        var bodyLines = new List<String>();

        if (shapes.Count > 0)
        {
            title = shapes[0].Text;
            for (var i = 1; i < shapes.Count; i++)
            {
                var text = shapes[i].Text.Trim();
                if (text.Length > 0) bodyLines.Add(text);
            }
        }
        else
        {
            // 回退：直接读取幻灯片文本
            var raw = reader.GetSlideText(slideNumber - 1);
            var lines = raw.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length > 0) title = lines[0].Trim();
            for (var i = 1; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (line.Length > 0) bodyLines.Add(line);
            }
        }

        // 顶部标题栏（矩形背景）
        DrawTitleBar(doc, slideNumber, title);

        // 正文
        foreach (var line in bodyLines)
        {
            // 跳过与标题相同的行（有时标题也出现在 body）
            if (line == title) continue;
            doc.AddText(line, BodyFontSize);
        }
    }

    /// <summary>绘制幻灯片标题栏</summary>
    /// <param name="doc">PDF 文档</param>
    /// <param name="slideNumber">幻灯片编号（1起始）</param>
    /// <param name="title">标题文本</param>
    private static void DrawTitleBar(PdfFluentDocument doc, Int32 slideNumber, String? title)
    {
        // 绘制深蓝色标题矩形（PDF 坐标 y=595-40-50=505，高50pt）
        var pageH = doc.PageHeight;
        var marginL = doc.MarginLeft;
        var marginR = doc.MarginRight;
        var contentW = doc.ContentWidth;

        // 蓝色背景块：从顶部边距开始，高 44pt
        var barY = pageH - 40f - 44f;      // PDF y（从底部起算）= 595-40-44 = 511
        doc.DrawRect(marginL, barY, contentW, 44f,
            fill: true, fillColor: "003087", borderColor: "003087", borderWidth: 0f);

        // 白色幻灯片标题文字（坐标从左下角，y = barY + 12）
        var textY = barY + 12f;
        if (title != null && title.Length > 0)
            doc.DrawText(title, marginL + 8f, textY, 18f);
        else
            doc.DrawText($"幻灯片 {slideNumber}", marginL + 8f, textY, 18f);

        // 推进 Y，使后续内容从标题栏下方开始（44 + 8 spacing）
        doc.AddEmptyLine(56f);
    }

    #endregion
}
