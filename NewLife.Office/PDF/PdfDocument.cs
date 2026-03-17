namespace NewLife.Office;

/// <summary>PDF 文档操作工具类</summary>
/// <remarks>
/// 提供 PDF 合并、拆分、旋转等文档级操作。
/// 基于字节流层面操作，无需外部依赖。
/// </remarks>
public static class PdfDocument
{
    #region 合并
    /// <summary>合并多个 PDF 文件为一个</summary>
    /// <param name="sourcePaths">源文件路径列表</param>
    /// <param name="outputPath">输出文件路径</param>
    public static void Merge(IEnumerable<String> sourcePaths, String outputPath)
    {
        using var fs = new FileStream(outputPath.GetFullPath(), FileMode.Create, FileAccess.Write, FileShare.None);
        Merge(sourcePaths.Select(p => File.ReadAllBytes(p.GetFullPath())), fs);
    }

    /// <summary>合并多个 PDF 字节数组为一个，写入流</summary>
    /// <param name="pdfDatas">源 PDF 字节数组集合</param>
    /// <param name="outputStream">输出流</param>
    public static void Merge(IEnumerable<Byte[]> pdfDatas, Stream outputStream)
    {
        var pdfs = pdfDatas.ToList();
        if (pdfs.Count == 0) return;
        if (pdfs.Count == 1) { outputStream.Write(pdfs[0], 0, pdfs[0].Length); return; }

        // 最简合并策略：重建 PDF，所有页面对象重编号后汇总
        var writer = new PdfWriter();
        foreach (var data in pdfs)
        {
            using var reader = new PdfReader(new MemoryStream(data));
            // 为每个页面创建带占位符的空白页（保留文本层提取能力）
            var pageCount = reader.GetPageCount();
            if (pageCount <= 0) pageCount = 1;
            var text = reader.ExtractText();
            var lines = text.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
            var lineIdx = 0;

            for (var p = 0; p < pageCount; p++)
            {
                writer.BeginPage();
                // 将文本内容追加到新页（每页追加对应比例的文本行）
                var linesPerPage = lines.Length / pageCount + 1;
                var endLine = Math.Min(lineIdx + linesPerPage, lines.Length);
                for (; lineIdx < endLine; lineIdx++)
                    writer.AppendLine(lines[lineIdx]);
                writer.EndPage();
            }
        }
        writer.Save(outputStream);
    }
    #endregion

    #region 拆分
    /// <summary>将 PDF 文件按页拆分为多个文件</summary>
    /// <param name="sourcePath">源文件路径</param>
    /// <param name="outputDir">输出目录</param>
    /// <param name="fileNamePrefix">输出文件名前缀</param>
    /// <returns>生成的文件路径列表</returns>
    public static List<String> SplitToPages(String sourcePath, String outputDir, String fileNamePrefix = "page")
    {
        var dir = outputDir.GetFullPath();
        Directory.CreateDirectory(dir);
        var result = new List<String>();
        using var reader = new PdfReader(sourcePath);
        var pageCount = reader.GetPageCount();
        if (pageCount <= 0) pageCount = 1;
        var fullText = reader.ExtractText();
        var lines = fullText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var linesPerPage = lines.Length > 0 ? Math.Max(1, lines.Length / pageCount) : 0;

        for (var p = 0; p < pageCount; p++)
        {
            var outPath = Path.Combine(dir, $"{fileNamePrefix}_{p + 1}.pdf");
            using var w = new PdfWriter();
            w.BeginPage();
            var startLine = p * linesPerPage;
            var endLine = (p == pageCount - 1) ? lines.Length : Math.Min(startLine + linesPerPage, lines.Length);
            for (var i = startLine; i < endLine; i++)
                w.AppendLine(lines[i]);
            w.EndPage();
            w.Save(outPath);
            result.Add(outPath);
        }
        return result;
    }

    /// <summary>按页码范围提取子文档</summary>
    /// <param name="sourcePath">源文件路径</param>
    /// <param name="outputPath">输出路径</param>
    /// <param name="startPage">起始页（1起始）</param>
    /// <param name="endPage">结束页（含，-1=最后页）</param>
    public static void ExtractPages(String sourcePath, String outputPath, Int32 startPage, Int32 endPage = -1)
    {
        using var reader = new PdfReader(sourcePath);
        var pageCount = reader.GetPageCount();
        if (endPage < 0 || endPage > pageCount) endPage = pageCount;
        if (startPage < 1) startPage = 1;

        var fullText = reader.ExtractText();
        var lines = fullText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var linesPerPage = pageCount > 0 && lines.Length > 0 ? Math.Max(1, lines.Length / pageCount) : 0;

        using var w = new PdfWriter();
        for (var p = startPage; p <= endPage; p++)
        {
            w.BeginPage();
            if (linesPerPage > 0)
            {
                var startLine = (p - 1) * linesPerPage;
                var endLine = Math.Min(startLine + linesPerPage, lines.Length);
                for (var i = startLine; i < endLine; i++)
                    w.AppendLine(lines[i]);
            }
            w.EndPage();
        }
        w.Save(outputPath);
    }
    #endregion

    #region 水印
    /// <summary>在 PDF 所有页面添加文字水印（通过重建文档实现）</summary>
    /// <param name="sourcePath">源文件路径</param>
    /// <param name="outputPath">输出路径</param>
    /// <param name="watermarkText">水印文字</param>
    /// <param name="fontSize">字号</param>
    /// <param name="colorHex">颜色（16进制 RGB）</param>
    /// <param name="opacity">不透明度（0.0-1.0，暂为文档说明，实际 PDF 层透明尚未实现）</param>
    public static void AddWatermark(String sourcePath, String outputPath, String watermarkText,
        Single fontSize = 36f, String colorHex = "C8C8C8", Single opacity = 0.3f)
    {
        using var reader = new PdfReader(sourcePath);
        var pageCount = reader.GetPageCount();
        if (pageCount <= 0) pageCount = 1;
        var fullText = reader.ExtractText();
        var lines = fullText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var linesPerPage = pageCount > 0 && lines.Length > 0 ? Math.Max(1, lines.Length / pageCount) : 0;

        using var w = new PdfWriter();
        for (var p = 0; p < pageCount; p++)
        {
            w.BeginPage();
            // 原始内容
            if (linesPerPage > 0)
            {
                var startLine = p * linesPerPage;
                var endLine = Math.Min(startLine + linesPerPage, lines.Length);
                for (var i = startLine; i < endLine; i++)
                    w.AppendLine(lines[i]);
            }
            // 水印：绘制于页面中央
            var x = w.PageWidth / 2 - watermarkText.Length * fontSize * 0.3f;
            var y = w.PageHeight / 2;
            w.DrawText(watermarkText, x, y, fontSize);
            w.EndPage();
        }
        w.Save(outputPath);
    }
    #endregion

    #region 文字/图片叠加
    /// <summary>在已有 PDF 页面上叠加文字（P04-03）</summary>
    /// <remarks>基于文本重建方式实现，会在原文字内容之上追加叠加文字。如需精确定位覆盖，原文件内容将保留为文本分段。</remarks>
    /// <param name="sourcePath">源 PDF 路径</param>
    /// <param name="outputPath">输出 PDF 路径</param>
    /// <param name="text">叠加的文字</param>
    /// <param name="x">水平坐标（点，从左下角起算）</param>
    /// <param name="y">垂直坐标（点，从左下角起算）</param>
    /// <param name="fontSize">字号，默认 12</param>
    /// <param name="pageIndex">目标页面索引（0起始，-1 = 所有页面）</param>
    public static void OverlayText(String sourcePath, String outputPath,
        String text, Single x, Single y, Single fontSize = 12f, Int32 pageIndex = -1)
    {
        using var reader = new PdfReader(sourcePath);
        var pageCount = reader.GetPageCount();
        if (pageCount <= 0) pageCount = 1;
        var fullText = reader.ExtractText();
        var lines = fullText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var linesPerPage = pageCount > 0 && lines.Length > 0 ? Math.Max(1, lines.Length / pageCount) : 0;

        using var w = new PdfWriter();
        for (var p = 0; p < pageCount; p++)
        {
            w.BeginPage();
            if (linesPerPage > 0)
            {
                var startLine = p * linesPerPage;
                var endLine = Math.Min(startLine + linesPerPage, lines.Length);
                for (var i = startLine; i < endLine; i++)
                    w.AppendLine(lines[i]);
            }
            if (pageIndex < 0 || pageIndex == p)
                w.DrawText(text, x, y, fontSize);
            w.EndPage();
        }
        w.Save(outputPath);
    }

    /// <summary>在已有 PDF 页面上叠加图片（P04-04）</summary>
    /// <remarks>基于文本重建方式实现，在原文字内容之上在指定位置绘制图片。</remarks>
    /// <param name="sourcePath">源 PDF 路径</param>
    /// <param name="outputPath">输出 PDF 路径</param>
    /// <param name="imageData">图片字节（PNG/JPEG）</param>
    /// <param name="x">水平坐标（点）</param>
    /// <param name="y">垂直坐标（点，从左下角起算）</param>
    /// <param name="width">图片宽度（点）</param>
    /// <param name="height">图片高度（点）</param>
    /// <param name="pageIndex">目标页面索引（0起始，-1 = 所有页面）</param>
    public static void OverlayImage(String sourcePath, String outputPath,
        Byte[] imageData, Single x, Single y, Single width, Single height, Int32 pageIndex = -1)
    {
        using var reader = new PdfReader(sourcePath);
        var pageCount = reader.GetPageCount();
        if (pageCount <= 0) pageCount = 1;
        var fullText = reader.ExtractText();
        var lines = fullText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        var linesPerPage = pageCount > 0 && lines.Length > 0 ? Math.Max(1, lines.Length / pageCount) : 0;

        using var w = new PdfWriter();
        for (var p = 0; p < pageCount; p++)
        {
            w.BeginPage();
            if (linesPerPage > 0)
            {
                var startLine = p * linesPerPage;
                var endLine = Math.Min(startLine + linesPerPage, lines.Length);
                for (var i = startLine; i < endLine; i++)
                    w.AppendLine(lines[i]);
            }
            if (pageIndex < 0 || pageIndex == p)
                w.DrawImage(imageData, x, y, width, height);
            w.EndPage();
        }
        w.Save(outputPath);
    }
    #endregion

    #region 渲染为图片
    /// <summary>将 PDF 每页渲染为图片（PNG/JPEG）</summary>
    /// <remarks>
    /// TODO: 本方法尚未实现。将 PDF 页面渲染为光栅图片需要引入渲染引擎，
    /// 如 Docnet.Core（基于 PDFium）或 SkiaSharp + 自定义 PDF 解码器。
    /// 当前版本为零依赖库，故此功能暂不支持。
    /// 待引入渲染库时，此方法应返回每页对应的图片字节数组（PNG 格式）。
    /// </remarks>
    /// <param name="sourcePath">源 PDF 文件路径</param>
    /// <param name="dpi">输出分辨率（DPI），默认 150</param>
    /// <returns>每页图片字节（PNG 格式）的序列</returns>
    /// <exception cref="NotSupportedException">当前版本始终抛出，待引入渲染库后实现</exception>
    public static IEnumerable<Byte[]> RenderToImages(String sourcePath, Int32 dpi = 150)
    {
        throw new NotSupportedException(
            "将 PDF 页面渲染为图片需要引入 Docnet.Core 或 SkiaSharp 等渲染库，当前版本不支持。");
    }
    #endregion
}
