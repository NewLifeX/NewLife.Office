namespace NewLife.Office.Ppt;

/// <summary>演示文稿模型扩展：保存与结构快照</summary>
public static class PresentationExtensions
{
    /// <summary>把演示文稿模型保存为 pptx 文件</summary>
    /// <remarks>
    /// 指定 <paramref name="templatePath"/> 时先加载该文件的母版/版式/主题，再重建幻灯片，
    /// 适合"读入已有文件 → 修改 → 另存"以保留原始视觉样式。
    /// </remarks>
    /// <param name="document">演示文稿模型</param>
    /// <param name="path">输出路径</param>
    /// <param name="templatePath">模板文件路径（通常是源文件）</param>
    public static void Save(this Presentation document, String path, String? templatePath = null)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (path == null) throw new ArgumentNullException(nameof(path));

        using var writer = templatePath.IsNullOrEmpty() ? new PptxWriter() : new PptxWriter(templatePath!);
        writer.Save(path, document);
    }

    /// <summary>把演示文稿模型写入流</summary>
    /// <param name="document">演示文稿模型</param>
    /// <param name="stream">目标流</param>
    /// <param name="templatePath">模板文件路径（通常是源文件）</param>
    public static void Save(this Presentation document, Stream stream, String? templatePath = null)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        using var writer = templatePath.IsNullOrEmpty() ? new PptxWriter() : new PptxWriter(templatePath!);
        writer.Save(stream, document);
    }

    /// <summary>构建结构快照</summary>
    /// <param name="document">演示文稿模型</param>
    /// <returns>结构快照</returns>
    public static PptxOutline ToOutline(this Presentation document) => PptxOutline.From(document);
}
