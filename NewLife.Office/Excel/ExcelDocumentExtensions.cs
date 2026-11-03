namespace NewLife.Office.Excel;

/// <summary>工作簿模型扩展：保存与按名称取工作表</summary>
public static class ExcelDocumentExtensions
{
    #region 保存
    /// <summary>把工作簿模型保存为 xlsx 文件</summary>
    /// <remarks>先序列化到内存，成功后才写文件，避免中途失败留下半个文件。</remarks>
    /// <param name="document">工作簿模型</param>
    /// <param name="path">输出路径</param>
    public static void Save(this ExcelDocument document, String path)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (path == null) throw new ArgumentNullException(nameof(path));

        using var ms = new MemoryStream();
        var writer = new ExcelWriter(ms);
        writer.WriteExcel(document);
        writer.Save();

        var file = path.GetFullPath();
        file.EnsureDirectory(true);
        File.WriteAllBytes(file, ms.ToArray());
    }

    /// <summary>把工作簿模型写入流（调用方负责流的生命周期）</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="stream">目标流</param>
    public static void Save(this ExcelDocument document, Stream stream)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var writer = new ExcelWriter(stream);
        writer.WriteExcel(document);
        writer.Save();
    }
    #endregion

    #region 查找
    /// <summary>按名称查找工作表（不区分大小写）</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="name">工作表名称，null 或空返回第一个工作表</param>
    /// <returns>工作表，未找到返回 null</returns>
    public static Worksheet? FindSheet(this ExcelDocument document, String? name)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.Worksheets.Count == 0) return null;
        if (name == null || name.Length == 0) return document.Worksheets[0];

        foreach (var sheet in document.Worksheets)
            if (String.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)) return sheet;
        return null;
    }

    /// <summary>按名称获取工作表，未找到抛出异常</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="name">工作表名称</param>
    /// <returns>工作表</returns>
    public static Worksheet GetSheet(this ExcelDocument document, String? name)
        => document.FindSheet(name) ?? throw new ArgumentException($"工作表不存在: {name}");

    /// <summary>构建结构快照</summary>
    /// <param name="document">工作簿模型</param>
    /// <param name="previewRows">每个工作表输出的预览行数</param>
    /// <returns>结构快照</returns>
    public static ExcelOutline ToOutline(this ExcelDocument document, Int32 previewRows = 5)
        => ExcelOutline.From(document, previewRows);
    #endregion
}
