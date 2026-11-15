namespace NewLife.Office;

/// <summary>不可信文档的资源边界，防止压缩炸弹、超大计数与深层嵌套拖垮进程</summary>
/// <remarks>
/// 默认为宽松但有限的上限，覆盖真实世界的大文档；宿主可按需下调。
/// 超限统一抛出 <see cref="InvalidDataException"/>，不会静默截断内容。
/// </remarks>
public static class OfficeLimits
{
    /// <summary>单个 ZIP 条目解压后的最大字节数（默认 512MB）</summary>
    public static Int64 MaxZipEntryBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>ZIP 条目总数上限（默认 10 万）</summary>
    public static Int32 MaxZipEntries { get; set; } = 100_000;

    /// <summary>单个文档所有 ZIP 条目解压后的总字节数上限（默认 1GB）</summary>
    public static Int64 MaxZipTotalBytes { get; set; } = 1024L * 1024 * 1024;

    /// <summary>XML 文档最大字符数（默认 2.56 亿）</summary>
    public static Int32 MaxXmlCharacters { get; set; } = 256 * 1024 * 1024;

    /// <summary>CFB 目录项数量上限（默认 10 万）</summary>
    public static Int32 MaxCfbDirectoryEntries { get; set; } = 100_000;

    /// <summary>CFB FAT 条目数量上限（默认 800 万，约 32MB）</summary>
    public static Int32 MaxCfbFatEntries { get; set; } = 8_388_608;

    /// <summary>CFB 目录树最大深度（默认 64）</summary>
    public static Int32 MaxCfbTreeDepth { get; set; } = 64;

    /// <summary>单个 CFB 流解压后的最大字节数（默认 512MB）</summary>
    public static Int64 MaxCfbStreamBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>PDF 交叉引用对象数量上限（默认 100 万）</summary>
    public static Int32 MaxPdfObjects { get; set; } = 1_000_000;

    /// <summary>PDF 页数上限（默认 10 万）</summary>
    public static Int32 MaxPdfPages { get; set; } = 100_000;

    /// <summary>PDF 单个流解压后的最大字节数（默认 256MB）</summary>
    public static Int64 MaxPdfStreamBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>PDF 对象/页树的最大递归深度（默认 64）</summary>
    public static Int32 MaxPdfDepth { get; set; } = 64;

    /// <summary>通用嵌套深度上限（默认 64）</summary>
    public static Int32 MaxNestingDepth { get; set; } = 64;

    /// <summary>Excel 最大行数（与 Excel 2016+ 一致）</summary>
    public static Int32 MaxExcelRows { get; set; } = 1_048_576;

    /// <summary>Excel 最大列数（与 Excel 2016+ 一致）</summary>
    public static Int32 MaxExcelColumns { get; set; } = 16_384;

    /// <summary>ODS 单元格/行重复次数上限（默认 104 万）</summary>
    public static Int32 MaxOdsRepeatCount { get; set; } = 1_048_576;

    /// <summary>BIFF8 共享字符串数量上限（默认 100 万）</summary>
    public static Int32 MaxBiffSstStrings { get; set; } = 1_000_000;

    /// <summary>MIME multipart 嵌套深度上限（默认 32）</summary>
    public static Int32 MaxMimeDepth { get; set; } = 32;

    /// <summary>Markdown 引用块嵌套深度上限（默认 64）</summary>
    public static Int32 MaxMarkdownQuoteDepth { get; set; } = 64;
}
