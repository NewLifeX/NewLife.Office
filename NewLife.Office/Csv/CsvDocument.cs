using System.Text;
using NewLife.IO;
using NewLife.Office.Excel;
using NewLife.Office.Markdown;

namespace NewLife.Office.Csv;

/// <summary>CSV 文档模型：原始行数据 + 分隔符，提供文本/Markdown 提取与工作簿转换</summary>
/// <remarks>
/// 解析与写入复用 NewLife.Core 的 <see cref="CsvFile"/>（RFC4180：引号包裹、字段内换行、双引号转义），
/// 本类只负责编码/分隔符识别、资源上限与向下游模型的桥接。
/// <para>示例：</para>
/// <code>
/// var csv = CsvDocument.ReadFile("sales.csv");
/// var markdown = csv.ExtractMarkdown();
/// var outline = csv.ToExcelDocument().ToOutline().ToJson();
/// </code>
/// </remarks>
public sealed class CsvDocument : ITextExtractable, IMarkdownExtractable, IMarkdownDocumentExtractable
{
    #region 属性
    private readonly List<String[]> _rows;

    /// <summary>文档名称（读取文件时为文件名，读取流时为空串）</summary>
    public String Name { get; }

    /// <summary>实际使用的字段分隔符</summary>
    public Char Separator { get; }

    /// <summary>原始行数据（含表头行，字段已去除包裹引号）</summary>
    public IReadOnlyList<String[]> Rows => _rows;

    /// <summary>数据行数（含表头行）</summary>
    public Int32 RowCount => _rows.Count;
    #endregion

    #region 构造
    private CsvDocument(String name, Char separator, List<String[]> rows)
    {
        Name = name;
        Separator = separator;
        _rows = rows;
    }

    /// <summary>读取 CSV 文件</summary>
    /// <param name="path">文件路径</param>
    /// <param name="options">读取选项，null 使用默认（自动识别编码与分隔符）</param>
    /// <returns>CSV 文档</returns>
    /// <exception cref="ArgumentNullException">path 为空</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    public static CsvDocument ReadFile(String path, CsvOptions? options = null)
    {
        if (String.IsNullOrWhiteSpace(path)) throw new ArgumentNullException(nameof(path));

        var fullPath = path.GetFullPath();
        if (!File.Exists(fullPath)) throw new FileNotFoundException(L.T("File not found: {0}", fullPath), fullPath);

        using var fs = File.OpenRead(fullPath);
        return ReadCore(fs, options, Path.GetFileName(fullPath));
    }

    /// <summary>读取 CSV 数据流</summary>
    /// <param name="stream">数据流，调用方负责释放；不可寻址时先复制到内存</param>
    /// <param name="options">读取选项，null 使用默认（自动识别编码与分隔符）</param>
    /// <returns>CSV 文档</returns>
    /// <exception cref="ArgumentNullException">stream 为空</exception>
    public static CsvDocument Read(Stream stream, CsvOptions? options = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        return ReadCore(stream, options, null);
    }

    private static CsvDocument ReadCore(Stream stream, CsvOptions? options, String? name)
    {
        MemoryStream? buffer = null;
        try
        {
            // 不可寻址流无法回退采样位置，先落到内存
            if (!stream.CanSeek)
            {
                buffer = OfficeStreams.CopyToMemory(stream, OfficeLimits.MaxCsvBytes);
                stream = buffer;
            }

            var sample = ReadSample(stream, 64 * 1024);
            var encoding = options?.Encoding ?? OfficeTextDecoder.DetectEncoding(sample, out _);
            var separator = options?.Separator ?? '\0';
            if (separator == '\0')
            {
                var bomLength = 0;
                if (options?.Encoding == null) OfficeTextDecoder.DetectEncoding(sample, out bomLength);
                var text = sample.Length > bomLength ? encoding.GetString(sample, bomLength, sample.Length - bomLength) : "";
                separator = CsvText.DetectSeparator(FirstLine(text));
            }

            stream.Position = 0;
            using var csv = new CsvFile(stream, leaveOpen: true) { Encoding = encoding, Separator = separator };

            var rows = new List<String[]>();
            foreach (var row in csv.ReadAll())
            {
                if (rows.Count >= OfficeLimits.MaxCsvRows)
                    throw new InvalidDataException(L.T("The CSV row count exceeds the resource limit ({0}).", OfficeLimits.MaxCsvRows));
                if (row.Length > OfficeLimits.MaxCsvColumns)
                    throw new InvalidDataException(L.T("The CSV column count {0} exceeds the resource limit ({1}).", row.Length, OfficeLimits.MaxCsvColumns));
                rows.Add(row);
            }

            return new CsvDocument(name ?? "", separator, rows);
        }
        finally
        {
            buffer?.Dispose();
        }
    }
    #endregion

    #region 提取
    /// <summary>提取规范化 CSV 文本（首行原样输出，字段按 RFC4180 转义）</summary>
    /// <returns>CSV 文本，无数据返回 null</returns>
    public String? ExtractText()
    {
        if (_rows.Count == 0) return null;

        var sb = new StringBuilder();
        foreach (var row in _rows)
            sb.Append(CsvText.BuildLine(row, Separator)).Append("\r\n");
        return sb.ToString();
    }

    /// <summary>提取 Markdown 文档对象（单张 GFM 表格，首行为表头）</summary>
    /// <returns>Markdown 文档，无数据返回 null</returns>
    public MarkdownDocument? ToMarkdownDocument()
    {
        if (_rows.Count == 0) return null;

        var rows = new List<String[]>(_rows.Count);
        foreach (var row in _rows)
        {
            var empty = true;
            foreach (var cell in row)
            {
                if (!String.IsNullOrWhiteSpace(cell))
                {
                    empty = false;
                    break;
                }
            }
            if (!empty) rows.Add(row);
        }
        if (rows.Count == 0) return null;

        var doc = new MarkdownDocument();
        doc.Blocks.Add(MarkdownBlockBuilder.Table(rows.ToArray(), true));
        return doc;
    }

    /// <summary>提取 Markdown 表格文本</summary>
    /// <returns>Markdown 字符串，无数据返回 null</returns>
    public String? ExtractMarkdown() => ToMarkdownDocument()?.ToMarkdown();

    /// <summary>转换为工作簿模型，供结构快照、区域读取与搜索复用</summary>
    /// <param name="sheetName">工作表名称，默认 Sheet1</param>
    /// <returns>仅含单张工作表的工作簿</returns>
    public ExcelDocument ToExcelDocument(String sheetName = "Sheet1")
    {
        var document = new ExcelDocument();
        var sheet = new Worksheet { Name = String.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName };
        foreach (var row in _rows)
        {
            var values = new Object?[row.Length];
            for (var i = 0; i < row.Length; i++)
                values[i] = row[i];
            sheet.Rows.Add(values);
        }
        document.Worksheets.Add(sheet);
        return document;
    }
    #endregion

    #region 辅助
    private static Byte[] ReadSample(Stream stream, Int32 size)
    {
        var origin = stream.Position;
        try
        {
            var length = (Int32)Math.Min(size, Math.Max(0, stream.Length - origin));
            var buffer = new Byte[length];
            var total = 0;
            while (total < length)
            {
                var n = stream.Read(buffer, total, length - total);
                if (n <= 0) break;
                total += n;
            }
            if (total < length) Array.Resize(ref buffer, total);
            return buffer;
        }
        finally
        {
            stream.Position = origin;
        }
    }

    private static String FirstLine(String text)
    {
        if (String.IsNullOrEmpty(text)) return "";

        var index = text.IndexOfAny(['\r', '\n']);
        var line = index >= 0 ? text.Substring(0, index) : text;
        return line.TrimStart('\uFEFF');
    }
    #endregion
}
