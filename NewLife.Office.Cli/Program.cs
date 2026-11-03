using System.Runtime.CompilerServices;
using System.Text;
using NewLife;
using NewLife.Office.Markdown;

namespace NewLife.Office.Cli;

/// <summary>NewLife.Office 命令行工具（GEN-2，对标 anydoc CLI/Agent Skill）</summary>
/// <remarks>
/// 提供 convert 命令将办公文档（Word/Excel/PPT/PDF/Markdown/ODS/RTF/EPUB 等）转换为 Markdown，
/// info 命令显示格式信息。内容识别 + 格式路由自动选择转换器，零外部依赖，可被 Agent/脚本直接调用。
/// <example>
/// <code>
/// NewLife.Office.Cli convert input.docx output.md
/// NewLife.Office.Cli convert input.xlsx
/// NewLife.Office.Cli info input.pdf
/// </code>
/// </example>
/// </remarks>
public static class Program
{
    static Program() => RuntimeHelpers.RunModuleConstructor(typeof(OfficeFactory).Module.ModuleHandle);

    /// <summary>程序入口</summary>
    /// <param name="args">命令行参数</param>
    /// <returns>退出码（0=成功，1=参数/文件错误，2=转换失败）</returns>
    public static Int32 Main(String[] args)
    {
        // PDF 文本提取需要 Latin-1/1252 编码支持
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try
        {
            if (args == null || args.Length == 0)
            {
                PrintUsage();
                return 1;
            }
            return args[0].ToLowerInvariant() switch
            {
                "convert" or "conv" => Convert(args.Skip(1).ToArray()),
                "doc2docx" or "d2x" => Doc2Docx(args.Skip(1).ToArray()),
                "info" => Info(args.Skip(1).ToArray()),
                "read" => ReadOutline(args.Skip(1).ToArray()),
                "apply" => ApplyPatch(args.Skip(1).ToArray()),
                "text" => ReadText(args.Skip(1).ToArray()),
                "range" => ReadRange(args.Skip(1).ToArray()),
                "search" => Search(args.Skip(1).ToArray()),
                "mcp" => McpServer.Run(),
                _ => PrintUsageAndReturn(),
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(Localization.Translate("Error: {0}", ex.Message));
            return 2;
        }
    }

    /// <summary>convert 命令：办公文档 → Markdown</summary>
    /// <param name="args">参数（&lt;input&gt; [output] [--images &lt;dir&gt;] [--header-footer]）</param>
    /// <returns>退出码</returns>
    private static Int32 Convert(String[] args)
    {
        // 选项解析：--images <dir> 图片落地目录，--header-footer 输出页眉页脚
        String? imageDir = null;
        var includeHeaderFooter = false;
        var positional = new List<String>();
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--images":
                    if (i + 1 < args.Length) imageDir = args[++i];
                    break;
                case "--header-footer":
                    includeHeaderFooter = true;
                    break;
                default:
                    positional.Add(args[i]);
                    break;
            }
        }

        if (positional.Count < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: convert <input> [output] [--images <dir>] [--header-footer]"));
            return 1;
        }
        var input = ResolvePath(positional[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", positional[0]));
            return 1;
        }

        // 仅在显式指定选项时构造，避免改变默认转换行为
        MarkdownConverterOptions? options = null;
        if (imageDir != null || includeHeaderFooter)
        {
            options = new MarkdownConverterOptions
            {
                IncludeHeadersFooters = includeHeaderFooter,
                ImageDirectory = imageDir != null ? ResolvePath(imageDir) : null,
                ImageHandling = imageDir != null ? MarkdownImageHandling.Reference : MarkdownImageHandling.Skip,
            };
        }

        // Markdown 直接透传，其余格式走转换中枢（失败时内容识别兜底，GEN-2b）
        String? md;
        var ext = Path.GetExtension(input).ToLowerInvariant();
        if (ext is ".md" or ".markdown")
        {
            md = File.ReadAllText(input);
        }
        else
        {
            try
            {
                md = FormatToMarkdown.Convert(input, options);
            }
            catch (ConvertErrorException)
            {
                // 扩展名缺失/错误：按文件头内容识别路由（GEN-2b）
                var reader = OfficeFactory.CreateReaderByContent(input);
                md = (reader as IMarkdownOptionExtractable)?.ToMarkdownDocument(options)?.ToMarkdown()
                    ?? (reader as IMarkdownExtractable)?.ExtractMarkdown();
            }
        }

        if (md == null)
        {
            Console.Error.WriteLine(Localization.Translate("Unable to convert: {0}", input));
            return 1;
        }

        if (positional.Count >= 2)
        {
            var output = ResolvePath(positional[1]);
            File.WriteAllText(output, md);
            Console.WriteLine(Localization.Translate("Converted: {0} → {1}", Path.GetFileName(input), output));
        }
        else
        {
            Console.Write(md);
        }
        return 0;
    }

    /// <summary>doc2docx 命令：doc（97-2003）→ docx（W20）</summary>
    /// <param name="args">参数（&lt;input.doc&gt; &lt;output.docx&gt;）</param>
    /// <returns>退出码</returns>
    private static Int32 Doc2Docx(String[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: doc2docx <input.doc> <output.docx>"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        var output = ResolvePath(args[1]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        OfficeFactory.ConvertDocToDocx(input, output);
        Console.WriteLine(Localization.Translate("Converted: {0} → {1}", Path.GetFileName(input), output));
        return 0;
    }

    /// <summary>info 命令：显示文件格式信息</summary>
    /// <param name="args">参数（&lt;input&gt;）</param>
    /// <returns>退出码</returns>
    private static Int32 Info(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: info <input>"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        var format = OfficeFactory.Detect(input);
        var info = new FileInfo(input);
        Console.WriteLine(Localization.Translate("File: {0}", input));
        Console.WriteLine(Localization.Translate("Format: {0}", format ?? Localization.Translate("Unknown")));
        Console.WriteLine(Localization.Translate("Size: {0} bytes", info.Length));
        return 0;
    }

    /// <summary>read 命令：输出文档结构快照（JSON），按扩展名/内容自动识别格式</summary>
    /// <param name="args">参数（&lt;input&gt;）</param>
    /// <returns>退出码</returns>
    private static Int32 ReadOutline(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: read <input>"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        var result = OfficeDocumentTool.ReadOutline(input);
        if (!result.Success)
        {
            Console.Error.WriteLine($"[{result.Code}] {result.Message}");
            return 2;
        }

        WriteStdout(result.Content!);
        return 0;
    }

    /// <summary>apply 命令：按补丁 JSON 修改文档；任一条失败则不写出文件（原子）</summary>
    /// <param name="args">参数（&lt;in&gt; &lt;patch.json&gt; &lt;out&gt;）</param>
    /// <returns>退出码（0=全部成功，1=参数/文件错误，2=格式/加密错误，3=有操作失败）</returns>
    private static Int32 ApplyPatch(String[] args)
    {
        if (args.Length < 3)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: apply <in> <patch.json> <out>"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        var patchPath = ResolvePath(args[1]);
        var output = ResolvePath(args[2]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }
        if (!File.Exists(patchPath))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[1]));
            return 1;
        }

        var result = OfficeDocumentTool.ApplyPatch(input, File.ReadAllText(patchPath), output);
        if (!result.Success)
        {
            Console.Error.WriteLine($"[{result.Code}] {result.Message}");
            return result.Code == OfficeToolErrorCode.AnchorMissing ? 3 : 2;
        }

        Console.WriteLine(Localization.Translate("Applied: {0} → {1}", Path.GetFileName(input), output));
        // 告警不影响成功状态，但必须让用户看到（否则"内容被静默丢弃"）
        foreach (var warning in result.Warnings)
            Console.Error.WriteLine(Localization.Translate("Warning: {0}", warning));
        return 0;
    }

    /// <summary>text 命令：分段读取纯文本</summary>
    /// <param name="args">参数（&lt;input&gt; [--max-chars N] [--cursor N]）</param>
    /// <returns>退出码</returns>
    private static Int32 ReadText(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: text <input> [--max-chars N] [--cursor N]"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        var maxChars = 0;
        var cursor = 0;
        for (var i = 1; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--max-chars" when i + 1 < args.Length:
                    Int32.TryParse(args[++i], out maxChars);
                    break;
                case "--cursor" when i + 1 < args.Length:
                    Int32.TryParse(args[++i], out cursor);
                    break;
            }
        }

        var result = OfficeDocumentTool.ReadText(input, maxChars, cursor);
        if (!result.Success)
        {
            Console.Error.WriteLine($"[{result.Code}] {result.Message}");
            return 2;
        }

        WriteStdout(result.Content ?? String.Empty);
        if (result.Truncated)
            Console.Error.WriteLine($"[Truncated] nextCursor={result.NextCursor}");
        return 0;
    }

    /// <summary>range 命令：分页读取工作簿区域</summary>
    /// <param name="args">参数（&lt;input&gt; &lt;range&gt; [--max-rows N] [--cursor N]）</param>
    /// <returns>退出码</returns>
    private static Int32 ReadRange(String[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: range <input> <range> [--max-rows N] [--cursor N]"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        var maxRows = 0;
        var cursor = 0;
        for (var i = 2; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--max-rows" when i + 1 < args.Length:
                    Int32.TryParse(args[++i], out maxRows);
                    break;
                case "--cursor" when i + 1 < args.Length:
                    Int32.TryParse(args[++i], out cursor);
                    break;
            }
        }

        var result = OfficeDocumentTool.ReadRange(input, args[1], maxRows, cursor);
        if (!result.Success)
        {
            Console.Error.WriteLine($"[{result.Code}] {result.Message}");
            return 2;
        }

        WriteStdout(result.Content!);
        return 0;
    }

    /// <summary>search 命令：在工作簿或 PDF 中搜索</summary>
    /// <param name="args">参数（&lt;input&gt; &lt;text&gt; [--regex] [--max N]）</param>
    /// <returns>退出码</returns>
    private static Int32 Search(String[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: search <input> <text> [--regex] [--max N]"));
            return 1;
        }
        var input = ResolvePath(args[0]);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", args[0]));
            return 1;
        }

        var regex = false;
        var max = 100;
        for (var i = 2; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--regex":
                    regex = true;
                    break;
                case "--max" when i + 1 < args.Length:
                    Int32.TryParse(args[++i], out max);
                    break;
            }
        }

        var result = OfficeDocumentTool.Search(input, args[1], regex, max);
        if (!result.Success)
        {
            Console.Error.WriteLine($"[{result.Code}] {result.Message}");
            return 2;
        }

        WriteStdout(result.Content!);
        return 0;
    }

    /// <summary>把用户传入的路径解析为绝对路径</summary>
    /// <remarks>
    /// 统一按“当前工作目录”解析相对路径。必须与 <see cref="File.Exists"/> 的基准一致：
    /// NewLife 的 <c>GetFullPath()</c> 扩展以程序目录为基准，直接用相对路径会出现
    /// “File.Exists 通过但真正打开时报找不到路径”的假故障。
    /// </remarks>
    /// <param name="path">用户输入的路径（绝对或相对）</param>
    /// <returns>绝对路径</returns>
    internal static String ResolvePath(String path) => Path.GetFullPath(path);

    /// <summary>向标准输出写文本；重定向/管道时按 UTF-8 写字节（避免控制台代码页转码乱码）</summary>
    /// <param name="text">输出文本</param>
    private static void WriteStdout(String text)
    {
        if (Console.IsOutputRedirected)
        {
            using var stdout = Console.OpenStandardOutput();
            using var writer = new StreamWriter(stdout, new UTF8Encoding(false)) { AutoFlush = true };
            writer.Write(text);
        }
        else
        {
            Console.Write(text);
        }
    }

    /// <summary>打印用法</summary>
    private static void PrintUsage()
    {
        Console.WriteLine(Localization.Translate("NewLife.Office CLI tool"));
        Console.WriteLine();
        Console.WriteLine(Localization.Translate("Usage:"));
        Console.WriteLine(Localization.Translate("  convert <input> [output]   Convert office documents to Markdown (output defaults to stdout)"));
        Console.WriteLine(Localization.Translate("      --images <dir>         Write images to <dir> and reference them by file name"));
        Console.WriteLine(Localization.Translate("      --header-footer        Include headers/footers as blockquotes (Word)"));
        Console.WriteLine(Localization.Translate("  doc2docx <in.doc> <out.docx>   Convert doc (97-2003) to docx (W20)"));
        Console.WriteLine(Localization.Translate("  info <input>               Show file format information"));
        Console.WriteLine(Localization.Translate("  read <input>               Dump structure snapshot as JSON (docx/xlsx/pptx/pdf, for agents)"));
        Console.WriteLine(Localization.Translate("  apply <in> <patch> <out>   Apply JSON patch to docx/xlsx/pptx/pdf (no output file if any op fails)"));
        Console.WriteLine(Localization.Translate("  text <input> [--max-chars N] [--cursor N]   Read plain text in chunks"));
        Console.WriteLine(Localization.Translate("  range <input> <range> [--max-rows N] [--cursor N]   Read a worksheet range page by page"));
        Console.WriteLine(Localization.Translate("  search <input> <text> [--regex] [--max N]   Search workbook cells or PDF text"));
        Console.WriteLine(Localization.Translate("  mcp                        Start MCP stdio server (tools for agents)"));
        Console.WriteLine();
        Console.WriteLine(Localization.Translate("Supported formats: docx/xlsx/pptx/pdf/md/ods/rtf/epub, etc."));
    }

    private static Int32 PrintUsageAndReturn()
    {
        PrintUsage();
        return 1;
    }
}
