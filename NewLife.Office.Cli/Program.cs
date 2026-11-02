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
        var input = positional[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
            return 1;
        }

        // 仅在显式指定选项时构造，避免改变默认转换行为
        MarkdownConverterOptions? options = null;
        if (imageDir != null || includeHeaderFooter)
        {
            options = new MarkdownConverterOptions
            {
                IncludeHeadersFooters = includeHeaderFooter,
                ImageDirectory = imageDir,
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
            File.WriteAllText(positional[1], md);
            Console.WriteLine(Localization.Translate("Converted: {0} → {1}", Path.GetFileName(input), positional[1]));
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
        var input = args[0];
        var output = args[1];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
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
        var input = args[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
            return 1;
        }

        var format = OfficeFactory.Detect(input);
        var info = new FileInfo(input);
        Console.WriteLine(Localization.Translate("File: {0}", input));
        Console.WriteLine(Localization.Translate("Format: {0}", format ?? Localization.Translate("Unknown")));
        Console.WriteLine(Localization.Translate("Size: {0} bytes", info.Length));
        return 0;
    }

    /// <summary>read 命令：输出 docx 结构快照（JSON），供 Agent 看清结构后按锚点修改</summary>
    /// <param name="args">参数（&lt;input.docx&gt;）</param>
    /// <returns>退出码</returns>
    private static Int32 ReadOutline(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: read <input.docx>"));
            return 1;
        }
        var input = args[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
            return 1;
        }

        using var reader = new NewLife.Office.Word.WordReader(input);
        WriteStdout(reader.ReadOutline().ToJson());
        return 0;
    }

    /// <summary>apply 命令：按补丁 JSON 修改 docx；任一条失败则不写出文件（原子）</summary>
    /// <param name="args">参数（&lt;in.docx&gt; &lt;patch.json&gt; &lt;out.docx&gt;）</param>
    /// <returns>退出码（0=全部成功，1=参数/文件错误，3=有操作失败）</returns>
    private static Int32 ApplyPatch(String[] args)
    {
        if (args.Length < 3)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: apply <in.docx> <patch.json> <out.docx>"));
            return 1;
        }
        var input = args[0];
        var patchPath = args[1];
        var output = args[2];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
            return 1;
        }
        if (!File.Exists(patchPath))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", patchPath));
            return 1;
        }

        var patch = NewLife.Office.Word.DocxPatch.ParseJson(File.ReadAllText(patchPath));
        using var reader = new NewLife.Office.Word.WordReader(input);
        var doc = reader.ReadDocument();
        doc.BeginEdit();

        var result = NewLife.Office.Word.DocxPatch.Apply(doc, patch);
        if (!result.Success)
        {
            foreach (var item in result.Failed)
                Console.Error.WriteLine(item);
            return 3;
        }

        using (var writer = new NewLife.Office.Word.WordWriter())
            writer.Save(output, doc);

        Console.WriteLine(Localization.Translate("Applied {0} operation(s): {1} → {2}", result.Applied.Count, Path.GetFileName(input), output));
        return 0;
    }

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
        Console.WriteLine(Localization.Translate("  read <input.docx>          Dump docx structure snapshot as JSON (for agents)"));
        Console.WriteLine(Localization.Translate("  apply <in> <patch> <out>   Apply JSON patch to docx (no output file if any op fails)"));
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
