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
    /// <param name="args">参数（&lt;input&gt; [output]）</param>
    /// <returns>退出码</returns>
    private static Int32 Convert(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine(Localization.Translate("Usage: convert <input> [output]"));
            return 1;
        }
        var input = args[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine(Localization.Translate("File not found: {0}", input));
            return 1;
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
                md = FormatToMarkdown.Convert(input);
            }
            catch (ConvertErrorException)
            {
                // 扩展名缺失/错误：按文件头内容识别路由（GEN-2b）
                var reader = OfficeFactory.CreateReaderByContent(input);
                md = (reader as IMarkdownExtractable)?.ExtractMarkdown();
            }
        }

        if (md == null)
        {
            Console.Error.WriteLine(Localization.Translate("Unable to convert: {0}", input));
            return 1;
        }

        if (args.Length >= 2)
        {
            File.WriteAllText(args[1], md);
            Console.WriteLine(Localization.Translate("Converted: {0} → {1}", Path.GetFileName(input), args[1]));
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

    /// <summary>打印用法</summary>
    private static void PrintUsage()
    {
        Console.WriteLine(Localization.Translate("NewLife.Office CLI tool"));
        Console.WriteLine();
        Console.WriteLine(Localization.Translate("Usage:"));
        Console.WriteLine(Localization.Translate("  convert <input> [output]   Convert office documents to Markdown (output defaults to stdout)"));
        Console.WriteLine(Localization.Translate("  doc2docx <in.doc> <out.docx>   Convert doc (97-2003) to docx (W20)"));
        Console.WriteLine(Localization.Translate("  info <input>               Show file format information"));
        Console.WriteLine();
        Console.WriteLine(Localization.Translate("Supported formats: docx/xlsx/pptx/pdf/md/ods/rtf/epub, etc."));
    }

    private static Int32 PrintUsageAndReturn()
    {
        PrintUsage();
        return 1;
    }
}
