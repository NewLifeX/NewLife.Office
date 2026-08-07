using System.Text;
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
            Console.Error.WriteLine($"错误: {ex.Message}");
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
            Console.Error.WriteLine("用法: convert <input> [output]");
            return 1;
        }
        var input = args[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine($"文件不存在: {input}");
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
            Console.Error.WriteLine($"无法转换: {input}");
            return 1;
        }

        if (args.Length >= 2)
        {
            File.WriteAllText(args[1], md);
            Console.WriteLine($"已转换: {Path.GetFileName(input)} → {args[1]}");
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
            Console.Error.WriteLine("用法: doc2docx <input.doc> <output.docx>");
            return 1;
        }
        var input = args[0];
        var output = args[1];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine($"文件不存在: {input}");
            return 1;
        }

        OfficeFactory.ConvertDocToDocx(input, output);
        Console.WriteLine($"已转换: {Path.GetFileName(input)} → {output}");
        return 0;
    }

    /// <summary>info 命令：显示文件格式信息</summary>
    /// <param name="args">参数（&lt;input&gt;）</param>
    /// <returns>退出码</returns>
    private static Int32 Info(String[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("用法: info <input>");
            return 1;
        }
        var input = args[0];
        if (!File.Exists(input))
        {
            Console.Error.WriteLine($"文件不存在: {input}");
            return 1;
        }

        var format = OfficeFactory.Detect(input);
        var info = new FileInfo(input);
        Console.WriteLine($"文件: {input}");
        Console.WriteLine($"格式: {format ?? "未知"}");
        Console.WriteLine($"大小: {info.Length} 字节");
        return 0;
    }

    /// <summary>打印用法</summary>
    private static void PrintUsage()
    {
        Console.WriteLine("NewLife.Office CLI 工具");
        Console.WriteLine();
        Console.WriteLine("用法:");
        Console.WriteLine("  convert <input> [output]   转换办公文档为 Markdown（output 缺省输出到 stdout）");
        Console.WriteLine("  doc2docx <in.doc> <out.docx>   doc（97-2003）转换为 docx（W20）");
        Console.WriteLine("  info <input>               显示文件格式信息");
        Console.WriteLine();
        Console.WriteLine("支持格式: docx/xlsx/pptx/pdf/md/ods/rtf/epub 等");
    }

    private static Int32 PrintUsageAndReturn()
    {
        PrintUsage();
        return 1;
    }
}
