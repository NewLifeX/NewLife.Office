using System.Text;
using BenchmarkDotNet.Attributes;
using NewLife.Office.Markdown;

namespace MarkdownBenchmark;

/// <summary>Markdown 解析/序列化/HTML 转换性能基准（MD10）</summary>
/// <remarks>
/// 覆盖大文档（200 组混合块）的解析吞吐、往返解析、序列化与 HTML 输出，含内存分配指标。
/// 运行方式：dotnet run -c Release --project Benchmark\MarkdownBenchmark
/// </remarks>
[MemoryDiagnoser]
public class MarkdownParseBenchmark
{
    private String _text = "";
    private String _englishText = "";
    private MarkdownDocument? _doc;

    [GlobalSetup]
    public void Setup()
    {
        var sb = new StringBuilder();
        for (var i = 0; i < 200; i++)
        {
            sb.Append($"## 标题 {i}\n\n");
            sb.Append($"段落 {i} 包含 **粗体**、*斜体*、`代码` 和 [链接](https://newlifex.com) 与 https://newlifex.com/docs 。\n\n");
            sb.Append("- 项目一\n- 项目二\n- 项目三\n\n");
            sb.Append("| 列1 | 列2 |\n| --- | --- |\n| A | B |\n\n");
            sb.Append("> 引用内容\n\n");
            sb.Append("```csharp\nvar x = 1;\nConsole.WriteLine(x);\n```\n\n");
        }
        _text = sb.ToString();

        // 纯英文段落：长字母 run + 常见词触发 h/f/w 裸 URL 尝试，但无任何标记字符
        var esb = new StringBuilder();
        for (var i = 0; i < 300; i++)
            esb.Append($"The quick brown fox jumps over the lazy dog near the river. This is a simple paragraph number {i} with many english words and numbers 12345.\n\n");
        _englishText = esb.ToString();
        _doc = MarkdownDocument.Parse(_text);
    }

    [Benchmark]
    public MarkdownDocument Parse() => MarkdownDocument.Parse(_text);

    /// <summary>纯英文段落（含 h/f/w 触发裸 URL/Email 尝试，无标记字符）——验证 MD19 email run 快速跳过</summary>
    [Benchmark]
    public MarkdownDocument ParseEnglish() => MarkdownDocument.Parse(_englishText);

    [Benchmark]
    public MarkdownDocument ParseRoundtrip() => MarkdownDocument.ParseRoundtrip(_text);

    [Benchmark]
    public String Serialize() => _doc!.ToMarkdown();

    [Benchmark]
    public String SerializeRoundtrip()
    {
        var doc = MarkdownDocument.ParseRoundtrip(_text);
        return doc.ToMarkdown();
    }

    [Benchmark]
    public String ToHtml() => _doc!.ToHtml();

    [Benchmark]
    public String ToHtmlPage() => _doc!.ToHtmlPage("基准测试");
}
