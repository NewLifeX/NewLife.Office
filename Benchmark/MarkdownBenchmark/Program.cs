using BenchmarkDotNet.Running;

namespace MarkdownBenchmark;

/// <summary>MarkdownBenchmark 入口（性能测试技能：Release 模式运行）</summary>
/// <remarks>dotnet run -c Release --project Benchmark\MarkdownBenchmark</remarks>
public static class Program
{
    public static void Main(String[] args)
    {
        BenchmarkRunner.Run<MarkdownParseBenchmark>();
    }
}
