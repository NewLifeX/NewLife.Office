# 性能测试报告 — Markdown 解析/序列化/HTML 转换

> 日期：2026-08-06 | 项目：NewLife.Office（MD10）| 基准项目：`Benchmark/MarkdownBenchmark`

## 测试目标

测量大文档（200 组混合块：标题/段落/列表/表格/引用/代码块，约 30KB）的解析、往返解析、序列化、往返序列化、HTML 转换与完整页面输出，覆盖吞吐量与内存分配，建立 MD10 优化后的基线，防止未来回归。

## 测试环境

- OS: Windows 10 22H2 (19045)
- CPU: Intel Core i9-10900K @ 3.70GHz（10 核 20 线程）
- .NET: .NET 10.0.10 (X64 RyuJIT AVX2)
- BenchmarkDotNet: v0.14.0（Release 模式，Short 模式运行，`[MemoryDiagnoser]`）

## 测试结果

| 方法 | Mean | Error | Median | Gen0 | Gen1 | Gen2 | Allocated |
|------|-----:|------:|-------:|-----:|-----:|-----:|----------:|
| Parse | 22,423.9 μs | 58.65 μs | 22,426.6 μs | 468.75 | 375.00 | - | 4810.29 KB |
| ParseRoundtrip | 22,245.2 μs | 47.02 μs | 22,245.7 μs | 468.75 | 375.00 | - | 4810.28 KB |
| Serialize | 781.9 μs | 14.31 μs | 774.3 μs | 124.02 | 124.02 | 124.02 | 610.23 KB |
| SerializeRoundtrip | 24,547.2 μs | 373.30 μs | 24,553.5 μs | 531.25 | 406.25 | - | 5694.46 KB |
| ToHtml | 500.6 μs | 9.29 μs | 492.3 μs | 71.29 | 71.29 | 71.29 | 780.84 KB |
| ToHtmlPage | 702.7 μs | 13.43 μs | 703.0 μs | 213.87 | 213.87 | 213.87 | 1267.15 KB |

> 注：本文档规模约 200 组混合块（日常 Markdown 文件的数倍），解析 ~22ms、序列化 ~0.8ms、HTML ~0.5ms，均处于可接受范围。

## 分析

- **解析 ~22ms / 4.8MB**：两遍扫描 + 分隔符栈解析（MD10 完整强调算法）。分配偏高源于逐行 `Split`、AST 对象实例与 `tokens.IndexOf` 引用查找；GC 压力集中在 Gen0/Gen1，无 Gen2（大对象堆无压力）。
- **Parse 与 ParseRoundtrip 几乎相同**：往返模式仅多捕获 `SourceText` 字符串，开销可忽略，可直接默认启用。
- **SerializeRoundtrip（24.5ms）**：包含重新解析 + 快照比较，与"解析+序列化"整体成本一致，属预期。
- **序列化 0.8ms / 610KB**：`StringBuilder` + 上下文相关转义，为五者中最快路径。
- **ToHtml 0.5ms / 781KB**：语义化输出 + 实体/链接处理，性能良好。
- **ToHtmlPage 0.7ms / 1.27MB**：额外页面骨架与内联样式，分配为 ToHtml 的 1.6 倍，符合预期。

## 优化建议（后续专项，本轮不实施）

1. **解析分配优化**：`Parse` 每次分配 4.8MB，主要来自 `String.Split` 与 AST 节点。可改用 `ArrayPool<Char>` 单行缓冲 + 节点复用，预计分配下降 30-50%。
2. **`tokens.IndexOf` 引用查找**：分隔符解析中 O(n²) 查找，可改为双向链表（CommonMark 参考实现方式），大文档强调密集场景有提升空间。
3. **预编译正则**：块级解析中多处 `Regex.Match` 未缓存，可静态化提升常量。

## 运行方式

```bash
# 先打包主库到本地源（规避 BDN Deterministic 与通配版本冲突）
dotnet pack NewLife.Office\NewLife.Office.csproj -c Release -o Bin\LocalFeed
# 运行基准（Release 模式）
dotnet run -c Release --project Benchmark\MarkdownBenchmark
```
