# 性能测试报告 — Markdown 解析/序列化/HTML 转换

> 日期：2026-08-06 | 项目：NewLife.Office（MD10→MD12）| 基准项目：`Benchmark/MarkdownBenchmark`

## 测试目标

测量大文档（200 组混合块：标题/段落/列表/表格/引用/代码块，约 30KB）的解析、往返解析、序列化、往返序列化、HTML 转换与完整页面输出，覆盖吞吐量与内存分配；对比 MD12 优化（往返捕获门控/子解析器行复用/强调索引化）前后效果。

## 测试环境

- OS: Windows 10 22H2 (19045)
- CPU: Intel Core i9-10900K @ 3.70GHz（10 核 20 线程）
- .NET: .NET 10.0.10 (X64 RyuJIT AVX2)
- BenchmarkDotNet: v0.14.0（Release 模式，Short 模式运行，`[MemoryDiagnoser]`）

## 测试结果（MD12 优化后）

| 方法 | Mean | Error | Median | Gen0 | Gen1 | Gen2 | Allocated |
|------|-----:|------:|-------:|-----:|-----:|-----:|----------:|
| Parse | 20,858.4 μs | 337.83 μs | 20,569.8 μs | 343.75 | 218.75 | - | 3761.21 KB |
| ParseRoundtrip | 21,492.4 μs | 81.98 μs | 21,490.2 μs | 437.50 | 343.75 | - | 4629.03 KB |
| Serialize | 743.3 μs | 7.23 μs | 741.8 μs | 124.02 | 124.02 | 124.02 | 610.64 KB |
| SerializeRoundtrip | 23,730.8 μs | 117.93 μs | 23,717.7 μs | 531.25 | 437.50 | - | 5513.23 KB |
| ToHtml | 475.8 μs | 2.89 μs | 476.2 μs | 71.29 | 71.29 | 71.29 | 780.85 KB |
| ToHtmlPage | 680.7 μs | 12.61 μs | 683.0 μs | 213.87 | 213.87 | 213.87 | 1267.16 KB |

### MD12 优化前后对比

| 方法 | 优化前 Mean | 优化后 Mean | 优化前分配 | 优化后分配 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|
| Parse | 22,423.9 μs | 20,858.4 μs | 4810.29 KB | 3761.21 KB | **-21.8%** |
| ParseRoundtrip | 22,245.2 μs | 21,492.4 μs | 4810.28 KB | 4629.03 KB | -3.8% |
| Serialize | 781.9 μs | 743.3 μs | 610.23 KB | 610.64 KB | ≈0 |
| SerializeRoundtrip | 24,547.2 μs | 23,730.8 μs | 5694.46 KB | 5513.23 KB | -3.2% |
| ToHtml | 500.6 μs | 475.8 μs | 780.84 KB | 780.85 KB | ≈0 |

> 注：本文档规模约 200 组混合块（日常 Markdown 文件的数倍），优化后解析 ~21ms、序列化 ~0.74ms、HTML ~0.48ms，均处于可接受范围。

## 分析

- **解析 ~21ms / 3.8MB（分配 -22%）**：主要优化来自**往返捕获门控（MD12）**——非往返 `Parse` 不再为每个块生成 `SourceText`/`SourceSnapshot`（快照=整块序列化），GC 压力 Gen0 由 469 降至 344。两遍扫描 + 分隔符栈解析为主成本；无 Gen2（大对象堆无压力）。
- **Parse 与 ParseRoundtrip 差距 3%**：往返模式多捕获原始源码与快照，代价有限；需要原始格式保真时直接用 `ParseRoundtrip`（MD10-10 公开 API）。
- **SerializeRoundtrip（23.7ms）**：包含重新解析 + 快照比较，与"解析+序列化"整体成本一致，属预期。
- **序列化 0.74ms / 611KB**：`StringBuilder` + 上下文相关转义，为最快路径。
- **ToHtml 0.48ms / 781KB**：语义化输出 + 实体/链接处理，性能良好。
- **ToHtmlPage 0.68ms / 1.27MB**：额外页面骨架与内联样式，分配为 ToHtml 的 1.6 倍，符合预期。

## 优化建议（后续专项）

1. **AST 字符串复制**：解析分配主成本为每个文本令牌的独立字符串（输入字符的 ~120 倍）。改用 `Memory<Char>` 切片令牌可进一步降分配，但涉及 net45 兼容与解析器大重构。
2. **`text.Split('\n')` 行缓冲**：主入口逐行 `Split` 可换 `ArrayPool<Char>` 单行扫描，预计再降 10-20%。
3. **深层嵌套子解析**：列表/引用的子解析器每次 `ToArray()` 复制行数组，可传递 `IReadOnlyList` 引用避免复制。

## 运行方式

```bash
# 先打包主库到本地源（规避 BDN Deterministic 与通配版本冲突）
dotnet pack NewLife.Office\NewLife.Office.csproj -c Release -o Bin\LocalFeed
# 清理本地包缓存后运行基准（Release 模式）
Remove-Item "$env:USERPROFILE\.nuget\packages\newlife.office\1.3.2026.806" -Recurse -Force
dotnet run -c Release --project Benchmark\MarkdownBenchmark
```
