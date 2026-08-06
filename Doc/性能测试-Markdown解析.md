# 性能测试报告 — Markdown 解析/序列化/HTML 转换

> 日期：2026-08-06 | 项目：NewLife.Office（MD10→MD12）| 基准项目：`Benchmark/MarkdownBenchmark`

## 测试目标

测量大文档（200 组混合块：标题/段落/列表/表格/引用/代码块，约 30KB）的解析、往返解析、序列化、往返序列化、HTML 转换与完整页面输出，覆盖吞吐量与内存分配；对比 MD12 优化（往返捕获门控/子解析器行复用/强调索引化）前后效果。

## 测试环境

- OS: Windows 10 22H2 (19045)
- CPU: Intel Core i9-10900K @ 3.70GHz（10 核 20 线程）
- .NET: .NET 10.0.10 (X64 RyuJIT AVX2)
- BenchmarkDotNet: v0.14.0（Release 模式，Short 模式运行，`[MemoryDiagnoser]`）

## 测试结果（MD14 优化后）

| 方法 | Mean | Error | Median | Gen0 | Gen1 | Gen2 | Allocated |
|------|-----:|------:|-------:|-----:|-----:|-----:|----------:|
| Parse | 16,603.2 μs | 203.07 μs | 16,522.4 μs | 281.25 | 187.50 | - | 3100.27 KB |
| ParseRoundtrip | 17,594.2 μs | 244.17 μs | 17,519.5 μs | 375.00 | 187.50 | - | 3968.09 KB |
| Serialize | 799.0 μs | 15.55 μs | 798.8 μs | 124.02 | 124.02 | 124.02 | 611.11 KB |
| SerializeRoundtrip | 20,193.9 μs | 349.16 μs | 19,873.6 μs | 468.75 | 343.75 | - | 4852.28 KB |
| ToHtml | 498.4 μs | 3.38 μs | 499.4 μs | 71.29 | 71.29 | 71.29 | 780.85 KB |
| ToHtmlPage | 745.5 μs | 14.70 μs | - | 213.87 | 213.87 | 213.87 | 1267.16 KB |

### MD14 优化前后对比（列表项快速路径）

| 方法 | 优化前 Mean | 优化后 Mean | 优化前分配 | 优化后分配 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|
| Parse | 20,858.4 μs | 16,603.2 μs | 3761.21 KB | 3100.27 KB | **-17.6%** |
| ParseRoundtrip | 21,492.4 μs | 17,594.2 μs | 4629.03 KB | 3968.09 KB | -14.3% |
| Serialize | 743.3 μs | 799.0 μs | 610.64 KB | 611.11 KB | ≈0 |
| SerializeRoundtrip | 23,730.8 μs | 20,193.9 μs | 5513.23 KB | 4852.28 KB | -12.0% |
| ToHtml | 475.8 μs | 498.4 μs | 780.85 KB | 780.85 KB | ≈0 |

> 注：本文档规模约 200 组混合块（日常 Markdown 文件的数倍），优化后解析 ~17ms、序列化 ~0.8ms、HTML ~0.5ms，均处于可接受范围。

## 分析

- **解析 ~17ms / 3.1MB（分配 -18%、耗时 -20%）**：MD14 通过 GC 分段量化定位真实热点——**每个列表项都创建子解析器 + 子文档 + FrontMatter 检查 + 引用预扫描 + ToArray**（列表输入 233x 放大）。单行普通内容快速路径直接行内解析，列表分配 -61%。Gen0 由 344 降至 281。
- **推翻 MD12 假设**：MD12 报告猜测「AST 字符串复制 ~120x 为主成本」，经 GC 分段实测文本令牌字符串仅占 1.3%（~49KB）；`Memory<Char>` 惰性切片（net45 经 System.Memory 支持）实测无收益且因对象膨胀与惰性转换使 Serialize/ToHtml 略增，**已回滚**。数据驱动决策：真实收益在列表子解析器消除。
- **Parse 与 ParseRoundtrip 差距 13%**：往返模式多捕获原始源码与快照，代价有限；需要原始格式保真时直接用 `ParseRoundtrip`（MD10-10 公开 API）。
- **SerializeRoundtrip（20.2ms）**：包含重新解析 + 快照比较，与"解析+序列化"整体成本一致，属预期。
- **序列化 0.8ms / 611KB**：`StringBuilder` + 上下文相关转义，为最快路径。
- **ToHtml 0.5ms / 781KB**：语义化输出 + 实体/链接处理，性能良好。

## 优化建议（后续专项）

1. **深层嵌套列表/引用行复制**：子解析器 `ParseLines` 仍 `ToArray()` 复制行数组，可传 `IReadOnlyList` 引用（实测影响较小，属代码质量改进）。
2. **块级行缓冲**：主入口 `text.Split('\n')` 每行独立子串，可 `ArrayPool<Char>` 单行扫描（预计收益有限，需基准验证）。

## 运行方式

```bash
# 先打包主库到本地源（规避 BDN Deterministic 与通配版本冲突）
dotnet pack NewLife.Office\NewLife.Office.csproj -c Release -o Bin\LocalFeed
# 清理本地包缓存后运行基准（Release 模式）
Remove-Item "$env:USERPROFILE\.nuget\packages\newlife.office\1.3.2026.806" -Recurse -Force
dotnet run -c Release --project Benchmark\MarkdownBenchmark
```
