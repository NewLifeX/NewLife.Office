# 性能测试报告 — Markdown 解析/序列化/HTML 转换

> 日期：2026-08-08 | 项目：NewLife.Office（MD10→MD26）| 基准项目：`Benchmark/MarkdownBenchmark`

## 测试目标

测量大文档（200 组混合块：标题/段落/列表/表格/引用/代码块，约 30KB）的解析、往返解析、序列化、往返序列化、HTML 转换与完整页面输出，覆盖吞吐量与内存分配；对比 MD12 优化（往返捕获门控/子解析器行复用/强调索引化）前后效果。

## 测试环境

- OS: Windows 10 22H2 (19045)
- CPU: Intel Core i9-10900K @ 3.70GHz（10 核 20 线程）
- .NET: .NET 10.0.10 (X64 RyuJIT AVX2)
- BenchmarkDotNet: v0.14.0（Release 模式，Short 模式运行，`[MemoryDiagnoser]`）

## 测试结果（MD26 优化后）

| 方法 | Mean | Error | Median | Gen0 | Gen1 | Gen2 | Allocated |
|------|-----:|------:|-------:|-----:|-----:|-----:|----------:|
| Parse | 8,647.6 μs | 72.96 μs | 56.96 μs | 171.88 | 93.75 | - | 1878.19 KB |
| ParseEnglish | 2,186.3 μs | 15.53 μs | 12.12 μs | 15.63 | 3.91 | - | 193.96 KB |
| ParseRoundtrip | 9,454.8 μs | 38.55 μs | 36.06 μs | 234.38 | 171.88 | - | 2405.42 KB |
| Serialize | 735.8 μs | 14.51 μs | 13.57 μs | 124.02 | 124.02 | 124.02 | 603.50 KB |
| SerializeRoundtrip | 11,004.6 μs | 26.30 μs | 23.32 μs | 281.25 | 171.88 | - | 2967.72 KB |
| ToHtml | 505.1 μs | 10.07 μs | 24.12 μs | 71.29 | 71.29 | 71.29 | 780.85 KB |
| ToHtmlPage | 646.9 μs | 12.91 μs | 21.21 μs | 213.87 | 213.87 | 213.87 | 1267.16 KB |

### MD25→MD26 优化前后对比（Roundtrip 快照池化）

| 方法 | MD25 Mean | MD26 Mean | MD25 分配 | MD26 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 8,828.0 μs | 8,647.6 μs | 1878.18 KB | 1878.19 KB | -2.0%（噪声） | ≈0 |
| ParseEnglish | 2,153.8 μs | 2,186.3 μs | 193.96 KB | 193.96 KB | +1.5%（噪声） | ≈0 |
| ParseRoundtrip | 9,642.8 μs | 9,454.8 μs | 2764.76 KB | 2405.42 KB | -2.0% | **-13.0%** |
| SerializeRoundtrip | 11,395.8 μs | 11,004.6 μs | 3671.30 KB | 2967.72 KB | -3.4% | **-19.2%** |

> 注：MD26 为 Roundtrip 场景里程碑——往返模式每块生成规范快照时 `new MarkdownWriter()` + `new StringBuilder()` + 完整序列化（~2000 块），改为 **writer 实例复用（状态在 finally 复位，单线程安全）+ `Pool.StringBuilder` 池化**；另单行块 `SourceText` 捕获免 `new String[1]` + `String.Join`。Parse/ParseEnglish（非往返）分配不变，佐证改动精准；`SerializeRoundtrip` 收益更大（IsModified 修改检测同样每块序列化，池化同样生效）。

### MD19-26 优化前后对比（11 类优化）

| 方法 | MD18 Mean | MD26 Mean | MD18 分配 | MD26 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 8,962.4 μs | 8,647.6 μs | 2276.33 KB | 1878.19 KB | -3.5%（混合中文） | **-17.5%** |
| ParseEnglish | 2,729.3 μs | 2,186.3 μs | 442.42 KB | 193.96 KB | **-19.9%**（纯英文） | **-56.2%** |
| ParseRoundtrip | 9,913.7 μs | 9,454.8 μs | 3162.90 KB | 2405.42 KB | -4.6% | **-23.9%** |
| SerializeRoundtrip | - | 11,004.6 μs | 5513.23 KB（MD14） | 2967.72 KB | -53.6%（MD14 起） | **-46.2%**（MD14 起） |

> 注：MD18/MD26 同机 Short 模式对比（MD18 为 stash 回退实测；SerializeRoundtrip 为 MD14 起累计）。十一类优化——email run、delims 延迟、IsPlainText 精确化、块调度正则快速失败、emoji 快速失败、ASCII 快速路径、表格 IndexOfAny/实体范围、DelimInfo struct、**List 直接接管（MD24 块/行内 Inlines + MD25 容器块 Children）**、**Roundtrip 快照池化（MD26：writer 实例复用 + StringBuilder 池化 + 单行 SourceText）**、BuildParagraphText 单行快路径——混合中文分配 **-17.5%**、英文 -56.2%。累计 MD14→MD26 混合中文分配 **-50.1%**（3761→1878KB）、耗时 **-58%**（20.9→8.8ms）；Roundtrip 解析分配 **-23.9%**（MD18 起）、往返序列化分配 **-46.2%**（MD14 起）。

### MD14 优化前后对比（列表项快速路径）

| 方法 | 优化前 Mean | 优化后 Mean | 优化前分配 | 优化后分配 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|
| Parse | 20,858.4 μs | 16,603.2 μs | 3761.21 KB | 3100.27 KB | **-17.6%** |
| ParseRoundtrip | 21,492.4 μs | 17,594.2 μs | 4629.03 KB | 3968.09 KB | -14.3% |
| Serialize | 743.3 μs | 799.0 μs | 610.64 KB | 611.11 KB | ≈0 |
| SerializeRoundtrip | 23,730.8 μs | 20,193.9 μs | 5513.23 KB | 4852.28 KB | -12.0% |
| ToHtml | 475.8 μs | 498.4 μs | 780.85 KB | 780.85 KB | ≈0 |

### MD15 优化前后对比（引用块快速路径 + 引用预扫描过滤）

| 方法 | MD14 Mean | MD15 Mean | MD14 分配 | MD15 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 16,603.2 μs | 11,394.8 μs | 3100.27 KB | 2916.96 KB | **-31%** | **-5.9%** |
| ParseRoundtrip | 17,594.2 μs | 12,393.6 μs | 3968.09 KB | 3784.78 KB | -30% | -4.6% |
| Serialize | 799.0 μs | 749.3 μs | 611.11 KB | 611.31 KB | -6% | ≈0 |
| SerializeRoundtrip | 20,193.9 μs | 14,303.5 μs | 4852.28 KB | 4669.45 KB | -29% | -3.8% |
| ToHtml | 498.4 μs | 484.5 μs | 780.85 KB | 780.85 KB | -3% | ≈0 |
| ToHtmlPage | 745.5 μs | 685.1 μs | 1267.16 KB | 1267.16 KB | -8% | ≈0 |

### MD16 优化前后对比（行内令牌复用 + 表格行拆分索引化）

| 方法 | MD15 Mean | MD16 Mean | MD15 分配 | MD16 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 11,394.8 μs | 10,618.5 μs | 2916.96 KB | 2604.46 KB | **-6.8%** | **-10.7%** |
| ParseRoundtrip | 12,393.6 μs | 11,779.0 μs | 3784.78 KB | 3472.29 KB | -5.0% | **-8.3%** |
| Serialize | 749.3 μs | 756.3 μs | 611.31 KB | 611.23 KB | ≈0 | ≈0 |
| SerializeRoundtrip | 14,303.5 μs | 13,458.9 μs | 4669.45 KB | 4356.47 KB | -5.9% | -6.7% |
| ToHtml | 484.5 μs | 487.3 μs | 780.85 KB | 780.85 KB | ≈0 | ≈0 |
| ToHtmlPage | 685.1 μs | 696.4 μs | 1267.16 KB | 1267.16 KB | ≈0 | ≈0 |

### MD17 优化前后对比（AST 集合延迟初始化 + 裸 URL 前缀 Span 比较）

| 方法 | MD16 Mean | MD17 Mean | MD16 分配 | MD17 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 10,618.5 μs | 10,654.2 μs | 2604.46 KB | 2332.58 KB | ≈0 | **-10.4%** |
| ParseRoundtrip | 11,779.0 μs | 11,325.5 μs | 3472.29 KB | 3219.16 KB | -3.9% | **-7.3%** |
| Serialize | 756.3 μs | 720.1 μs | 611.31 KB | 611.32 KB | -4.8% | ≈0 |
| SerializeRoundtrip | 13,458.9 μs | 13,696.9 μs | 4356.47 KB | 4103.34 KB | +1.8% | -5.8% |
| ToHtml | 487.3 μs | 464.1 μs | 780.85 KB | 780.85 KB | -4.8% | ≈0 |
| ToHtmlPage | 696.4 μs | 631.8 μs | 1267.16 KB | 1267.16 KB | -9.3% | ≈0 |

### MD18 优化前后对比（行内纯文本快速路径 + 表格行快速拆分）

| 方法 | MD17 Mean | MD18 Mean | MD17 分配 | MD18 分配 | 耗时降幅 | 分配降幅 |
|------|-----------:|-----------:|----------:|----------:|--------:|--------:|
| Parse | 10,654.2 μs | 9,186.3 μs | 2332.58 KB | 2276.33 KB | **-13.8%** | -2.4% |
| ParseRoundtrip | 11,325.5 μs | 9,834.5 μs | 3219.16 KB | 3162.90 KB | **-13.2%** | -1.7% |
| Serialize | 720.1 μs | 745.7 μs | 611.32 KB | 611.24 KB | +3.6% | ≈0 |
| SerializeRoundtrip | 13,696.9 μs | 11,697.4 μs | 4103.34 KB | 4047.09 KB | **-14.6%** | -1.4% |
| ToHtml | 464.1 μs | 467.2 μs | 780.85 KB | 780.85 KB | ≈0 | ≈0 |
| ToHtmlPage | 631.8 μs | 662.1 μs | 1267.16 KB | 1267.16 KB | +4.8% | ≈0 |

> 注：本文档规模约 200 组混合块（日常 Markdown 文件的数倍），优化后解析 ~9.2ms、序列化 ~0.75ms、HTML ~0.47ms，均处于可接受范围。

## 分析

- **解析 ~11ms / 2.9MB（MD15 再降耗时 31%、分配 5.9%）**：MD15 延续 MD14 的子解析器消除思路——**引用块快速路径**（单行普通内容直接行内解析，不再创建子解析器；基准 200 个引用块全部命中）叠加 **引用预扫描行首过滤**（引用定义必须以 `[` 开头，≤3 空格缩进；普通文档绝大多数行直接跳过，不再逐行 Trim + 正则）。两者消除子解析器创建、FrontMatter 检查、预扫描开销，与 MD14 列表项快速路径（列表分配 -61%）叠加后 Parse 累计分配 -22%（3761→2917KB）、耗时累计 -45%（20.9→11.4ms）。
- **解析 ~10.6ms / 2.6MB（MD16 再降耗时 6.8%、分配 10.7%）**：MD16 经 GC 分段诊断定位两个高频小对象热点——①**行内令牌复用**：`ParseInlineCore` 的 tokens 临时列表改为别名复用 result（result 恒为空列表传入），消除每行内解析 1 个 List 分配 + 末尾 `AddRange` 全量复制；②**表格行拆分索引化**：`SplitTableRow` 用索引边界剥离首尾空白与管道符（替代 `Trim` + `[1..]`/`[..^1]` 子串），并移除每行 `StringBuilder`，普通单元格直接 `Substring` 零额外分配（转义管道 `\|` 仅命中时走 `ProcessCell`）。**表格是最大单项扩展成本**：诊断实测 422KB→298KB（占 Parse 分配 11%）。
- **GC 分段诊断结论（MD16）**：41.8KB 输入下 Parse 分配构成——预处理 Split 164KB（6%）、AST 无惰性构造（Parse 与 Parse+Touch 差值≈0）、扩展分支合计 ~420KB（表格为最大单项）。再次验证数据驱动决策：优化应优先命中子解析器消除（MD14/15）与高频小对象分配（MD16），而非字符串切片假设。
- **解析 ~10.7ms / 2.3MB（MD17 再降分配 10.4%）**：MD17 聚焦 **AST 集合延迟初始化**——`MarkdownInline.Children` 与 `MarkdownBlock.Children/Inlines` 从 `= []` 属性初始化器改为 `_children ??= []` 惰性创建（叶子文本令牌/叶子块不再各带一个空 List）。基准 200 组混合块约 3000 令牌 + 2800 块，每空 List ~40B → 省 ~250KB。附带 **裸 URL 前缀 Span 比较**（`TryParseBareUrl` 用 `text.AsSpan(...).StartsWith(prefix, OrdinalIgnoreCase)` 替代 `text[start..end]` 子串，每次调用省一个子串）。Gen0 250→219，GC 压力显著下降；Serialize/ToHtml/ToHtmlPage 耗时亦降 5-9%（延迟初始化减少序列化遍历对象）。
- **块类型分配分布（MD17 probe）**：纯段落 714KB/200 个（最重，行内令牌+自动链接）、纯表格 564KB、纯列表 339KB、纯标题 216KB、纯代码块 181KB、纯引用 142KB——段落行内解析与表格为两大结构性成本（AST 对象+文本子串，难以进一步消除）。
- **解析 ~9.2ms / 2.3MB（MD18 首次大幅降耗时 13.8%）**：耗时诊断（Stopwatch probe）显示**表格是最高耗时密度块类型**（0.56ms/KB vs 段落 0.23ms/KB），且自动链接/email 尝试各 ~0.15-0.2ms。MD18 实施——①**行内纯文本快速路径**：`ParseInline`/`ParseInlineWithRefs` 入口先 `IsPlainText` 扫描（排除转义/实体/代码/图片/链接/强调/删除线/<>/$/:/@/换行/裸URL前缀 hfw 等触发字符），无标记字符时直接 `CreateText` 单令牌，跳过 TokenizeInline 逐字符全分支检查；表格单元格/列表项/引用等简单文本全部命中。②**SplitTableRow 快速拆分**：行内无反引号/反斜杠时直接按 `|` 切分（省逐字符代码段/转义判断）。Parse 首次突破 10ms 大关，ParseRoundtrip/SerializeRoundtrip 耗时同步降 13-15%。
- **耗时诊断方法（MD18 probe）**：`Stopwatch` 均值测量 + `MarkdownPipeline` 开关对比（关闭自动链接/删除线识别分支成本）；块类型同字符量对比（表格 0.56ms/KB 最高）。自动链接成本 ~7%、email 尝试成本 ~6%（英文无 @ 文本）。
- **解析 ~9.0ms / 2.3MB（MD19-21 英文场景再降 19%）**：MD19 系列五个优化——①**Email run 快速跳过**：`TokenizeInline` 原对每个字母数字字符调用 `TryParseBareEmail`（内部再 `IndexOf('@')` 全段扫描），英文无 @ 文本呈 O(n²)；改为按连续 email 字符 run（含 `@`）单次 `IndexOf('@')`，无 @ 整段跳过。②**delims 延迟创建**：`ParseInlineCore` 原无条件 `new List<DelimInfo>()`，改为 `TokenizeInline` 仅遇 `*`/`_` 分隔符时才分配并返回（无强调行内解析省一个空 List，ParseEnglish 分配 -2.1%、Gen0 -3.9）。③**IsPlainText 精确化（MD19c）**：原保守排除 h/f/w（防裸 URL），纯英文段落（如 "The quick brown fox..." 含 h/f/w 但非 URL）被误排除快速路径 → 全走完整 TokenizeInline；改为带索引扫描 + `IsUrlPrefixAt`（精确匹配 `https://`/`http://`/`ftp://`/`www.` 前缀，与 `TryParseBareUrl` 前缀集一致），非 URL 的 h/f/w 放行——URL 必以前缀开头故不会漏识别，纯英文段落命中快速路径单令牌直达。④**块调度正则快速失败（MD20）**：`TryParseAtxHeading`/`IsOrderedListMarker`/`TryParseFootnoteDefinition` 三个每行调用的正则方法加前缀字符检查（`#`/数字/`[^`，均为正则必然前缀），混合文档普通行不再触发正则引擎，只做一次字符比较。⑤**emoji 快速失败（MD21）**：耗时 probe 定位 emoji 分支省 12%——`:` 触发时先检查 `text[i+1]` 是否为字母/数字/下划线（emoji 短码 `:word:` 必以字母开头），URL 协议分隔符 `://`、时间 `12:30` 等直接跳过；且 `IndexOf(':', i+1)` 扫描范围限制 ≤31（code 长度 ≤30 才有效，避免扫到段尾）。**踩坑**：a) email run 界定用了 `IsEmailChar`（含 `_`），会把 `_` 强调标记吞进 run 跳过 → 强调解析失效（3 测试失败）；修复为 run 内含 `_` 时回退原逻辑（逐字符尝试），保证 `john_doe@x.com` 与强调兼容。b) delims 参数按值传递，方法内 `??=` 只改参数副本导致调用方仍为 null（36 测试失败）；改为 `TokenizeInline` 返回列表。混合中文基准英文 run 多被中文/标记打断收益不明显，纯英文长 run 场景 ParseEnglish -19.0%（stash 回退同机对比实测）。
- **推翻 MD12 假设**：MD12 报告猜测「AST 字符串复制 ~120x 为主成本」，经 GC 分段实测文本令牌字符串仅占 1.3%（~49KB）；`Memory<Char>` 惰性切片（net45 经 System.Memory 支持）实测无收益且因对象膨胀与惰性转换使 Serialize/ToHtml 略增，**已回滚**。数据驱动决策：真实收益在列表/引用子解析器消除。
- **Parse 与 ParseRoundtrip 差距 8%**：往返模式多捕获原始源码与快照，代价有限；需要原始格式保真时直接用 `ParseRoundtrip`（MD10-10 公开 API）。
- **SerializeRoundtrip（14.3ms）**：包含重新解析 + 快照比较，与"解析+序列化"整体成本一致，属预期。
- **序列化 0.75ms / 611KB**：`StringBuilder` + 上下文相关转义，为最快路径。
- **ToHtml 0.5ms / 781KB**：语义化输出 + 实体/链接处理，性能良好。

## 优化建议（后续专项）

1. **深层嵌套列表/引用行复制**：子解析器 `ParseLines` 仍 `ToArray()` 复制行数组，可传 `IReadOnlyList` 引用（MD15 后仅多行引用/多行列表项仍创建子解析器；实测影响较小，属代码质量改进）。
2. **块级行缓冲**：主入口 `text.Split('\n')` 每行独立子串（诊断实测预处理 164KB/6%），可 `ArrayPool<Char>` 单行扫描（预计收益有限，需基准验证）。
3. **表格单元格快速路径（MD16 后续）**：基准 200 表格 × 4 单元格 ≈800 次 `ParseInlineWithRefs`（每次独立 List + Text 对象）。单 token 单元格（如 `A`/`B`）可直达文本构造，但 AST 对象为必需结构，预计收益有限；若需进一步压降可评估。
4. **DelimInfo class→struct（评估中）**：强调分隔符 run 每项一个引用对象（~48B），改 struct 需处理 `ResolveEmphasis` 的可空 opener 与写回，风险中等、预计收益 ~1-2%，暂缓。

> MD24 已完成（里程碑）：块级/行内节点**直接接管 Parse 创建的 List**（工厂 `List` 重载 + `SetInlines`/`_children` 直赋，省每块/每节点 1 个新 List + AddRange 复制）+ `BuildParagraphText` 单行快路径——混合中文分配 -15.6%、ParseEnglish 分配 -55%；累计 MD14→MD24 分配 -49.5%、耗时 -58%。

## 运行方式

```bash
# 先打包主库到本地源（规避 BDN Deterministic 与通配版本冲突）
dotnet pack NewLife.Office\NewLife.Office.csproj -c Release -o Bin\LocalFeed
# 清理本地包缓存后运行基准（Release 模式）
Remove-Item "$env:USERPROFILE\.nuget\packages\newlife.office\1.3.2026.807" -Recurse -Force
# 基准失败（CS0246 NewLife 找不到 / 0 runs）时：重新 pack + 删缓存 + 删 Bin\Benchmark 下 BDN 临时目录
dotnet run -c Release --project Benchmark\MarkdownBenchmark
```
