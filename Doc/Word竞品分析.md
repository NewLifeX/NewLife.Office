# Word 竞品分析

> 版本：v2.0 | 日期：2026-08-06
> 返回：[竞品分析报告](竞品分析报告.md)

## 1. 概述

本报告针对 .NET 生态中 Word 文档操作的主流开源/商业库，从**功能覆盖度、高保真读写、API 易用性、性能、体积、许可证、框架兼容性**七个维度进行深度对比，为 NewLife.Office Word 模块的功能定位与差异化竞争提供依据。

### 1.1 竞品全景

| 库名 | 许可证 | 最新版本 | 包体积 | 支持格式 | Stars | 下载量 |
|------|--------|---------|--------|---------|-------|--------|
| **Open XML SDK** | MIT | 3.x | ~2MB | docx/xlsx/pptx | ~4k | 3000万+ |
| **NPOI** | Apache 2.0 | 2.7.x | ~10MB | doc/docx/xls/xlsx/ppt/pptx | ~5k | 5000万+ |
| **MiniWord** | Apache 2.0 | 0.9.x | <500KB | docx（仅模板写入） | ~600+ | 50万+ |
| **OfficeIMO** | MIT | 0.17.x | ~1MB | docx | ~300+ | 10万+ |
| **TemplateEngine.Docx** | MIT | 1.x | <1MB | docx（仅模板） | ~200+ | 20万+ |
| **GemBox.Document** | 免费受限/商业 | 3.7.x | ~5MB | docx/pdf/html/rtf | N/A | 100万+ |
| **Aspose.Words** | 商业 | 25.x | ~40MB | doc/docx/rtf/html/pdf 等 | N/A | 1000万+ |
| **Spire.Doc** | 免费受限/商业 | 12.x | ~15MB | doc/docx/pdf/html/rtf | N/A | 200万+ |
| **NewLife.Office** | MIT | 1.3.x | <500KB | docx/doc | — | — |

### 1.2 竞品分层定位

```mermaid
graph TD
    subgraph 底层
        A[Open XML SDK]
    end
    subgraph 中层封装
        B[NPOI]
        C[OfficeIMO]
        D[GemBox.Doc Free]
    end
    subgraph 轻量专用
        E[MiniWord]
        F[TemplateEngine.Docx]
    end
    subgraph 商业旗舰
        G[Aspose.Words]
        H[Spire.Doc Pro]
    end
    subgraph 全栈MIT
        I[NewLife.Office]
    end
    A --> B
    A --> C
    A --> D
    I --> J[NewLife.Core]
```

> **NewLife.Office 定位**：MIT 许可下功能最全面的 Word 文档库，以文档模型（`Document`）为核心实现 docx 高保真读写往返，覆盖程序化创建、模板填充、对象映射、邮件合并与格式转换的完整链路。唯一在 MIT 许可下同时支持 docx 读写模型 + doc 读取 + 模板 + 转换的开源方案。

## 2. Word 功能对比矩阵（92 项）

> 标记说明：✅ 完整支持 | ⚠️ 部分可用（含 RawXml 透传兜底）| ❌ 不支持。NewLife.Office 列基于 v1.3.x 已实现能力；竞品列基于其公开能力合理评估。

### 2.1 基础读写

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W01 | 读取 docx 文本 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W02 | 写入/创建 docx | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W03 | 文档模型（对象模型读写往返） | ✅ | ❌ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W04 | 读取 doc（97-2003） | ✅ | ❌ | ✅ | ❌ | ❌ | ✅ | ✅ |
| W05 | 写入 doc | ❌ | ❌ | ✅ | ❌ | ❌ | ✅ | ✅ |
| W06 | RawXml 透传保真 | ✅ | ✅ | ⚠️ | ❌ | ❌ | ✅ | ✅ |
| W07 | ZIP 部件完整透传 | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |

### 2.2 段落与文本格式

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W08 | 粗体/斜体 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W09 | 下划线（多样式） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W10 | 删除线 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W11 | 上标/下标 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W12 | 字体名称/大小 | ✅ | ✅ | ✅ | ⚠️ | ✅ | ✅ | ✅ |
| W13 | 字体颜色 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W14 | 字符间距/字符缩放 | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W15 | 文字效果（发光/阴影） | ✅ | ✅ | ❌ | ❌ | ❌ | ⚠️ | ✅ |
| W16 | 多 Run 富文本段落 | ✅ | ✅ | ✅ | ⚠️ | ✅ | ✅ | ✅ |
| W17 | 段落内嵌图片 | ✅ | ✅ | ⚠️ | ✅ | ⚠️ | ✅ | ✅ |

### 2.3 段落与排版

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W18 | 标题 1-6 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W19 | 段落对齐（左/中/右/两端） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W20 | 段落缩进（左/右/首行/悬挂） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W21 | 段前/段后间距 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W22 | 行距（单倍/1.5/双倍/固定值） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W23 | 段落边框（四边） | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W24 | 段落背景色 | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W25 | 制表位（Tab Stop） | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W26 | 分页符 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W27 | 分页控制（keepNext/keepLines/widowControl） | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W28 | 分节符（页面/连续/奇偶页） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W29 | 首字下沉 | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ | ✅ |
| W30 | 段落大纲级别 | ⚠️ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |

### 2.4 列表

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W31 | 无序列表（项目符号） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W32 | 有序列表（编号） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W33 | 多级嵌套列表 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W34 | 自定义项目符号/编号格式 | ⚠️ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W35 | 列表续号 | ⚠️ | ✅ | ❌ | ❌ | ⚠️ | ✅ | ✅ |

### 2.5 表格

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W36 | 创建表格 | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W37 | 读取表格 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W38 | 单元格合并（横向/纵向） | ✅ | ✅ | ✅ | ⚠️ | ✅ | ✅ | ✅ |
| W39 | 表格边框 | ✅ | ✅ | ✅ | ⚠️ | ✅ | ✅ | ✅ |
| W40 | 单元格背景色 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W41 | 表头行样式 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W42 | 斑马纹（交替行色） | ✅ | ✅ | ⚠️ | ❌ | ❌ | ✅ | ✅ |
| W43 | 列宽精确控制 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W44 | 单元格垂直对齐 | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W45 | 单元格内多段落 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W46 | 嵌套表格 | ⚠️ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |

### 2.6 图片与绘图

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W47 | 插入图片（PNG/JPEG） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W48 | 图片尺寸/位置控制 | ✅ | ✅ | ✅ | ⚠️ | ✅ | ✅ | ✅ |
| W49 | 图片环绕方式 | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W50 | 文本框/形状 | ⚠️ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W51 | 自选图形（线条/矩形/圆等） | ❌ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W52 | SVG/复杂矢量图形（透传保留） | ⚠️ | ✅ | ❌ | ❌ | ❌ | ✅ | ✅ |

### 2.7 页面与布局

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W53 | 页面尺寸/方向/边距 | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W54 | 页眉（default/first/even 三类型） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W55 | 页脚（含页码） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W56 | 奇偶页/首页不同页眉页脚 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W57 | 多节不同页面设置 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W58 | 分栏（ColumnCount） | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W59 | 行号 | ✅ | ✅ | ❌ | ❌ | ✅ | ✅ | ✅ |
| W60 | 页面边框 | ✅ | ✅ | ❌ | ❌ | ✅ | ✅ | ✅ |
| W61 | 文字水印 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |

### 2.8 导航与引用

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W62 | 超链接（URL/邮箱） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| W63 | 书签（读写） | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W64 | 内部交叉引用 REF | ✅ | ✅ | ❌ | ❌ | ⚠️ | ✅ | ✅ |
| W65 | 目录（TOC 域） | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W66 | 脚注 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W67 | 尾注 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |

### 2.9 文档属性与保护

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W68 | 文档属性（标题/作者/主题） | ✅ | ✅ | ✅ | ❌ | ✅ | ✅ | ✅ |
| W69 | 自定义属性 | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W70 | 文档变量 | ✅ | ✅ | ❌ | ❌ | ✅ | ✅ | ✅ |
| W71 | 文档保护（只读） | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W72 | 文档加密（AES-128） | ✅ | ✅ | ⚠️ | ❌ | ✅ | ✅ | ✅ |
| W73 | 修订追踪（Track Changes） | ❌ | ✅ | ❌ | ❌ | ⚠️ | ✅ | ✅ |

### 2.10 高级内容

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W74 | 批注（Comment） | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W75 | 域代码（MERGEFIELD/PAGE/DATE） | ✅ | ✅ | ⚠️ | ❌ | ⚠️ | ✅ | ✅ |
| W76 | 内容控件 SDT（RichText/ComboBox/PlainText/Date/DropDownList） | ✅ | ✅ | ❌ | ❌ | ⚠️ | ✅ | ✅ |
| W77 | 自定义 XML 部件 | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ | ✅ |
| W78 | 公式（OMML，透传保留） | ⚠️ | ✅ | ❌ | ❌ | ❌ | ✅ | ✅ |
| W79 | 图表（Chart，透传保留） | ⚠️ | ✅ | ❌ | ❌ | ⚠️ | ✅ | ✅ |
| W80 | 嵌入对象（OLE，透传保留） | ⚠️ | ✅ | ⚠️ | ❌ | ❌ | ✅ | ✅ |
| W81 | SmartArt（透传保留） | ⚠️ | ✅ | ❌ | ❌ | ❌ | ✅ | ⚠️ |
| W82 | ActiveX 控件 | ❌ | ✅ | ❌ | ❌ | ❌ | ✅ | ⚠️ |

### 2.11 模板与数据绑定

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W83 | 占位符替换 `{{Key}}` | ✅ | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| W84 | 表格区域循环填充 | ✅ | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| W85 | 图片占位符替换 | ✅ | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| W86 | 嵌套 XML 拆分占位符合并 | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ | ❌ |
| W87 | 对象映射（WriteObjects/ReadObjects） | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| W88 | 邮件合并 MailMerge（单/多记录） | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |

### 2.12 格式转换

| 编号 | 功能 | NewLife.Office | Open XML SDK | NPOI | MiniWord | OfficeIMO | GemBox.Document | Aspose.Words |
|------|------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| W89 | docx → PDF | ✅ 内容映射 | ❌ | ❌ | ❌ | ❌ | ✅ Pro | ✅ 高保真 |
| W90 | docx → HTML | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ Pro | ✅ |
| W91 | docx → Markdown | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ |
| W92 | docx → 图片渲染 | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ Pro | ✅ |

---

## 3. 高保真读写能力对比

### 3.1 保真策略分级

| 级别 | 策略 | 代表库 | 说明 |
|------|------|--------|------|
| **L1 语义级** | 仅提取文本内容 | MiniWord | 忽略所有格式，只保留文字 |
| **L2 模型级** | 构建对象模型，覆盖主要属性 | NPOI、OfficeIMO | 模型覆盖 60-80% 常用属性 |
| **L3 透传级** | 模型 + RawXml 兜底 | **NewLife.Office**、GemBox | 已建模属性走模型，未建模透传原始 XML |
| **L4 引擎级** | 完整 DOM + 渲染引擎 | Aspose.Words | 商业级完整实现，理解全部 OOXML 语义 |

### 3.2 NewLife.Office 保真架构

```
docx 文件
    ↓ WordReader.ReadDocument()
Document (模型)
    ├── Elements[]          程序化访问（段落/表格/图片）
    │   ├── Paragraph   StyleId + Alignment + Indent + Spacing + Runs
    │   ├── Table       完整表格模型
    │   └── RawXml          每个元素保留原始 XML ★
    ├── Images[]             图片二进制
    ├── Hyperlinks[]         超链接关系
    ├── Headers[]            富文本页眉
    ├── Footers[]            富文本页脚
    ├── Comments[]           批注
    ├── StylesXml            样式表原始 XML ★
    ├── NumberingXml         编号定义原始 XML ★
    ├── SettingsXml          文档设置原始 XML ★
    ├── SectPrXml            节属性原始 XML ★
    ├── DocumentXmlNsDecls   命名空间声明 ★
    └── OtherParts[]         所有非 document.xml 部件 ★
    ↓ WordWriter.Save(doc)
docx 文件（视觉保真）
```

> **★ 标记的字段**是保真的关键——Reader 原样捕获，Writer 优先使用它们而非重新生成。这使得 NewLife.Office 在修改已有 docx 时能达到 L3 级保真。

### 3.3 往返保真测试场景

| 测试场景 | NewLife.Office | NPOI | Open XML SDK | Aspose.Words |
|----------|:---:|:---:|:---:|:---:|
| 纯文本段落往返 | ✅ 无损 | ✅ 无损 | ✅ 无损 | ✅ 无损 |
| 多字体/颜色混排 | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 复杂表格（合并单元格/边框） | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 段落边框/底纹 | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 页眉页脚（多类型） | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 多节不同页面设置 | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 脚注/尾注 | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 目录（TOC） | ✅ RawXml 保真 | 部分损失 | ✅ 无损 | ✅ 无损 |
| 内嵌公式（OMML） | ✅ 透传 | ❌ 丢失 | ✅ 无损 | ✅ 无损 |
| 内嵌图表 | ✅ 透传 | ❌ 丢失 | ✅ 无损 | ✅ 无损 |
| 修订追踪 | ❌ 丢失 | ❌ 丢失 | ✅ 无损 | ✅ 无损 |
| 内容控件 | ❌ 丢失 | ❌ 丢失 | ✅ 无损 | ✅ 无损 |
| 自定义 XML 部件 | ❌ 丢失 | ❌ 丢失 | ✅ 无损 | ✅ 无损 |

---

## 4. 非功能对比

### 4.1 依赖与体积

| 库 | 外部依赖 | 包体积 | 内存占用 | 启动开销 |
|---|---------|--------|---------|---------|
| **NewLife.Office** | 1（NewLife.Core） | ~200KB（Word 部分） | 低 | 极小 |
| Open XML SDK | 1（System.IO.Packaging） | ~2MB | 中 | 低 |
| NPOI | 5+（SharpZipLib 等） | ~10MB | 高 | 中 |
| MiniWord | 0 | ~300KB | 极低 | 极小 |
| OfficeIMO | 2（含 Open XML SDK） | ~1MB | 中 | 低 |
| Aspose.Words | 0 | ~40MB | 高 | 中 |
| GemBox.Document | 0 | ~5MB | 中 | 低 |

### 4.2 框架兼容性

| 库 | net45 | net461 | netstandard2.0 | net6.0+ | net8.0+ |
|---|:---:|:---:|:---:|:---:|:---:|
| **NewLife.Office** | ✅ | ✅ | ✅ | ✅ | ✅ |
| Open XML SDK | ❌ v3+ | ❌ v3+ | ✅ | ✅ | ✅ |
| NPOI | ✅ | ✅ | ✅ | ✅ | ✅ |
| MiniWord | ❌ | ❌ | ✅ | ✅ | ✅ |
| OfficeIMO | ❌ | ❌ | ✅ | ✅ | ✅ |
| Aspose.Words | ✅ | ✅ | ✅ | ✅ | ✅ |
| GemBox | ✅ | ✅ | ✅ | ✅ | ✅ |

### 4.3 许可证风险矩阵

| 库 | 许可类型 | 商业闭源使用 | SaaS 使用 | 修改源码 | 风险等级 |
|---|---------|:---:|:---:|:---:|:---:|
| **NewLife.Office** | MIT | ✅ 免费 | ✅ 免费 | ✅ 允许 | 🟢 无 |
| Open XML SDK | MIT | ✅ 免费 | ✅ 免费 | ✅ 允许 | 🟢 无 |
| NPOI | Apache 2.0 | ✅ 免费 | ✅ 免费 | ✅ 允许 | 🟢 低 |
| MiniWord | Apache 2.0 | ✅ 免费 | ✅ 免费 | ✅ 允许 | 🟢 低 |
| OfficeIMO | MIT | ✅ 免费 | ✅ 免费 | ✅ 允许 | 🟢 无 |
| Aspose.Words | 商业 | ❌ 需付费 | ❌ 需付费 | ❌ 禁止 | 🔴 高 |
| GemBox | 免费受限/商业 | ❌ 超页数需付费 | ❌ 超页数需付费 | ❌ 禁止 | 🟡 中 |
| Spire.Doc | 免费受限/商业 | ❌ 超500段需付费 | ❌ 超500段需付费 | ❌ 禁止 | 🟡 中 |

---

## 5. API 代码易用性对比

### 5.1 创建文档并写入内容

```csharp
// ✅ NewLife.Office — 语义化 API
using var writer = new WordWriter();
writer.AppendHeading("年度报告", 1);
writer.AppendParagraph("这是报告正文内容……");
writer.AppendTable(new[] { new[] { "项目", "金额" }, new[] { "收入", "100万" } }, firstRowHeader: true);
writer.Save("output.docx");
```

```csharp
// Open XML SDK — 底层冗长（约 40 行等效代码）
using var doc = WordprocessingDocument.Create("output.docx", WordprocessingDocumentType.Document);
var mainPart = doc.AddMainDocumentPart();
mainPart.Document = new Document();
var body = new Body();
var para = new Paragraph();
var run = new Run();
run.AppendChild(new Text("年度报告"));
para.AppendChild(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }));
para.AppendChild(run);
body.AppendChild(para);
// ... 表格创建约 20 行
mainPart.Document.AppendChild(body);
```

```csharp
// NPOI — Java 风格 API
var doc = new XWPFDocument();
var para = doc.CreateParagraph();
para.Style = "Heading1";
var run = para.CreateRun();
run.SetText("年度报告");
// 表格创建需要 XWPFTable 等
using var fs = new FileStream("output.docx", FileMode.Create);
doc.Write(fs);
```

```csharp
// Aspose.Words — 功能最全但需付费
var doc = new Document();
var builder = new DocumentBuilder(doc);
builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
builder.Write("年度报告");
doc.Save("output.docx");
```

### 5.2 模板填充与邮件合并

```csharp
// ✅ NewLife.Office — 占位符填充 + 多记录 MailMerge
var tpl = new WordTemplate("template.docx");
tpl.Fill("output.docx", new Dictionary<String, Object?>
{
    ["Name"] = "张三",
    ["Date"] = DateTime.Today,
    ["Amount"] = 9800m
});
// 多记录邮件合并（合同/通知批量生成）
tpl.MailMerge("output.docx", new[]
{
    new Dictionary<String, Object?> { ["Name"] = "张三", ["Amount"] = 100m },
    new Dictionary<String, Object?> { ["Name"] = "李四", ["Amount"] = 200m },
});
```

```csharp
// MiniWord — 类似 API，但仅模板填充、无多记录合并
MiniWord.SaveAsByTemplate("output.docx", "template.docx",
    new { Name = "张三", Date = DateTime.Today });
```

```csharp
// Aspose.Words — Mail Merge 引擎
var doc = new Document("template.docx");
doc.MailMerge.Execute(new[] { "Name" }, new[] { "张三" });
doc.Save("output.docx");
```

### 5.3 读取与修改（往返编辑）

```csharp
// ✅ NewLife.Office — 模型级读写往返
using var reader = new WordReader("source.docx");
var doc = reader.ReadDocument();
doc.DocumentProperties.Title = "新标题";
// 在开头插入新段落
doc.Elements.Insert(0, new Element
{
    Type = ElementType.Paragraph,
    Paragraph = new Paragraph
    {
        Style = ParagraphStyle.Heading1,
        Runs = { new Run { Text = "新增章节" } }
    }
});
using var writer = new WordWriter();
writer.Save("modified.docx", doc);
```

```csharp
// Open XML SDK — 直接操作 XML DOM
using var doc = WordprocessingDocument.Open("source.docx", true);
var body = doc.MainDocumentPart.Document.Body;
var para = new Paragraph();
var run = new Run(new Text("新增章节"));
para.AppendChild(new ParagraphProperties(
    new ParagraphStyleId { Val = "Heading1" }));
para.AppendChild(run);
body.InsertBefore(para, body.FirstChild);
// 手动处理其他部件……
```

```csharp
// NPOI — 部分模型，无完整往返
var doc = new XWPFDocument(File.OpenRead("source.docx"));
// 只能操作有限属性，很多格式信息会在 re-save 时丢失
```

---

## 6. 各竞品深度分析

### 6.1 Open XML SDK

**定位**：微软官方 OOXML 底层 SDK，所有 docx 操作的基础层。

| 维度 | 评价 |
|------|------|
| 优势 | 完整 OOXML 规范支持；类型安全；长期维护；MIT 许可 |
| 劣势 | API 极端底层；简单操作需大量代码；无文档模型；无便捷 API |
| 适用场景 | 需要精确控制 XML 结构；作为其他库的底层依赖 |
| 保真度 | L2+（直接操作 XML，可达 L4 但需手写所有细节） |

**与 NewLife.Office 关系**：互补。NewLife.Office 在其之上提供高层模型和便捷 API，同时保持通过 RawXml 直接操作底层 XML 的能力。

### 6.2 NPOI

**定位**：Java POI 的 .NET 移植，老牌开源办公文档库。

| 维度 | 评价 |
|------|------|
| 优势 | 唯一同时免费支持 doc/docx 的开源库；社区成熟；Apache 2.0 |
| 劣势 | 包体积大（10MB）；API 非 .NET 原生风格；内存占用高；无完整模型往返 |
| 保真度 | L2（模型级，部分属性丢失） |
| 适用场景 | 需要读取 doc 格式且不能使用商业库；遗留系统维护 |

**与 NewLife.Office 对比**：NewLife.Office 在 docx 读写保真度、API 简洁性、体积上均优于 NPOI；NPOI 唯一优势是 doc 写入（NewLife.Office 仅 doc 读取）。

### 6.3 MiniWord

**定位**：极致轻量的 docx 模板填充工具。

| 维度 | 评价 |
|------|------|
| 优势 | 体积极小；API 简洁；模板填充开箱即用 |
| 劣势 | 仅支持模板写入；不支持读取；无任何格式模型；依赖 Open XML SDK |
| 保真度 | L1（纯文本替换） |
| 适用场景 | 简单的模板填充（合同、通知、证书等） |

**与 NewLife.Office 对比**：NewLife.Office 的 `WordTemplate` 提供相同甚至更丰富的模板填充能力（含表格循环、图片替换、对象映射、XML 拆分占位符合并、多记录 MailMerge），同时具备完整读写模型。MiniWord 仅在"只需要模板填充且不需要任何其他功能"的场景有体积优势。

### 6.4 OfficeIMO

**定位**：Open XML SDK 的现代化高层封装。

| 维度 | 评价 |
|------|------|
| 优势 | API 比 Open XML SDK 友好；MIT 许可；代码可读性好 |
| 劣势 | 项目年轻，社区小；依赖 Open XML SDK；功能覆盖不完整 |
| 保真度 | L2（模型级） |
| 适用场景 | 不想直接用 Open XML SDK 的简单 docx 操作 |

**与 NewLife.Office 对比**：NewLife.Office 功能覆盖远超 OfficeIMO（模板、doc 读取、PDF/HTML/Markdown 转换、透传保真），且无 Open XML SDK 依赖。

### 6.5 GemBox.Document

**定位**：商业级 Word 库的"免费试用"入口。

| 维度 | 评价 |
|------|------|
| 优势 | API 设计优秀；功能全面；文档丰富；性能好 |
| 劣势 | 免费版限制页数/段落数；商业许可价格不菲；闭源 |
| 保真度 | L3（模型+透传） |
| 适用场景 | 商业项目预算充足时的一站式方案 |

**与 NewLife.Office 对比**：NewLife.Office 是 MIT 下唯一能达到 GemBox 同类功能密度的方案。GemBox 在 PDF 渲染保真度上有优势。

### 6.6 Aspose.Words

**定位**：.NET 生态 Word 文档操作的"金标准"。

| 维度 | 评价 |
|------|------|
| 优势 | 功能最全；保真度最高；内置渲染引擎；格式互转最强 |
| 劣势 | 商业许可价格高（$999+/开发者/年）；闭源；包体积大 |
| 保真度 | L4（引擎级，完整 DOM + 渲染） |
| 适用场景 | 对保真度有极致要求且预算充足的商业项目 |

**与 NewLife.Office 对比**：Aspose.Words 是终极方案，NewLife.Office 是 MIT 下的最优方案。两者剩余差距主要在：PDF 渲染保真度（内容映射 vs 像素级）、修订追踪模型级操作、doc 写入、docx→图片渲染。NewLife.Office 在 API 简洁性、零成本与轻量性上胜出。

---

## 7. 差异化竞争力总结

### 7.1 NewLife.Office Word 模块核心优势

| 优势 | 说明 |
|------|------|
| **MIT 全免费** | 无任何商用限制，对比 GemBox/Spire 免费版有页数/段落数限制 |
| **文档模型往返** | `WordReader.ReadDocument()` → 修改 → `WordWriter.Save(doc)`，唯一 MIT 下支持此模式 |
| **L3 双兜底保真** | RawXml 逐元素透传 + DocumentXml 全文透传，编辑已有文档不失视觉效果 |
| **ZIP 部件透传** | 所有非 document.xml 部件原样保留（主题/字体/设置/脚注/尾注等） |
| **模板与数据绑定** | 占位符/表格区域/图片/对象映射/XML 拆分合并 + 多记录 MailMerge，超越 MiniWord |
| **格式转换** | docx→PDF、docx→HTML、docx→Markdown，均为 MIT 免费 |
| **doc 格式读取** | 自研 OLE2/CFB + MS-DOC 解析器，零依赖读取旧版 .doc |
| **极致轻量** | Word 模块 <200KB，仅依赖 NewLife.Core |
| **全框架兼容** | net45 → net9.0+，无框架版本盲区 |

**功能覆盖率**：对比矩阵 92 项中，NewLife.Office 完整支持 77 项、RawXml 透传兜底 10 项，可用率 **94.6%**（87/92），在 MIT/Apache 开源库中排名第一；仅 5 项不支持（见 7.2）。

### 7.2 当前差距（真实未实现项）

| 差距 | 说明 | 影响 |
|------|------|------|
| 修订追踪（Track Changes） | 仅可通过 DocumentXml 全文透传保留现有修订，无模型级读取/接受/拒绝 | 🟡 中 |
| doc 写入 | 仅支持 doc 读取（OLE2 纯文本 + 段落结构），不支持写回 .doc；**已补 doc→docx 转换**（`DocReader.ToDocument/SaveAsDocx`，段落/表格/格式映射），覆盖历史文档迁移主场景 | 🟢 低 |
| doc 图片/页眉页脚提取 | MS-DOC DrawingGroup/PICF 与 STSHI/Plcfhdd 二进制解析复杂度高，暂缓 | 🟢 低 |
| docx → 图片渲染 | PDF 转换为内容映射型，非像素级渲染引擎，无法渲染成图片 | 🟡 中 |
| 公式/图表/SmartArt 程序化创建 | 现有公式/图表/SmartArt 通过部件透传完整保留，但不支持从零创建 | 🟢 低 |
| 自选图形/ActiveX | 不支持形状对象创建与 ActiveX 控件 | 🟢 低 |

> **策略**：上述差距均为设计取舍——保持 MIT 纯净性、零外部依赖与轻量化是更高优先级。PDF 像素级渲染可通过可选的外部工具（LibreOffice headless）集成提供，不引入核心库重依赖；修订追踪模型、doc 写入等将按用户需求优先级排期。doc 图片/页眉页脚提取在 doc→docx 转换覆盖主场景后收益有限，暂缓。

---

（完）
