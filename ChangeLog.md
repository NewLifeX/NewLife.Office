# NewLife.Office 版本更新记录

## v1.5.2026.1024 (2026-10-24)

### 多语言本地化（31 种语言）
- **运行时词表**：接入 NewLife.Core 12 本地化体系，异常消息统一英文源文，词表由 11 种语言扩展至 31 种，共 130 条已登记词条（62 条运行时消息 + 64 条文档输出标签 + 4 条程序集元数据）
- **文档输出标签**：文本提取、模板默认值、EPUB 等文档输出标签按当前区域翻译，导出内容随语言变化
- **程序集元数据**：AssemblyTitle/Description 改英文源文并接入词表，主包与渲染扩展启用 PackageReadmeFile（NuGet 包内附带英文 README）
- **质量门禁**：新增 tools/verify-localization.ps1 一键校验（词表双向覆盖、英文源文守卫、待翻译清单、目标书写系统检查、README 完整性），支持 -FailOnPending 发布门禁

### NativeAOT 与多平台
- **AOT 兼容改造**：主库目标框架扩展为 net45/netstandard2.0/netstandard2.1/net8.0/net10.0，net8.0+ 开启 IsAotCompatible，泛型反射入口补齐 DynamicallyAccessedMembers 标注，IL 告警清零
- **AOT 冒烟**：新增 AotSmoke 工程，23 个场景覆盖 Markdown/PDF/Excel/Word/PPT/RTF/ODS/EML/ICS/vCard/EPUB/XPS/OLE2 读写与 31 种语言表
- **CI 矩阵**：win-x64/win-x86/win-arm64、linux-x64/linux-arm64/linux-loongarch64（QEMU）、osx-x64/osx-arm64；新增 riscv64/s390x 指令集冒烟与 Android/iOS 移动端冒烟工程

### 排版能力增强
- **Word**：支持注入 styles.xml 样式表并美化默认排版；表格默认单元格内边距与垂直居中，避免文字贴边框
- **Excel**：支持设置工作簿默认字体；写出器识别货币与千分位数值并保留显示格式
- **PPT**：支持主题字体引用与图表主题色

### 兼容性修复
- **PDF**：补齐流对象与对象流（含 Flate 解压）解析、嵌套页树遍历、十六进制字符串 UTF-16 BOM/小端识别；元数据超出 WinAnsi 表示范围时按 UTF-16BE 写出，CJK 标题不再丢失；修复 Tokenize 遇未识别字节死循环
- **Markdown**：列表项子块裁剪结尾空行，紧凑列表序列化不动点成立
- **PPT/Word/Excel 回读**：回读连接器/动画/批注/文档属性/页眉页脚/节/版式/母版并深度对比，修复动画时长单位与目标序号偏移；补齐编号定义、富文本段落与透视表字段回读
- **Word**：修复 KPI 行之后正文未排版

### 工程与测试
- **测试提速**：xunit 并行度收敛至 8 线程、往返用例受限并发、逐字节断言改用 SequenceEqual，CI 增加 NuGet 缓存与挂起超时保护
- **工程**：解决方案迁移为 slnx 格式；基准项目改用项目引用，兼容 BenchmarkDotNet 确定性构建
- **依赖**：NewLife.Core 升级至 12.0.2026.1024（正式版），NewLife.UnitTest 升级至 1.1.2026.301-beta0617

---

## v1.4.2026.0902 (2026-09-02)

### Word 增强
- **读取保真修复**：RunProperties 三态化（支持 w:val=0 显式关闭）、ParseRun 多 w:t/tab/br/cr 拼接、列表 numId→numFmt 映射判定、新增高亮/小型大写/全大写/隐藏/东亚字体/语言属性、样式继承 docDefaults+basedOn 链、空段落保留
- **图片与页眉页脚**：SVG/VML/浮动 anchor 图片解析与写回、页眉页脚富文本模型（Reader 三类型填充 + Writer 部件生成）
- **多节支持**：Section 模型 + Reader 切分 + Writer 程序化多节输出
- **日常编辑 API**：FindText/ReplaceText 查找替换（跨 Run 匹配保留首 Run 格式）、Table 行列增删操作、单元格便利 API（GetCell/SetCellText）
- **转换器增强**：WordHtmlConverter 列表分组/图片/删除线/上下标/合并单元格，WordPdfConverter 列表前缀/字体/分页符/段落内图片
- **doc 二进制格式提取**：ReadParagraphsWithFormat 基于 PlcBte+CHPX/PAPX 解析段落格式
- **脚注/尾注**：WordWriter 程序化创建 + Find/Replace 覆盖 + Markdown 转换、脚注占位符合并替换
- **文本框与修订**：文本框模型化与 Find/Replace 覆盖、修订追踪模型级读取
- **doc→docx 转换**：DocReader.ToDocument/SaveAsDocx（段落+表格+格式，保持文档顺序）+ CLI 子命令
- **自选图形/公式/图表**：自选图形创建、公式与图表创建
- **写回保真优化**：页眉页脚显式重建、模板占位符合并、嵌套表格支持、ClearRawXml 解决读改写时 RawXml 兜底覆盖修改的痛点

### Excel 增强
- **xls 全面增强**：样式读取（FONT/FORMAT/XF）、批注读取（NOTE+TXO）、命名范围/打印区域写入、保护/数据验证、BiffReader 合并单元格/页面设置
- **条件格式 dxf**：字体样式/填充色/边框颜色读写 + AddConditionalFormat 新增 fontColor/bold/borderColor 重载、图案前景色调色板还原
- **图表增强**：doughnut/radar/stock/surface 类型识别与写入、图表读写往返修复
- **批注增强**：Comment 模型（尺寸/可见性/vml 解析）、富文本增强、线程化批注
- **切片器**：切片器 + 透视表切片器关联
- **流式读取性能**：ReadRows/ReadRowsWithIndex XmlReader 流式解析（10 万行 995ms，内存平坦）
- **显示与打印**：SetGridlines/SetZoomScale/SetDefaultRowHeight、打印区域/分页符读写往返、隐藏行列写入
- **读取健壮性**：1904 日期系统、t=e/t=d 单元格类型、富文本共享字符串、主题色解析、缺失部件容错

### Markdown 引擎优化
- **CommonMark 规范补全**：规范级完善、边界修复（setext 单横线/反引号等长分隔/表格管线/嵌套链接/紧松列表/编码检测）
- **解析性能大幅优化**：MD15-MD26 多轮优化（引用块快速路径、AST 延迟初始化、行内令牌复用、DelimInfo 改 struct、快照序列化池化等），Parse 耗时与分配显著下降
- **转换中枢扩展**：ODS/RTF/EPUB→Markdown 一键转换
- **读写完善**：写入转义/列表续行/引用链接/HTML 块/管线开关

### PPT 增强
- **形状往返保真**：组内图片读取/写入、组内形状属性（旋转/翻转/渐变/虚线/线宽/圆角）读写、CornerRadius 精度修复
- **图表增强**：AddAreaChart 公开方法、图表系列颜色/散点 X 值/轴范围/图例位置读回
- **超链接**：Shape 级超链接（URL/文件跳转）+ 组内形状超链接读写
- **PptWriter**：最小可用写入 + Merge/CopySlideFrom 保真修复 + 批注关系/图表命名空间修复
- **竞品功能 S17-S20**：对标 Aspose.Slides/ShapeCrawler

### PDF 增强
- **结构化表格提取**：对标 iText 7/Aspose.PDF
- **嵌入字体子集化**：对标 PdfSharp
- **文档属性测试修复**：S14 审计修复

### CLI 工具链
- **内容识别**：OfficeFactory magic bytes 识别文件格式
- **转换工具链**：命令行转换工具，doc→docx 子命令

### 测试与质量
- **Rendering 测试补齐**：RenderingTests 21 项覆盖 REN01-13 + 修复 3 个生产缺陷
- **交叉兼容测试**：NPOI + OpenXml SDK 竞品包交叉验证、真实文件 Fixtures 回归、合并单元格交叉验证

---

## v1.3.2026.0802 (2026-08-02)

### 新增渲染扩展包
- **NewLife.Office.Rendering**：全新渲染扩展包（首个版本），基于 SkiaSharp 跨平台 2D 渲染，PDF/Word/PPT/Markdown 文档页面渲染为图片，支持图片缩放/裁剪/旋转/水印/格式转换、文字→图片、二维码→图片、缩略图网格
- **DocumentPreview**：文档预览渲染入口，统一调度各格式页面渲染

### Markdown 引擎全面升级
- **MarkdownPipeline 管线**：新增 Markdown 处理管线，支持扩展挂载与阶段化处理
- **全格式转换中枢 FormatToMarkdown**：Excel/Word/PPT/PDF 等任意支持格式一键转换为 Markdown
- **HTML→Markdown 反向转换**：`HtmlToMarkdownConverter` + 配置选项，31 项测试
- **Word→Markdown 增强**：格式保留、GFM 表格、图片 data URI 内联
- **高级解析扩展**：AutoLinks 自动链接、YAML FrontMatter、脚注 Footnotes、数学公式 Math
- **Emoji 短码**：`:smile:`→😄，内置 50+ 常用映射
- **自定义属性**：`{.class #id}` 解析 + HTML class/id 渲染
- **定义列表**：DefinitionList/Term/Description 解析与 HTML 渲染
- **缩写支持**：`*[ABBR]:` 解析 + HTML abbr 渲染
- **AST 强类型化**：Markdown AST 升级为强类型子类体系，源码位置追踪

### 模型命名空间重构
- **各格式模型规范化**：Excel/Word/PPT/PDF/邮件/iCal/Epub 类型命名与命名空间全面统一，iCal/Epub 归类至专属子命名空间，提升 API 可发现性

### 测试与质量
- **模型属性断言补全**：PDF D5-D8、Excel E13-E19、Word W7-W14、PPT P7-P12 模型属性断言补全
- **往返测试加固**：PPT/Word/Excel/PDF 往返测试增强，修复 ParseFillColor XPath 解析 bug

---

## v1.2.2026.0702 (2026-07-02)

### Excel 增强
- **xls BiffWriter 全面补全（M25）**：样式（字体/填充/边框/颜色/数字格式）、合并单元格、冻结窗格、行高、超链接读写、自动筛选、页面设置、页眉页脚、默认列宽、公式写入——xls 与 xlsx 功能平齐
- **xls BiffReader 增强**：COLINFO(0x007D) 列宽解析、HYPERLINK 记录解析
- **xlsx 高保真补齐（M18-M24）**：命名范围（Defined Names）、结构化表格（Table）、富文本与高级样式（渐变填充/图案填充/四边独立边框/字体增强）、图表集成（Writer/Reader）、条件格式增强（图标集/自定义公式）、布局增强（文本旋转/缩进/缩小填充/行列分组）、企业特性（标签颜色/工作簿保护/计算选项）
- **xlsx Reader 端高保真补齐**：完整快照与无损还原
- **迷你图 Sparklines**：`AddSparklineGroup` + `ReadSparklines` 读写支持
- **数据透视表增强**：`AddFilterField` 筛选字段
- **文档属性**：`DocumentProperties` core.xml 读写往返
- **分页符与打印区域**：`SetPageBreak` + `SetPrintArea` + rowBreaks/colBreaks
- **对角线边框**：单元格对角线边框样式支持

### Word 增强
- **文档加密 WordEncryptor**：AES-128 加密 docx（ECMA-376 Agile Encryption，OLE2 CFB 容器），加密/解密双向
- **邮件合并引擎**：`AppendMergeField` + `WordTemplate.MailMerge` 单/多记录合并
- **高保真读写往返**：L3 级保真架构（模型 + RawXml 透传），ZIP 部件完整透传
- **多级列表模型 WordNumbering**：Writer 生成 + Reader 解析 + 多级嵌套列表
- **表格合并**：gridSpan/vMerge 写入 + Reader 回读
- **段落分页控制**：keepNext/keepLines/widowControl 孤行控制
- **页面设置增强**：TitlePage + EvenAndOddHeaders 首页/奇偶页不同页眉页脚
- **文档变量 DocumentVariables**：读写 settings.xml w:docVars 企业模板元数据
- **自定义 XML 部件 CustomXmlParts**：读写 customXml/item*.xml
- **内容控件 SDT**：RichText/ComboBox/PlainText/Date/DropDownList 读写
- **内部交叉引用 REF 域**：`AppendCrossRef` 引用书签显示页码
- **文字发光/阴影效果**：GlowColor/ShadowColor + w14:glow/w14:shadow XML
- **有序列表增强**：OOXML 编号 + ListStartOverride 从任意值开始
- **首字下沉 DropCap**：framePr w:dropCap 模型 + 读写往返
- **行号 LineNumbering**：WordLineNumberSettings 模型 + 读写往返
- **页面边框 PageBorder**：pgBorders 读写
- **自定义文档属性 CustomProperties**：读写
- **分栏 Columns**：WordPageSettings ColumnCount/ColumnSpacing
- **水平分隔线**：`AppendHorizontalRule` 段落底部边框样式分隔符
- **对象映射增强**：嵌套对象属性展开，`WriteObjects` maxDepth 参数

### PPT 增强
- **图表体系**：雷达图、股价图、曲面图、散点图 c:xVal 数值 X 轴
- **形状操作增强（S16）**：Z-Order 图层控制、取消组合、渐变填充、翻转（FlipH/FlipV）、文本方向/自动适应/内边距/虚线样式、Anchor + Clone 补全 14 属性、Alt Text、圆角矩形
- **图片增强**：图片旋转、SVG 读写往返（asvg:svgBlip）、图片圆角（CornerRadius）、形状图片填充（FillImage）、图片原地替换
- **表格增强**：动态添加/删除行列、上标/下标、渐变背景
- **Section 管理**：PptSection 模型读写，按节组织幻灯片
- **动画增强**：Morph 切换类型 + 元素动画写入
- **形状克隆/重复**：DuplicateShape + PptShape.Clone 模型级深拷贝
- **批注/连接器/文档属性**：comments + connectors + docProps 完整读写

### PDF 增强
- **数字签名 PdfSigner**：PKCS#7 分离签名 + 可见签名域，手动 ASN.1 DER 构建，零外部依赖
- **FluentDocument 绘图原语完整透传**：DrawEllipse/RoundedRect/Arc/Bezier/Polygon/Polyline、渐变矩形、虚线样式、透明度控制、图片旋转、QR 码
- **JPEG DCTDecode 直通**：JPEG 流不解码直接嵌入，保持原始品质
- **自定义页码格式**：`{page}/{total}` 占位符 + 字符/词间距 Tc/Tw
- **注释系统**：ReadAnnotations 读取 + PdfWriter.AddAnnotation 通用写入接口（全部注释类型）
- **PdfTable 模型**：DrawTable 集成到 PdfWriter
- **表格提取 ExtractTables**：位置聚类启发式算法
- **PDF/A 合规**：PDF/A-1B/2B/3B 输出（XMP 元数据 + OutputIntent）
- **AES 加密**：AES-128/256 加密（PdfCipherRevision 枚举 + PdfEncryptor 策略模式）
- **高保真解析架构**：xref 交叉引用表 + PdfObjectParser + PdfContentStream 内容流重建
- **压缩解码**：LZWDecode + RunLengthDecode + ASCII85/Hex 全过滤器链
- **嵌入附件**：EmbedFile + EmbeddedFiles 名树 + GetAttachments 读取
- **字体信息读取**：ReadFonts + PdfFontInfo 模型
- **QR 码生成**：PdfQRCode + DrawQRCode/AppendQRCode，纯 C# 零依赖

### 工程与质量
- **测试**：测试集从 415 项增长至 694 项（+279），统一命名空间结构与临时文件生成方式
- **文档**：需求文档/功能模块清单/四大竞品分析全面审计同步，修正 30+ 处过时标记
- **csproj 元数据**：修正 AssemblyTitle/Description/PackageProjectUrl
- **README**：新增英文版 README_EN.MD，中英文互链
- **代码质量**：修正空引用警告，统一字节读写方式，优化二进制文档解析

---

## v1.1.2026.0501 (2026-05-01)

### 统一文本提取接口
- **ITextExtractable**：新增统一文本提取接口，所有支持格式均可通过 `ExtractText()` 获取纯文本
- **IMarkdownExtractable**：新增 Markdown 提取接口，结构化格式可通过 `ExtractMarkdown()` 导出带格式的 Markdown
- **全格式覆盖**：Excel、Word、PPT、PDF、RTF、ODS、EML、iCalendar、vCard、EPUB、XPS、Markdown 均已实现两个接口

### OfficeFactory 增强
- **ReadText()**：新增 `OfficeFactory.ReadText(filePath)` / `ReadText(stream, ext)` 静态方法，一行代码提取任意格式纯文本
- **ReadMarkdown()**：新增 `OfficeFactory.ReadMarkdown(filePath)` / `ReadMarkdown(stream, ext)` 静态方法，一行代码获取 Markdown 内容

---

## v1.0.2026.0403 (2026-04-03)

### 项目说明
- **NewLife.Office**：一个功能全面的 .NET 办公自动化库，覆盖 Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 格式，零外部依赖，MIT 许可
- **PDF 字体增强**：增强 PDF 字体支持，优化文本提取与测试覆盖
- **程序集标题与描述**：更新程序集标题和描述，完善功能说明

---
