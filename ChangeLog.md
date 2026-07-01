# NewLife.Office 版本更新记录

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
