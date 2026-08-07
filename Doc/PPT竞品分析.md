# PPT 竞品分析

> 版本：v2.0 | 日期：2026-08-06
> 返回：[竞品分析报告](竞品分析报告.md)

---

## 1. PPT 竞品概览

| 竞品 | 类型 | 许可 | 定位 |
|------|------|------|------|
| **本项目（NewLife.Office）** | 开源库（v1.3.x） | MIT | 零外部依赖全能型：pptx 读写（L2 保真）+ 模板填充 + PDF 转换 + 自动排版，含 ppt 老格式读取 |
| **Open XML SDK** | 开源 SDK（底层） | MIT | OOXML 规范底层访问，微软官方维护，作为上层库的地基，开发效率低 |
| **NPOI** | 开源库 | Apache 2.0（v2.8+ 商用受限 ⚠️） | Java POI 移植，以 Excel/Word 为核心；PPT 模块（HSLF/XSLF）位于 scratchpad 草稿区，成熟度不足 |
| **ShapeCrawler** | 开源库 | MIT | 依赖 Open XML SDK，API 友好，pptx 读取为主，无写入能力（只读为主），功能持续完善中 |
| **Aspose.Slides** | 商业库 | 商业（约数千美元/年） | 功能最全，唯一支持高保真渲染/完整动画/SmartArt/导出视频，包体积约 40 MB |
| **Spire.Presentation** | 商业库（含免费版） | 免费受限（10 页）/ 商业 | API 风格类似 Aspose.Slides，免费版限制严格，不推荐生产选型 |

> ⚠️ **NPOI 许可证变化（2025 年起）**：NPOI v2.8.0 引入 OSMFEULA，要求营利性组织支付月度维护费，不再完全免费商用。且 NPOI 的 PPT 模块（HSLF/XSLF）长期位于 `scratchpad/` 草稿区，功能远未完成，非其核心竞争力。

### 1.1 PPT 模块成熟度评级

| 库名 | PPT 模块成熟度 | 核心优势 |
|------|:---:|------|
| **Aspose.Slides** | ⭐⭐⭐⭐⭐ | 功能最全，唯一支持完整渲染/动画/SmartArt/导出视频 |
| **NewLife.Office** | ⭐⭐⭐⭐ | 零依赖/MIT，pptx 读写完整（L2 保真），含 ppt 读取/PDF 转换/自动排版/模板填充 |
| **ShapeCrawler** | ⭐⭐⭐⭐ | API 最友好，pptx 读取完善，v0.79+ 稳定，写入能力有限（只读为主） |
| **Open XML SDK** | ⭐⭐⭐ | 底层完全访问，需大量样板代码，开发效率低 |
| **Spire.Presentation Free** | ⭐⭐⭐ | 免费版限 10 页，商业版功能全但价格高 |
| **NPOI** | ⭐⭐ | PPT 模块残缺，不建议选型 |

---

## 2. 功能对比矩阵

> 标记说明：✅ 完整支持 | ⚠️ 部分可用 | ❌ 不支持。NewLife.Office 列基于 v1.3.x 当前能力实测标注；ShapeCrawler 以只读为主，写入/创建类能力整体受限。

### 2.1 文件格式与读写

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 读取 pptx | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 创建/写入 pptx | ✅ | ✅ | ⚠️ | ⚠️ | ✅ |
| 读取 ppt（97-2003） | ✅ | ❌ | ✅（HSLF） | ❌ | ✅ |
| 写入 ppt（97-2003） | ❌ | ❌ | ⚠️ | ❌ | ✅ |
| PPT → PPTX 转换 | ❌ | ❌ | ❌ | ❌ | ✅ |
| 读取 ODP（OpenDocument） | ❌ | ❌ | ❌ | ❌ | ✅ |
| 写入 ODP | ❌ | ❌ | ❌ | ❌ | ✅ |
| Stream 读写接口 | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 模板填充（{{Key}}） | ✅ | ❌ | ❌ | ⚠️ | ✅ |
| 零外部依赖 | ✅ | ❌ | ❌ | ❌ | ✅ |

### 2.2 幻灯片操作

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 添加幻灯片 | ✅ | ✅ | ⚠️ | ⚠️ | ✅ |
| 删除幻灯片 | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| 复制/移动幻灯片（同文件） | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| 跨文件复制幻灯片（含媒体自动复制） | ✅ | ⚠️（手动） | ❌ | ❌ | ✅（自动复制母版） |
| 合并演示文稿（Merge） | ✅ | ⚠️（手动） | ❌ | ❌ | ✅ |
| 隐藏/显示幻灯片 | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| Section 分组管理 | ✅ | ✅ | ❌ | ⚠️（v0.30） | ✅ |
| 幻灯片大小/比例 | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| 演讲者备注 | ✅ | ✅ | ❌ | ⚠️（v0.53） | ✅ |
| 页脚/幻灯片编号 | ✅ | ✅ | ❌ | ⚠️（v0.47） | ✅ |
| 全局页眉页脚设置 | ✅ | ✅ | ❌ | ⚠️ | ✅ |

### 2.3 形状与内容

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 文本框 | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 自选形状（30+ 种预置） | ✅ | ✅ | ⚠️ | ✅（20+ 种） | ✅ |
| 形状精确位置（EMU） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 形状旋转角度 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 形状克隆/重复 | ✅ | ⚠️（手动） | ❌ | ⚠️（Duplicate） | ✅ |
| 形状 Alt Text | ✅ | ✅ | ❌ | ✅（v0.57） | ✅ |
| 组形状（创建/操作） | ✅ | ✅ | ❌ | ⚠️（v0.67） | ✅ |
| 连接器/线条 | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| 占位符操作（Role） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 渐变/图案填充 | ✅ | ✅ | ❌ | ✅（IShapeFill） | ✅ |
| 翻转（FlipH/FlipV）/Z-Order/取消组合 | ✅ | ✅ | ❌ | ❌ | ✅ |
| SmartArt | ❌ | ✅（底层） | ❌ | ❌ | ✅ |
| OLE 嵌入对象 | ❌ | ✅ | ✅ | ❌ | ✅ |
| 数学公式（MathML） | ❌ | ✅ | ❌ | ❌ | ✅ |

### 2.4 文本格式（Run/Paragraph 级）

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| Run 级字体（Latin/EastAsian/ComplexScript/Symbol） | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 字体颜色（RGB） | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 字体颜色（主题色 schemeClr） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 东亚/西文字体独立设置 | ✅ | ✅ | ❌ | ✅（v0.41） | ✅ |
| 下划线 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 渐变文字填充 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 上标/下标 | ✅ | ✅ | ❌ | ✅（OffsetEffect） | ✅ |
| 超链接（Run 级） | ✅ | ✅ | ❌ | ✅（v0.29） | ✅ |
| 段落对齐 | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 行距（百分比/磅值）/段前后间距 | ✅ | ✅ | ❌ | ✅（v0.62） | ✅ |
| 项目符号 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 缩进级别（多级大纲 0-8） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 文本方向（垂直/旋转） | ✅ | ✅ | ❌ | ✅（v0.71） | ✅ |
| 文本内边距 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 自动适应（AutoFit） | ✅ | ✅ | ❌ | ✅ | ✅ |

### 2.5 表格

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 创建表格 | ✅ | ✅ | ❌ | ⚠️ | ✅ |
| 单元格文本/富文本/独立样式 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 单元格背景色 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 单元格字体格式 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 单元格边框样式（四边独立） | ✅ | ✅ | ❌ | ✅（v0.54+） | ✅ |
| 合并单元格（gridSpan/rowSpan/vMerge 读写双向） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 列宽控制 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 动态添加/删除行列 | ✅ | ✅ | ❌ | ✅（v0.57） | ✅ |
| 表格主题样式引用 | ✅ | ✅ | ❌ | ✅（v0.59） | ✅ |
| 首行表头样式 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 读取嵌套表格内容 | ✅ | ✅ | ❌ | ✅ | ✅ |

### 2.6 图表

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 柱状图/条形图 | ✅ | ✅ | ❌ | ✅（v0.67） | ✅ |
| 折线图 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 饼图/环形图 | ✅ | ✅ | ❌ | ✅（v0.64） | ✅ |
| 面积图 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 散点图/气泡图 | ✅ | ✅ | ❌ | ✅（v0.78） | ✅ |
| 雷达图 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 股价图/曲面图 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 图表标题 | ✅ | ✅ | ❌ | ✅（v0.73） | ✅ |
| 多系列数据读写 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 坐标轴 Min/Max | ✅ | ✅ | ❌ | ✅（v0.45） | ✅ |
| 散点图数值 X 轴（c:xVal） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 图表数据修改（UpdateChartData） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 图表样式/配色 | ❌ | ✅ | ❌ | ❌ | ✅ |

### 2.7 图片与媒体

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 插入图片（PNG/JPEG） | ✅ | ✅ | ✅ | ⚠️ | ✅ |
| 图片批量提取 | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 图片原地替换（ReplaceImage） | ✅ | ✅ | ❌ | ✅（Image.Update） | ✅ |
| SVG 图片读写（asvg:svgBlip） | ✅ | ✅ | ❌ | ✅（v0.52） | ✅ |
| 图片圆角 | ✅ | ✅ | ❌ | ✅（CornerSize） | ✅ |
| 嵌入视频（mp4/wmv，含缩略图） | ✅ | ✅ | ✅（Movie） | ✅（v0.25） | ✅ |
| 嵌入音频（mp3/wav） | ✅ | ✅ | ❌ | ✅（v0.24） | ✅ |
| 背景图片 | ✅ | ✅ | ⚠️ | ✅ | ✅ |
| 背景渐变填充 | ✅ | ✅ | ❌ | ✅（IShapeFill） | ✅ |
| 形状图片填充 | ❌ | ✅ | ❌ | ✅ | ✅ |

### 2.8 母版、版式与主题

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 从模板加载完整母版/版式/主题（L2 保真透传） | ✅ | ⚠️（手动） | ❌ | ⚠️ | ✅ |
| 编程式创建母版（CreateMaster） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 跨文件复制母版 | ✅ | ⚠️（手动） | ❌ | ❌ | ✅（自动） |
| 版式索引控制 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 读取/枚举版式详情 | ✅ | ✅ | ❌ | ✅ | ✅ |
| 主题色（6 强调色）编程修改 | ✅ | ✅ | ❌ | ✅（ITheme，v0.40） | ✅ |
| 嵌入字体加载/透传（字节一致） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 原始 XML 访问 | ✅ | ✅ | ❌ | ⚠️ | ✅ |

### 2.9 动画与切换

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 幻灯片切换效果（fade/push/wipe 等） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 切换时长/方向控制 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 形态转换（Morph Transition） | ✅ | ✅（底层） | ❌ | ❌ | ✅（唯一高层 API） |
| 元素进入/退出动画 | ✅ | ✅ | ❌ | ❌ | ✅（150+ 效果） |
| 元素强调动画 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 自定义动画路径 | ⚠️（MotionPath 模型已建） | ✅ | ❌ | ❌ | ✅ |
| 动画触发/延迟/时长/序列（Timeline） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 交互式动画（点击触发形状） | ❌ | ✅ | ❌ | ❌ | ✅ |
| 读取切换效果设置 | ✅ | ✅ | ❌ | ❌ | ✅ |

### 2.10 保护、属性与其他

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 修改密码保护（SHA-512 modifyVerifier） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 打开密码（AES 加密） | ❌ | ⚠️ | ❌ | ❌ | ✅ |
| 文档属性（core.xml/app.xml）写入 | ✅ | ✅ | ❌ | ✅（v0.59） | ✅ |
| 批注读取（作者/日期/正文/坐标） | ✅ | ✅ | ❌ | ❌ | ✅ |
| 批注写入 | ✅ | ✅ | ❌ | ❌ | ✅ |
| 超链接（文本/形状跳转） | ✅ | ✅ | ❌ | ✅ | ✅ |
| 超链接（跳转到文件） | ❌ | ✅ | ❌ | ✅（v0.58） | ✅ |
| Section 分组 | ✅ | ✅ | ❌ | ⚠️（v0.30） | ✅ |
| VBA 宏（pptm） | ❌ | ⚠️ | ❌ | ❌ | ✅ |

### 2.11 导出与转换

| 功能 | NewLife.Office | Open XML SDK | NPOI | ShapeCrawler | Aspose.Slides |
|------|:---:|:---:|:---:|:---:|:---:|
| 导出 PDF | ✅（文本映射级，PptxPdfConverter+LayoutEngine） | ❌ | ❌ | ❌ | ✅（高保真+PDF/A/UA） |
| 导出 PNG/JPEG（幻灯片渲染） | ❌ | ❌ | ❌ | ❌ | ✅ |
| 导出 SVG | ❌ | ❌ | ❌ | ❌ | ✅ |
| 导出 HTML（含 HTML5 动画） | ❌ | ❌ | ❌ | ❌ | ✅ |
| 导出 XPS | ❌ | ❌ | ❌ | ❌ | ✅ |
| 导出视频（GIF/MP4） | ❌ | ❌ | ❌ | ❌ | ✅（唯一） |
| Markdown 导出 | ✅ | ❌ | ❌ | ✅（v0.65） | ✅ |
| 自动排版引擎（LayoutEngine，5 种策略） | ✅ | ❌ | ❌ | ❌ | ❌ |

---

## 3. 典型场景开发体验对比

以下通过具体开发场景评估各库的 API 友好度与代码复杂度（★ 越多代码越简洁）。

### 3.1 创建一张含标题文本框的幻灯片

| 库名 | 复杂度 | 说明 |
|------|:---:|------|
| **NewLife.Office** | ★★★★★ | `var w = new PptxWriter(); w.AddSlide(); w.AddTextBox(0, "标题", 2, 1, 20, 3);` 约 3 行 |
| **ShapeCrawler** | ★★☆☆☆ | 以读取为主，写入/创建能力有限，链式创建仅部分场景可用 |
| **Open XML SDK** | ★★☆☆☆ | 需手动创建 PresentationPart/SlidePart/ShapeTree/Shape/TextBody/Run 等约 25 行 |
| **Aspose.Slides** | ★★★★☆ | `presentation.Slides[0].Shapes.AddTextFrame(rect, "标题")` 需配置许可证 |
| **NPOI** | ★★☆☆☆ | XSLF 模块不完整，实际功能受限 |

### 3.2 从现有 pptx 读取所有文本和形状

| 库名 | 复杂度 | 说明 |
|------|:---:|------|
| **NewLife.Office** | ★★★★★ | `using var r = new PptxReader(path); var slides = r.ReadAllSlides();` 高层模型完整 |
| **ShapeCrawler** | ★★★★☆ | `using var p = new Presentation(path); foreach(var s in p.Slides)` 枚举 Shapes |
| **Open XML SDK** | ★★★☆☆ | 需手动遍历 SlidePart/ShapeTree/TextBody 等 XML 元素 |
| **Aspose.Slides** | ★★★★☆ | 功能最全但需商业许可 |

### 3.3 创建含数据的柱状图

| 库名 | 复杂度 | 说明 |
|------|:---:|------|
| **NewLife.Office** | ★★★★★ | `w.AddBarChart(0, cats, new[]{series}, 2, 4, 20, 10)` 单行调用 |
| **ShapeCrawler** | ★★★☆☆ | `shapes.AddBarChart(ChartType.BarClustered)` 可创建图表，但整体写入能力有限 |
| **Open XML SDK** | ★☆☆☆☆ | 需手动构建 ChartPart + 内嵌 xlsx EmbeddedPackage + 完整图表 XML 约 60 行 |
| **Aspose.Slides** | ★★★★☆ | API 完整，需商业许可 |

### 3.4 从企业模板 pptx 生成演示文稿

| 库名 | 复杂度 | 说明 |
|------|:---:|------|
| **NewLife.Office** | ★★★★★ | `new PptxWriter(templatePath)` 一行加载，自动复用母版/版式/主题/嵌入字体 |
| **ShapeCrawler** | ★★★☆☆ | 需手动处理 SlideMasterPart 关系 |
| **Open XML SDK** | ★★☆☆☆ | 手动复制 SlideMasterPart/SlideLayoutPart/ThemePart 及关系 |
| **Aspose.Slides** | ★★★★☆ | 支持模板，自动处理母版复制 |

### 3.5 替换模板占位符并填充表格数据

| 库名 | 复杂度 | 说明 |
|------|:---:|------|
| **NewLife.Office** | ★★★★★ | `new PptxTemplate(path).FillTable(out, data, lists)` 一行完成，支持占位符/表格行扩展/图片占位符/对象集合导出表格 |
| **ShapeCrawler** | ★★★☆☆ | 需手动遍历 Shapes 找占位符并逐一替换 |
| **Open XML SDK** | ★★☆☆☆ | 需手动操作 XML 字符串查找替换 |
| **Aspose.Slides** | ★★★★☆ | 支持，Mail Merge 风格 API |
| **NPOI** | ❌ | 不支持 |

---

## 4. 竞品优劣势详析

### 4.1 Open XML SDK

**版本**：v3.x（2024+），由微软维护  
**许可**：MIT，长期维护保障  
**GitHub**：~4k Stars，微软官方仓库，随 Office 规范同步更新

**优势**：
- 微软官方出品，对 OOXML 规范最完整的 .NET 实现
- 类型安全：所有 XML 元素均有对应 .NET 类型，编译期检查，IDE 智能提示完善
- 完全控制底层 XML 结构，任何 OOXML 规范描述的功能均可实现
- 可作为其他库的底层（ShapeCrawler 基于其构建）
- 长期维护保障，与微软 Office 格式规范生命周期绑定

**劣势**：
- API 极为底层冗长：创建一张简单幻灯片需要实例化 PresentationDocument、PresentationPart、SlidePart、SlideLayoutPart、ShapeTree、Shape、TextBody、Paragraph、Run 等约 20+ 个对象
- 无任何业务级封装，开发者需深入理解 OOXML 规范，开发效率低
- 无渲染能力，不能导出 PDF 或图片
- 不支持 ppt（97-2003）老格式

**适用场景**：需要精确控制 XML 结构、构建上层库的基础层；不推荐在业务代码中直接大量使用。

---

### 4.2 NPOI（PPT 模块）

**版本**：v2.7.x（Apache 2.0）/ v2.8.0+（OSMFEULA，商用付费）  
**许可**：⚠️ v2.8.0 起营利性组织商用需付费  
**GitHub**：~6.2k Stars，以 Excel/Word 模块为核心

> ⚠️ **重要说明**：NPOI 的 PPT 功能模块（HSLF 处理 .ppt，XSLF 处理 .pptx）长期位于仓库的 `scratchpad/` 草稿目录，标记为实验性/未完成。NPOI 的 README 的"主要功能"列表甚至不包含 PPT，其核心价值在 Excel 和 Word 模块。

**优势**：
- HSLF 模块（.ppt 97-2003）相对完整，支持 Slide/SlideMaster/Shape/Picture/Movie/Background/TextRun
- 在同一库中同时处理 Excel/Word/PPT，降低依赖数量

**劣势**：
- XSLF 模块（.pptx）严重残缺：缺少表格、图表、动画、评论、大多数格式控制，成熟度不足
- API 移植自 Java POI，不够 .NET 原生，使用体验差
- v2.8.0 起商用需付费，对企业项目带来法律合规风险
- 包体积约 10 MB，相比其功能性价比低

**适用场景**：已有大量历史代码使用 NPOI 的 Excel/Word 且需要同时处理老版 .ppt 文本的项目。**PPT 场景不建议新项目选用 NPOI。**

---

### 4.3 ShapeCrawler

**版本**：v0.79.x（2025年），活跃开发中  
**许可**：MIT  
**GitHub**：~1.5k Stars，社区增长迅速，版本迭代频繁

**优势**：
- **API 设计最现代友好**：链式调用、`using Presentation`、强类型 Shape 继承层次（IAutoShape/IPicture/ITable/IGroupShape/IVideoShape 等）
- 功能覆盖全面：形状/文本/表格/图表/图片/音视频/母版/版式/主题色/超链接/页眉页脚/备注/Section/组形状
- 持续活跃开发，几乎每个月发布新版本，功能稳步扩充
  - v0.52：SVG 图片；v0.54：表格边框；v0.57：表格列操作；v0.67：组形状/柱状图；v0.73：图表标题；v0.79：公司属性
- `Rows.Add(insertIndex, templateRow)` 基于模板行动态插入表格行——设计简洁，实用性强
- `AsMarkdown()` 导出（v0.65）：pptx 内容一键转 Markdown

**劣势**：
- **依赖 Open XML SDK**：项目引入了额外依赖，与 NewLife.Office 零依赖原则不同
- **写入/创建能力有限（只读为主）**：无法完整创建新演示文稿、复制幻灯片、合并文件等
- 不支持 ppt（97-2003）老格式
- 无任何渲染/导出能力（PDF/PNG/HTML 均无）
- **不支持幻灯片切换效果和元素动画**（重要缺失，ShapeCrawler 的显著弱项）
- 无评论/批注支持
- 无文档保护/加密
- v0.x 版本号意味着 API 仍可能有 Breaking Change

**亮点 API 设计**：
- `IShapeFill.SetHexSolidColor()` / `SetPicture()` / `SetNoFill()` — 形状填充最简洁
- `ITable.MergeCells()` / `AddColumn()` / `InsertColumnAfter()` — 表格动态操作最方便
- `ISlideMaster.ITheme` — 主题颜色访问最简洁
- `IAutoShape.Duplicate()` — 形状克隆最直观

**适用场景**：纯 pptx 读取、API 友好度优先、不需要写入/渲染/动画/老格式的 .NET 项目。是目前免费 pptx 库中读取 API 设计最好的。

---

### 4.4 Aspose.Slides for .NET

**版本**：持续更新（商业产品）  
**许可**：商业，Developer OEM License 约数千美元，需年度订阅  
**NuGet 下载量**：500 万+，商业生态中主导地位

**优势**：
- **功能最全的 .NET PPT 操作库**，无功能盲区
- **唯一支持高保真渲染**：pptx → PDF/PNG/JPEG/SVG/HTML5/XPS/BMP/TIFF，内置布局引擎
- **唯一支持 Morph 形态转换**（3D 对象在幻灯片间平滑变形）
- **150+ 动画效果**，自定义动画路径，完整 Timeline/Trigger/Sequence 系统
- **导出视频**（逐帧渲染 → 视频流，Animated GIF）——业界唯一
- SmartArt 创建和编辑、数学公式（MathML）、3D 效果
- ODP/FODP/OTP（OpenDocument）完整读写，PPT/PPTX 互转
- PDF/A-2/3、PDF/UA 合规导出（归档和无障碍访问标准）
- 字体嵌入/子集化精细控制，跨平台字体回退规则
- AI 插件支持（Aspose.Slides.AI 扩展）

**劣势**：
- **价格高昂**：对中小企业和开源项目不可用
- **闭源**：无法审计内部实现，依赖供应商
- 包体积约 40 MB，影响部署镜像大小
- 不配置许可证会在输出文件中加水印

**适用场景**：企业级高保真 PDF/图片导出、动画控制、SmartArt 编辑、全格式互转等需要全功能且预算充足的商业项目。若只需 pptx 读写无需导出图片，用 NewLife.Office 即可替代且免费。

---

### 4.5 Spire.Presentation Free

**版本**：持续更新  
**许可**：免费受限版（10页严格限制）/ 商业版  
**包大小**：约 20 MB

**优势**：免费版可快速验证 API 形态；API 设计风格类似 Aspose.Slides，迁移成本低。

**劣势**：
- 免费版 10 页限制在生产环境几乎不可用
- 商业版价格与 Aspose.Slides 相当，但功能深度、社区支持、文档质量均差距明显
- 闭源，供应商锁定

**适用场景**：评估/原型验证，不推荐生产选型。

---

## 5. NewLife.Office 差距清单

基于竞品深度对比，将 PPT 模块能力分为**已全部完成的差距项**与**仍存在的真实差距项**两类。

### 5.1 已全部完成的差距项 ✅

以下此前标为缺失的能力已全部完成（模型层 + Writer XML 生成侧均已落地）：

| 差距项 | 对标竞品 | 能力说明 |
|--------|---------|---------|
| docProps/core.xml + app.xml 写入 | 所有竞品 | 文档属性完整读写 |
| 表格单元格合并（读写双向） | ShapeCrawler/Aspose/OpenXML | `gridSpan`/`rowSpan`/`vMerge` 往返 |
| 表格单元格四边独立边框 | ShapeCrawler(v0.54+)/Aspose | `<a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>` |
| 连接器 XML 写入 | 所有竞品 | `ConnectionShape` → `<p:cxnSp>` |
| 批注 XML 读写 | 所有竞品 | `comments{N}.xml` 含作者/日期/正文/坐标 |
| 元素动画 XML 写入 | Aspose/OpenXML | `<p:timing>` 进入/强调/退出 + 触发/延迟/时长/序列 |
| 全局页眉页脚 | 所有竞品 | `<p:hf>` 元素写入 |
| 幻灯片隐藏/显示 | ShapeCrawler/Aspose | `show="0"` |
| 图表散点图/气泡图/股价/曲面 | ShapeCrawler(v0.78)/Aspose | 10 类图表创建 + 读取 + 修改 |
| 图表坐标轴 Min/Max | ShapeCrawler(v0.45)/Aspose | 已覆盖 |
| 散点图数值 X 轴（c:xVal） | Aspose/OpenXML | 已覆盖 |
| 形状旋转角度 | 所有竞品 | `<a:xfrm rot>` |
| 图片原地替换 | ShapeCrawler/Aspose | `PptxWriter.ReplaceImage()` |
| Section 管理 | ShapeCrawler(v0.30)/Aspose | `Presentation.Sections` 读写 |
| SVG 图片读写 | ShapeCrawler(v0.52)/Aspose | `asvg:svgBlip` |
| 背景渐变填充 | 所有竞品 | `<p:bgPr><a:gradFill>` |
| 上标/下标 | ShapeCrawler/Aspose | Run 级 |
| 形状 Alt Text | ShapeCrawler(v0.57)/Aspose | 已覆盖 |
| 文本方向 | ShapeCrawler(v0.71)/Aspose | 垂直/旋转 |
| 渐变文字填充 | Aspose/OpenXML | Run 级渐变 |
| 编程式创建母版 | Aspose/OpenXML | `CreateMaster` + 跨文件复制母版 + 版式索引控制 |
| 主题色（6 强调色） | ShapeCrawler/Aspose | 已覆盖 |
| 原始 XML 访问 | OpenXML/Aspose | 直接访问/修改 XML 部件 |
| 视频/音频嵌入（含缩略图） | Aspose/OpenXML | `PptVideo` |
| 图片圆角 | ShapeCrawler/Aspose | 已覆盖 |
| 写入 ppt（97-2003） | Aspose | `PptWriter` 最小可用：文本幻灯片（多行/中文/长文本续记录），合法 OLE2（PowerPoint Document + Current User 流） |
| 打开密码（AES 加密） | Aspose / OpenXML | `SaveEncrypted` OOXML Agile Encryption + `PptxReader(password)` 解密 |
| 形状图片填充 | Aspose / ShapeCrawler | `Shape.FillImage` BlipFill 读写（含组内形状） |
| 图表样式/配色 | Aspose / OpenXML | `StyleIndex` + 系列颜色读写 + 数值轴 min/max/图例位置读回 |
| 超链接跳转到文件 | Aspose / ShapeCrawler | Run/TextBox/Shape/组内形状跳转文件（TargetMode=External）读写 |

### 5.2 仍存在的真实差距项

以下为 PPT 模块当前真实未实现的能力，选择时需明确边界：

| 差距项 | 对标竞品 | 说明 |
|--------|---------|------|
| **SmartArt 创建/编辑** | Aspose / OpenXML（底层） | 不支持 SmartArt 模型，无法生成/修改 SmartArt 图形 |
| **完整动画时间轴系统** | Aspose（150+ 效果） | 支持进入/强调/退出三类基础动画，缺交互式动画（点击触发）、自定义动画路径、效果库 |
| **pptx → PNG/JPEG 幻灯片渲染** | Aspose | 需外部渲染库（SkiaSharp 等），与零外部依赖定位冲突 |
| **高保真 PDF 导出** | Aspose | 当前为文本映射级（`PptxPdfConverter` + `LayoutEngine`），非像素级 |
| **导出 SVG/HTML/XPS/视频** | Aspose | 不计划 |
| **ODP/FODP 读写** | Aspose | 不计划 |
| **VBA 宏（pptm）** | Aspose / OpenXML | 不计划 |
| **OLE 嵌入对象** | Aspose / OpenXML | 不计划 |
| **数学公式（MathML）** | Aspose / OpenXML | 不计划 |

---

## 6. 结论与差异化定位

### 6.1 真实未实现项说明

NewLife.Office 的 PPT 模块在**读写与业务自动化**层面已全面落地，但以下高级能力真实未实现：

| 未实现项 | 影响 |
|---------|------|
| **SmartArt 创建/编辑** | 无法生成 SmartArt 图形 |
| **完整动画时间轴系统** | 支持进入/强调/退出三类基础动画，但缺交互式动画（点击触发）、自定义动画路径、150+ 效果库 |
| **pptx → PNG/JPEG 幻灯片渲染** | 无像素级渲染；PDF 导出为文本映射级（`PptxPdfConverter` + `LayoutEngine` 自动排版） |

### 6.2 差异化定位

- **vs ShapeCrawler**：
  - ✅ 零外部依赖（ShapeCrawler 依赖 Open XML SDK）
  - ✅ 支持写入与创建（ShapeCrawler 以只读为主）
  - ✅ 支持 ppt（97-2003）老格式读取
  - ✅ 内置 PDF 转换（文本映射级）+ 自动排版引擎（`LayoutEngine`，5 种布局策略）
  - ✅ 模板填充（`{{Key}}` 占位符/表格行扩展/图片替换/对象集合导出表格 `WriteObjects`）
  - ✅ 幻灯片切换 + 元素动画读写（ShapeCrawler 两者均缺失）
  - ✅ 批注读写 / 连接器 / 全局页眉页脚 / 文档属性（ShapeCrawler 缺失）
  - ✅ 表格合并/边框/动态行列完成
  - ❌ 无图片渲染/PDF 高保真（与 ShapeCrawler 同等缺失）

- **vs NPOI（PPT 模块）**：
  - ✅ PPT 功能完整性远超（NPOI PPT 在 scratchpad，功能严重残缺）
  - ✅ API 设计更现代，符合 .NET 惯例
  - ✅ MIT 许可，v2.8+ NPOI 存在商用风险

- **vs Open XML SDK**：
  - ✅ 高层封装，开发效率数倍提升
  - ✅ 零依赖（Open XML SDK 本身是外部依赖）
  - ✅ 内置 PDF 转换、自动排版等上层业务能力
  - ⚠️ Open XML SDK 支持 SmartArt/OLE/公式等底层能力，NewLife.Office 不覆盖

- **vs Aspose.Slides**：
  - ✅ 完全免费，MIT 许可，无许可证管理风险
  - ✅ 零依赖，部署简单，镜像体积小
  - ❌ 无像素级渲染（PDF/图片导出精度有限）
  - ❌ 无完整动画时间轴系统、SmartArt、导出视频等高级能力

### 6.3 结论

.NET 生态中免费可用的 PPT 操作库极为有限。NewLife.Office 在**零依赖 + MIT 许可 + ppt/pptx 双格式 + 写入/创建完整 + 模板填充 + PDF 转换 + 自动排版**的组合能力上具有独特优势，是企业内部文档自动化、模板批处理生成和 SaaS 应用的优选方案。

**选型建议**：

| 需求场景 | 推荐选型 |
|---------|---------|
| pptx 写入/创建、模板批量生成、表格图表填充、免费零依赖 | **NewLife.Office** |
| 高保真渲染（PDF/PNG）、SmartArt、完整动画 | **Aspose.Slides**（商业预算充足） |
| 仅需读取 pptx 文本/形状、偏好现代 API | **ShapeCrawler** |
| 底层精确控制 OOXML 结构 | **Open XML SDK** |
| 老 .ppt 文本提取 | **NewLife.Office** 或 **NPOI（HSLF）** |

---

← 返回 [竞品分析报告](竞品分析报告.md)
