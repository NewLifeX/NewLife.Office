# NewLife.Office 版本更新记录

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
