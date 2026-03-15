# NewLife.Office - 办公自动化库

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

办公自动化库  

源码： https://github.com/NewLifeX/NewLife.Office  
Nuget：NewLife.Office  

## 项目简介

`NewLife.Office` 是一个功能全面的 .NET 办公自动化库，覆盖 **Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 以及基于 OLE2/BIFF8 的老式二进制格式，致力于为 .NET 社区提供零依赖、MIT 许可的高质量办公自动化能力。

- 仅依赖 `NewLife.Core`，零外部依赖
- 支持 `net45/netstandard2.0/netstandard2.1`
- 直接基于 Open XML（ZIP + XML）/ PDF / OLE2 文件结构实现
- 统一命名空间 `NewLife.Office`，API 风格一致，学习成本低

> CSV 读写能力由 `NewLife.Core` 的 `CsvFile` 提供，本仓库不重复实现 CSV。

## 已实现能力（Excel）

### ExcelWriter（写入）

- 多工作表写入、表头与多行数据写入
- 单元格样式：字体/背景/边框/对齐/换行/数字格式
- 布局能力：列宽、行高、合并单元格、冻结窗格、自动筛选
- 高级能力：超链接、下拉数据验证、条件格式、插入图片、图表（柱/折/饼/散点）、数据透视表
- 页面能力：页面方向、纸张大小、页边距、页眉页脚、打印标题行
- 安全能力：工作表保护
- 便捷能力：`WriteObjects<T>()`、`WriteDataTable()`

### ExcelReader（读取）

- 读取多工作表行数据（自动类型转换）
- `ReadObjects<T>()` 对象映射导入（支持 `DisplayName/Description`）
- `ReadDataTable()` 导入为 `DataTable`
- `GetMergeRanges()` 读取合并区域

### ExcelTemplate（模板填充）

- 基于现有 xlsx 模板执行 `{{Key}}` 占位符替换
- 支持输出到文件或流

### BiffReader（xls 安全读取）

- 基于 OLE2/CFB 容器，解析 BIFF8 格式工作表流
- `ReadSheet()` 读取工作表行数据
- `ReadObjects<T>()` 对象映射导入
- 支持字符串、数字、布尔、日期、RK 编码等多种单元格类型

## 已实现能力（Word）

### WordWriter / WordReader / WordTemplate

- docx 全面读写：段落、表格、图片、标题、列表、超链接、页眉页脚、页面设置
- 模板填充：`{{Key}}` 占位符替换，保留原样式
- `WordHtmlConverter`：docx 转语义化 HTML（段落→`<p>`，表格→`<table>`，标题→`<h1>`-`<h6>`）
- `WordPdfConverter`：docx 转 PDF（内容映射型，纳入 PdfFluentDocument）
- `DocReader`：读取 doc 格式（OLE2 容器，提取纯文本和段落结构）

## 已实现能力（PPT）

### PptxWriter / PptxReader / PptxTemplate

- pptx 全面读写：幻灯片管理、文本框、表格、图片、基本形状、超链接、备注、模板填充
- 布局设计：母版、版式、背景、尺寸、幻灯片切换动画
- `PptReader`：读取 ppt 格式（OLE2 容器，提取幻灯片文本）

## 已实现能力（PDF）

### PdfWriter / PdfReader / PdfDocument

- 创建 PDF：文本排版、表格、图片、页眉页脚、中文字体支持
- `PdfFluentDocument`：声明式布局 API（Row/Column/Text/Image/Table），自动分页
- PDF 操作：合并/拆分、水印、页面旋转、文字覆盖、可点击链接、书签导航
- `PdfEncryptor`：加密/权限控制
- 文本提取：逐页/全文，带坐标信息提取

## 已实现能力（Markdown / RTF / ODS）

### Markdown

- `MarkdownParser`：解析 CommonMark + GFM（标题/段落/代码块/列表/引用/粗斜体/图片/表格/任务列表）
- `MarkdownWriter`：将 `MarkdownDocument` 序列化为 .md 文本（包括 GFM 扩展）
- `MarkdownHtmlConverter`：Markdown 转换为语义化 HTML

### RTF

- `RtfReader`：解析 RTF 文件，提取段落（字体/颜色/对齐）、表格、文档属性
- `RtfWriter`：创建 RTF 文件（段落格式/表格/模板填充）

### ODS

- `OdsReader`：读取 LibreOffice ODS 格式，提取多工作表单元格数据
- `OdsWriter`：生成 ODS 文件（多工作表、字符串/数字/日期/布尔数据）

## OLE2/CFB 容器

- `CfbReader`：解析 Compound File Binary 容器（安全寻找 Stream，支持 v3/v4 水印）
- `CfbWriter`：生成合法 CFB 文件（写入流和目录树）
- 支撟 xls（Excel 97-2003）、doc（Word 97-2003）、ppt（PPT 97-2003）共用基础层

## 快速开始

```csharp
using NewLife.Office;

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);

writer.WriteHeader("Data", new[] { "编号", "姓名", "金额" }, CellStyle.Header);
writer.WriteRow("Data", new Object?[] { 1, "张三", 1234.56m }, new CellStyle
{
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00"
});

writer.FreezePane("Data", 1);
writer.SetAutoFilter("Data", "A1:C1");
writer.Save();

ms.Position = 0;
var reader = new ExcelReader(ms, Encoding.UTF8);
var rows = reader.ReadRows("Data").ToList();
```

### 对象集合导出

```csharp
public class UserInfo
{
    [DisplayName("编号")]
    public Int32 Id { get; set; }

    [DisplayName("姓名")]
    public String Name { get; set; }

    [DisplayName("金额")]
    public Decimal Amount { get; set; }
}

var list = new List<UserInfo>
{
    new() { Id = 1, Name = "张三", Amount = 1234.56m },
    new() { Id = 2, Name = "李四", Amount = 9876.00m },
};

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);
writer.WriteObjects("用户列表", list);
writer.Save();
```

### 对象集合导入

```csharp
using var fs = File.OpenRead("users.xlsx");
var reader = new ExcelReader(fs);
var users = reader.ReadObjects<UserInfo>("用户列表");
```

### 模板填充

```csharp
// 基于 xlsx 模板替换 {{Key}} 占位符，保留原有样式
var template = new ExcelTemplate("report_template.xlsx");
template.Fill(new Dictionary<String, Object>
{
    ["Title"] = "2025年度报告",
    ["Author"] = "张三",
    ["Date"] = DateTime.Today.ToString("yyyy-MM-dd"),
});
template.Save("report_output.xlsx");
```

### 单元格样式

```csharp
var style = new CellStyle
{
    Bold = true,
    FontColor = "FFFFFF",        // 白色字体
    BackgroundColor = "4472C4",  // 蓝色背景
    HAlign = HorizontalAlignment.Center,
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00",
};

writer.WriteHeader("Sheet1", new[] { "编号", "姓名", "金额" }, style);
writer.MergeCell("Sheet1", "A1:C1");  // 合并标题行
writer.SetRowHeight("Sheet1", 1, 30); // 设置行高
```

## 测试状态

- 当前测试集：`246` 项
- 通过：`246` 项
- 失败：`0` 项

## 已实现格式概览

| 格式 | 读取 | 写入 | 模板 | 转换 | 说明 |
|------|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | 含样式/图表/透视表/条件格式/图片 |
| **Excel xls** | ✅ | 🔄规划 | — | — | OLE2/BIFF8，BiffReader 已实现读取 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | 含段落/表格/图片/页眉页脚 |
| **Word doc** | ✅ | — | — | — | MS-DOC，DocReader 提取纯文本 |
| **PPT pptx** | ✅ | ✅ | ✅ | 🔄规划 | 含文本框/图片/表格/图表/母版 |
| **PPT ppt** | ✅ | — | — | — | MS-PPT，PptReader 提取幻灯片文本 |
| **PDF** | ✅ | ✅ | — | — | 含 Fluent API/合并/拆分/加密 |
| **Markdown** | ✅ | ✅ | — | HTML | CommonMark + GFM，含 AST 模型 |
| **RTF** | ✅ | ✅ | ✅ | — | 段落/表格/模板填充 |
| **ODS** | ✅ | ✅ | — | — | LibreOffice 原生格式 |
| **EML** | 🔄规划 | 🔄规划 | — | — | RFC 5322 + MIME |
| **MSG** | 🔄规划 | — | — | — | Outlook，OLE2 容器 |
| **iCalendar** | 🔄规划 | 🔄规划 | — | — | RFC 5545 .ics 格式 |
| **vCard** | 🔄规划 | 🔄规划 | — | — | RFC 6350 .vcf 格式 |
| **EPUB** | 🔄规划 | 🔄规划 | — | — | EPUB 2/3 电子书 |
| **XPS** | 🔄规划 | 🔄规划 | — | — | XPS/OXPS 固定版式 |

## 代码结构

```
NewLife.Office/
├── Excel/        # xlsx 读写：ExcelWriter / ExcelReader / ExcelTemplate / CellStyle / ExcelHelper / BiffReader
├── Word/         # docx 读写：WordWriter / WordReader / WordTemplate / DocReader / WordHtmlConverter / WordPdfConverter
├── PPT/          # pptx 读写：PptxWriter / PptxReader / PptxTemplate / PptReader
├── PDF/          # PDF 读写：PdfWriter / PdfReader / PdfDocument / PdfFluentDocument / PdfEncryptor
├── OLE2/         # OLE2/CFB 容器：CfbReader / CfbWriter
├── Markdown/     # Markdown 读写：MarkdownParser / MarkdownWriter / MarkdownHtmlConverter
├── Rtf/          # RTF 读写：RtfReader / RtfWriter
├── Ods/          # ODS 读写：OdsReader / OdsWriter
├── Email/        # EML/MSG 读写（规划中）
├── Calendar/     # iCalendar .ics 读写（规划中）
├── Contact/      # vCard .vcf 读写（规划中）
├── Epub/         # EPUB 读写（规划中）
└── Xps/          # XPS 读写（规划中）
```

## 规划中

以下功能已在路线图中，但尚未实现：

### 写入 xls（BIFF8，Excel 97-2003）

- 生成合法 BIFF8 工作簿流并写入 OLE2 容器

### PPT → PDF / 图片

- pptx 每张幻灯片转换为 PDF 页面或 PNG/JPEG 图像

### Markdown → Word / PDF

- Markdown 标题/段落映射为 Word 样式或通过 PDF 渲染

### RTF 图片嵌入 / ODS 合并单元格

- RtfWriter 支持 PNG/JPEG 图片嵌入
- OdsReader 识别合并单元格（number-columns-spanned）

### Email / iCalendar / vCard

- EML（RFC 5322 + MIME）读写，附件提取
- MSG（Outlook，基于 OLE2）读取
- iCalendar .ics 读写（RFC 5545）
- vCard .vcf 读写（RFC 6350）

### EPUB / XPS

- EPUB 2/3 电子书读写
- XPS/OXPS 固定版式文档读写

## 新生命项目矩阵

各项目默认支持net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5，旧版（2024.0801）支持net4.0/net2.0

|                               项目                               | 年份  | 说明                                                                                        |
| :--------------------------------------------------------------: | :---: | ------------------------------------------------------------------------------------------- |
|                             基础组件                             |       | 支撑其它中间件以及产品项目                                                                  |
|          [NewLife.Core](https://github.com/NewLifeX/X)           | 2002  | 核心库，日志、配置、缓存、网络、序列化、APM性能追踪                                         |
|    [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode)    | 2005  | 大数据中间件，单表百亿级，MySql/SQLite/SqlServer/Oracle/PostgreSql/达梦，自动分表，读写分离 |
|      [NewLife.Net](https://github.com/NewLifeX/NewLife.Net)      | 2005  | 网络库，单机千万级吞吐率（2266万tps），单机百万级连接（400万Tcp长连接）                     |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011  | 协议通信库，提供CS应用通信框架，支持Http/RPC通信框架，高吞吐，物联网设备低开销易接入        |
|     [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube)     | 2010  | 魔方快速开发平台，集成了用户权限、SSO登录、OAuth服务端等，单表100亿级项目验证               |
|    [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent)    | 2008  | 服务管理组件，把应用安装成为操作系统守护进程，Windows服务、Linux的Systemd                   |
|     [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero)     | 2020  | Zero零代脚手架，基于NewLife组件生态的项目模板NewLife.Templates，Web、WebApi、Service        |
|                              中间件                              |       | 对接知名中间件平台                                                                          |
|    [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis)    | 2017  | Redis客户端，微秒级延迟，百万级吞吐，丰富的消息队列，百亿级数据量项目验证                   |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018  | RocketMQ纯托管客户端，支持Apache RocketMQ和阿里云消息队列，十亿级项目验                     |
|     [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT)     | 2019  | 物联网消息协议，MqttClient/MqttServer，客户端支持阿里云物联网                               |
|      [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT)      | 2022  | IoT标准库，定义物联网领域的各种通信协议标准规范                                             |
|   [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus)   | 2022  | ModbusTcp/ModbusRTU/ModbusASCII，基于IoT标准库实现，支持ZeroIoT平台和IoTEdge网关            |
|  [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens)  | 2022  | 西门子PLC协议，基于IoT标准库实现，支持IoT平台和IoTEdge                                      |
|      [NewLife.Map](https://github.com/NewLifeX/NewLife.Map)      | 2022  | 地图组件库，封装百度地图、高德地图、腾讯地图、天地图                                        |
|    [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio)    | 2023  | 音频编解码库，PCM/ADPCMA/G711A/G722U/WAV/AAC                                                |
|                             产品平台                             |       | 产品平台级，编译部署即用，个性化自定义                                                      |
|         [Stardust](https://github.com/NewLifeX/Stardust)         | 2018  | 星尘，分布式服务平台，节点管理、APM监控中心、配置中心、注册中心、发布中心                   |
|           [AntJob](https://github.com/NewLifeX/AntJob)           | 2019  | 蚂蚁调度，分布式大数据计算平台（实时/离线），蚂蚁搬家分片思想，万亿级数据量项目验证         |
|      [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP)      | 2021  | 企业ERP，产品管理、客户管理、销售管理、供应商管理                                           |
|         [CrazyCoder](https://github.com/NewLifeX/XCoder)         | 2006  | 码神工具，众多开发者工具，网络、串口、加解密、正则表达式、Modbus、MQTT                      |
|           [EasyIO](https://github.com/NewLifeX/EasyIO)           | 2023  | 简易文件存储，支持分布式系统中文件集中存储。                                                |
|           [XProxy](https://github.com/NewLifeX/XProxy)           | 2005  | 产品级反向代理，NAT代理、Http代理                                                           |
|        [HttpMeter](https://github.com/NewLifeX/HttpMeter)        | 2022  | Http压力测试工具                                                                            |
|         [GitCandy](https://github.com/NewLifeX/GitCandy)         | 2015  | Git源代码管理系统                                                                           |
|          [SmartOS](https://github.com/NewLifeX/SmartOS)          | 2014  | 嵌入式操作系统，完全独立自主，支持ARM Cortex-M芯片架构                                      |
|          [SmartA2](https://github.com/NewLifeX/SmartA2)          | 2019  | 嵌入式工业计算机，物联网边缘网关，高性能.NET8主机，应用于工业、农业、交通、医疗             |
|                          FIoT物联网平台                          | 2020  | 物联网整体解决方案，建筑、环保、农业，软硬件及大数据分析一体化，单机十万级点位项目验证      |
|                        UWB高精度室内定位                         | 2020  | 厘米级（10~20cm）高精度室内定位，软硬件一体化，与其它系统联动，大型展厅项目验证             |

## 新生命开发团队

![XCode](https://newlifex.com/logo.png)  

新生命团队（NewLife）成立于2002年，是新时代物联网行业解决方案提供者，致力于提供软硬件应用方案咨询、系统架构规划与开发服务。  
团队主导的80多个开源项目已被广泛应用于各行业，Nuget累计下载量高达400余万次。  
团队开发的大数据中间件NewLife.XCode、蚂蚁调度计算平台AntJob、星尘分布式平台Stardust、缓存队列组件NewLife.Redis以及物联网平台FIoT，均成功应用于电力、高校、互联网、电信、交通、物流、工控、医疗、文博等行业，为客户提供了大量先进、可靠、安全、高质量、易扩展的产品和系统集成服务。  

我们将不断通过服务的持续改进，成为客户长期信赖的合作伙伴，通过不断的创新和发展，成为国内优秀的IoT服务供应商。  

`新生命团队始于2002年，部分开源项目具有20年以上漫长历史，源码库保留有2010年以来所有修改记录`  
网站：<https://newlifex.com>  
开源：<https://github.com/newlifex>  
QQ群：1600800/1600838  
微信公众号：  
![智能大石头](https://newlifex.com/stone.jpg)  
