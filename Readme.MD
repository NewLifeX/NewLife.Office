# NewLife.Office - 办公自动化库

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

功能全面的 .NET 办公自动化库，一库覆盖 Excel、Word、PPT、PDF 等 15+ 文档格式

[English](Readme_EN.MD) | 简体中文 | [繁體中文](Readme_TW.MD) | [日本語](Readme_JA.MD) | [한국어](Readme_KO.MD) | [Español](Readme_ES.MD) | [Français](Readme_FR.MD) | [Deutsch](Readme_DE.MD) | [Português](Readme_PT.MD) | [Русский](Readme_RU.MD)

源码： https://github.com/NewLifeX/NewLife.Office  
Nuget：NewLife.Office（渲染扩展包：NewLife.Office.Rendering）  

## 项目简介

`NewLife.Office` 是一个功能全面的 .NET 办公自动化库，覆盖 **Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 以及基于 OLE2/BIFF8 的老式二进制格式，并配套 **Rendering** 渲染扩展包（文档 → 图片）。

- 零外部依赖：仅依赖 `NewLife.Core`，直接基于 Open XML（ZIP + XML）/ PDF / OLE2 文件结构实现
- 全框架支持：`net45` / `netstandard2.0` / `netstandard2.1` 等
- MIT 免费：无商业许可风险
- 统一 API：命名空间 `NewLife.Office`，各格式 API 风格一致，学习成本低

> CSV 读写能力由 `NewLife.Core` 的 `CsvFile` 提供，本仓库不重复实现。

## 亮点

- **全格式覆盖**：四大核心格式 + 十种轻量格式 + OLE2/BIFF8 老式格式，一库搞定
- **零外部依赖**：直接操作文件底层结构，包体积 <500KB
- **高保真往返**：L2/L3 级保真架构（模型 + RawXml/部件透传），读→改→写不丢格式
- **Markdown 全格式转换中枢**：MD↔docx/PDF/HTML/Excel/PPT/Word 双向转换，为 AI 知识库喂入提供统一入口
- **渲染扩展包**：`NewLife.Office.Rendering`（独立 NuGet）提供文档 → 图片渲染、图片处理、二维码

## 支持的格式

| 格式 | 读取 | 写入 | 模板 | 转换 | 渲染 | 说明 |
|------|:---:|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | 含样式/图表/透视表/条件格式/图片/命名范围/结构化表格/批注/迷你图/切片器/线程化批注/隐藏行列，流式读取与 1904 日期支持 |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8，BiffReader+BiffWriter 读写双全，含样式/合并/超链接/筛选/页面设置 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | 含段落/表格/图片/页眉页脚/批注/内容控件/AES 加密/邮件合并/查找替换（含文本框）/脚注模型/多节，NPOI+OpenXmlValidator 交叉验证 |
| **Word doc** | ✅ | — | — | ✅ doc→docx | — | MS-DOC，DocReader 纯文本 + 段落格式提取（粗体/字号/对齐等）+ doc→docx 转换 |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | 含文本框/图片/表格/图表/母版/动画/批注/Section；跨文件合并（含图表/批注/备注保真） |
| **PPT ppt** | ✅ | ✅ | — | — | — | MS-PPT：PptReader 文本提取 + PptWriter 文本幻灯片写入 |
| **PDF** | ✅ | ✅ | — | — | — | 含 Fluent API/合并/拆分/加密/签名/表单/注释/PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/各格式→MD | — | CommonMark + GFM + 强类型 AST，全格式转换中枢 |
| **RTF** | ✅ | ✅ | ✅ | — | — | 段落/表格/模板填充/图片嵌入（PNG/JPEG） |
| **ODS** | ✅ | ✅ | — | — | — | LibreOffice 原生格式，含合并单元格/样式/公式 |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME，附件/内嵌图片 |
| **MSG** | ✅ | — | — | — | — | Outlook，OLE2 + MAPI，提取主题/正文/附件 |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics，VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf，vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3，OPF/Nav/章节/封面 |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML，Glyphs 文本提取/多页写入/图片/元数据 |
| **Rendering** | — | — | — | — | ✅ | 独立扩展包（SkiaSharp）：PDF/Word/PPT/Markdown → 图片 |

## 快速开始

### 统一文本 / Markdown 提取

```csharp
using NewLife.Office;

// 一行代码提取任意格式纯文本
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// 提取为 Markdown 格式（保留结构信息）
String? md = OfficeFactory.ReadMarkdown("slides.pptx");

// doc（97-2003）→ docx 转换
OfficeFactory.ConvertDocToDocx("legacy.doc", "legacy.docx");
```

### 命令行工具（NewLife.Office.Cli）

```bash
# 任意办公文档 → Markdown
dotnet run --project NewLife.Office.Cli -- convert input.docx output.md

# doc（97-2003）→ docx（历史文档迁移）
dotnet run --project NewLife.Office.Cli -- doc2docx legacy.doc legacy.docx
```

### Excel 写入与读取

```csharp
using NewLife.Office;

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);
writer.WriteHeader("Data", new[] { "编号", "姓名", "金额" }, CellStyle.Header);
writer.WriteRow("Data", new Object?[] { 1, "张三", 1234.56m }, new CellStyle
{
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00",
});
writer.Save();

ms.Position = 0;
var reader = new ExcelReader(ms, Encoding.UTF8);
var rows = reader.ReadRows("Data").ToList();
```

> 更多能力（对象映射导入导出、模板填充、格式转换、渲染等）见 [需求文档](Doc/需求文档.md) 与源码示例。

## 文档索引

- [需求文档](Doc/需求文档.md)
- [功能清单](Doc/功能清单.md)
- [架构设计](Doc/架构设计.md)
- [竞品分析](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

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
