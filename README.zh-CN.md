# NewLife.Office - 办公自动化库

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)


![GitHub License](https://img.shields.io/github/license/NewLifeX/NewLife.Office?logo=github)
![NuGet Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![NuGet](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![NuGet prerelease](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)
[![test](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml/badge.svg?branch=master)](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml)
[![locales](https://img.shields.io/badge/locales-31-2ea44f?logo=dotnet)](Doc/多语言本地化.md)
[![AOT smoke](https://img.shields.io/badge/AOT%20smoke-23%2F23-brightgreen?logo=dotnet)](.github/workflows/aot.yml)
![platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS%20%7C%20Android%20%7C%20iOS-blue)
![arch](https://img.shields.io/badge/arch-x64%20%7C%20x86%20%7C%20arm64%20%7C%20loongarch64%20%7C%20riscv64%20%7C%20s390x-orange)
[![contributors](https://img.shields.io/github/contributors/NewLifeX/NewLife.Office?logo=github)](https://github.com/NewLifeX/NewLife.Office/graphs/contributors)

<!-- lang:start -->
🌐 [English](README.md) | [繁體中文](README.zh-TW.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

功能全面的 .NET 办公自动化库，一库覆盖 Excel、Word、PPT、PDF 等 15+ 文档格式


源码： https://github.com/NewLifeX/NewLife.Office  
Nuget：NewLife.Office（渲染扩展包：NewLife.Office.Rendering）  

> ℹ️ 本中文版与 [English README](README.md) 同批维护；若两者内容不一致，以英文版为对外口径。

## 项目简介

`NewLife.Office` 是一个功能全面的 .NET 办公自动化库，覆盖 **Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 以及基于 OLE2/BIFF8 的老式二进制格式，并配套 **Rendering** 渲染扩展包（文档 → 图片）。

- 零外部依赖：仅依赖 `NewLife.Core`，直接基于 Open XML（ZIP + XML）/ PDF / OLE2 文件结构实现
- 全框架支持：`net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`，net8.0+ 开启 AOT 兼容分析（IsAotCompatible，零 IL 告警）
- MIT 免费：无商业许可风险
- 统一 API：命名空间 `NewLife.Office`，各格式 API 风格一致，学习成本低

> CSV 读写能力由 `NewLife.Core` 的 `CsvFile` 提供，本仓库不重复实现。

## 亮点

- **全格式覆盖**：四大核心格式 + 十种轻量格式 + OLE2/BIFF8 老式格式，一库搞定
- **零外部依赖**：直接操作文件底层结构，包体积 <500KB
- **高保真往返**：L2/L3 级保真架构（模型 + RawXml/部件透传），读→改→写不丢格式
- **Markdown 全格式转换中枢**：MD↔docx/PDF/HTML/Excel/PPT/Word 双向转换，为 AI 知识库喂入提供统一入口
- **渲染扩展包**：`NewLife.Office.Rendering`（独立 NuGet）提供文档 → 图片渲染、图片处理、二维码
- **AOT/跨平台兼容**：`net8.0` / `net10.0` 零 IL 告警，CI 覆盖 win-x64 / win-x86 / win-arm64 / linux-x64 / linux-arm64 / osx-x64 / osx-arm64 / linux-loongarch64（龙芯 LA64 经 QEMU），23 项 AOT 冒烟场景（含 31 种语言表）
- **31 种语言本地化**：异常消息、文档输出标签与程序集元数据按当前区域自动翻译，缺失译文回退英文；`Locales/*.yml` 编译期生成，NativeAOT 下同样可用

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

## 生态版图

NewLife 生态按依赖分层，各项目独立版本、独立发版。下面的版图在所有 NewLife 仓库中逐字一致；每个项目链接到代码仓库与主 NuGet 包。

| 项目 | 一句话定位 | 下载量 |
| --- | --- | --- |
| **基础底座** | | |
| [NewLife.Core](https://github.com/NewLifeX/X) | 日志、配置、缓存、网络、序列化、定时调度与 APM 追踪。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Core?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Core) |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 数据中间件：单表百亿级、自动分表、读写分离，支持 10 多种数据库。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCode?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCode) |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 高性能网络库：单机 177 万 pps、14.5 Gbps（TCP 1KB，2026-10 实测），400 万 Tcp 长连接。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Net?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Net) |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 协议通信与 RPC 框架，面向 CS 应用与物联网设备。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Remoting?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Remoting) |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 把 .NET 应用安装为操作系统守护进程：Windows 服务与 Linux systemd。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Agent?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Agent) |
| [NewLife.Templates](https://github.com/NewLifeX/NewLife.Zero) | `dotnet new` 项目模板：控制台、Web、WebApi、服务、数据与物联网。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Templates?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Templates) |
| **数据存储与消息** | | |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | Redis 客户端：微秒级延迟、百万级吞吐，内置消息队列。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Redis?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Redis) |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 纯托管 RocketMQ 客户端，支持 Apache RocketMQ 与阿里云消息队列。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.RocketMQ?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.RocketMQ) |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | MQTT 客户端与服务端，客户端支持阿里云物联网。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MQTT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MQTT) |
| [NewLife.NovaDb](https://github.com/NewLifeX/NewLife.NovaDb) | C# 混合数据库：嵌入/网络双模，关系型、时序、消息队列与 KV 于一体。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.NovaDb?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.NovaDb) |
| [NewLife.MySql](https://github.com/NewLifeX/NewLife.MySql) | MySql 驱动与数据访问辅助层，基于 NewLife 数据栈实现。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MySql?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MySql) |
| **Web、AI 与分布式** | | |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 魔方快速开发平台：用户权限、SSO、OAuth 服务端与 CRUD 脚手架。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Cube?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Cube) |
| [Stardust](https://github.com/NewLifeX/Stardust) | 星尘分布式服务平台：节点管理、APM、配置中心、注册中心与发布中心。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Stardust?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Stardust) |
| [AntJob](https://github.com/NewLifeX/AntJob) | 蚂蚁调度：分布式任务与实时数据平台，支持分片、重试与回溯补算。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AntJob?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AntJob) |
| [NewLife.AI](https://github.com/NewLifeX/NewLife.AI) | 统一接入 46 家大模型服务商，附带开箱即用的 ChatAI Web 对话应用。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AI?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AI) |
| **工业与物联网** | | |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | IoT 标准库，定义物联网通信协议契约与设备抽象。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IoT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IoT) |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | Modbus TCP/RTU/ASCII，基于 NewLife.IoT 实现。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Modbus?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Modbus) |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 西门子 PLC 协议，基于 NewLife.IoT 实现。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Siemens?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Siemens) |
| [NewLife.Omron](https://github.com/NewLifeX/NewLife.Omron) | 欧姆龙 PLC 协议，基于 NewLife.IoT 实现。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Omron?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Omron) |
| [NewLife.Melsec](https://github.com/NewLifeX/NewLife.Melsec) | 三菱 Melsec PLC 协议，基于 NewLife.IoT 实现。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Melsec?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Melsec) |
| [NewLife.OPC](https://github.com/NewLifeX/NewLife.OPC) | OPC 客户端，用于工业数据采集。 | — |
| [NewLife.BACnet](https://github.com/NewLifeX/NewLife.BACnet) | BACnet 楼宇自控协议栈。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.BACnet?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.BACnet) |
| [ZeroIoT](https://github.com/NewLifeX/ZeroIoT) | 开源轻量级 IoT 平台，含网关与设备管理。 | — |
| **车联网** | | |
| [NewLife.JT808](https://github.com/NewLifeX/NewLife.JT808) | JT/T 808 车辆定位协议：70+ 消息体，覆盖 2011/2013/2019 三个版本。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT808?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT808) |
| [NewLife.JT809](https://github.com/NewLifeX/NewLife.JT809) | JT/T 809 部标平台互联协议，面向运输管理部门。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT809?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT809) |
| [NewLife.JT1078](https://github.com/NewLifeX/NewLife.JT1078) | JT/T 1078 车载视频：RTP 解析、FLV/fMP4 转封装与直播推送。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT1078?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT1078) |
| **领域组件** | | |
| [NewLife.Office](https://github.com/NewLifeX/NewLife.Office) | 一个包覆盖 Excel、Word、PPT、PDF 等 15+ 文档格式。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Office?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Office) |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 跨平台音频：编解码、设备 IO、DSP、容器与流式处理。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Audio?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Audio) |
| [NewLife.Video](https://github.com/NewLifeX/NewLife.Video) | 视频编解码、封装格式、帧处理与流媒体。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Video?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Video) |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 百度/高德/腾讯/天地图统一接口，支持地理编码与坐标转换。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Map?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Map) |
| [NewLife.IP](https://github.com/NewLifeX/NewLife.IP) | IPv4 归属地：离线内存映射库 + 在线 API 双源。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IP?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IP) |
| **工具与产品** | | |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 码神工具：网络、串口、加解密、正则、Modbus 与 MQTT。 | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCoder?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCoder) |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 自托管 Git 源代码管理系统。 | — |
| [LuckyClover](https://github.com/NewLifeX/LuckyClover) | 从低版本宿主启动时，自动安装匹配的 .NET 运行时。 | — |

> 徽章取各项目主包；主题包、扩展包与卫星包见对应 NuGet 页面。未在表中列出的 30 多个包（FTP、LoRa、Office 扩展、各类协议驱动等）见 [NewLifeX 组织](https://github.com/NewLifeX) 与 [newlifex.com](https://newlifex.com)。

## 新生命开发团队

<img src="https://newlifex.com/logo.png" width="48" align="right" alt="NewLife" />

新生命团队成立于 2002 年，二十余年专注 .NET 与物联网技术生态，主导开源项目近百个（生态版图收录 34 个主干项目），为电力、物流、医疗、工控、互联网等行业提供软硬件一体化解决方案；以“学无先后达者为师”为信条。

开源组件在 NuGet 累计下载 <!--nuget-total:start-->6.48M<!--nuget-total:end--> 次（全生态口径，截至 2026-10）。

网站：https://newlifex.com · 开源组织：https://github.com/NewLifeX

<details>
<summary>微信公众号：智能大石头（扫码关注）</summary>

![智能大石头](https://newlifex.com/stone.jpg)

</details>
