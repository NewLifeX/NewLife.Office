# NewLife.Office - Office Automation Library

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
🌐 [简体中文](README.zh-CN.md) | [繁體中文](README.zh-TW.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

A comprehensive .NET office automation library covering Excel, Word, PPT, PDF and 15+ document formats in one package


Source: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (Rendering package: NewLife.Office.Rendering)  

## Overview

`NewLife.Office` is a comprehensive .NET office automation library covering **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS, and legacy binary formats based on OLE2/BIFF8, backed by the **Rendering** extension package (documents → images).

- Zero external dependencies: only depends on `NewLife.Core`, built directly on Open XML (ZIP + XML) / PDF / OLE2 file structures
- Full framework support: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, with AOT compatibility analysis (IsAotCompatible, zero IL warnings) on net8.0+
- MIT licensed: free with no commercial license risk
- Unified API: `NewLife.Office` namespace with a consistent API style, low learning cost

> CSV read/write is provided by `CsvFile` in `NewLife.Core`; this library does not duplicate CSV.

## Highlights

- **Full format coverage**: four core formats + ten lightweight formats + OLE2/BIFF8 legacy formats, all in one library
- **Zero external dependencies**: operates directly on file structures, package size <500KB
- **High-fidelity roundtrip**: L2/L3 fidelity architecture (model + RawXml/part passthrough), read → modify → write without losing formatting
- **Markdown conversion hub**: bidirectional MD↔docx/PDF/HTML/Excel/PPT/Word conversion, a unified entry for AI knowledge base ingestion
- **Rendering extension**: `NewLife.Office.Rendering` (separate NuGet) provides document → image rendering, image processing, and QR codes
- **AOT/cross-platform**: zero IL warnings on `net8.0` / `net10.0`; CI covers win-x64 / win-x86 / win-arm64 / linux-x64 / linux-arm64 / osx-x64 / osx-arm64 / linux-loongarch64 (LoongArch via QEMU), with 23 AOT smoke scenarios including all 31 language tables
- **31-language localization**: exception messages, document output labels and assembly metadata are translated for the current culture and fall back to English; `Locales/*.yml` are compiled into the assembly and work under NativeAOT

## Supported Formats

| Format | Read | Write | Template | Conversion | Render | Notes |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Styles, charts, pivot tables, conditional formatting, images, named ranges, structured tables, comments, sparklines, slicers, threaded comments, hidden rows/columns, streaming read and 1904 date support |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter read/write, styles/merges/hyperlinks/filters/page setup |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Paragraphs, tables, images, headers/footers, comments, content controls, AES encryption, mail merge, find/replace (including text boxes), footnote model, multiple sections; cross-validated with NPOI+OpenXmlValidator |
| **Word doc** | ✅ | — | — | ✅ doc→docx | — | MS-DOC, DocReader plain text + paragraph formatting extraction (bold/font size/alignment, etc.) + doc→docx conversion |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Text boxes, images, tables, charts, masters, animations, comments, sections; cross-file merge (with chart/comment/notes fidelity) |
| **PPT ppt** | ✅ | ✅ | — | — | — | MS-PPT: PptReader text extraction + PptWriter text slide writing |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, merge/split, encryption, signatures, forms, annotations, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/all→MD | — | CommonMark + GFM + typed AST, full-format conversion hub |
| **RTF** | ✅ | ✅ | ✅ | — | — | Paragraphs, tables, template fill, image embedding (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | LibreOffice native format, merged cells, styles, formulas |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, attachments, inline images |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, subject/body/attachments |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/chapters/cover |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, Glyphs text extraction, multi-page write, images, metadata |
| **Rendering** | — | — | — | — | ✅ | Separate package (SkiaSharp): PDF/Word/PPT/Markdown → images |

## Quick Start

### Unified Text / Markdown Extraction

```csharp
using NewLife.Office;

// One-liner: extract plain text from any supported format
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Extract as Markdown (preserves structure)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");

// doc (97-2003) → docx conversion
OfficeFactory.ConvertDocToDocx("legacy.doc", "legacy.docx");
```

### Command-Line Tool (NewLife.Office.Cli)

```bash
# Any office document → Markdown
dotnet run --project NewLife.Office.Cli -- convert input.docx output.md

# doc (97-2003) → docx (legacy document migration)
dotnet run --project NewLife.Office.Cli -- doc2docx legacy.doc legacy.docx
```

### Excel Write and Read

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

> More capabilities (object mapping import/export, template fill, format conversion, rendering, etc.) see [Requirements](Doc/需求文档.md) and the source examples.

## Documentation

- [Requirements](Doc/需求文档.md)
- [Feature List](Doc/功能清单.md)
- [Architecture](Doc/架构设计.md)
- [Competitive Analysis](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## Ecosystem Map

NewLife is a layered ecosystem of independently versioned open-source projects. The map below is shared verbatim by every NewLife repository; each project links to its repository and its primary NuGet package.

| Project | What it does | Downloads |
| --- | --- | --- |
| **Foundation** | | |
| [NewLife.Core](https://github.com/NewLifeX/X) | Logging, configuration, caching, networking, serialization, scheduling and APM tracing. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Core?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Core) |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | Data middleware: single tables at the tens-of-billions scale, automatic sharding, read-write splitting, 10+ databases. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCode?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCode) |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | High-performance network library: 1.77M pps / 14.5 Gbps on a single machine (TCP 1KB, 2026-10), 4M long-lived TCP connections. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Net?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Net) |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | Protocol communication and RPC framework for client/server applications and IoT devices. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Remoting?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Remoting) |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | Installs .NET applications as OS daemons: Windows services and Linux systemd. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Agent?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Agent) |
| [NewLife.Templates](https://github.com/NewLifeX/NewLife.Zero) | `dotnet new` project templates for console, web, API, service, data and IoT projects. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Templates?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Templates) |
| **Data & Messaging** | | |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | Redis client with microsecond latency, millions of ops/s and built-in message queues. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Redis?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Redis) |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | Pure managed RocketMQ client for Apache RocketMQ and Alibaba Cloud Message Queue. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.RocketMQ?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.RocketMQ) |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | MQTT client and server, including Alibaba Cloud IoT connectivity. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MQTT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MQTT) |
| [NewLife.NovaDb](https://github.com/NewLifeX/NewLife.NovaDb) | Hybrid database in C#: embedded or networked, relational, time-series, message queue and KV in one. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.NovaDb?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.NovaDb) |
| [NewLife.MySql](https://github.com/NewLifeX/NewLife.MySql) | MySql driver and helper layer built on the NewLife data stack. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MySql?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MySql) |
| **Web, AI & Distributed** | | |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | Rapid admin-platform framework: permissions, SSO, OAuth server and CRUD scaffolding. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Cube?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Cube) |
| [Stardust](https://github.com/NewLifeX/Stardust) | Distributed service platform: node management, APM, configuration center, registry and release center. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Stardust?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Stardust) |
| [AntJob](https://github.com/NewLifeX/AntJob) | Distributed task and real-time data scheduling platform with sharding, retries and backfill. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AntJob?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AntJob) |
| [NewLife.AI](https://github.com/NewLifeX/NewLife.AI) | Unified client for 46 LLM providers plus ChatAI, a ready-to-run web chat application. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AI?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AI) |
| **Industrial & IoT** | | |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | IoT standard library defining protocol contracts and device abstractions. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IoT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IoT) |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | Modbus TCP/RTU/ASCII implementation on top of NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Modbus?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Modbus) |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | Siemens PLC protocol on top of NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Siemens?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Siemens) |
| [NewLife.Omron](https://github.com/NewLifeX/NewLife.Omron) | Omron PLC protocol on top of NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Omron?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Omron) |
| [NewLife.Melsec](https://github.com/NewLifeX/NewLife.Melsec) | Mitsubishi Melsec PLC protocol on top of NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Melsec?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Melsec) |
| [NewLife.OPC](https://github.com/NewLifeX/NewLife.OPC) | OPC client for industrial data access. | — |
| [NewLife.BACnet](https://github.com/NewLifeX/NewLife.BACnet) | BACnet building-automation protocol stack. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.BACnet?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.BACnet) |
| [ZeroIoT](https://github.com/NewLifeX/ZeroIoT) | Lightweight open-source IoT platform with gateway and device management. | — |
| **Vehicle Telematics** | | |
| [NewLife.JT808](https://github.com/NewLifeX/NewLife.JT808) | JT/T 808 vehicle positioning protocol: 70+ message bodies, 2011/2013/2019. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT808?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT808) |
| [NewLife.JT809](https://github.com/NewLifeX/NewLife.JT809) | JT/T 809 platform-to-platform protocol for transport authorities. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT809?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT809) |
| [NewLife.JT1078](https://github.com/NewLifeX/NewLife.JT1078) | JT/T 1078 in-vehicle video: RTP parsing, FLV/fMP4 remuxing and live streaming. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT1078?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT1078) |
| **Domain Components** | | |
| [NewLife.Office](https://github.com/NewLifeX/NewLife.Office) | One package for Excel, Word, PPT, PDF and 15+ document formats. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Office?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Office) |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | Cross-platform audio: codecs, device I/O, DSP, containers and streaming. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Audio?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Audio) |
| [NewLife.Video](https://github.com/NewLifeX/NewLife.Video) | Video codecs, container formats, frame processing and streaming. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Video?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Video) |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | Unified interface over Baidu, AMap, Tencent and Tianditu with geocoding and coordinate conversion. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Map?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Map) |
| [NewLife.IP](https://github.com/NewLifeX/NewLife.IP) | IPv4 geolocation with an offline memory-mapped database and an online API fallback. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IP?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IP) |
| **Tools & Products** | | |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | Developer toolbox: network, serial port, crypto, regex, Modbus and MQTT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCoder?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCoder) |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | Self-hosted Git source code management system. | — |
| [LuckyClover](https://github.com/NewLifeX/LuckyClover) | Installs the matching .NET runtime automatically when launched from an older host. | — |

> Badges show the primary package of each project; theme, extension and satellite packages are listed on their NuGet pages.The 30+ packages not listed here (FTP, LoRa, Office extensions, protocol drivers and more) are available under the [NewLifeX organization](https://github.com/NewLifeX) and on [newlifex.com](https://newlifex.com).

## Team

<img src="https://newlifex.com/logo.png" width="48" align="right" alt="NewLife" />

Founded in 2002, the NewLife team has spent more than two decades on the .NET and IoT ecosystem, maintains nearly 100 open-source projects (34 of them core to this map), and delivers integrated software and hardware solutions for power, logistics, healthcare, industrial control and internet industries. We live by the motto "learning has no seniority — mastery earns respect".

Open-source components have been downloaded <!--nuget-total:start-->6.48M<!--nuget-total:end--> times from NuGet (ecosystem-wide, as of 2026-10).

Website: https://newlifex.com · Open source organization: https://github.com/NewLifeX

<details>
<summary>WeChat Official Account: 智能大石头 (scan to follow)</summary>

![智能大石头](https://newlifex.com/stone.jpg)

</details>
