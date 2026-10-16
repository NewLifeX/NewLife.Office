# NewLife.Office - Office Automation Library

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

A comprehensive .NET office automation library covering Excel, Word, PPT, PDF and 15+ document formats in one package

[简体中文](Readme.MD) | English | [繁體中文](Readme_TW.MD) | [日本語](Readme_JA.MD) | [한국어](Readme_KO.MD) | [Español](Readme_ES.MD) | [Français](Readme_FR.MD) | [Deutsch](Readme_DE.MD) | [Português](Readme_PT.MD) | [Русский](Readme_RU.MD)

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

## Supported Formats

| Format | Read | Write | Template | Conversion | Render | Notes |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Styles, charts, pivot tables, conditional formatting, images, named ranges, structured tables |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter read/write, styles/merges/hyperlinks/filters/page setup |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Paragraphs, tables, images, headers/footers, comments, content controls, AES encryption, mail merge |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader extracts plain text |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Text boxes, images, tables, charts, masters, animations, comments, sections; cross-file merge |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader extracts slide text |
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

## NewLife Project Ecosystem

All projects target net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5 by default. Legacy builds (2024.0801) support net4.0/net2.0.

| Project | Year | Description |
|:-------:|:----:|-------------|
| **Core Libraries** | | Foundation supporting middleware and products |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Core library: logging, configuration, caching, networking, serialization, APM tracing |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Big data middleware, 10B+ rows per table, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, auto-sharding, read/write separation |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Network library, 22.66M tps on a single machine, 4M TCP long connections |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Protocol communication library, CS application communication framework, HTTP/RPC, high throughput, low-cost IoT device access |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Rapid development platform, user permissions, SSO login, OAuth server, proven at 10B+ row scale |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Service management component, installs apps as OS daemons (Windows Service, Linux systemd) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero scaffolding, NewLife.Templates project templates, Web, WebApi, Service |
| **Middleware** | | Integrations with major middleware platforms |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis client, microsecond latency, million-level throughput, rich message queues, validated at 10B+ scale |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Pure managed RocketMQ client, Apache RocketMQ and Alibaba Cloud MQ, validated in 1B+ projects |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | IoT messaging protocol, MqttClient/MqttServer, Alibaba Cloud IoT support |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT standard library, defines communication protocol standards for IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, based on IoT standard library, ZeroIoT platform and IoTEdge gateway |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Siemens PLC protocols, based on IoT standard library, IoT platform and IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Map component library, wrapping Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Audio codec library, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Platforms & Products** | | Production-ready, compile-deploy-and-use |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Distributed service platform: node management, APM monitoring, config center, registry, publish center |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Distributed big data computing platform (real-time/offline), ant-migration sharding, validated at trillion-row scale |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | Enterprise ERP: product, customer, sales, supplier management |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Developer toolbox: networking, serial, crypto, regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Simple file storage for centralized file storage in distributed systems |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Production-grade reverse proxy, NAT proxy, HTTP proxy |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP stress testing tool |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git source code management system |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Embedded operating system, fully independent, ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Embedded industrial computer, IoT edge gateway, high-performance .NET 8 host, for industrial/agricultural/transportation/medical |
| **FIoT IoT Platform** | 2020 | End-to-end IoT solution: construction, environmental, agriculture, software/hardware + big data analytics, validated at 100K+ device points |
| **UWB Indoor Positioning** | 2020 | Centimeter-level (10-20cm) high-precision indoor positioning, integrated with other systems, validated in large exhibition projects |

## NewLife Development Team

![NewLife](https://newlifex.com/logo.png)

The NewLife team, founded in 2002, is a next-generation IoT industry solution provider, dedicated to software/hardware solution consulting, system architecture planning, and development services.

More than 80 open-source projects led by the team are widely used across industries, with cumulative NuGet downloads exceeding 4 million. The big data middleware NewLife.XCode, the distributed scheduling platform AntJob, the distributed platform Stardust, the cache queue component NewLife.Redis, and the FIoT IoT platform have been successfully applied in power utilities, universities, internet, telecom, transportation, logistics, industrial control, medical, and cultural heritage industries, delivering advanced, reliable, secure, high-quality, and easily extensible products and system integration services.

We will continue to improve our services to become a long-term trusted partner for our customers, and through continuous innovation and development, become an excellent IoT service provider in China.

`The NewLife team was founded in 2002. Some open-source projects have a history of over 20 years, and the source repositories preserve all modification records since 2010.`

Website: <https://newlifex.com>  
Open Source: <https://github.com/newlifex>  
QQ Group: 1600800 / 1600838  
WeChat Official Account:  
![Smart Stone](https://newlifex.com/stone.jpg)
