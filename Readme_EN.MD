# NewLife.Office - Office Automation Library

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

Office automation library for .NET

[简体中文](Readme.MD) | English

Source: https://github.com/NewLifeX/NewLife.Office  
NuGet: `NewLife.Office`

## Overview

`NewLife.Office` is a comprehensive .NET office automation library covering **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS, and legacy binary formats based on OLE2/BIFF8. It is dedicated to providing the .NET community with a high-quality, zero-dependency, MIT-licensed office automation solution.

- Single dependency on `NewLife.Core` — zero external dependencies
- Targets `net45/netstandard2.0/netstandard2.1`
- Built directly on Open XML (ZIP + XML) / PDF / OLE2 file structures
- Unified `NewLife.Office` namespace with consistent API style

> CSV read/write is provided by `NewLife.Core`'s `CsvFile`; this library does not duplicate CSV.

## Supported Formats

| Format | Read | Write | Template | Conversion | Notes |
|--------|:----:|:-----:|:--------:|:----------:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | Styles, charts, pivot tables, conditional formatting, images |
| **Excel xls** | ✅ | ✅ | — | — | OLE2/BIFF8, BiffReader+BiffWriter, styles/merges/links/filters/page setup |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | Paragraphs, tables, images, headers/footers, AES encryption |
| **Word doc** | ✅ | — | — | — | MS-DOC, DocReader extracts plain text & paragraphs |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | Text boxes, images, tables, charts, masters, animations, comments, sections |
| **PPT ppt** | ✅ | — | — | — | MS-PPT, PptReader extracts slide text |
| **PDF** | ✅ | ✅ | — | — | Fluent API, merge/split, encryption, digital signatures |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF | CommonMark + GFM, AST model, →Word/→PDF |
| **RTF** | ✅ | ✅ | ✅ | — | Paragraphs, tables, template fill, image embedding (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | LibreOffice native format |
| **EML** | ✅ | ✅ | — | — | RFC 5322 + MIME, attachments, inline images |
| **MSG** | ✅ | — | — | — | Outlook, OLE2 + MAPI, subject/body/attachments |
| **iCalendar** | ✅ | ✅ | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | EPUB 2/3, OPF/Nav/chapters/cover |
| **XPS** | ✅ | ✅ | — | — | ZIP+XML, Glyphs text extraction, multi-page write, images, metadata |

## Quick Start

### Unified Text/Markdown Extraction

```csharp
using NewLife.Office;

// One-liner: extract plain text from any supported format
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Extract as Markdown (preserves structural information)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");

// Extract from a stream
using var fs = File.OpenRead("data.pdf");
String? pdfText = OfficeFactory.ReadText(fs, ".pdf");
```

You can also use the interfaces directly:

```csharp
using var reader = new WordReader("document.docx");
String? text = (reader as ITextExtractable)?.ExtractText();
String? md = (reader as IMarkdownExtractable)?.ExtractMarkdown();
```

### Excel Read/Write

```csharp
using NewLife.Office;

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);

writer.WriteHeader("Data", new[] { "ID", "Name", "Amount" }, CellStyle.Header);
writer.WriteRow("Data", new Object?[] { 1, "John", 1234.56m }, new CellStyle
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

### Object Collection Export

```csharp
public class UserInfo
{
    [DisplayName("ID")]
    public Int32 Id { get; set; }

    [DisplayName("Name")]
    public String Name { get; set; }

    [DisplayName("Amount")]
    public Decimal Amount { get; set; }
}

var list = new List<UserInfo>
{
    new() { Id = 1, Name = "John", Amount = 1234.56m },
    new() { Id = 2, Name = "Jane", Amount = 9876.00m },
};

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);
writer.WriteObjects("Users", list);
writer.Save();
```

### Object Collection Import

```csharp
using var fs = File.OpenRead("users.xlsx");
var reader = new ExcelReader(fs);
var users = reader.ReadObjects<UserInfo>("Users");
```

### Template Fill

```csharp
// Replace {{Key}} placeholders in an xlsx template, preserving styles
var template = new ExcelTemplate("report_template.xlsx");
template.Fill(new Dictionary<String, Object>
{
    ["Title"] = "2025 Annual Report",
    ["Author"] = "John",
    ["Date"] = DateTime.Today.ToString("yyyy-MM-dd"),
});
template.Save("report_output.xlsx");
```

### Cell Styles

```csharp
var style = new CellStyle
{
    Bold = true,
    FontColor = "FFFFFF",        // White text
    BackgroundColor = "4472C4",  // Blue background
    HAlign = HorizontalAlignment.Center,
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00",
};

writer.WriteHeader("Sheet1", new[] { "ID", "Name", "Amount" }, style);
writer.MergeCell("Sheet1", "A1:C1");  // Merge header row
writer.SetRowHeight("Sheet1", 1, 30); // Set row height
```

## Excel Capabilities

### ExcelWriter

- Multi-sheet writing, headers and multi-row data
- Cell styles: font/background/border/alignment/wrap/number format
- Layout: column width, row height, merged cells, freeze panes, auto-filter
- Advanced: hyperlinks, dropdown validation, conditional formatting, images, charts (column/line/pie/scatter), pivot tables
- Page: orientation, paper size, margins, headers/footers, print title rows
- Security: sheet protection
- Convenience: `WriteObjects<T>()`, `WriteDataTable()`

### ExcelReader

- Read multi-sheet row data (auto type conversion)
- `ReadObjects<T>()` object mapping import (supports `DisplayName/Description`)
- `ReadDataTable()` import as `DataTable`
- `GetMergeRanges()` read merged regions

### ExcelTemplate

- `{{Key}}` placeholder replacement on existing xlsx templates
- Output to file or stream

### BiffReader (xls Safe Read)

- OLE2/CFB container-based BIFF8 worksheet stream parsing
- `ReadSheet()` reads worksheet row data
- `ReadObjects<T>()` object mapping import
- Supports string, number, boolean, date, RK-encoded cell types

### BiffWriter (xls Write)

- Generate BIFF8 files with styles (font/fill/border/color), merged cells, row height
- Hyperlinks, auto-filter, page setup, headers/footers, freeze panes
- Formula writing, column width, sheet protection

## Word Capabilities

### WordWriter / WordReader / WordTemplate

- Full docx read/write: paragraphs, tables, images, headings, lists, hyperlinks, headers/footers, page settings
- Template fill: `{{Key}}` placeholder replacement, preserving original styles
- `WordHtmlConverter`: docx to semantic HTML (paragraph→`<p>`, table→`<table>`, heading→`<h1>`-`<h6>`)
- `WordPdfConverter`: docx to PDF (content-mapping style, via PdfFluentDocument)
- `WordEncryptor`: AES-128 encrypt docx (ECMA-376 Agile Encryption, OLE2 CFB container)
- `DocReader`: read .doc format (OLE2 container, extract plain text and paragraph structure)

## PowerPoint Capabilities

### PptxWriter / PptxReader / PptxTemplate

- Full pptx read/write: slide management, text boxes, tables, images, basic shapes, hyperlinks, notes, template fill
- Layout design: masters, layouts, backgrounds, dimensions, slide transitions
- `PptReader`: read .ppt format (OLE2 container, extract slide text)

## PDF Capabilities

### PdfWriter / PdfReader / PdfDocument

- Create PDFs: text layout, tables, images, headers/footers, CJK font support (Type0/STSong-Light, UniGB-UCS2-H encoding)
- `PdfFluentDocument`: declarative layout API (Row/Column/Text/Image/Table), auto-pagination
- PDF operations: merge/split, watermark, page rotation, text overlay, clickable links, bookmark navigation
- `PdfEncryptor`: encryption/permission control (RC4/AES-128/AES-256)
- `PdfSigner`: digital signatures (PKCS#7 detached signature + visible signature field, manual ASN.1 DER, no external deps)
- Text extraction: per-page/full document with coordinate information

## Markdown / RTF / ODS

### Markdown

- `MarkdownParser`: parse CommonMark + GFM (headings/paragraphs/code blocks/lists/blockquotes/bold+italic/images/tables/task lists)
- `MarkdownWriter`: serialize `MarkdownDocument` to .md text (including GFM extensions)
- `MarkdownHtmlConverter`: Markdown to semantic HTML
- `MarkdownWordConverter`: Markdown → Word (docx), headings/paragraphs/lists/tables mapped to Word styles
- `MarkdownPdfConverter`: Markdown → PDF (direct via PdfFluentDocument)
- `MarkdownDocument.ToWord()` / `ToPdf()` convenience methods

### RTF

- `RtfReader`: parse RTF files, extract paragraphs (font/color/alignment), tables, document properties, embedded images (PNG/JPEG/WMF/EMF)
- `RtfWriter`: create RTF files (paragraph formatting/tables/template fill); `AddImage()` supports embedding PNG/JPEG (hex-encoded)

### ODS

- `OdsReader`: read LibreOffice ODS format, extract multi-sheet cell data, merged cells, cell styles (font/color/background/alignment)
- `OdsWriter`: generate ODS files (multi-sheet, string/number/date/boolean/formula, cell style writing)

## Email / Calendar / Contacts / eBooks

### EML (Email)

- `EmlReader`: parse RFC 5322 + MIME emails, including multipart/mixed, multipart/alternative
- `EmlWriter`: generate valid MIME emails (text/HTML body, attachments, inline images)
- RFC 2047 Base64/Q-Encoding header decoding, quoted-printable body decoding

### MSG (Outlook Email)

- `MsgReader`: parse Outlook .msg files via OLE2/CFB container, extract MAPI properties
- Supports subject (PR_SUBJECT), sender, recipients (To/Cc/Bcc)
- Supports plain text body (PR_BODY) and HTML body (PR_HTML_BODY)
- Supports multi-attachment extraction (`__attach_version1.0_#NN` sub-storages, with filename and binary data)

### iCalendar

- `ICalReader`: parse RFC 5545 .ics files, extract VEVENT/VTODO/VJOURNAL with recurrence rules
- `ICalWriter`: generate RFC 5545-compliant .ics files, RRULE serialization and 75-byte line folding

### vCard

- `VCardReader`: parse RFC 6350 .vcf files, supports vCard 3.0/4.0, FN/N/ORG/TEL/EMAIL/ADR/BDAY/PHOTO
- `VCardWriter`: generate vCard 4.0 format, multi-contact batch writing

### EPUB

- `EpubReader`: read EPUB 2/3 files (ZIP + OPF + XHTML), extract metadata and chapter content
- `EpubWriter`: generate EPUB 3 compliant eBooks (mimetype/container.xml/OPF/Nav/chapter XHTML/CSS/cover image)

## XPS

### XpsReader / XpsWriter

- `XpsReader`: parse XPS (XML Paper Specification) files (ZIP + XML), traverse FixedDocumentSequence → FixedDocument → FixedPage
- Text extraction: from `<Glyphs UnicodeString="...">` elements
- Image extraction: `ExtractImages()` returns all embedded PNG/JPEG resources
- Page info: `Width`/`Height` from `<FixedPage>` attributes
- Metadata: `ReadProperties()` parses `docProps/core.xml` (title/author/subject)
- `XpsWriter`: generate valid XPS files (`[Content_Types].xml` + `_rels/.rels` + FDS + .fpage)
- `AddPage(width, height, IEnumerable<(text, x, y, fontSize)>)` multi-page write
- `SetProperties(XpsProperties)` write document properties
- `AddImage(name, data, mime)` embed image resources
- `Save(path|stream)` / `ToBytes()` output

## OLE2/CFB Container

- `CfbReader`: parse Compound File Binary container (safe Stream lookup, v3/v4 support)
- `CfbWriter`: generate valid CFB files (write streams and directory tree)
- Shared foundation for xls (Excel 97-2003), doc (Word 97-2003), ppt (PPT 97-2003)

## Test Status

- Total tests: `694`
- Passed: `694`
- Failed: `0`

## Code Structure

```
NewLife.Office/
├── Excel/        # xlsx read/write: ExcelWriter / ExcelReader / ExcelTemplate / CellStyle / ExcelHelper / BiffReader
├── Word/         # docx read/write: WordWriter / WordReader / WordTemplate / DocReader / WordHtmlConverter / WordPdfConverter
├── PPT/          # pptx read/write: PptxWriter / PptxReader / PptxTemplate / PptReader
├── PDF/          # PDF read/write: PdfWriter / PdfReader / PdfDocument / PdfFluentDocument / PdfEncryptor
├── OLE2/         # OLE2/CFB container: CfbReader / CfbWriter
├── Markdown/     # Markdown read/write: MarkdownParser / MarkdownWriter / MarkdownHtmlConverter
├── Rtf/          # RTF read/write: RtfReader / RtfWriter
├── Ods/          # ODS read/write: OdsReader / OdsWriter
├── Mail/         # EML email read/write: EmlReader / EmlWriter / EmlMessage
├── Calendar/     # iCalendar .ics read/write: ICalReader / ICalWriter / ICalDocument
├── VCard/        # vCard .vcf read/write: VCardReader / VCardWriter / VCardContact
├── Epub/         # EPUB read/write: EpubReader / EpubWriter / EpubDocument
└── Xps/          # XPS read/write: XpsReader / XpsWriter
```

## Roadmap

The following features are planned but not yet implemented:

### PPT → Image

- Render each slide as PNG/JPEG image (requires external rendering library such as SkiaSharp/Docnet.Core; currently throws `NotSupportedException`)

### PDF Page Rendering to Image

- Render each PDF page as PNG/JPEG (requires external rendering library)

### ODS Merged Cells

- OdsReader recognize merged cells (number-columns-spanned)

## NewLife Project Ecosystem

All projects target net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5 by default. Legacy builds (2024.0801) support net4.0/net2.0.

| Project | Year | Description |
|:-------:|:----:|-------------|
| **Core Libraries** | | Foundation for middleware and products |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Core library: logging, config, cache, networking, serialization, APM tracing |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Big data middleware, 10B+ rows per table, MySql/SQLite/SqlServer/Oracle/PostgreSql/DaMeng, auto-sharding, read/write separation |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Network library, 22.66M tps throughput, 4M concurrent TCP connections |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | RPC framework, HTTP/RPC communication, high throughput, IoT-friendly |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Rapid development platform, user/role/permission, SSO, OAuth, proven at 10B+ row scale |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Service management, install apps as OS daemons (Windows Service, Linux systemd) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero scaffolding, project templates for Web, WebApi, Service |
| **Middleware** | | Integrations with major platforms |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis client, microsecond latency, million-level throughput, rich message queues |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Pure managed RocketMQ client, Apache RocketMQ & Alibaba Cloud MQ |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | IoT messaging protocol, MqttClient/MqttServer, Alibaba Cloud IoT support |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT standard library, protocol definitions for IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, IoT platform & edge gateway |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Siemens PLC protocols, IoT platform & edge gateway |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Map component library, Baidu/Amap/Tencent/Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Audio codec library, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Platforms & Products** | | Ready-to-deploy solutions |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Distributed service platform, node management, APM, config/registry/publish center |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Distributed big data computing platform (real-time/batch), trillion-row proven |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | Enterprise ERP: product, customer, sales, supplier management |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Developer toolset: network, serial, crypto, regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Simple file storage for distributed systems |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Production-grade reverse proxy, NAT proxy, HTTP proxy |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP stress testing tool |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git source code management system |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Embedded operating system, ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Embedded industrial computer, IoT edge gateway, .NET 8 |
| **FIoT IoT Platform** | 2020 | End-to-end IoT solution: construction, environmental, agriculture, 100K+ device points |
| **UWB Indoor Positioning** | 2020 | Centimeter-level (10-20cm) high-precision indoor positioning |

## NewLife Development Team

![NewLife](https://newlifex.com/logo.png)

The NewLife team, founded in 2002, is a next-generation IoT solutions provider dedicated to software/hardware consulting, system architecture planning, and development services.

Over 80 open-source projects led by the team are widely used across industries, with NuGet downloads exceeding 4 million. Our big data middleware NewLife.XCode, distributed scheduling platform AntJob, distributed platform Stardust, cache queue component NewLife.Redis, and FIoT IoT platform have been successfully deployed in power utilities, universities, internet, telecom, transportation, logistics, industrial control, healthcare, and cultural heritage sectors.

`The NewLife team was founded in 2002. Some open-source projects have a history spanning over 20 years, with source repositories preserving all modification records since 2010.`

Website: <https://newlifex.com>  
Open Source: <https://github.com/newlifex>  
QQ Group: 1600800 / 1600838  
WeChat Official Account:  
![Smart Stone](https://newlifex.com/stone.jpg)
