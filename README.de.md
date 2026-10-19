# NewLife.Office - Büroautomatisierungsbibliothek

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)


![GitHub License](https://img.shields.io/github/license/NewLifeX/NewLife.Office?logo=github)
![NuGet Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![NuGet](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![NuGet prerelease](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)
[![test](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml/badge.svg?branch=master)](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml)
![platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS%20%7C%20Android%20%7C%20iOS-blue)
![arch](https://img.shields.io/badge/arch-x64%20%7C%20arm64%20%7C%20loongarch64%20%7C%20riscv64%20%7C%20s390x-orange)
[![contributors](https://img.shields.io/github/contributors/NewLifeX/NewLife.Office?logo=github)](https://github.com/NewLifeX/NewLife.Office/graphs/contributors)

<!-- lang:start -->
🌐 [中文](Readme.MD) | [繁體中文](README.zh-TW.md) | [English](README.en.md) | [Español](README.es.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

Eine umfassende .NET-Büroautomatisierungsbibliothek, die Excel, Word, PPT, PDF und mehr als 15 Dokumentformate in einem Paket abdeckt


Quellcode: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (Render-Erweiterungspaket: NewLife.Office.Rendering)  

## Überblick

`NewLife.Office` ist eine umfassende .NET-Büroautomatisierungsbibliothek, die **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, E-Mail (eml/msg), iCalendar, vCard, EPUB, XPS sowie die auf OLE2/BIFF8 basierenden Legacy-Binärformate abdeckt, ergänzt durch das **Rendering**-Erweiterungspaket (Dokumente → Bilder).

- Null externe Abhängigkeiten: hängt nur von `NewLife.Core` ab, direkt auf Open-XML- (ZIP + XML) / PDF- / OLE2-Dateistrukturen aufgebaut
- Volle Framework-Unterstützung: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, mit AOT-Kompatibilitätsanalyse (IsAotCompatible, keine IL-Warnungen) ab net8.0
- MIT-Lizenz: kostenlos, ohne kommerzielles Lizenzrisiko
- Einheitliche API: Namespace `NewLife.Office`, konsistenter API-Stil, geringe Lernkosten

> CSV-Lesen/-Schreiben wird von `CsvFile` in `NewLife.Core` bereitgestellt; diese Bibliothek dupliziert CSV nicht.

## Highlights

- **Vollständige Formatauswahl**: vier Kernformate + zehn Leichtgewichtsformate + OLE2/BIFF8-Legacyformate, alles in einer Bibliothek
- **Null externe Abhängigkeiten**: arbeitet direkt auf Dateistrukturen, Paketgröße <500KB
- **Hochpräziser Roundtrip**: L2/L3-Treuearchitektur (Modell + RawXml/Teil-Durchreichung), Lesen → Ändern → Schreiben ohne Formatverlust
- **Markdown-Konvertierungszentrum**: bidirektionale Konvertierung MD↔docx/PDF/HTML/Excel/PPT/Word, ein einheitlicher Einstieg für die KI-Wissensdatenbank-Aufnahme
- **Render-Erweiterung**: `NewLife.Office.Rendering` (separates NuGet) bietet Dokument→Bild-Rendering, Bildverarbeitung und QR-Codes

## Unterstützte Formate

| Format | Lesen | Schreiben | Vorlage | Konvertierung | Rendern | Hinweise |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Stile, Diagramme, Pivot-Tabellen, bedingte Formatierung, Bilder, benannte Bereiche, strukturierte Tabellen |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter, Stile/Zusammenführungen/Hyperlinks/Filter/Seiteneinrichtung |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Absätze, Tabellen, Bilder, Kopf-/Fußzeilen, Kommentare, Inhaltssteuerelemente, AES-Verschlüsselung, Seriendruck |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader extrahiert Klartext |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Textfelder, Bilder, Tabellen, Diagramme, Master, Animationen, Kommentare, Abschnitte; dateiübergreifendes Zusammenführen |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader extrahiert Folientext |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, Zusammenführen/Aufteilen, Verschlüsselung, Signaturen, Formulare, Anmerkungen, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/alle→MD | — | CommonMark + GFM + typisiertes AST, Formatkonvertierungszentrum |
| **RTF** | ✅ | ✅ | ✅ | — | — | Absätze, Tabellen, Vorlagenbefüllung, Bildeinbettung (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | Natives LibreOffice-Format, verbundene Zellen, Stile, Formeln |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, Anhänge, Inline-Bilder |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, Betreff/Text/Anhänge |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/Kapitel/Cover |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, Glyphs-Textauszug, Mehrseiten-Schreiben, Bilder, Metadaten |
| **Rendering** | — | — | — | — | ✅ | Separates Paket (SkiaSharp): PDF/Word/PPT/Markdown → Bilder |

## Schnellstart

### Einheitliche Text- / Markdown-Extraktion

```csharp
using NewLife.Office;

// Eine Zeile: Klartext aus jedem unterstützten Format extrahieren
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Als Markdown extrahieren (behält Strukturinformationen)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Excel schreiben und lesen

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

> Weitere Funktionen (Objektzuordnungs-Import/-Export, Vorlagenbefüllung, Formatumwandlung, Rendering usw.) siehe [Anforderungen](Doc/需求文档.md) und die Quellcode-Beispiele.

## Dokumentation

- [Anforderungen](Doc/需求文档.md)
- [Funktionsliste](Doc/功能清单.md)
- [Architektur](Doc/架构设计.md)
- [Wettbewerbsanalyse](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## NewLife-Projektökosystem

Alle Projekte unterstützen standardmäßig net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5. Alte Versionen (2024.0801) unterstützen net4.0/net2.0.

| Projekt | Jahr | Beschreibung |
|:---:|:---:|------|
| **Basis-Komponenten** | | Grundlage für andere Middleware und Produkte |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Kernbibliothek: Protokollierung, Konfiguration, Cache, Netzwerk, Serialisierung, APM-Tracing |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Big-Data-Middleware, 10 Mrd.+ Zeilen pro Tabelle, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, automatische Sharding, Lese-/Schreibtrennung |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Netzwerkbibliothek, 22,66 Mio. tps auf einer Maschine, 4 Mio. TCP-Langverbindungen |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Protokoll-Kommunikationsbibliothek, CS-Framework, HTTP/RPC, hoher Durchsatz, kostengünstiger IoT-Zugang |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Schnellentwicklungsplattform, Benutzerrechte, SSO-Login, OAuth-Server, validiert bei 10 Mrd. Zeilen |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Dienstverwaltungskomponente, installiert Apps als OS-Daemons (Windows-Dienst, Linux-systemd) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero-Gerüst, Projektvorlagen NewLife.Templates, Web, WebApi, Service |
| **Middleware** | | Integration mit bekannten Middleware-Plattformen |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis-Client, Mikrosekunden-Latenz, Millionendurchsatz, umfangreiche Nachrichtenwarteschlangen, validiert bei 10 Mrd. |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Reiner verwalteter RocketMQ-Client, Apache RocketMQ und Alibaba Cloud MQ, validiert in Milliardenprojekten |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | IoT-Messaging-Protokoll, MqttClient/MqttServer, Alibaba-Cloud-IoT-Unterstützung |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT-Standardbibliothek, definiert Kommunikationsprotokollstandards für IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, basierend auf IoT-Standardbibliothek, ZeroIoT-Plattform und IoTEdge-Gateway |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Siemens-PLC-Protokolle, basierend auf IoT-Standardbibliothek, IoT-Plattform und IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Karten-Komponentenbibliothek, Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Audio-Codec-Bibliothek, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Produktplattformen** | | Produktplattform-Level, kompilieren, bereitstellen, nutzen |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Verteilte Dienstplattform: Knotenverwaltung, APM-Überwachung, Konfigurationszentrum, Registrierung, Veröffentlichungszentrum |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Verteilte Big-Data-Berechnungsplattform (Echtzeit/Offline), Ameisenumzug-Sharding, validiert bei Billionen |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | Unternehmens-ERP: Produkt-, Kunden-, Vertriebs-, Lieferantenverwaltung |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Entwickler-Werkzeugkasten: Netzwerk, Seriell, Verschlüsselung, Regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Einfacher Dateispeicher, zentrale Dateispeicherung in verteilten Systemen |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Produktionsreifer Reverse-Proxy, NAT-Proxy, HTTP-Proxy |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP-Lasttest-Tool |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git-Quellcode-Verwaltungssystem |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Eingebettetes Betriebssystem, völlig unabhängig, ARM-Cortex-M-Architektur |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Eingebetteter Industriecomputer, IoT-Edge-Gateway, leistungsstarker .NET-8-Host, für Industrie/Landwirtschaft/Verkehr/Medizin |
| **FIoT-IoT-Plattform** | 2020 | Komplettes IoT: Bau, Umwelt, Landwirtschaft, Software/Hardware + Big-Data-Analyse, validiert mit 100.000+ Punkten |
| **UWB-Innenortung** | 2020 | Zentimetergenaue (10~20cm) hochpräzise Innenortung, mit anderen Systemen gekoppelt, in großen Ausstellungen validiert |

## NewLife-Entwicklungsteam

![NewLife](https://newlifex.com/logo.png)  

Das NewLife-Team, gegründet 2002, ist ein Lösungsanbieter der IoT-Branche der nächsten Generation, der sich der Beratung von Software-/Hardware-Lösungen, der Systemarchitekturplanung und Entwicklungsdiensten widmet.  
Mehr als 80 vom Team geleitete Open-Source-Projekte werden branchenweit eingesetzt, mit über 4 Millionen kumulierten NuGet-Downloads.  
Die Big-Data-Middleware NewLife.XCode, die verteilte Berechnungsplattform AntJob, die verteilte Plattform Stardust, die Cache-Warteschlangenkomponente NewLife.Redis und die IoT-Plattform FIoT wurden erfolgreich in Energieversorgung, Universitäten, Internet, Telekommunikation, Verkehr, Logistik, Industrieautomation, Medizin und Kulturerbe eingesetzt und bieten Kunden fortschrittliche, zuverlässige, sichere, qualitativ hochwertige und leicht erweiterbare Produkte sowie Systemintegrationsdienste.  

Wir werden unsere Dienste kontinuierlich verbessern, um ein langfristig vertrauenswürdiger Partner für unsere Kunden zu werden und durch kontinuierliche Innovation und Entwicklung ein ausgezeichneter IoT-Dienstanbieter in China zu werden.  

`Das NewLife-Team wurde 2002 gegründet. Einige Open-Source-Projekte haben eine über 20-jährige Geschichte, und die Quell-Repositories bewahren alle Änderungsdatensätze seit 2010.`  
Website: <https://newlifex.com>  
Open Source: <https://github.com/newlifex>  
QQ-Gruppe: 1600800 / 1600838  
WeChat-Offizielles Konto:  
![Smart Stone](https://newlifex.com/stone.jpg)
