# NewLife.Office - 辦公自動化程式庫

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
🌐 [English](README.md) | [简体中文](README.zh-CN.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

功能全面的 .NET 辦公自動化程式庫，一庫涵蓋 Excel、Word、PPT、PDF 等 15+ 文件格式


原始碼： https://github.com/NewLifeX/NewLife.Office  
NuGet：NewLife.Office（渲染擴充套件：NewLife.Office.Rendering）  

## 專案簡介

`NewLife.Office` 是一個功能全面的 .NET 辦公自動化程式庫，涵蓋 **Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 以及基於 OLE2/BIFF8 的舊式二進位格式，並搭配 **Rendering** 渲染擴充套件（文件 → 圖片）。

- 零外部依賴：僅依賴 `NewLife.Core`，直接基於 Open XML（ZIP + XML）/ PDF / OLE2 檔案結構實作
- 全框架支援：`net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`，net8.0+ 開啟 AOT 相容分析（IsAotCompatible，零 IL 警告）
- MIT 免費：無商業授權風險
- 統一 API：命名空間 `NewLife.Office`，各格式 API 風格一致，學習成本低

> CSV 讀寫能力由 `NewLife.Core` 的 `CsvFile` 提供，本倉庫不重複實作。

## 亮點

- **全格式涵蓋**：四大核心格式 + 十種輕量格式 + OLE2/BIFF8 舊式格式，一庫搞定
- **零外部依賴**：直接操作檔案底層結構，套件體積 <500KB
- **高保真往返**：L2/L3 級保真架構（模型 + RawXml/元件透傳），讀→改→寫不丟格式
- **Markdown 全格式轉換中樞**：MD↔docx/PDF/HTML/Excel/PPT/Word 雙向轉換，為 AI 知識庫餵入提供統一入口
- **渲染擴充套件**：`NewLife.Office.Rendering`（獨立 NuGet）提供文件 → 圖片渲染、圖片處理、QR Code

## 支援的格式

| 格式 | 讀取 | 寫入 | 範本 | 轉換 | 渲染 | 說明 |
|------|:---:|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | 含樣式/圖表/樞紐分析表/條件格式/圖片/命名範圍/結構化表格 |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8，BiffReader+BiffWriter 讀寫雙全，含樣式/合併/超連結/篩選/頁面設定 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | 含段落/表格/圖片/頁首頁尾/註解/內容控制項/AES 加密/郵件合併 |
| **Word doc** | ✅ | — | — | — | — | MS-DOC，DocReader 提取純文字 |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | 含文字方塊/圖片/表格/圖表/母片/動畫/註解/Section；跨檔案合併 |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT，PptReader 提取投影片文字 |
| **PDF** | ✅ | ✅ | — | — | — | 含 Fluent API/合併/拆分/加密/簽章/表單/註釋/PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/各格式→MD | — | CommonMark + GFM + 強型別 AST，全格式轉換中樞 |
| **RTF** | ✅ | ✅ | ✅ | — | — | 段落/表格/範本填充/圖片嵌入（PNG/JPEG） |
| **ODS** | ✅ | ✅ | — | — | — | LibreOffice 原生格式，含合併儲存格/樣式/公式 |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME，附件/內嵌圖片 |
| **MSG** | ✅ | — | — | — | — | Outlook，OLE2 + MAPI，提取主旨/本文/附件 |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics，VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf，vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3，OPF/Nav/章節/封面 |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML，Glyphs 文字提取/多頁寫入/圖片/中繼資料 |
| **Rendering** | — | — | — | — | ✅ | 獨立擴充套件（SkiaSharp）：PDF/Word/PPT/Markdown → 圖片 |

## 快速開始

### 統一文字 / Markdown 提取

```csharp
using NewLife.Office;

// 一行程式碼提取任意格式純文字
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// 提取為 Markdown 格式（保留結構資訊）
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Excel 寫入與讀取

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

> 更多能力（物件對應匯入匯出、範本填充、格式轉換、渲染等）見 [需求文件](Doc/需求文档.md) 與原始碼範例。

## 文件索引

- [需求文件](Doc/需求文档.md)
- [功能清單](Doc/功能清单.md)
- [架構設計](Doc/架构设计.md)
- [競品分析](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## 新生命專案矩陣

各專案預設支援 net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5，舊版（2024.0801）支援 net4.0/net2.0

| 專案 | 年份 | 說明 |
|:---:|:---:|------|
| **基礎元件** | | 支撐其它中介軟體及產品專案 |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | 核心庫，日誌、設定、快取、網路、序列化、APM 效能追蹤 |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | 大數據中介軟體，單表百億級，MySQL/SQLite/SQL Server/Oracle/PostgreSQL/達夢，自動分表、讀寫分離 |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | 網路庫，單機千萬級吞吐率（2266萬 tps），單機百萬級連線（400萬 Tcp 長連線） |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | 協定通訊庫，提供 CS 應用通訊框架，支援 HTTP/RPC，高吞吐、物聯網裝置低開銷易接入 |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | 魔方快速開發平台，整合使用者權限、SSO 登入、OAuth 服務端等，單表 100 億級專案驗證 |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | 服務管理元件，把應用安裝成為作業系統守護程序，Windows 服務、Linux 的 systemd |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero 零代腳手架，基於 NewLife 元件生態的專案範本 NewLife.Templates，Web、WebApi、Service |
| **中介軟體** | | 對接知名中介軟體平台 |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis 用戶端，微秒級延遲，百萬級吞吐，豐富的訊息佇列，百億級資料量專案驗證 |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | RocketMQ 純受控用戶端，支援 Apache RocketMQ 和阿里雲訊息佇列，十億級專案驗證 |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | 物聯網訊息協定，MqttClient/MqttServer，用戶端支援阿里雲物聯網 |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT 標準庫，定義物聯網領域的各種通訊協定標準規範 |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII，基於 IoT 標準庫實作，支援 ZeroIoT 平台和 IoTEdge 閘道 |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | 西門子 PLC 協定，基於 IoT 標準庫實作，支援 IoT 平台和 IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | 地圖元件庫，封裝百度地圖、高德地圖、騰訊地圖、天地圖 |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | 音訊編解碼庫，PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **產品平台** | | 產品平台級，編譯部署即用，個人化自訂 |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | 星塵，分散式服務平台，節點管理、APM 監控中心、設定中心、註冊中心、發佈中心 |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | 螞蟻調度，分散式大數據計算平台（即時/離線），螞蟻搬家分片思想，萬億級資料量專案驗證 |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | 企業 ERP，產品管理、客戶管理、銷售管理、供應商管理 |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | 碼神工具，眾多開發者工具，網路、序列埠、加解密、正規表示式、Modbus、MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | 簡易檔案儲存，支援分散式系統中檔案集中儲存 |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | 產品級反向代理，NAT 代理、HTTP 代理 |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP 壓力測試工具 |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git 原始碼管理系統 |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | 嵌入式作業系統，完全獨立自主，支援 ARM Cortex-M 晶片架構 |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | 嵌入式工業電腦，物聯網邊緣閘道，高效能 .NET 8 主機，應用於工業、農業、交通、醫療 |
| **FIoT 物聯網平台** | 2020 | 物聯網整體解決方案，建築、環保、農業，軟硬體及大數據分析一體化，單機十萬級點位專案驗證 |
| **UWB 高精度室內定位** | 2020 | 公分級（10~20cm）高精度室內定位，軟硬體一體化，與其它系統聯動，大型展廳專案驗證 |

## 新生命開發團隊

![NewLife](https://newlifex.com/logo.png)  

新生命團隊（NewLife）成立於 2002 年，是新時代物聯網行業解決方案提供者，致力於提供軟硬體應用方案諮詢、系統架構規劃與開發服務。  
團隊主導的 80 多個開源專案已被廣泛應用於各行業，NuGet 累計下載量高達 400 餘萬次。  
團隊開發的大數據中介軟體 NewLife.XCode、螞蟻調度計算平台 AntJob、星塵分散式平台 Stardust、快取佇列元件 NewLife.Redis 以及物聯網平台 FIoT，均成功應用於電力、高校、網際網路、電信、交通、物流、工控、醫療、文博等行業，為客戶提供了大量先進、可靠、安全、高品質、易擴充的產品和系統整合服務。  

我們將不斷透過服務的持續改進，成為客戶長期信賴的合作夥伴，透過不斷的創新和發展，成為國內優秀的 IoT 服務供應商。  

`新生命團隊始於 2002 年，部分開源專案具有 20 年以上漫長歷史，原始碼庫保留有 2010 年以來所有修改記錄`  
網站：<https://newlifex.com>  
開源：<https://github.com/newlifex>  
QQ 群：1600800/1600838  
微信公眾號：  
![智慧大石頭](https://newlifex.com/stone.jpg)
