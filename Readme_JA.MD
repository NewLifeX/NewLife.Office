# NewLife.Office - オフィス自動化ライブラリ

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

Excel、Word、PPT、PDF など 15 以上のドキュメント形式を1つのライブラリでカバーする、機能豊富な .NET オフィス自動化ライブラリ

[English](Readme_EN.MD) | [简体中文](Readme.MD) | [繁體中文](Readme_TW.MD) | 日本語 | [한국어](Readme_KO.MD) | [Español](Readme_ES.MD) | [Français](Readme_FR.MD) | [Deutsch](Readme_DE.MD) | [Português](Readme_PT.MD) | [Русский](Readme_RU.MD)

ソース: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office（レンダリング拡張パッケージ: NewLife.Office.Rendering）  

## プロジェクト概要

`NewLife.Office` は、**Excel（xlsx/xls）、Word（docx/doc）、PowerPoint（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS に加え、OLE2/BIFF8 ベースのレガシーバイナリ形式もカバーする、機能豊富な .NET オフィス自動化ライブラリです。**Rendering** 拡張パッケージ（ドキュメント → 画像）も提供します。

- 外部依存ゼロ: `NewLife.Core` のみに依存し、Open XML（ZIP + XML）/ PDF / OLE2 ファイル構造を直接操作
- 全フレームワーク対応: `net45` / `netstandard2.0` / `netstandard2.1` など
- MIT ライセンス: 無料、商用ライセンスリスクなし
- 統一 API: 名前空間 `NewLife.Office`、全形式で API スタイルが統一され学習コストが低い

> CSV の読み書きは `NewLife.Core` の `CsvFile` が提供するため、本ライブラリでは重複実装しません。

## ハイライト

- **全形式カバー**: 4つのコア形式 + 10の軽量形式 + OLE2/BIFF8 レガシー形式を1つのライブラリで
- **外部依存ゼロ**: ファイルの低レベル構造を直接操作、パッケージサイズ <500KB
- **高忠実度ラウンドトリップ**: L2/L3 レベルの忠実度アーキテクチャ（モデル + RawXml/パーツ透過）、読む→変更→書くで書式を失わない
- **Markdown 変換ハブ**: MD↔docx/PDF/HTML/Excel/PPT/Word 双方向変換、AI ナレッジベース投入の統一エントリ
- **レンダリング拡張**: `NewLife.Office.Rendering`（独立 NuGet）がドキュメント → 画像レンダリング、画像処理、QR コードを提供

## 対応形式

| 形式 | 読取 | 書込 | テンプレート | 変換 | レンダリング | 説明 |
|------|:---:|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | スタイル/グラフ/ピボットテーブル/条件付き書式/画像/名前付き範囲/構造化テーブル |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8、BiffReader+BiffWriter で読み書き、スタイル/結合/ハイパーリンク/フィルタ/ページ設定 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | 段落/表/画像/ヘッダー・フッター/コメント/コンテンツコントロール/AES 暗号化/差し込み印刷 |
| **Word doc** | ✅ | — | — | — | — | MS-DOC、DocReader がプレーンテキストを抽出 |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | テキストボックス/画像/表/グラフ/マスター/アニメーション/コメント/セクション；ファイル間マージ |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT、PptReader がスライドテキストを抽出 |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API/結合/分割/暗号化/署名/フォーム/注釈/PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/各形式→MD | — | CommonMark + GFM + 型付き AST、全形式変換ハブ |
| **RTF** | ✅ | ✅ | ✅ | — | — | 段落/表/テンプレート埋め込み/画像埋め込み（PNG/JPEG） |
| **ODS** | ✅ | ✅ | — | — | — | LibreOffice ネイティブ形式、結合セル/スタイル/数式 |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME、添付/インライン画像 |
| **MSG** | ✅ | — | — | — | — | Outlook、OLE2 + MAPI、件名/本文/添付を抽出 |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics、VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf、vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3、OPF/Nav/章/表紙 |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML、Glyphs テキスト抽出/複数ページ書き込み/画像/メタデータ |
| **Rendering** | — | — | — | — | ✅ | 独立パッケージ（SkiaSharp）: PDF/Word/PPT/Markdown → 画像 |

## クイックスタート

### 統一テキスト / Markdown 抽出

```csharp
using NewLife.Office;

// 一行で任意形式のプレーンテキストを抽出
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Markdown として抽出（構造情報を保持）
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Excel 書き込みと読み取り

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

> その他の機能（オブジェクトマッピングのインポート/エクスポート、テンプレート埋め込み、形式変換、レンダリングなど）は [要件ドキュメント](Doc/需求文档.md) とソースコードの例を参照してください。

## ドキュメント

- [要件ドキュメント](Doc/需求文档.md)
- [機能リスト](Doc/功能清单.md)
- [アーキテクチャ設計](Doc/架构设计.md)
- [競合分析](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## 新生命プロジェクトマトリクス

全プロジェクトはデフォルトで net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5 をサポート。旧バージョン（2024.0801）は net4.0/net2.0 をサポート。

| プロジェクト | 年 | 説明 |
|:---:|:---:|------|
| **基本コンポーネント** | | その他の中間ウェアと製品を支える基盤 |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | コアライブラリ: ログ、設定、キャッシュ、ネットワーク、シリアライゼーション、APM トレース |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | ビッグデータ中間ウェア、テーブルあたり100億行、MySQL/SQLite/SQL Server/Oracle/PostgreSQL/達夢、自動シャーディング、読み書き分離 |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | ネットワークライブラリ、単機 2266万 tps、400万 TCP 長接続 |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | プロトコル通信ライブラリ、CS アプリ通信フレームワーク、HTTP/RPC、高スループット、IoT デバイス低コスト接続 |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | 高速開発プラットフォーム、ユーザー権限、SSO ログイン、OAuth サーバー、100億行規模で検証 |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | サービス管理コンポーネント、アプリを OS デーモンとしてインストール（Windows サービス、Linux systemd） |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero ゼロコード足場、NewLife エコシステムのプロジェクトテンプレート（NewLife.Templates）、Web、WebApi、Service |
| **中間ウェア** | | 主要中間ウェアプラットフォームとの連携 |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis クライアント、マイクロ秒レイテンシ、百万級スループット、豊富なメッセージキュー、100億級データで検証 |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | 純マネージド RocketMQ クライアント、Apache RocketMQ と Alibaba Cloud MQ、10億級プロジェクトで検証 |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | IoT メッセージングプロトコル、MqttClient/MqttServer、Alibaba Cloud IoT 対応 |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT 標準ライブラリ、IoT の通信プロトコル標準を定義 |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII、IoT 標準ライブラリベース、ZeroIoT プラットフォームと IoTEdge ゲートウェイ |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | シーメンス PLC プロトコル、IoT 標準ライブラリベース、IoT プラットフォームと IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | 地図コンポーネントライブラリ、百度/高徳/騰訊/天地図をラップ |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | オーディオコーデックライブラリ、PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **製品プラットフォーム** | | コンパイルしてデプロイすればすぐ使える製品レベル |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | 分散サービスプラットフォーム、ノード管理、APM 監視センター、設定センター、レジストリ、公開センター |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | 分散ビッグデータ計算プラットフォーム（リアルタイム/オフライン）、蟻の引越しシャーディング、兆級データで検証 |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | 企業 ERP、製品管理、顧客管理、販売管理、仕入先管理 |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | 開発者ツールボックス、ネットワーク、シリアル、暗号化、正規表現、Modbus、MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | 分散システム向け簡易ファイルストレージ、ファイル集中ストレージをサポート |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | 製品級リバースプロキシ、NAT プロキシ、HTTP プロキシ |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP 負荷テストツール |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git ソースコード管理システム |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | 組み込み OS、完全独立、ARM Cortex-M チップアーキテクチャ対応 |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | 組み込み産業用コンピュータ、IoT エッジゲートウェイ、高性能 .NET 8 ホスト、産業・農業・交通・医療に応用 |
| **FIoT モノのインターネットプラットフォーム** | 2020 | モノのインターネット全体ソリューション、建築、環境、農業、ソフト/ハード+ビッグデータ分析一体型、単機10万点規模で検証 |
| **UWB 高精度屋内測位** | 2020 | センチメートル級（10~20cm）高精度屋内測位、ソフト/ハード一体型、他システムと連携、大型展示会プロジェクトで検証 |

## 新生命開発チーム

![NewLife](https://newlifex.com/logo.png)  

新生命チーム（NewLife）は2002年に設立され、次世代 IoT 業界ソリューション提供者として、ソフトウェア/ハードウェアソリューションコンサルティング、システムアーキテクチャ設計と開発サービスを提供しています。  
チームが主導する80以上のオープンソースプロジェクトは様々な業界で広く使用され、NuGet 累計ダウンロード数は400万を超えています。  
ビッグデータ中間ウェア NewLife.XCode、分散スケジューリングプラットフォーム AntJob、分散プラットフォーム Stardust、キャッシュキューコンポーネント NewLife.Redis、そして FIoT モノのインターネットプラットフォームは、電力、大学、インターネット、通信、交通、物流、産業制御、医療、文化博覧などの業界で成功裏に応用され、先進的で信頼性が高く、安全で高品質、拡張性に優れた製品とシステム統合サービスを提供しています。  

私たちはサービスの継続的改善を通じて、顧客から長期的に信頼されるパートナーとなり、継続的なイノベーションと発展を通じて、中国を代表する優れた IoT サービスプロバイダーになります。  

`新生命チームは2002年に設立されました。一部のオープンソースプロジェクトは20年以上の歴史を持ち、ソースリポジトリには2010年以降のすべての変更記録が保存されています。`  
ウェブサイト: <https://newlifex.com>  
オープンソース: <https://github.com/newlifex>  
QQ グループ: 1600800 / 1600838  
微信公式アカウント:  
![スマートストーン](https://newlifex.com/stone.jpg)
