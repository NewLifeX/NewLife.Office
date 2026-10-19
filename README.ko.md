# NewLife.Office - 오피스 자동화 라이브러리

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
🌐 [中文](Readme.MD) | [繁體中文](README.zh-TW.md) | [English](README.en.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [العربية](README.ar.md)
<!-- lang:end -->

Excel, Word, PPT, PDF 등 15개 이상의 문서 형식을 하나의 라이브러리로 지원하는 종합 .NET 오피스 자동화 라이브러리


소스: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (렌더링 확장 패키지: NewLife.Office.Rendering)  

## 프로젝트 소개

`NewLife.Office`는 **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS 및 OLE2/BIFF8 기반의 레거시 바이너리 형식을 지원하는 종합 .NET 오피스 자동화 라이브러리입니다. **Rendering** 확장 패키지(문서 → 이미지)도 제공합니다.

- 외부 의존성 제로: `NewLife.Core`에만 의존하며 Open XML (ZIP + XML) / PDF / OLE2 파일 구조를 직접 조작
- 전체 프레임워크 지원: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, net8.0+에서 AOT 호환 분석(IsAotCompatible, IL 경고 없음) 활성화
- MIT 라이선스: 무료, 상업용 라이선스 리스크 없음
- 통일된 API: `NewLife.Office` 네임스페이스, 모든 형식의 API 스타일 일관, 학습 비용 낮음

> CSV 읽기/쓰기는 `NewLife.Core`의 `CsvFile`이 제공하므로 이 라이브러리에서는 중복 구현하지 않습니다.

## 하이라이트

- **전체 형식 커버**: 4개 핵심 형식 + 10개 경량 형식 + OLE2/BIFF8 레거시 형식을 하나의 라이브러리로
- **외부 의존성 제로**: 파일 저수준 구조 직접 조작, 패키지 크기 <500KB
- **고충실도 라운드트립**: L2/L3 충실도 아키텍처(모델 + RawXml/파츠 투과), 읽기→수정→쓰기 시 서식 손실 없음
- **Markdown 변환 허브**: MD↔docx/PDF/HTML/Excel/PPT/Word 양방향 변환, AI 지식 베이스 주입을 위한 통일된 진입점
- **렌더링 확장**: `NewLife.Office.Rendering`(별도 NuGet)이 문서 → 이미지 렌더링, 이미지 처리, QR 코드 제공

## 지원 형식

| 형식 | 읽기 | 쓰기 | 템플릿 | 변환 | 렌더링 | 설명 |
|------|:---:|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | 스타일/차트/피벗 테이블/조건부 서식/이미지/이름 정의/구조화 테이블 |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter 읽기/쓰기, 스타일/병합/하이퍼링크/필터/페이지 설정 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | 문단/표/이미지/머리글·바닥글/주석/콘텐츠 컨트롤/AES 암호화/메일 머지 |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader가 일반 텍스트 추출 |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | 텍스트 상자/이미지/표/차트/마스터/애니메이션/주석/Section; 파일 간 병합 |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader가 슬라이드 텍스트 추출 |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API/병합/분할/암호화/서명/폼/주석/PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/각 형식→MD | — | CommonMark + GFM + 타입형 AST, 전체 형식 변환 허브 |
| **RTF** | ✅ | ✅ | ✅ | — | — | 문단/표/템플릿 채우기/이미지 임베드 (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | LibreOffice 네이티브 형식, 병합 셀/스타일/수식 |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, 첨부/인라인 이미지 |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, 제목/본문/첨부 추출 |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/챕터/표지 |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, Glyphs 텍스트 추출/다중 페이지 쓰기/이미지/메타데이터 |
| **Rendering** | — | — | — | — | ✅ | 별도 패키지 (SkiaSharp): PDF/Word/PPT/Markdown → 이미지 |

## 빠른 시작

### 통합 텍스트 / Markdown 추출

```csharp
using NewLife.Office;

// 한 줄로 모든 지원 형식의 일반 텍스트 추출
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Markdown으로 추출 (구조 정보 보존)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Excel 쓰기와 읽기

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

> 더 많은 기능(객체 매핑 가져오기/내보내기, 템플릿 채우기, 형식 변환, 렌더링 등)은 [요구사항 문서](Doc/需求文档.md)와 소스 코드 예제를 참조하세요.

## 문서

- [요구사항 문서](Doc/需求文档.md)
- [기능 목록](Doc/功能清单.md)
- [아키텍처 설계](Doc/架构设计.md)
- [경쟁사 분석](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## 신생명 프로젝트 매트릭스

모든 프로젝트는 기본적으로 net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5를 지원합니다. 구버전(2024.0801)은 net4.0/net2.0을 지원합니다.

| 프로젝트 | 연도 | 설명 |
|:---:|:---:|------|
| **기본 컴포넌트** | | 다른 미들웨어와 제품 프로젝트를 지원하는 기반 |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | 코어 라이브러리: 로그, 설정, 캐시, 네트워크, 직렬화, APM 성능 추적 |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | 빅데이터 미들웨어, 테이블당 100억 행, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/다멍, 자동 샤딩, 읽기/쓰기 분리 |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | 네트워크 라이브러리, 단일 머신 2266만 tps, 400만 TCP 장기 연결 |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | 프로토콜 통신 라이브러리, CS 애플리케이션 통신 프레임워크, HTTP/RPC, 고처리량, IoT 디바이스 저비용 접속 |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | 고속 개발 플랫폼, 사용자 권한, SSO 로그인, OAuth 서버, 100억 행 규모 검증 |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | 서비스 관리 컴포넌트, 애플리케이션을 OS 데몬으로 설치 (Windows 서비스, Linux systemd) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Zero 제로 코드 스캐폴드, NewLife 생태계 기반 프로젝트 템플릿 NewLife.Templates, Web, WebApi, Service |
| **미들웨어** | | 주요 미들웨어 플랫폼 연동 |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Redis 클라이언트, 마이크로초 레이턴시, 백만급 처리량, 풍부한 메시지 큐, 100억급 데이터 검증 |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | 순수 매니지드 RocketMQ 클라이언트, Apache RocketMQ 및 알리바바 클라우드 MQ, 10억급 프로젝트 검증 |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | IoT 메시징 프로토콜, MqttClient/MqttServer, 알리바바 클라우드 IoT 지원 |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | IoT 표준 라이브러리, IoT 분야의 각종 통신 프로토콜 표준 정의 |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, IoT 표준 라이브러리 기반, ZeroIoT 플랫폼 및 IoTEdge 게이트웨이 지원 |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | 지멘스 PLC 프로토콜, IoT 표준 라이브러리 기반, IoT 플랫폼 및 IoTEdge 지원 |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | 지도 컴포넌트 라이브러리, 바이두/아맵/텐센트/톈디투 래핑 |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | 오디오 코덱 라이브러리, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **제품 플랫폼** | | 제품 플랫폼급, 컴파일 후 배포 즉시 사용 |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | 분산 서비스 플랫폼, 노드 관리, APM 모니터링 센터, 설정 센터, 레지스트리, 배포 센터 |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | 분산 빅데이터 계산 플랫폼 (실시간/오프라인), 개미 이사 샤딩, 조 단위 데이터 검증 |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | 기업 ERP, 제품 관리, 고객 관리, 판매 관리, 공급업체 관리 |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | 개발자 도구 모음, 네트워크, 시리얼, 암호화, 정규식, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | 간편 파일 스토리지, 분산 시스템의 파일 집중 저장 지원 |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | 제품급 리버스 프록시, NAT 프록시, HTTP 프록시 |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | HTTP 부하 테스트 도구 |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Git 소스 코드 관리 시스템 |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | 임베디드 OS, 완전 독립, ARM Cortex-M 칩 아키텍처 지원 |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | 임베디드 산업용 컴퓨터, IoT 엣지 게이트웨이, 고성능 .NET 8 호스트, 산업/농업/교통/의료에 응용 |
| **FIoT 사물인터넷 플랫폼** | 2020 | 사물인터넷 종합 솔루션, 건설, 환경, 농업, 소프트/하드웨어 + 빅데이터 분석 통합, 단일 머신 10만 포인트 검증 |
| **UWB 고정밀 실내 위치추적** | 2020 | 센티미터급(10~20cm) 고정밀 실내 위치추적, 소프트/하드웨어 통합, 타 시스템 연동, 대형 전시장 프로젝트 검증 |

## 신생명 개발팀

![NewLife](https://newlifex.com/logo.png)  

신생명 팀(NewLife)은 2002년에 설립되었으며, 차세대 IoT 산업 솔루션 제공자로서 소프트웨어/하드웨어 솔루션 컨설팅, 시스템 아키텍처 설계 및 개발 서비스를 제공합니다.  
팀이 주도하는 80여 개의 오픈소스 프로젝트는 다양한 산업에 널리 사용되고 있으며, NuGet 누적 다운로드는 400만 회를 넘습니다.  
빅데이터 미들웨어 NewLife.XCode, 분산 스케줄링 플랫폼 AntJob, 분산 플랫폼 Stardust, 캐시 큐 컴포넌트 NewLife.Redis, 그리고 IoT 플랫폼 FIoT는 전력, 대학, 인터넷, 통신, 교통, 물류, 공업 제어, 의료, 문화 박물관 등 산업에 성공적으로 적용되어, 고객에게 선진적이고 안정적이며 안전하고 고품질이며 확장이 쉬운 제품과 시스템 통합 서비스를 제공해 왔습니다.  

우리는 서비스의 지속적 개선을 통해 고객이 장기적으로 신뢰하는 파트너가 되고, 지속적인 혁신과 발전을 통해 중국을 대표하는 우수한 IoT 서비스 공급자가 되겠습니다.  

`신생명 팀은 2002년에 시작되었으며, 일부 오픈소스 프로젝트는 20년 이상의 오랜 역사를 가지고 있고, 소스 저장소에는 2010년 이후의 모든 수정 기록이 보존되어 있습니다.`  
웹사이트: <https://newlifex.com>  
오픈소스: <https://github.com/newlifex>  
QQ 그룹: 1600800/1600838  
위챗 공식 계정:  
![스마트 스톤](https://newlifex.com/stone.jpg)
