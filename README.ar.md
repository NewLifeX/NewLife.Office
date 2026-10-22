# NewLife.Office - مكتبة أتمتة مستندات Office

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
🌐 [English](README.md) | [简体中文](README.zh-CN.md) | [繁體中文](README.zh-TW.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md)
<!-- lang:end -->

> قد تتأخر هذه الترجمة عن [نسخة اللغة الإنجليزية](README.md)؛ وفي حال وجود اختلاف، تكون النسخة الإنجليزية هي المرجع.

المصدر: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (حزمة العرض: NewLife.Office.Rendering)  

## نظرة عامة

`NewLife.Office` مكتبة شاملة لأتمتة مستندات Office في .NET تغطي **Excel (xlsx/xls) وWord (docx/doc) وPowerPoint (pptx/ppt) وPDF**، إضافة إلى Markdown وRTF وODS وEmail (eml/msg) وiCalendar وvCard وEPUB وXPS والصيغ الثنائية القديمة المبنية على OLE2/BIFF8، وتدعمها حزمة **Rendering** الإضافية (مستندات → صور).

- صفر تبعيات خارجية: تعتمد على `NewLife.Core` فقط، ومبنية مباشرة على بنى ملفات Open XML (ZIP + XML) / PDF / OLE2
- دعم كامل لأطر .NET: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`، مع تحليل توافق AOT (IsAotCompatible، صفر تحذيرات IL) على net8.0+
- ترخيص MIT: مجانًا دون مخاطر ترخيص تجاري
- واجهة برمجة موحّدة: مساحة الأسماء `NewLife.Office` بأسلوب API متسق وتكلفة تعلّم منخفضة

> تُوفَّر قراءة/كتابة CSV عبر `CsvFile` في `NewLife.Core`؛ هذه المكتبة لا تكرّر CSV.

## أبرز الميزات

- **تغطية كاملة للصيغ**: أربع صيغ أساسية + عشر صيغ خفيفة + صيغ OLE2/BIFF8 القديمة، في مكتبة واحدة
- **صفر تبعيات خارجية**: تعمل مباشرة على بنى الملفات، وحجم الحزمة أقل من 500KB
- **استدارة عالية الدقة**: بنية وفاء L2/L3 (نموذج + تمرير RawXml/الأجزاء)، قراءة → تعديل → كتابة دون فقدان التنسيق
- **مركز تحويل Markdown**: تحويل ثنائي الاتجاه بين MD وdocx/PDF/HTML/Excel/PPT/Word، مدخل موحّد لتغذية قواعد معرفة الذكاء الاصطناعي
- **امتداد العرض**: `NewLife.Office.Rendering` (حزمة NuGet منفصلة) يوفّر تحويل المستندات إلى صور، ومعالجة الصور، ورموز QR
- **AOT/متعدد المنصات**: صفر تحذيرات IL على `net8.0` / `net10.0`؛ تغطي CI منصات win-x64 / win-x86 / win-arm64 / linux-x64 / linux-arm64 / osx-x64 / osx-arm64 / linux-loongarch64 (LoongArch عبر QEMU)، مع 23 سيناريو اختبار AOT تشمل جداول اللغات الـ31
- **دعم 31 لغة**: تُترجم رسائل الاستثناءات وتسميات مخرجات المستندات وبيانات تعريف التجميع حسب الثقافة الحالية مع الرجوع إلى الإنجليزية؛ تُجمَّع ملفات `Locales/*.yml` داخل التجميع وتعمل تحت NativeAOT

## الصيغ المدعومة

| الصيغة | قراءة | كتابة | قالب | تحويل | عرض | ملاحظات |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | الأنماط والرسوم البيانية والجداول المحورية والتنسيق الشرطي والصور والنطاقات المسماة والجداول المنظمة والتعليقات والمخططات المصغرة والمقاطع والمحادثات الموضوعية وإخفاء الصفوف/الأعمدة والقراءة المتدفقة ودعم تواريخ 1904 |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8، قراءة/كتابة BiffReader+BiffWriter، الأنماط/الدمج/الارتباطات التشعبية/المرشحات/إعداد الصفحة |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | الفقرات والجداول والصور والرؤوس/التذييلات والتعليقات وعناصر التحكم بالمحتوى وتشفير AES ودمج البريد والبحث/الاستبدال (بما في ذلك مربعات النص) ونموذج الحواشي وأقسام متعددة؛ تحقق متقاطع مع NPOI+OpenXmlValidator |
| **Word doc** | ✅ | — | — | ✅ doc→docx | — | MS-DOC، استخراج النص العادي وتنسيق الفقرات عبر DocReader (عريض/حجم الخط/المحاذاة، إلخ) + تحويل doc→docx |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | مربعات النص والصور والجداول والرسوم البيانية والشرائح الرئيسية والحركات والتعليقات والأقسام؛ دمج عبر الملفات (مع الحفاظ على الرسوم/التعليقات/الملاحظات) |
| **PPT ppt** | ✅ | ✅ | — | — | — | MS-PPT: استخراج النص عبر PptReader + كتابة شرائح نصية عبر PptWriter |
| **PDF** | ✅ | ✅ | — | — | — | واجهة Fluent API، دمج/تقسيم، تشفير، توقيعات، نماذج، تعليقات توضيحية، PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/all→MD | — | CommonMark + GFM + AST مُنمَّط، مركز تحويل لجميع الصيغ |
| **RTF** | ✅ | ✅ | ✅ | — | — | فقرات، جداول، تعبئة قوالب، تضمين صور (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | صيغة LibreOffice الأصلية، خلايا مدموجة، أنماط، صيغ رياضية |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME، مرفقات، صور مضمنة |
| **MSG** | ✅ | — | — | — | — | Outlook، OLE2 + MAPI، الموضوع/النص/المرفقات |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics، VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf، vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3، OPF/Nav/الفصول/الغلاف |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML، استخراج نص Glyphs، كتابة متعددة الصفحات، صور، بيانات وصفية |
| **Rendering** | — | — | — | — | ✅ | حزمة منفصلة (SkiaSharp): PDF/Word/PPT/Markdown → صور |

## البدء السريع

### استخراج موحّد للنص / Markdown

```csharp
using NewLife.Office;

// سطر واحد: استخراج النص العادي من أي صيغة مدعومة
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// الاستخراج بصيغة Markdown (مع الحفاظ على البنية)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");

// تحويل doc (97-2003) → docx
OfficeFactory.ConvertDocToDocx("legacy.doc", "legacy.docx");
```

### أداة سطر الأوامر (NewLife.Office.Cli)

```bash
# أي مستند Office → Markdown
dotnet run --project NewLife.Office.Cli -- convert input.docx output.md

# doc (97-2003) → docx (ترحيل المستندات القديمة)
dotnet run --project NewLife.Office.Cli -- doc2docx legacy.doc legacy.docx
```

### كتابة وقراءة Excel

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

> لمزيد من القدرات (استيراد/تصدير تعيين الكائنات، تعبئة القوالب، تحويل الصيغ، العرض، إلخ) راجع [متطلبات المشروع](Doc/需求文档.md) وأمثلة المصدر.

## التوثيق

> 📖 التوثيق الكامل (بالصينية): [مستندات المستودع](Doc)

## خريطة المنظومة

NewLife منظومة متعدّدة الطبقات من مشاريع مفتوحة المصدر تُصدَر بإصدارات مستقلة. الخريطة أدناه مشتركة حرفيًا بين جميع مستودعات NewLife؛ وكل مشروع يرتبط بمستودعه وبحزمته الرئيسية على NuGet.

| المشروع | الوظيفة | التنزيلات |
| --- | --- | --- |
| **الأساس** | | |
| [NewLife.Core](https://github.com/NewLifeX/X) | التسجيل والتهيئة والتخزين المؤقت والشبكات والتسلسل والجدولة وتتبّع أداء APM. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Core?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Core) |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | وسيط بيانات: جداول واحدة بمستوى عشرات المليارات من الصفوف، تقسيم تلقائي للجداول، فصل القراءة عن الكتابة، وأكثر من 10 قواعد بيانات. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCode?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCode) |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | مكتبة شبكات عالية الأداء: 1.77 مليون pps / 14.5 Gbps على جهاز واحد (TCP 1KB، 2026-10)، و4 ملايين اتصال TCP طويل الأمد. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Net?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Net) |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | إطار اتصال بروتوكولي وRPC لتطبيقات العميل/الخادم وأجهزة إنترنت الأشياء. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Remoting?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Remoting) |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | يثبّت تطبيقات .NET كخدمات حارسة لنظام التشغيل: خدمات Windows وsystemd في Linux. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Agent?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Agent) |
| [NewLife.Templates](https://github.com/NewLifeX/NewLife.Zero) | قوالب مشاريع `dotnet new` للطرفية وweb وAPI والخدمات والبيانات وإنترنت الأشياء. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Templates?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Templates) |
| **البيانات والمراسلة** | | |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | عميل Redis بزمن استجابة بالميكروثانية، وملايين العمليات في الثانية، وقوائم رسائل مدمجة. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Redis?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Redis) |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | عميل RocketMQ مُدار بالكامل لـ Apache RocketMQ وخدمة رسائل Alibaba Cloud. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.RocketMQ?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.RocketMQ) |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | عميل وخادم MQTT، بما في ذلك الاتصال بـ Alibaba Cloud IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MQTT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MQTT) |
| [NewLife.NovaDb](https://github.com/NewLifeX/NewLife.NovaDb) | قاعدة بيانات هجينة بلغة C#: مدمجة أو شبكية، وعلائقية، وسلاسل زمنية، وقائمة رسائل، وKV في واحد. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.NovaDb?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.NovaDb) |
| [NewLife.MySql](https://github.com/NewLifeX/NewLife.MySql) | برنامج تشغيل MySql وطبقة مساعدة مبنية على مكدّس بيانات NewLife. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.MySql?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.MySql) |
| **Web والذكاء الاصطناعي والأنظمة الموزّعة** | | |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | إطار منصة إدارة سريعة: الصلاحيات، والدخول الموحّد SSO، وخادم OAuth، وتوليد واجهات CRUD. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Cube?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Cube) |
| [Stardust](https://github.com/NewLifeX/Stardust) | منصة خدمات موزّعة: إدارة العقد، وAPM، ومركز الإعدادات، ومركز التسجيل، ومركز النشر. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Stardust?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Stardust) |
| [AntJob](https://github.com/NewLifeX/AntJob) | منصة موزّعة للمهام وجدولة البيانات الفورية، مع التقسيم وإعادة المحاولة وإعادة التعبئة. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AntJob?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AntJob) |
| [NewLife.AI](https://github.com/NewLifeX/NewLife.AI) | عميل موحّد لـ 46 مزوّد LLM بالإضافة إلى ChatAI، تطبيق ويب جاهز للمحادثة. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.AI?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.AI) |
| **الصناعة وإنترنت الأشياء** | | |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | مكتبة معايير إنترنت الأشياء التي تعرّف عقود البروتوكولات وتجريدات الأجهزة. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IoT?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IoT) |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | تنفيذ Modbus TCP/RTU/ASCII فوق NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Modbus?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Modbus) |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | بروتوكول Siemens PLC فوق NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Siemens?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Siemens) |
| [NewLife.Omron](https://github.com/NewLifeX/NewLife.Omron) | بروتوكول Omron PLC فوق NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Omron?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Omron) |
| [NewLife.Melsec](https://github.com/NewLifeX/NewLife.Melsec) | بروتوكول Mitsubishi Melsec PLC فوق NewLife.IoT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Melsec?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Melsec) |
| [NewLife.OPC](https://github.com/NewLifeX/NewLife.OPC) | عميل OPC للوصول إلى البيانات الصناعية. | — |
| [NewLife.BACnet](https://github.com/NewLifeX/NewLife.BACnet) | مكدّس بروتوكول BACnet لأتمتة المباني. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.BACnet?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.BACnet) |
| [ZeroIoT](https://github.com/NewLifeX/ZeroIoT) | منصة إنترنت أشياء مفتوحة المصدر وخفيفة، مع إدارة البوابات والأجهزة. | — |
| **تليماتيكا المركبات** | | |
| [NewLife.JT808](https://github.com/NewLifeX/NewLife.JT808) | بروتوكول تحديد مواقع المركبات JT/T 808: أكثر من 70 جسم رسالة، 2011/2013/2019. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT808?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT808) |
| [NewLife.JT809](https://github.com/NewLifeX/NewLife.JT809) | بروتوكول JT/T 809 بين المنصات لجهات إدارة النقل. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT809?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT809) |
| [NewLife.JT1078](https://github.com/NewLifeX/NewLife.JT1078) | فيديو داخل المركبة JT/T 1078: تحليل RTP وإعادة التجميع إلى FLV/fMP4 والبث المباشر. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.JT1078?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.JT1078) |
| **مكوّنات المجال** | | |
| [NewLife.Office](https://github.com/NewLifeX/NewLife.Office) | حزمة واحدة لـ Excel وWord وPPT وPDF وأكثر من 15 صيغة مستندات. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Office?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Office) |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | صوت عبر المنصات: مرمّزات، وإدخال/إخراج الأجهزة، وDSP، وحاويات، وبث. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Audio?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Audio) |
| [NewLife.Video](https://github.com/NewLifeX/NewLife.Video) | مرمّزات الفيديو وصيغ الحاويات ومعالجة الإطارات والبث. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Video?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Video) |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | واجهة موحّدة فوق Baidu وAMap وTencent وTianditu مع الترميز الجغرافي وتحويل الإحداثيات. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.Map?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.Map) |
| [NewLife.IP](https://github.com/NewLifeX/NewLife.IP) | تحديد الموقع الجغرافي لعناوين IPv4 بقاعدة بيانات غير متصلة مُعيّنة في الذاكرة، وواجهة API عبر الإنترنت كبديل. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.IP?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.IP) |
| **الأدوات والمنتجات** | | |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | صندوق أدوات المطوّر: الشبكات والمنفذ التسلسلي والتشفير والتعبيرات النمطية وModbus وMQTT. | [![NuGet](https://img.shields.io/nuget/dt/NewLife.XCoder?style=flat-square&logo=nuget&label=downloads)](https://www.nuget.org/packages/NewLife.XCoder) |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | نظام استضافة ذاتية لإدارة الشيفرة المصدرية Git. | — |
| [LuckyClover](https://github.com/NewLifeX/LuckyClover) | يثبّت بيئة تشغيل .NET المطابقة تلقائيًا عند التشغيل من مضيف أقدم. | — |

> تعرض الشارات الحزمة الرئيسية لكل مشروع؛ أما حزم السمات والإضافات والحزم التابعة فمذكورة في صفحاتها على NuGet. الحزم التي تتجاوز 30 حزمة وغير المذكورة هنا (FTP وLoRa وإضافات Office وبرامج تشغيل البروتوكولات وغيرها) متاحة في [منظمة NewLifeX](https://github.com/NewLifeX) وعلى [newlifex.com](https://newlifex.com).

## الفريق

<img src="https://newlifex.com/logo.png" width="48" align="right" alt="NewLife" />

تأسّست NewLife في عام 2002، وأمضت أكثر من عقدين في منظومة .NET وإنترنت الأشياء، وتدير ما يقارب 100 مشروع مفتوح المصدر (34 منها مشاريع رئيسية)، وتقدّم حلولًا متكاملة للبرمجيات والأجهزة لقطاعات الكهرباء والخدمات اللوجستية والرعاية الصحية والتحكّم الصناعي والإنترنت. ونلتزم بمبدأ "لا أسبقية في التعلّم — الإتقان يكسب الاحترام".

نُزّلت المكوّنات مفتوحة المصدر <!--nuget-total:start-->6.48M<!--nuget-total:end--> مرة من NuGet (على مستوى المنظومة بأكملها، حتى 2026-10).

الموقع: https://newlifex.com · المنظمة مفتوحة المصدر: https://github.com/NewLifeX

<details>
<summary>حساب WeChat الرسمي: 智能大石头 (امسح للاشتراك)</summary>

![智能大石头](https://newlifex.com/stone.jpg)

</details>

## الترخيص

MIT License © 2002-2026 فريق NewLife

[GitHub](https://github.com/NewLifeX/NewLife.Office) | [Website](https://newlifex.com)
