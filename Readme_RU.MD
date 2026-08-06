# NewLife.Office - Библиотека автоматизации Office

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

Комплексная библиотека автоматизации Office для .NET, покрывающая Excel, Word, PPT, PDF и более 15 форматов документов в одном пакете

[English](Readme_EN.MD) | [简体中文](Readme.MD) | [繁體中文](Readme_TW.MD) | [日本語](Readme_JA.MD) | [한국어](Readme_KO.MD) | [Español](Readme_ES.MD) | [Français](Readme_FR.MD) | [Deutsch](Readme_DE.MD) | [Português](Readme_PT.MD) | Русский

Исходный код: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (пакет рендеринга: NewLife.Office.Rendering)  

## Обзор

`NewLife.Office` — это комплексная библиотека автоматизации Office для .NET, покрывающая **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS, а также унаследованные бинарные форматы на основе OLE2/BIFF8, дополненная пакетом расширения **Rendering** (документы → изображения).

- Ноль внешних зависимостей: зависит только от `NewLife.Core`, построена непосредственно на структурах файлов Open XML (ZIP + XML) / PDF / OLE2
- Полная поддержка фреймворков: `net45` / `netstandard2.0` / `netstandard2.1` и другие
- Лицензия MIT: бесплатно, без риска коммерческой лицензии
- Единый API: пространство имён `NewLife.Office`, согласованный стиль API, низкий порог входа

> Чтение/запись CSV предоставляется классом `CsvFile` в `NewLife.Core`; данная библиотека не дублирует CSV.

## Ключевые возможности

- **Полное покрытие форматов**: четыре основных формата + десять лёгких форматов + унаследованные форматы OLE2/BIFF8 — всё в одной библиотеке
- **Ноль внешних зависимостей**: работает непосредственно со структурами файлов, размер пакета <500KB
- **Высокоточный roundtrip**: архитектура точности L2/L3 (модель + сквозная передача RawXml/частей), чтение → изменение → запись без потери форматирования
- **Центр конвертации Markdown**: двунаправленная конвертация MD↔docx/PDF/HTML/Excel/PPT/Word, единая точка входа для загрузки баз знаний ИИ
- **Расширение рендеринга**: `NewLife.Office.Rendering` (отдельный NuGet) предоставляет рендеринг документов → изображения, обработку изображений и QR-коды

## Поддерживаемые форматы

| Формат | Чтение | Запись | Шаблон | Конвертация | Рендер | Примечания |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Стили, диаграммы, сводные таблицы, условное форматирование, изображения, именованные диапазоны, структурированные таблицы |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter, стили/объединения/гиперссылки/фильтры/настройка страницы |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Абзацы, таблицы, изображения, колонтитулы, комментарии, элементы управления содержимым, шифрование AES, слияние почты |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader извлекает обычный текст |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Текстовые поля, изображения, таблицы, диаграммы, мастер-слайды, анимации, комментарии, разделы; слияние между файлами |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader извлекает текст слайдов |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, слияние/разделение, шифрование, подписи, формы, аннотации, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/все→MD | — | CommonMark + GFM + типизированный AST, центр конвертации форматов |
| **RTF** | ✅ | ✅ | ✅ | — | — | Абзацы, таблицы, заполнение шаблонов, встраивание изображений (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | Нативный формат LibreOffice, объединённые ячейки, стили, формулы |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, вложения, встроенные изображения |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, тема/текст/вложения |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/главы/обложка |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, извлечение текста Glyphs, многостраничная запись, изображения, метаданные |
| **Rendering** | — | — | — | — | ✅ | Отдельный пакет (SkiaSharp): PDF/Word/PPT/Markdown → изображения |

## Быстрый старт

### Единое извлечение текста / Markdown

```csharp
using NewLife.Office;

// Одна строка: извлечь обычный текст из любого поддерживаемого формата
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Извлечь как Markdown (сохраняет структурную информацию)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Запись и чтение Excel

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

> Дополнительные возможности (импорт/экспорт объектного сопоставления, заполнение шаблонов, конвертация форматов, рендеринг и т.д.) см. в [Требованиях](Doc/需求文档.md) и примерах исходного кода.

## Документация

- [Требования](Doc/需求文档.md)
- [Список функций](Doc/功能清单.md)
- [Архитектура](Doc/架构设计.md)
- [Анализ конкурентов](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## Экосистема проектов NewLife

Все проекты по умолчанию поддерживают net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5. Старые версии (2024.0801) поддерживают net4.0/net2.0.

| Проект | Год | Описание |
|:---:|:---:|------|
| **Базовые компоненты** | | Основа, поддерживающая другие промежуточные и продуктовые проекты |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Базовая библиотека: логирование, конфигурация, кэш, сеть, сериализация, трассировка APM |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Промежуточное ПО больших данных, 10 млрд+ строк на таблицу, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, автоматическое шардирование, разделение чтения/записи |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Сетевая библиотека, 22,66 млн tps на одной машине, 4 млн TCP-долгоживущих соединений |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Библиотека протокольной связи, фреймворк CS, HTTP/RPC, высокая пропускная способность, низкопороговое подключение IoT |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Платформа быстрой разработки, права пользователей, SSO-вход, OAuth-сервер, проверено на масштабе 10 млрд строк |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Компонент управления службами, установка приложений как демонов ОС (Windows Service, Linux systemd) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Каркас Zero, шаблоны проектов NewLife.Templates, Web, WebApi, Service |
| **Промежуточное ПО** | | Интеграция с популярными платформами промежуточного ПО |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Клиент Redis, микросекундная задержка, миллионная пропускная способность, богатые очереди сообщений, проверено на 10 млрд данных |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Чистый управляемый клиент RocketMQ, Apache RocketMQ и Alibaba Cloud MQ, проверено в проектах на миллиарды |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | Протокол обмена сообщениями IoT, MqttClient/MqttServer, поддержка Alibaba Cloud IoT |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | Стандартная библиотека IoT, определяет стандарты протоколов связи для IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, на основе стандартной библиотеки IoT, платформа ZeroIoT и шлюз IoTEdge |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Протоколы Siemens PLC, на основе стандартной библиотеки IoT, платформа IoT и IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Библиотека картографических компонентов, Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Библиотека аудиокодеков, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Продуктовые платформы** | | Продуктовый уровень, компилируй, разворачивай и используй |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Распределённая сервисная платформа: управление узлами, APM-мониторинг, центр конфигурации, реестр, центр публикации |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Распределённая платформа вычислений больших данных (реального времени/офлайн), шардирование по миграции муравьёв, проверено на триллионах |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | Корпоративная ERP: управление продуктами, клиентами, продажами, поставщиками |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Набор инструментов разработчика: сеть, COM-порт, шифрование, регулярные выражения, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Простое файловое хранилище, централизованное хранение файлов в распределённых системах |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Продуктовый обратный прокси, NAT-прокси, HTTP-прокси |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | Инструмент нагрузочного тестирования HTTP |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Система управления исходным кодом Git |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Встраиваемая ОС, полностью независимая, архитектура ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Встраиваемый промышленный компьютер, IoT-шлюз на периферии, высокопроизводительный хост .NET 8, для промышленности/сельского хозяйства/транспорта/медицины |
| **Платформа IoT FIoT** | 2020 | Комплексное решение IoT: строительство, экология, сельское хозяйство, интеграция ПО/аппаратного обеспечения + анализ больших данных, проверено на 100 тыс.+ точках |
| **UWB-позиционирование в помещении** | 2020 | Высокоточное позиционирование в помещении с точностью до сантиметра (10~20см), интеграция с другими системами, проверено на крупных выставках |

## Команда разработчиков NewLife

![NewLife](https://newlifex.com/logo.png)  

Команда NewLife, основанная в 2002 году, является поставщиком решений для IoT-индустрии нового поколения и занимается консалтингом программных/аппаратных решений, планированием системной архитектуры и разработкой.  
Более 80 проектов с открытым исходным кодом, возглавляемых командой, широко используются в различных отраслях, суммарное количество загрузок NuGet превышает 4 миллиона.  
Промежуточное ПО больших данных NewLife.XCode, распределённая платформа вычислений AntJob, распределённая платформа Stardust, компонент кэш-очередей NewLife.Redis и IoT-платформа FIoT успешно применяются в энергетике, университетах, интернете, телекоммуникациях, транспорте, логистике, промышленной автоматизации, медицине и культурном наследии, предоставляя клиентам передовые, надёжные, безопасные, высококачественные и легко расширяемые продукты и услуги системной интеграции.  

Мы будем постоянно улучшать наши услуги, чтобы стать долгосрочным надёжным партнёром для наших клиентов, и благодаря постоянным инновациям и развитию стать отличным поставщиком IoT-услуг в Китае.  

`Команда NewLife была основана в 2002 году. Некоторые проекты с открытым исходным кодом имеют более чем 20-летнюю историю, а репозитории исходного кода сохраняют все записи изменений с 2010 года.`  
Сайт: <https://newlifex.com>  
Открытый код: <https://github.com/newlifex>  
Группа QQ: 1600800 / 1600838  
Официальный аккаунт WeChat:  
![Умный камень](https://newlifex.com/stone.jpg)
