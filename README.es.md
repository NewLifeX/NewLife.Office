# NewLife.Office - Biblioteca de automatización de Office

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)


![GitHub License](https://img.shields.io/github/license/NewLifeX/NewLife.Office?logo=github)
![NuGet Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![NuGet](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![NuGet prerelease](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)
[![test](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml/badge.svg?branch=master)](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml)
![platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-blue)
[![contributors](https://img.shields.io/github/contributors/NewLifeX/NewLife.Office?logo=github)](https://github.com/NewLifeX/NewLife.Office/graphs/contributors)

<!-- lang:start -->
🌐 [中文](Readme.MD) | [繁體中文](README.zh-TW.md) | [English](README.en.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Português](README.pt.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

Una biblioteca integral de automatización de Office para .NET que cubre Excel, Word, PPT, PDF y más de 15 formatos de documentos en un solo paquete


Código fuente: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (paquete de renderizado: NewLife.Office.Rendering)  

## Descripción general

`NewLife.Office` es una biblioteca integral de automatización de Office para .NET que cubre **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS y los formatos binarios heredados basados en OLE2/BIFF8, complementada con el paquete de extensión **Rendering** (documentos → imágenes).

- Cero dependencias externas: solo depende de `NewLife.Core`, construido directamente sobre las estructuras de archivo Open XML (ZIP + XML) / PDF / OLE2
- Soporte completo de frameworks: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, con análisis de compatibilidad AOT (IsAotCompatible, cero advertencias IL) en net8.0+
- Licencia MIT: gratuito, sin riesgo de licencia comercial
- API unificada: espacio de nombres `NewLife.Office`, estilo de API consistente, bajo costo de aprendizaje

> La lectura/escritura de CSV la proporciona `CsvFile` en `NewLife.Core`; esta biblioteca no duplica CSV.

## Destacados

- **Cobertura completa de formatos**: cuatro formatos principales + diez formatos ligeros + formatos heredados OLE2/BIFF8, todo en una biblioteca
- **Cero dependencias externas**: opera directamente sobre las estructuras de archivo, tamaño de paquete <500KB
- **Roundtrip de alta fidelidad**: arquitectura de fidelidad L2/L3 (modelo + paso directo de RawXml/partes), leer → modificar → escribir sin perder formato
- **Centro de conversión Markdown**: conversión bidireccional MD↔docx/PDF/HTML/Excel/PPT/Word, un punto de entrada unificado para la ingesta de bases de conocimiento de IA
- **Extensión de renderizado**: `NewLife.Office.Rendering` (NuGet independiente) proporciona renderizado de documentos → imágenes, procesamiento de imágenes y códigos QR

## Formatos compatibles

| Formato | Leer | Escribir | Plantilla | Conversión | Render | Notas |
|---------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Estilos, gráficos, tablas dinámicas, formato condicional, imágenes, rangos con nombre, tablas estructuradas |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter, estilos/combinaciones/hipervínculos/filtros/configuración de página |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Párrafos, tablas, imágenes, encabezados/pies de página, comentarios, controles de contenido, cifrado AES, combinación de correspondencia |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader extrae texto plano |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Cuadros de texto, imágenes, tablas, gráficos, patrones, animaciones, comentarios, secciones; fusión entre archivos |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader extrae texto de diapositivas |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, fusión/división, cifrado, firmas, formularios, anotaciones, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/todos→MD | — | CommonMark + GFM + AST tipado, centro de conversión de formatos |
| **RTF** | ✅ | ✅ | ✅ | — | — | Párrafos, tablas, relleno de plantillas, incrustación de imágenes (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | Formato nativo de LibreOffice, celdas combinadas, estilos, fórmulas |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, adjuntos, imágenes integradas |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, asunto/cuerpo/adjuntos |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/capítulos/portada |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, extracción de texto Glyphs, escritura multipágina, imágenes, metadatos |
| **Rendering** | — | — | — | — | ✅ | Paquete independiente (SkiaSharp): PDF/Word/PPT/Markdown → imágenes |

## Inicio rápido

### Extracción unificada de texto / Markdown

```csharp
using NewLife.Office;

// Una línea: extraer texto plano de cualquier formato compatible
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Extraer como Markdown (conserva la información estructural)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Escritura y lectura de Excel

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

> Más capacidades (importación/exportación de mapeo de objetos, relleno de plantillas, conversión de formatos, renderizado, etc.) consulte [Requisitos](Doc/需求文档.md) y los ejemplos del código fuente.

## Documentación

- [Requisitos](Doc/需求文档.md)
- [Lista de funciones](Doc/功能清单.md)
- [Arquitectura](Doc/架构设计.md)
- [Análisis de la competencia](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## Ecosistema de proyectos NewLife

Todos los proyectos admiten por defecto net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5. Las versiones antiguas (2024.0801) admiten net4.0/net2.0.

| Proyecto | Año | Descripción |
|:---:|:---:|------|
| **Componentes base** | | Base que soporta otros middleware y productos |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Biblioteca principal: registro, configuración, caché, red, serialización, trazado APM |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Middleware de big data, 10 mil millones+ de filas por tabla, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, partición automática, separación lectura/escritura |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Biblioteca de red, 22.66M tps en una sola máquina, 4M conexiones TCP largas |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Biblioteca de comunicación por protocolo, framework CS, HTTP/RPC, alto rendimiento, acceso IoT de bajo costo |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Plataforma de desarrollo rápido, permisos de usuario, inicio de sesión SSO, servidor OAuth, validado a escala de 10 mil millones de filas |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Componente de gestión de servicios, instala aplicaciones como daemons del SO (Servicio de Windows, systemd de Linux) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Andamiaje Zero, plantillas de proyecto NewLife.Templates, Web, WebApi, Service |
| **Middleware** | | Integración con plataformas de middleware populares |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Cliente Redis, latencia de microsegundos, rendimiento de millones, colas de mensajes, validado a escala de 10 mil millones |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Cliente RocketMQ puro gestionado, Apache RocketMQ y Alibaba Cloud MQ, validado en proyectos de mil millones |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | Protocolo de mensajería IoT, MqttClient/MqttServer, soporte de Alibaba Cloud IoT |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | Biblioteca estándar de IoT, define estándares de protocolos de comunicación para IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, basado en la biblioteca estándar IoT, plataforma ZeroIoT y puerta de enlace IoTEdge |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Protocolos de PLC Siemens, basado en la biblioteca estándar IoT, plataforma IoT e IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Biblioteca de componentes de mapas, envoltura de Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Biblioteca de códecs de audio, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Plataformas y productos** | | Listas para compilar, implementar y usar |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Plataforma de servicios distribuidos: gestión de nodos, monitoreo APM, centro de configuración, registro, publicación |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Plataforma de computación distribuida de big data (tiempo real/offline), partición por migración de hormigas, validado a escala de billones |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | ERP empresarial: gestión de productos, clientes, ventas, proveedores |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Caja de herramientas de desarrollador: red, serial, cifrado, regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Almacenamiento de archivos simple, soporta almacenamiento centralizado en sistemas distribuidos |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Proxy inverso de grado de producción, proxy NAT, proxy HTTP |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | Herramienta de pruebas de estrés HTTP |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Sistema de gestión de código fuente Git |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Sistema operativo embebido, totalmente independiente, arquitectura ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Computadora industrial embebida, puerta de enlace IoT de borde, host .NET 8 de alto rendimiento, para industria/agricultura/transporte/medicina |
| **Plataforma IoT FIoT** | 2020 | Solución integral de IoT: construcción, medio ambiente, agricultura, software/hardware + análisis de big data, validado con 100 mil+ puntos |
| **Posicionamiento interior UWB** | 2020 | Posicionamiento interior de alta precisión a nivel de centímetros (10~20cm), integrado con otros sistemas, validado en grandes exposiciones |

## Equipo de desarrollo NewLife

![NewLife](https://newlifex.com/logo.png)  

El equipo NewLife, fundado en 2002, es un proveedor de soluciones de la industria IoT de nueva generación, dedicado a la consultoría de soluciones de software/hardware, planificación de arquitectura de sistemas y servicios de desarrollo.  
Más de 80 proyectos de código abierto liderados por el equipo se utilizan ampliamente en diversas industrias, con descargas acumuladas de NuGet que superan los 4 millones.  
El middleware de big data NewLife.XCode, la plataforma de computación distribuida AntJob, la plataforma distribuida Stardust, el componente de cola de caché NewLife.Redis y la plataforma IoT FIoT se han aplicado con éxito en energía eléctrica, universidades, internet, telecomunicaciones, transporte, logística, control industrial, medicina y patrimonio cultural, ofreciendo a los clientes productos avanzados, confiables, seguros, de alta calidad y fácilmente extensibles, así como servicios de integración de sistemas.  

Continuaremos mejorando nuestros servicios para convertirnos en un socio de confianza a largo plazo para nuestros clientes y, a través de la innovación y el desarrollo continuos, convertirnos en un excelente proveedor de servicios IoT en China.  

`El equipo NewLife comenzó en 2002. Algunos proyectos de código abierto tienen más de 20 años de historia, y los repositorios de código fuente conservan todos los registros de modificaciones desde 2010.`  
Sitio web: <https://newlifex.com>  
Código abierto: <https://github.com/newlifex>  
Grupo QQ: 1600800 / 1600838  
Cuenta oficial de WeChat:  
![Piedra inteligente](https://newlifex.com/stone.jpg)
