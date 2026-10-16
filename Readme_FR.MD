# NewLife.Office - Bibliothèque d'automatisation bureautique

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)
![GitHub License](https://img.shields.io/github/license/newlifex/NewLife.Office?logo=github)
![Nuget Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![Nuget](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![Nuget (with prereleases)](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)

Une bibliothèque complète d'automatisation bureautique .NET couvrant Excel, Word, PPT, PDF et plus de 15 formats de documents en un seul package

[English](Readme_EN.MD) | [简体中文](Readme.MD) | [繁體中文](Readme_TW.MD) | [日本語](Readme_JA.MD) | [한국어](Readme_KO.MD) | [Español](Readme_ES.MD) | Français | [Deutsch](Readme_DE.MD) | [Português](Readme_PT.MD) | [Русский](Readme_RU.MD)

Source : https://github.com/NewLifeX/NewLife.Office  
NuGet : NewLife.Office (package de rendu : NewLife.Office.Rendering)  

## Présentation

`NewLife.Office` est une bibliothèque complète d'automatisation bureautique .NET couvrant **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS ainsi que les formats binaires hérités basés sur OLE2/BIFF8, complétée par le package d'extension **Rendering** (documents → images).

- Zéro dépendance externe : dépend uniquement de `NewLife.Core`, construit directement sur les structures de fichiers Open XML (ZIP + XML) / PDF / OLE2
- Prise en charge complète des frameworks : `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, avec analyse de compatibilité AOT (IsAotCompatible, zéro avertissement IL) sur net8.0+
- Licence MIT : gratuit, sans risque de licence commerciale
- API unifiée : espace de noms `NewLife.Office`, style d'API cohérent, faible coût d'apprentissage

> La lecture/écriture CSV est fournie par `CsvFile` dans `NewLife.Core` ; cette bibliothèque ne duplique pas le CSV.

## Points forts

- **Couverture complète des formats** : quatre formats principaux + dix formats légers + formats hérités OLE2/BIFF8, le tout en une bibliothèque
- **Zéro dépendance externe** : opère directement sur les structures de fichiers, taille du package <500KB
- **Aller-retour haute fidélité** : architecture de fidélité L2/L3 (modèle + passage direct RawXml/parties), lire → modifier → écrire sans perdre le formatage
- **Centre de conversion Markdown** : conversion bidirectionnelle MD↔docx/PDF/HTML/Excel/PPT/Word, un point d'entrée unifié pour l'ingestion de bases de connaissances IA
- **Extension de rendu** : `NewLife.Office.Rendering` (NuGet indépendant) fournit le rendu documents → images, le traitement d'images et les codes QR

## Formats pris en charge

| Format | Lecture | Écriture | Modèle | Conversion | Rendu | Notes |
|--------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Styles, graphiques, tableaux croisés dynamiques, mise en forme conditionnelle, images, plages nommées, tableaux structurés |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter, styles/fusions/hyperliens/filtres/mise en page |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Paragraphes, tableaux, images, en-têtes/pieds de page, commentaires, contrôles de contenu, chiffrement AES, publipostage |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader extrait le texte brut |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Zones de texte, images, tableaux, graphiques, masques, animations, commentaires, sections ; fusion inter-fichiers |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader extrait le texte des diapositives |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, fusion/division, chiffrement, signatures, formulaires, annotations, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/tous→MD | — | CommonMark + GFM + AST typé, centre de conversion de formats |
| **RTF** | ✅ | ✅ | ✅ | — | — | Paragraphes, tableaux, remplissage de modèles, incorporation d'images (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | Format natif LibreOffice, cellules fusionnées, styles, formules |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, pièces jointes, images intégrées |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, objet/corps/pièces jointes |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/chapitres/couverture |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, extraction de texte Glyphs, écriture multipage, images, métadonnées |
| **Rendering** | — | — | — | — | ✅ | Package indépendant (SkiaSharp) : PDF/Word/PPT/Markdown → images |

## Démarrage rapide

### Extraction unifiée de texte / Markdown

```csharp
using NewLife.Office;

// Une ligne : extraire le texte brut de tout format pris en charge
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Extraire en Markdown (conserve les informations structurelles)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Écriture et lecture Excel

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

> Plus de capacités (import/export de mappage d'objets, remplissage de modèles, conversion de formats, rendu, etc.) voir [Exigences](Doc/需求文档.md) et les exemples du code source.

## Documentation

- [Exigences](Doc/需求文档.md)
- [Liste des fonctionnalités](Doc/功能清单.md)
- [Architecture](Doc/架构设计.md)
- [Analyse concurrentielle](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## Écosystème de projets NewLife

Tous les projets prennent en charge par défaut net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5. Les anciennes versions (2024.0801) prennent en charge net4.0/net2.0.

| Projet | Année | Description |
|:---:|:---:|------|
| **Composants de base** | | Base soutenant les autres middleware et produits |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Bibliothèque principale : journalisation, configuration, cache, réseau, sérialisation, traçage APM |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Middleware big data, 10 milliards+ de lignes par table, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, partitionnement automatique, séparation lecture/écriture |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Bibliothèque réseau, 22,66M tps sur une seule machine, 4M connexions TCP longues |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Bibliothèque de communication par protocole, framework CS, HTTP/RPC, haut débit, accès IoT à faible coût |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Plateforme de développement rapide, permissions utilisateur, SSO, serveur OAuth, validé à l'échelle de 10 milliards de lignes |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Composant de gestion de services, installe les applications comme démons OS (Service Windows, systemd Linux) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Échafaudage Zero, modèles de projet NewLife.Templates, Web, WebApi, Service |
| **Middleware** | | Intégration avec les plateformes de middleware populaires |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Client Redis, latence en microsecondes, débit de millions, files de messages, validé à l'échelle de 10 milliards |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Client RocketMQ pur géré, Apache RocketMQ et Alibaba Cloud MQ, validé dans des projets de milliards |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | Protocole de messagerie IoT, MqttClient/MqttServer, support Alibaba Cloud IoT |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | Bibliothèque standard IoT, définit les normes de protocoles de communication IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, basé sur la bibliothèque standard IoT, plateforme ZeroIoT et passerelle IoTEdge |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Protocoles PLC Siemens, basé sur la bibliothèque standard IoT, plateforme IoT et IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Bibliothèque de composants cartographiques, Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Bibliothèque de codecs audio, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Plateformes et produits** | | Prêts à compiler, déployer et utiliser |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Plateforme de services distribués : gestion de nœuds, surveillance APM, centre de configuration, registre, publication |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Plateforme de calcul distribué big data (temps réel/hors ligne), partitionnement par migration de fourmis, validé à l'échelle de milliers de milliards |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | ERP d'entreprise : gestion des produits, clients, ventes, fournisseurs |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Boîte à outils du développeur : réseau, série, chiffrement, regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Stockage de fichiers simple, stockage centralisé dans les systèmes distribués |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Proxy inverse de qualité production, proxy NAT, proxy HTTP |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | Outil de test de charge HTTP |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Système de gestion de code source Git |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Système d'exploitation embarqué, totalement indépendant, architecture ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Ordinateur industriel embarqué, passerelle IoT de périphérie, hôte .NET 8 haute performance, pour industrie/agriculture/transport/médecine |
| **Plateforme IoT FIoT** | 2020 | Solution IoT complète : construction, environnement, agriculture, logiciel/matériel + analyse big data, validé avec 100 000+ points |
| **Positionnement intérieur UWB** | 2020 | Positionnement intérieur haute précision au centimètre (10~20cm), intégré à d'autres systèmes, validé dans de grandes expositions |

## Équipe de développement NewLife

![NewLife](https://newlifex.com/logo.png)  

L'équipe NewLife, fondée en 2002, est un fournisseur de solutions de l'industrie IoT de nouvelle génération, dédié au conseil en solutions logiciel/matériel, à la planification d'architecture système et aux services de développement.  
Plus de 80 projets open source dirigés par l'équipe sont largement utilisés dans divers secteurs, avec des téléchargements NuGet cumulés dépassant 4 millions.  
Le middleware big data NewLife.XCode, la plateforme de calcul distribué AntJob, la plateforme distribuée Stardust, le composant de file d'attente de cache NewLife.Redis et la plateforme IoT FIoT ont été appliqués avec succès dans l'énergie électrique, les universités, l'internet, les télécommunications, les transports, la logistique, le contrôle industriel, la médecine et le patrimoine culturel, offrant aux clients des produits avancés, fiables, sécurisés, de haute qualité et facilement extensibles, ainsi que des services d'intégration de systèmes.  

Nous continuerons à améliorer nos services pour devenir un partenaire de confiance à long terme pour nos clients et, grâce à l'innovation et au développement continus, devenir un excellent fournisseur de services IoT en Chine.  

`L'équipe NewLife a commencé en 2002. Certains projets open source ont plus de 20 ans d'histoire, et les dépôts de code source conservent tous les enregistrements de modifications depuis 2010.`  
Site web : <https://newlifex.com>  
Open source : <https://github.com/newlifex>  
Groupe QQ : 1600800 / 1600838  
Compte officiel WeChat :  
![Pierre intelligente](https://newlifex.com/stone.jpg)
