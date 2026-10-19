# NewLife.Office - Biblioteca de automação de Office

![GitHub top language](https://img.shields.io/github/languages/top/newlifex/NewLife.Office?logo=github)


![GitHub License](https://img.shields.io/github/license/NewLifeX/NewLife.Office?logo=github)
![NuGet Downloads](https://img.shields.io/nuget/dt/NewLife.Office?logo=nuget)
![NuGet](https://img.shields.io/nuget/v/NewLife.Office?logo=nuget)
![NuGet prerelease](https://img.shields.io/nuget/vpre/NewLife.Office?label=dev%20nuget&logo=nuget)
[![test](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml/badge.svg?branch=master)](https://github.com/NewLifeX/NewLife.Office/actions/workflows/test.yml)
[![locales](https://img.shields.io/badge/locales-31-2ea44f?logo=dotnet)](Doc/多语言本地化.md)
[![AOT smoke](https://img.shields.io/badge/AOT%20smoke-23%2F23-brightgreen?logo=dotnet)](.github/workflows/aot.yml)
![platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS%20%7C%20Android%20%7C%20iOS-blue)
![arch](https://img.shields.io/badge/arch-x64%20%7C%20arm64%20%7C%20loongarch64%20%7C%20riscv64%20%7C%20s390x-orange)
[![contributors](https://img.shields.io/github/contributors/NewLifeX/NewLife.Office?logo=github)](https://github.com/NewLifeX/NewLife.Office/graphs/contributors)

<!-- lang:start -->
🌐 [中文](Readme.MD) | [繁體中文](README.zh-TW.md) | [English](README.en.md) | [Español](README.es.md) | [Deutsch](README.de.md) | [日本語](README.ja.md) | [Français](README.fr.md) | [Русский](README.ru.md) | [한국어](README.ko.md) | [العربية](README.ar.md)
<!-- lang:end -->

Uma biblioteca abrangente de automação de Office para .NET que cobre Excel, Word, PPT, PDF e mais de 15 formatos de documentos em um único pacote


Código-fonte: https://github.com/NewLifeX/NewLife.Office  
NuGet: NewLife.Office (pacote de renderização: NewLife.Office.Rendering)  

## Visão geral

`NewLife.Office` é uma biblioteca abrangente de automação de Office para .NET que cobre **Excel (xlsx/xls), Word (docx/doc), PowerPoint (pptx/ppt), PDF**, Markdown, RTF, ODS, Email (eml/msg), iCalendar, vCard, EPUB, XPS e os formatos binários legados baseados em OLE2/BIFF8, complementada pelo pacote de extensão **Rendering** (documentos → imagens).

- Zero dependências externas: depende apenas de `NewLife.Core`, construído diretamente sobre as estruturas de arquivo Open XML (ZIP + XML) / PDF / OLE2
- Suporte completo a frameworks: `net45` / `netstandard2.0` / `netstandard2.1` / `net8.0` / `net10.0`, com análise de compatibilidade AOT (IsAotCompatible, zero avisos IL) no net8.0+
- Licença MIT: gratuito, sem risco de licença comercial
- API unificada: namespace `NewLife.Office`, estilo de API consistente, baixo custo de aprendizado

> A leitura/escrita de CSV é fornecida por `CsvFile` em `NewLife.Core`; esta biblioteca não duplica CSV.

## Destaques

- **Cobertura completa de formatos**: quatro formatos principais + dez formatos leves + formatos legados OLE2/BIFF8, tudo em uma biblioteca
- **Zero dependências externas**: opera diretamente nas estruturas de arquivo, tamanho do pacote <500KB
- **Roundtrip de alta fidelidade**: arquitetura de fidelidade L2/L3 (modelo + passagem direta de RawXml/partes), ler → modificar → escrever sem perder formatação
- **Centro de conversão Markdown**: conversão bidirecional MD↔docx/PDF/HTML/Excel/PPT/Word, um ponto de entrada unificado para ingestão de bases de conhecimento de IA
- **Extensão de renderização**: `NewLife.Office.Rendering` (NuGet independente) fornece renderização de documentos → imagens, processamento de imagens e códigos QR

## Formatos suportados

| Formato | Ler | Escrever | Modelo | Conversão | Render | Notas |
|---------|:---:|:---:|:---:|:---:|:---:|-------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | — | Estilos, gráficos, tabelas dinâmicas, formatação condicional, imagens, intervalos nomeados, tabelas estruturadas |
| **Excel xls** | ✅ | ✅ | — | — | — | OLE2/BIFF8, BiffReader+BiffWriter, estilos/combinações/hiperlinks/filtros/configuração de página |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | — | Parágrafos, tabelas, imagens, cabeçalhos/rodapés, comentários, controles de conteúdo, criptografia AES, mala direta |
| **Word doc** | ✅ | — | — | — | — | MS-DOC, DocReader extrai texto simples |
| **PPT pptx** | ✅ | ✅ | ✅ | PDF | — | Caixas de texto, imagens, tabelas, gráficos, mestres, animações, comentários, seções; fusão entre arquivos |
| **PPT ppt** | ✅ | — | — | — | — | MS-PPT, PptReader extrai texto de slides |
| **PDF** | ✅ | ✅ | — | — | — | Fluent API, fusão/divisão, criptografia, assinaturas, formulários, anotações, PDF-A |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF/todos→MD | — | CommonMark + GFM + AST tipado, centro de conversão de formatos |
| **RTF** | ✅ | ✅ | ✅ | — | — | Parágrafos, tabelas, preenchimento de modelos, incorporação de imagens (PNG/JPEG) |
| **ODS** | ✅ | ✅ | — | — | — | Formato nativo do LibreOffice, células mescladas, estilos, fórmulas |
| **EML** | ✅ | ✅ | — | — | — | RFC 5322 + MIME, anexos, imagens incorporadas |
| **MSG** | ✅ | — | — | — | — | Outlook, OLE2 + MAPI, assunto/corpo/anexos |
| **iCalendar** | ✅ | ✅ | — | — | — | RFC 5545 .ics, VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | — | RFC 6350 .vcf, vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | — | EPUB 2/3, OPF/Nav/capítulos/capa |
| **XPS** | ✅ | ✅ | — | — | — | ZIP+XML, extração de texto Glyphs, escrita multipágina, imagens, metadados |
| **Rendering** | — | — | — | — | ✅ | Pacote independente (SkiaSharp): PDF/Word/PPT/Markdown → imagens |

## Início rápido

### Extração unificada de texto / Markdown

```csharp
using NewLife.Office;

// Uma linha: extrair texto simples de qualquer formato suportado
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// Extrair como Markdown (preserva informações estruturais)
String? md = OfficeFactory.ReadMarkdown("slides.pptx");
```

### Escrita e leitura de Excel

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

> Mais recursos (importação/exportação de mapeamento de objetos, preenchimento de modelos, conversão de formatos, renderização, etc.) consulte [Requisitos](Doc/需求文档.md) e os exemplos do código-fonte.

## Documentação

- [Requisitos](Doc/需求文档.md)
- [Lista de recursos](Doc/功能清单.md)
- [Arquitetura](Doc/架构设计.md)
- [Análise da concorrência](Doc/竞品分析报告.md)
- [ChangeLog](ChangeLog.md)

## Ecossistema de projetos NewLife

Todos os projetos suportam por padrão net9.0/netstandard2.1/netstandard2.0/net4.62/net4.5. Versões antigas (2024.0801) suportam net4.0/net2.0.

| Projeto | Ano | Descrição |
|:---:|:---:|------|
| **Componentes base** | | Base que suporta outros middleware e produtos |
| [NewLife.Core](https://github.com/NewLifeX/X) | 2002 | Biblioteca principal: log, configuração, cache, rede, serialização, rastreamento APM |
| [NewLife.XCode](https://github.com/NewLifeX/NewLife.XCode) | 2005 | Middleware de big data, 10 bilhões+ de linhas por tabela, MySQL/SQLite/SQL Server/Oracle/PostgreSQL/DaMeng, particionamento automático, separação leitura/escrita |
| [NewLife.Net](https://github.com/NewLifeX/NewLife.Net) | 2005 | Biblioteca de rede, 22,66M tps em uma única máquina, 4M conexões TCP longas |
| [NewLife.Remoting](https://github.com/NewLifeX/NewLife.Remoting) | 2011 | Biblioteca de comunicação por protocolo, framework CS, HTTP/RPC, alto rendimento, acesso IoT de baixo custo |
| [NewLife.Cube](https://github.com/NewLifeX/NewLife.Cube) | 2010 | Plataforma de desenvolvimento rápido, permissões de usuário, login SSO, servidor OAuth, validado na escala de 10 bilhões de linhas |
| [NewLife.Agent](https://github.com/NewLifeX/NewLife.Agent) | 2008 | Componente de gerenciamento de serviços, instala aplicativos como daemons do SO (Serviço do Windows, systemd do Linux) |
| [NewLife.Zero](https://github.com/NewLifeX/NewLife.Zero) | 2020 | Scaffolding Zero, modelos de projeto NewLife.Templates, Web, WebApi, Service |
| **Middleware** | | Integração com plataformas de middleware populares |
| [NewLife.Redis](https://github.com/NewLifeX/NewLife.Redis) | 2017 | Cliente Redis, latência de microssegundos, rendimento de milhões, filas de mensagens, validado na escala de 10 bilhões |
| [NewLife.RocketMQ](https://github.com/NewLifeX/NewLife.RocketMQ) | 2018 | Cliente RocketMQ puro gerenciado, Apache RocketMQ e Alibaba Cloud MQ, validado em projetos de bilhões |
| [NewLife.MQTT](https://github.com/NewLifeX/NewLife.MQTT) | 2019 | Protocolo de mensageria IoT, MqttClient/MqttServer, suporte Alibaba Cloud IoT |
| [NewLife.IoT](https://github.com/NewLifeX/NewLife.IoT) | 2022 | Biblioteca padrão IoT, define padrões de protocolos de comunicação para IoT |
| [NewLife.Modbus](https://github.com/NewLifeX/NewLife.Modbus) | 2022 | ModbusTcp/ModbusRTU/ModbusASCII, baseado na biblioteca padrão IoT, plataforma ZeroIoT e gateway IoTEdge |
| [NewLife.Siemens](https://github.com/NewLifeX/NewLife.Siemens) | 2022 | Protocolos Siemens PLC, baseado na biblioteca padrão IoT, plataforma IoT e IoTEdge |
| [NewLife.Map](https://github.com/NewLifeX/NewLife.Map) | 2022 | Biblioteca de componentes de mapas, Baidu, Amap, Tencent, Tianditu |
| [NewLife.Audio](https://github.com/NewLifeX/NewLife.Audio) | 2023 | Biblioteca de codecs de áudio, PCM/ADPCMA/G711A/G722U/WAV/AAC |
| **Plataformas e produtos** | | Nível de plataforma de produto, compilar, implantar e usar |
| [Stardust](https://github.com/NewLifeX/Stardust) | 2018 | Plataforma de serviços distribuídos: gerenciamento de nós, monitoramento APM, centro de configuração, registro, publicação |
| [AntJob](https://github.com/NewLifeX/AntJob) | 2019 | Plataforma de computação distribuída de big data (tempo real/offline), particionamento por migração de formigas, validado na escala de trilhões |
| [NewLife.ERP](https://github.com/NewLifeX/NewLife.ERP) | 2021 | ERP empresarial: gerenciamento de produtos, clientes, vendas, fornecedores |
| [CrazyCoder](https://github.com/NewLifeX/XCoder) | 2006 | Caixa de ferramentas do desenvolvedor: rede, serial, criptografia, regex, Modbus, MQTT |
| [EasyIO](https://github.com/NewLifeX/EasyIO) | 2023 | Armazenamento de arquivos simples, armazenamento centralizado em sistemas distribuídos |
| [XProxy](https://github.com/NewLifeX/XProxy) | 2005 | Proxy reverso de nível de produção, proxy NAT, proxy HTTP |
| [HttpMeter](https://github.com/NewLifeX/HttpMeter) | 2022 | Ferramenta de teste de carga HTTP |
| [GitCandy](https://github.com/NewLifeX/GitCandy) | 2015 | Sistema de gerenciamento de código-fonte Git |
| [SmartOS](https://github.com/NewLifeX/SmartOS) | 2014 | Sistema operacional embarcado, totalmente independente, arquitetura ARM Cortex-M |
| [SmartA2](https://github.com/NewLifeX/SmartA2) | 2019 | Computador industrial embarcado, gateway IoT de borda, host .NET 8 de alto desempenho, para indústria/agricultura/transporte/medicina |
| **Plataforma IoT FIoT** | 2020 | Solução completa de IoT: construção, meio ambiente, agricultura, software/hardware + análise de big data, validado com 100 mil+ pontos |
| **Posicionamento interno UWB** | 2020 | Posicionamento interno de alta precisão em nível de centímetros (10~20cm), integrado a outros sistemas, validado em grandes exposições |

## Equipe de desenvolvimento NewLife

![NewLife](https://newlifex.com/logo.png)  

A equipe NewLife, fundada em 2002, é um provedor de soluções da indústria IoT de nova geração, dedicado à consultoria de soluções de software/hardware, planejamento de arquitetura de sistemas e serviços de desenvolvimento.  
Mais de 80 projetos de código aberto liderados pela equipe são amplamente utilizados em diversos setores, com downloads acumulados de NuGet superiores a 4 milhões.  
O middleware de big data NewLife.XCode, a plataforma de computação distribuída AntJob, a plataforma distribuída Stardust, o componente de fila de cache NewLife.Redis e a plataforma IoT FIoT foram aplicados com sucesso em energia elétrica, universidades, internet, telecomunicações, transporte, logística, controle industrial, medicina e patrimônio cultural, oferecendo aos clientes produtos avançados, confiáveis, seguros, de alta qualidade e facilmente extensíveis, além de serviços de integração de sistemas.  

Continuaremos melhorando nossos serviços para nos tornarmos um parceiro de confiança de longo prazo para nossos clientes e, por meio de inovação e desenvolvimento contínuos, nos tornarmos um excelente provedor de serviços IoT na China.  

`A equipe NewLife começou em 2002. Alguns projetos de código aberto têm mais de 20 anos de história, e os repositórios de código-fonte preservam todos os registros de modificações desde 2010.`  
Site: <https://newlifex.com>  
Código aberto: <https://github.com/newlifex>  
Grupo QQ: 1600800 / 1600838  
Conta oficial do WeChat:  
![Pedra inteligente](https://newlifex.com/stone.jpg)
