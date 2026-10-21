# tools

多语言本地化与 README 校验脚本。仓库根目录只保留配置（如 `readme-languages.json`），脚本统一放在这里。

| 脚本 | 作用 |
|---|---|
| `check-locales.ps1` | 词表键集合、占位符、格式说明符、代码↔词表双向覆盖、翻译完整性；支持待翻译清单与覆盖报告 |
| `check-english-exceptions.ps1` | 检查异常、日志、控制台消息是否为英文源文，禁止字符串插值 |
| `sync-readme-languages.ps1` | 按 `readme-languages.json` 生成/校验 README 语言切换行 |
| `check-readme-languages.ps1` | 校验 11 种 README 的结构、代码块、目标语言书写系统与整段英文残留 |
| `pending-translations.json` | 迁移期「英文与简体中文已完成、其余语言待补」的键清单 |

## 常用命令

```powershell
# 迁移期：待翻译键只告警
pwsh -File tools/check-locales.ps1 -LocaleDir NewLife.Office/Locales -SourceDir "NewLife.Office,NewLife.Office.Rendering,NewLife.Office.Cli" -RequireTranslated -AllowPending -PendingFile tools/pending-translations.json

# 发布门禁：待翻译清单必须为空
pwsh -File tools/check-locales.ps1 -LocaleDir NewLife.Office/Locales -SourceDir "NewLife.Office,NewLife.Office.Rendering,NewLife.Office.Cli" -RequireTranslated -FailOnPending

# 导出各语言覆盖报告
pwsh -File tools/check-locales.ps1 -LocaleDir NewLife.Office/Locales -SourceDir "NewLife.Office,NewLife.Office.Rendering,NewLife.Office.Cli" -RequireTranslated -AllowPending -Report Bin/localization-report.tsv

pwsh -File tools/check-english-exceptions.ps1
pwsh -File tools/sync-readme-languages.ps1 -Check
pwsh -File tools/check-readme-languages.ps1
```

## 约定

- 英文是源文与兜底；词表 key 必须与代码/csproj 中的英文字符串逐字一致。
- 新增消息：先加英文源文与 `zh-Hans.yml`，再登记到 `pending-translations.json`；批量翻译阶段补齐后移除。
- 阶段三前冻结英文源文：改词会导致全部译文失配。
- 词表脚本的单一来源为 NewLife.Core `Tools/Scripts`；本目录是同步副本，修改后应回灌上游，避免各仓再次分叉。
