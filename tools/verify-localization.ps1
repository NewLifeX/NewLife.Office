# 一键运行多语言门禁：词表、英文运行时消息、README 切换行/完整性、NuGet 包版 README
#
# 用法：
#   pwsh -File tools/verify-localization.ps1                    # 迁移期，待翻译键只告警
#   pwsh -File tools/verify-localization.ps1 -FailOnPending     # 发布门禁，待翻译清单必须为空
#   pwsh -File tools/verify-localization.ps1 -Report Bin/localization-report.tsv
#   pwsh -File tools/verify-localization.ps1 -SkipReadme -SkipPackage
param(
    [Switch]$FailOnPending,
    [Switch]$StrictScript,
    [Switch]$SkipReadme,
    [Switch]$SkipPackage,
    [String]$Report
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$failed = $false

function Invoke-Check([String]$Name, [String]$Script, [String[]]$StepArgs) {
    Write-Host "==> $Name" -ForegroundColor Cyan
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot $Script) @StepArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ $Name" -ForegroundColor Red
        $script:failed = $true
    }
}

$localeArgs = @(
    '-LocaleDir', 'NewLife.Office/Locales',
    '-SourceDir', 'NewLife.Office,NewLife.Office.Rendering,NewLife.Office.Cli',
    '-RequireTranslated',
    '-PendingFile', 'tools/pending-translations.json',
    '-ScriptCheck'
)
if ($FailOnPending) { $localeArgs += '-FailOnPending' } else { $localeArgs += '-AllowPending' }
if ($StrictScript) { $localeArgs += '-StrictScript' }
if ($Report) { $localeArgs += @('-Report', $Report) }

Invoke-Check '词表一致性' 'check-locales.ps1' $localeArgs
Invoke-Check '英文运行时消息' 'check-english-exceptions.ps1' @()

if (!$SkipReadme) {
    Invoke-Check 'README 语言切换行' 'sync-readme-languages.ps1' @('-Check')
    Invoke-Check 'README 翻译完整性' 'check-readme-languages.ps1' @()
}

if (!$SkipPackage) {
    Invoke-Check 'NuGet 包版 README' 'sync-readme-package.ps1' @('-Check')
}

if ($failed) {
    Write-Host "多语言门禁未通过" -ForegroundColor Red
    exit 1
}

Write-Host "多语言门禁全部通过" -ForegroundColor Green
