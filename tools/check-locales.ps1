# 多语言词表一致性校验：键集合、占位符、格式说明符、代码→词表覆盖、翻译完整性
# 以基准语言文件（默认 zh-Hans.yml）为准，所有语言必须提供完全相同的键集合
#
# 用法：
#   pwsh -File tools/check-locales.ps1                                      # 本库默认目录
#   pwsh -File tools/check-locales.ps1 -LocaleDir src\Lib\Locales -SourceDir src\Lib
#   pwsh -File tools/check-locales.ps1 -RequireTranslated                   # 禁止英文回退，强制全译
#   pwsh -File tools/check-locales.ps1 -RequireTranslated -AllowPending     # 迁移期：待翻译键降为告警
#   pwsh -File tools/check-locales.ps1 -RequireTranslated -FailOnPending    # 发布门禁：待翻译清单必须为空
#   pwsh -File tools/check-locales.ps1 -SkipOrphan                          # 只校验词表内部一致性
#   pwsh -File tools/check-locales.ps1 -ScriptCheck                         # 检查非拉丁语言译文是否含目标书写系统
#   pwsh -File tools/check-locales.ps1 -Report Bin/localization-report.tsv  # 导出各语言覆盖报告
param(
    [String]$LocaleDir = "NewLife.Office/Locales",
    [String]$Baseline = "zh-Hans.yml",
    [String[]]$SourceDir = @("NewLife.Office", "NewLife.Office.Rendering", "NewLife.Office.Cli"),
    [String]$MetadataKeys,
    [String]$PendingFile = "tools/pending-translations.json",
    [String]$Report,
    [Switch]$SkipOrphan,
    [Switch]$RequireTranslated,
    [Switch]$AllowPending,
    [Switch]$FailOnPending,
    [Switch]$ScriptCheck,
    [Switch]$StrictScript
)

$ErrorActionPreference = "Stop"

# 支持用逗号/分号分隔多个源码目录（CI 以 -File 传参时数组会退化为字符串）
$SourceDir = @($SourceDir | ForEach-Object { $_ -split '[,;]' } | Where-Object { $_ })

# 脚本位于 tools，仓库根目录在其上一级
$root = Split-Path -Parent $PSScriptRoot
$localeDir = if ([IO.Path]::IsPathRooted($LocaleDir)) { $LocaleDir } else { Join-Path $root $LocaleDir }
if (!(Test-Path $localeDir)) {
    Write-Host "未找到语言目录：$localeDir" -ForegroundColor Red
    exit 1
}

# 非源码来源键（属性元数据、外部 Model.xml 等）豁免孤儿键判定，但仍要求全语言齐全
if (!$MetadataKeys) {
    $candidate = Join-Path $localeDir "metadata-keys.json"
    if (Test-Path -LiteralPath $candidate) { $MetadataKeys = $candidate }
}
$metadataKeySet = [Collections.Generic.HashSet[String]]::new([StringComparer]::Ordinal)
if ($MetadataKeys) {
    $metadataPath = if ([IO.Path]::IsPathRooted($MetadataKeys)) { $MetadataKeys } else { Join-Path $root $MetadataKeys }
    if (!(Test-Path -LiteralPath $metadataPath)) {
        Write-Host "未找到非源码来源键清单：$metadataPath" -ForegroundColor Red
        exit 1
    }
    foreach ($item in (Get-Content -LiteralPath $metadataPath -Encoding UTF8 -Raw | ConvertFrom-Json).keys) {
        if (![String]::IsNullOrEmpty($item)) { [void]$metadataKeySet.Add($item) }
    }
}

# 迁移期待翻译清单：阶段一只维护英文源文与简体中文，其余语言缺失键登记在此
$pendingKeySet = [Collections.Generic.HashSet[String]]::new([StringComparer]::Ordinal)
$pendingPath = if ($PendingFile) {
    if ([IO.Path]::IsPathRooted($PendingFile)) { $PendingFile } else { Join-Path $root $PendingFile }
}
if ($pendingPath -and (Test-Path -LiteralPath $pendingPath)) {
    $pendingData = Get-Content -LiteralPath $pendingPath -Encoding UTF8 -Raw | ConvertFrom-Json
    foreach ($item in $pendingData.keys) {
        if (![String]::IsNullOrEmpty($item)) { [void]$pendingKeySet.Add($item) }
    }
}

$pattern = '^\s*"((?:[^"\\]|\\.)*)"\s*:\s*"((?:[^"\\]|\\.)*)"\s*$'
$placeholder = '\{(\d+)(?:,-?\d+)?(:[^}]*)?\}'

function Unescape-Yaml([string]$text) {
    # 与源生成器保持一致：还原 \\ \" \n \r \t \uXXXX
    return [regex]::Replace($text, '\\(u[0-9a-fA-F]{4}|.)', {
        param($m)
        $v = $m.Groups[1].Value
        if ($v.Length -eq 5) { return [string][char][Convert]::ToInt32($v.Substring(1), 16) }
        switch -CaseSensitive ($v) {
            'n' { return "`n" }
            'r' { return "`r" }
            't' { return "`t" }
            '\\' { return '\' }
            '"' { return '"' }
            default { return $m.Value }
        }
    })
}

function Get-Placeholders([string]$text) {
    $set = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($m in [regex]::Matches($text, $placeholder)) {
        [void]$set.Add([int]$m.Groups[1].Value)
    }
    return ,$set
}

function Get-PlaceholderSpecs([string]$text) {
    $map = @{}
    foreach ($m in [regex]::Matches($text, $placeholder)) {
        $index = [int]$m.Groups[1].Value
        if (!$map.ContainsKey($index)) { $map[$index] = $m.Groups[2].Value }
    }
    return $map
}

# 目标书写系统：只对整句/短语做启发式检查，短标签与技术词不参与
function Test-TargetScript([String]$culture, [String]$key, [String]$value) {
    $pattern = switch -Wildcard ($culture) {
        'zh-*' { '[\u4e00-\u9fff]' }
        'ja' { '[\u3040-\u30ff\u4e00-\u9fff]' }
        'ko' { '[\uac00-\ud7af]' }
        'ru' { '[\u0400-\u04ff]' }
        'uk' { '[\u0400-\u04ff]' }
        'el' { '[\u0370-\u03ff]' }
        'ar' { '[\u0600-\u06ff]' }
        'fa' { '[\u0600-\u06ff]' }
        'he' { '[\u0590-\u05ff]' }
        'hi' { '[\u0900-\u097f]' }
        'bn' { '[\u0980-\u09ff]' }
        'th' { '[\u0e00-\u0e7f]' }
        default { '' }
    }
    if (!$pattern) { return $true }
    if ($key -match '\*\*') { return $true }
    $letters = ([regex]::Matches($key, '[A-Za-z]')).Count
    if ($letters -lt 8 -or $key -notmatch ' ') { return $true }
    if ([regex]::IsMatch($value, $pattern)) { return $true }
    # 含其它非 ASCII 字符时不判失败，避免误伤音译与符号
    if ([regex]::IsMatch($value, '[^\x00-\x7F]')) { return $true }
    return $false
}

function Read-Locale([string]$path) {
    $entries = [ordered]@{}
    $errors = @()
    $lineNo = 0
    foreach ($line in Get-Content -LiteralPath $path -Encoding UTF8) {
        $lineNo++
        if ($line.Trim().Length -eq 0 -or $line.TrimStart().StartsWith("#")) { continue }
        $m = [regex]::Match($line, $pattern)
        if (!$m.Success) {
            $errors += "第 $lineNo 行格式非法：$line"
            continue
        }
        $key = Unescape-Yaml $m.Groups[1].Value
        if ($entries.Contains($key)) {
            $errors += "第 $lineNo 行键重复：$key"
            continue
        }
        $entries[$key] = Unescape-Yaml $m.Groups[2].Value
    }
    return @{ Entries = $entries; Errors = $errors }
}

$referencePath = Join-Path $localeDir $Baseline
if (!(Test-Path -LiteralPath $referencePath)) {
    Write-Host "未找到基准语言文件：$referencePath" -ForegroundColor Red
    exit 1
}
$reference = Read-Locale $referencePath
$allErrors = @($reference.Errors | ForEach-Object { "$($Baseline): $_" })
$allWarnings = @()
$pendingMissingCount = 0

foreach ($file in Get-ChildItem $localeDir -Filter *.yml | Sort-Object Name) {
    if ($file.Name -eq $Baseline) { continue }
    $locale = Read-Locale $file.FullName
    foreach ($e in $locale.Errors) { $allErrors += "$($file.Name): $e" }

    foreach ($key in $reference.Entries.Keys) {
        if (!$locale.Entries.Contains($key)) {
            if ($pendingKeySet.Contains($key)) {
                if ($FailOnPending) {
                    $allErrors += "$($file.Name): 缺少键（待翻译） $key"
                }
                else {
                    $pendingMissingCount++
                }
            }
            else {
                $allErrors += "$($file.Name): 缺少键 $key"
            }
        }
    }
    foreach ($key in $locale.Entries.Keys) {
        if (!$reference.Entries.Contains($key)) {
            $allErrors += "$($file.Name): 多余键 $key"
            continue
        }

        $value = $locale.Entries[$key]
        if ($RequireTranslated -and $value -ceq $key) {
            $allErrors += "$($file.Name): 未翻译（英文回退）→ $key"
        }
        if ($ScriptCheck -and !(Test-TargetScript $file.BaseName $key $value)) {
            $message = "$($file.Name): 译文不含目标语言书写系统 → $key = $value"
            if ($StrictScript) { $allErrors += $message } else { $allWarnings += $message }
        }
        $expected = Get-Placeholders $key
        $actual = Get-Placeholders $value
        if (!$expected.SetEquals($actual)) {
            $allErrors += "$($file.Name): 占位符不一致 → $key"
            continue
        }

        # 格式说明符只在译文保留该占位符时比对，缺失时仅告警
        $enSpecs = Get-PlaceholderSpecs $key
        $trSpecs = Get-PlaceholderSpecs $value
        foreach ($index in $enSpecs.Keys) {
            if ($trSpecs.ContainsKey($index) -and $trSpecs[$index] -ne $enSpecs[$index]) {
                $allWarnings += "$($file.Name): 格式说明符不一致 → $key"
                break
            }
        }
    }
}

# 孤儿键与未登记键：词表与源码双向覆盖
$orphans = @()
if (!$SkipOrphan) {
    $sourceDirs = @($SourceDir | ForEach-Object {
        if ([IO.Path]::IsPathRooted($_)) { $_ } else { Join-Path $root $_ }
    } | Where-Object { Test-Path $_ })

    $sourceFiles = @($sourceDirs | ForEach-Object { Get-ChildItem $_ -Recurse -Filter *.cs } |
        Where-Object { $_.FullName -notmatch '\\(obj|bin|Locales)\\' })
    # 源文范围含 .csproj：程序集标题/描述等元数据写在 csproj 中，同样是词表源文
    $sourceTextFiles = @($sourceDirs | ForEach-Object { Get-ChildItem $_ -Recurse -Include *.cs, *.csproj } |
        Where-Object { $_.FullName -notmatch '\\(obj|bin|Locales)\\' })
    $sourceText = ($sourceTextFiles | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"

    foreach ($key in $reference.Entries.Keys) {
        # 非源码来源键（属性元数据、外部 Model.xml 等）由外部数据引用，不参与源码归属判定
        if ($metadataKeySet.Contains($key)) { continue }

        # 源码中的字面量形式：制表符写作 \t，引号写作 \"
        $escaped = $key.Replace('\', '\\').Replace('"', '\"').Replace("`t", '\t').Replace("`r", '\r').Replace("`n", '\n')
        if (!$sourceText.Contains($key) -and !$sourceText.Contains($escaped)) {
            $orphans += $key
        }
    }
    foreach ($o in $orphans) { $allErrors += "孤儿键（源码中找不到）：$o" }

    # 反向覆盖：源码中 L.T / Localization.Translate 的英文源文必须已登记，否则运行时静默回退英文
    $usedPattern = '(?:L\.T|L\.TLog|Localization\.Translate(?:ForLog)?)\s*\(\s*(?:[A-Za-z_][\w\.]*\s*,\s*)?"((?:[^"\\]|\\.)*)"'
    $reported = @{}
    foreach ($file in $sourceFiles) {
        $text = [IO.File]::ReadAllText($file.FullName)
        $text = [regex]::Replace($text, '(?s)/\*.*?\*/', '')
        $text = [regex]::Replace($text, '(?m)^\s*///.*$', '')
        $text = [regex]::Replace($text, '(?m)//.*$', '')
        foreach ($m in [regex]::Matches($text, $usedPattern)) {
            $key = Unescape-Yaml $m.Groups[1].Value
            if (!$reference.Entries.Contains($key) -and !$reported.ContainsKey($key)) {
                $reported[$key] = $true
                $rel = $file.FullName
                if ($rel.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { $rel = $rel.Substring($root.Length).TrimStart("\", "/") }
                $allErrors += "未登记词条（源码使用但词表缺失）：$rel → $key"
            }
        }
    }
}

# 迁移期待翻译清单：存在待翻译键时必须显式放行，防止静默长期回退英文
if ($pendingKeySet.Count -gt 0) {
    if ($FailOnPending) {
        $allErrors += "待翻译清单仍有 $($pendingKeySet.Count) 条，发布前必须清零：$pendingPath"
    }
    elseif (!$AllowPending) {
        $allErrors += "存在 $($pendingKeySet.Count) 条待翻译键，请使用 -AllowPending 继续开发，或在发布前补齐：$pendingPath"
    }
    else {
        $allWarnings += "待翻译清单 $($pendingKeySet.Count) 条（-AllowPending 放行，缺失 $pendingMissingCount 处）：$pendingPath"
    }
}

# 覆盖报告：各语言键数、缺失键、英文回退数，供批量翻译阶段核对
if ($Report) {
    $reportPath = if ([IO.Path]::IsPathRooted($Report)) { $Report } else { Join-Path $root $Report }
    $reportDir = Split-Path -Parent $reportPath
    if ($reportDir -and !(Test-Path $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
    $lines = @("culture`tkeys`tmissing`tfallback")
    foreach ($file in Get-ChildItem $localeDir -Filter *.yml | Sort-Object Name) {
        $entry = Read-Locale $file.FullName
        $missing = 0
        $fallback = 0
        foreach ($key in $reference.Entries.Keys) {
            if (!$entry.Entries.Contains($key)) { $missing++; continue }
            if ($entry.Entries[$key] -ceq $key) { $fallback++ }
        }
        $lines += "{0}`t{1}`t{2}`t{3}" -f $file.BaseName, $entry.Entries.Count, $missing, $fallback
    }
    [IO.File]::WriteAllLines($reportPath, $lines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "覆盖报告已写入：$reportPath"
}

foreach ($e in $allErrors) { Write-Host $e -ForegroundColor Red }
foreach ($w in $allWarnings) { Write-Host $w -ForegroundColor Yellow }

$count = (Get-ChildItem $localeDir -Filter *.yml).Count
if ($allErrors.Count -gt 0) {
    Write-Host "多语言词表校验未通过：$count 个语言，$($allErrors.Count) 个错误" -ForegroundColor Red
    exit 1
}

Write-Host "多语言词表校验通过：$count 个语言，基准 $($reference.Entries.Count) 条" -ForegroundColor Green
