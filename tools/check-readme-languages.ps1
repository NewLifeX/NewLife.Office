# README 多语言完整性校验：语种文件、结构、代码块、目标语言书写系统、整段英文残留
# 只做结构性与启发式检查，语义是否准确仍需人工/子代理通读。
#
# 用法：
#   pwsh -File tools/check-readme-languages.ps1                 # 校验并输出报告
#   pwsh -File tools/check-readme-languages.ps1 -Strict         # 告警也失败
#   pwsh -File tools/check-readme-languages.ps1 -Root D:\src\MyLib
param(
    [String]$Root = (Split-Path -Parent $PSScriptRoot),
    [String]$Config = "readme-languages.json",
    [Switch]$Strict
)

$ErrorActionPreference = "Stop"

$configPath = if ([IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $Root $Config }
if (!(Test-Path -LiteralPath $configPath)) { throw "未找到语言配置：$configPath" }

$langs = @((Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json).languages)
if ($langs.Count -eq 0) { throw "$Config 未配置任何语种" }

# 目标语言书写系统：用于识别整份英文复制的低质量翻译
$scriptMap = @{
    'zh-CN' = '[\u4e00-\u9fff]'
    'zh-TW' = '[\u4e00-\u9fff]'
    'ja'    = '[\u3040-\u30ff\u4e00-\u9fff]'
    'ko'    = '[\uac00-\ud7af]'
    'ru'    = '[\u0400-\u04ff]'
    'uk'    = '[\u0400-\u04ff]'
    'el'    = '[\u0370-\u03ff]'
    'ar'    = '[\u0600-\u06ff]'
    'fa'    = '[\u0600-\u06ff]'
    'he'    = '[\u0590-\u05ff]'
    'hi'    = '[\u0900-\u097f]'
    'bn'    = '[\u0980-\u09ff]'
    'th'    = '[\u0e00-\u0e7f]'
}

function Get-FenceLanguages([string]$text) {
    $langs = @()
    foreach ($m in [regex]::Matches($text, '(?m)^```([A-Za-z0-9_+-]*)\s*$')) {
        $langs += $m.Groups[1].Value.ToLowerInvariant()
    }
    return ,$langs
}

function Get-HeadingCounts([string]$text) {
    $counts = @{}
    foreach ($m in [regex]::Matches($text, '(?m)^(#{2,4})\s+\S')) {
        $level = $m.Groups[1].Value.Length
        if (!$counts.ContainsKey($level)) { $counts[$level] = 0 }
        $counts[$level]++
    }
    return $counts
}

function Get-ProseLines([string]$text) {
    # 去掉围栏代码块，只保留正文行，避免把代码/命令误判为未翻译
    $clean = [regex]::Replace($text, '(?ms)^```.*?^```\s*$', '')
    return @($clean -split "`r?`n")
}

$english = $langs | Where-Object { $_.code -eq 'en' } | Select-Object -First 1
if (!$english) { throw "readme-languages.json 缺少 English（code=en）条目" }
$englishPath = Join-Path $Root $english.file
$englishText = [IO.File]::ReadAllText($englishPath)
$englishFences = Get-FenceLanguages $englishText
$englishHeadings = Get-HeadingCounts $englishText
$englishProse = [Collections.Generic.HashSet[String]]::new([StringComparer]::Ordinal)
foreach ($line in Get-ProseLines $englishText) {
    $t = $line.Trim()
    if ($t.Length -ge 120) { [void]$englishProse.Add($t) }
}

$errors = @()
$warnings = @()

foreach ($item in $langs) {
    $path = Join-Path $Root $item.file
    if (!(Test-Path -LiteralPath $path)) {
        $errors += "$($item.file): 文件不存在"
        continue
    }
    $text = [IO.File]::ReadAllText($path)
    if ($text.Trim().Length -eq 0) {
        $errors += "$($item.file): 文件为空"
        continue
    }

    $fences = Get-FenceLanguages $text
    if ($fences.Count -lt $englishFences.Count) {
        $errors += "$($item.file): 代码块数量 $($fences.Count) < 英文 $($englishFences.Count)，可能缺少示例"
    }
    elseif ($fences.Count -gt $englishFences.Count) {
        $warnings += "$($item.file): 代码块数量 $($fences.Count) > 英文 $($englishFences.Count)"
    }
    elseif (($fences -join ',') -ne ($englishFences -join ',')) {
        $warnings += "$($item.file): 代码块语言序列与英文不一致"
    }

    $headings = Get-HeadingCounts $text
    foreach ($level in $englishHeadings.Keys) {
        $actual = if ($headings.ContainsKey($level)) { $headings[$level] } else { 0 }
        if ($actual -lt $englishHeadings[$level]) {
            $warnings += "$($item.file): 标题层级 $('#' * $level) 数量 $actual < 英文 $($englishHeadings[$level])，可能缺少章节"
        }
    }

    if ($item.code -ne 'en' -and $scriptMap.ContainsKey($item.code)) {
        $matches = [regex]::Matches($text, $scriptMap[$item.code]).Count
        if ($matches -lt 50) {
            $errors += "$($item.file): 目标语言字符仅 $matches 个，疑似整份英文复制"
        }
    }

    if ($item.code -ne 'en') {
        $copied = 0
        foreach ($line in Get-ProseLines $text) {
            $t = $line.Trim()
            if ($t.Length -lt 120) { continue }
            if ($t.StartsWith('|') -or $t.StartsWith('#') -or $t.StartsWith('<') -or $t.StartsWith('[') -or $t.StartsWith('!') -or $t.StartsWith('>')) { continue }
            if ($t.Contains('http://') -or $t.Contains('https://')) { continue }
            if ($englishProse.Contains($t)) { $copied++ }
        }
        if ($copied -gt 0) {
            $warnings += "$($item.file): 发现 $copied 行与英文完全相同的长正文，可能未翻译"
        }
    }
}

foreach ($e in $errors) { Write-Host $e -ForegroundColor Red }
foreach ($w in $warnings) { Write-Host $w -ForegroundColor Yellow }

if ($errors.Count -gt 0 -or ($Strict -and $warnings.Count -gt 0)) {
    Write-Host "README 多语言校验未通过：$($langs.Count) 个语种，$($errors.Count) 个错误，$($warnings.Count) 个告警" -ForegroundColor Red
    exit 1
}

Write-Host "README 多语言校验通过：$($langs.Count) 个语种，$($warnings.Count) 个告警" -ForegroundColor Green
