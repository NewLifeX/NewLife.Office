# 构建 office 命令的独立单文件产物（每个 RID 一个 office.exe，无附带 dll）
#
# 设计取舍：不采用 NativeAOT。实测 NativeAOT 产物 > 10MB，超过约定的 5MB 阈值，
# 而"依赖 .NET 8 的单文件"仅约 3.6MB 且启动即用；AOT 兼容性由 AotSmoke 单独守护（库侧）。
#
# 用法：
#   pwsh -File tools/publish-cli.ps1                                  # 全部 RID
#   pwsh -File tools/publish-cli.ps1 -Runtime win-x64,linux-x64        # 指定 RID
#   pwsh -File tools/publish-cli.ps1 -OutDir out -VersionSuffix 2026.1103
param(
    [String[]]$Runtime = @('win-x64', 'win-x86', 'win-arm64', 'linux-x64', 'linux-arm64', 'osx-x64', 'osx-arm64'),
    [String]$OutDir = 'Bin/Cli-dist',
    [String]$VersionSuffix = ''
)

$ErrorActionPreference = 'Stop'

# pwsh -File 传参不会按逗号拆分，这里统一切开
$Runtime = @($Runtime | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'NewLife.Office.Cli/NewLife.Office.Cli.csproj'
$out = if ([IO.Path]::IsPathRooted($OutDir)) { $OutDir } else { Join-Path $root $OutDir }
New-Item -ItemType Directory -Path $out -Force | Out-Null

$rows = @()
foreach ($rid in $Runtime) {
    $dir = Join-Path $out $rid
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }

    $args = @('publish', $project, '-c', 'Release', '-r', $rid, '--self-contained', 'false',
        '-p:PublishSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true',
        '-p:DebugType=None', '-p:PublishDocumentationFiles=false', '-o', $dir)
    if ($VersionSuffix) { $args += @('-p:VersionSuffix=' + $VersionSuffix) }

    & dotnet @args
    if ($LASTEXITCODE -ne 0) { throw "publish 失败：$rid" }

    # 单文件产物之外可能还有 pdb/xml，清掉保证"只有一个可执行文件"
    Get-ChildItem $dir -File | Where-Object { $_.Name -notmatch '^office(\.exe)?$' } | Remove-Item -Force
    $exe = Get-ChildItem $dir -File | Select-Object -First 1
    if (!$exe) { throw "未找到可执行文件：$rid" }

    $zip = Join-Path $out "office-$rid.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path $exe.FullName -DestinationPath $zip

    $mb = [Math]::Round($exe.Length / 1MB, 2)
    $rows += [pscustomobject]@{ Runtime = $rid; File = $exe.Name; MB = $mb; Zip = (Split-Path -Leaf $zip) }
    if ($mb -gt 5) { Write-Warning "$rid 产物 $mb MB 超过 5MB 约定阈值" }
}

$rows | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "产物目录：$out"
