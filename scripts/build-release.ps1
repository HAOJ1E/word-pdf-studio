[CmdletBinding()]
param(
    [switch]$SkipEnginePreparation
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$engineArchive = Join-Path $repoRoot 'engine\libreoffice.zip'
$project = Join-Path $repoRoot 'WordPdfStudio\WordPdfStudio.csproj'
$publishDirectory = Join-Path $repoRoot 'artifacts\publish'
$releaseDirectory = Join-Path $repoRoot 'artifacts\release'

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw '未找到 dotnet。请先安装 .NET 10 SDK。'
}

if (-not (Test-Path -LiteralPath $engineArchive)) {
    if ($SkipEnginePreparation) {
        throw '缺少 engine\libreoffice.zip，无法生成完整离线版。'
    }

    & (Join-Path $PSScriptRoot 'prepare-engine.ps1')
}

New-Item -ItemType Directory -Path $publishDirectory, $releaseDirectory -Force | Out-Null

dotnet publish $project `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output $publishDirectory `
    -p:PublishSingleFile=true

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish 失败，退出代码 $LASTEXITCODE"
}

$publishedExe = Join-Path $publishDirectory 'WordPdfStudio.exe'
$releaseExe = Join-Path $releaseDirectory '文页PDF_离线便携版_v1.0.exe'
Copy-Item -LiteralPath $publishedExe -Destination $releaseExe -Force
$hash = (Get-FileHash -LiteralPath $releaseExe -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $releaseDirectory 'SHA256.txt') -Value "$hash  文页PDF_离线便携版_v1.0.exe" -Encoding utf8

Write-Host "构建完成：$releaseExe"
Write-Host "SHA-256：$hash"
