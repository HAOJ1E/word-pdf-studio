[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repoRoot = Split-Path -Parent $PSScriptRoot
$engineDirectory = Join-Path $repoRoot 'engine'
$engineArchive = Join-Path $engineDirectory 'libreoffice.zip'
$cacheDirectory = Join-Path $repoRoot '.cache'
$msiPath = Join-Path $cacheDirectory 'LibreOffice_26.2.6_Win_x86-64.msi'
$downloadUrl = 'https://download.documentfoundation.org/libreoffice/stable/26.2.6/win/x86_64/LibreOffice_26.2.6_Win_x86-64.msi'
$expectedMsiSha256 = 'F9877032FD908BEB9C0DDF06DF4AF5C2E85F419C42E14876C4CCE5AAE5FB2660'

New-Item -ItemType Directory -Path $engineDirectory, $cacheDirectory -Force | Out-Null

if ((Test-Path -LiteralPath $engineArchive) -and -not $Force) {
    Write-Host "转换内核已存在：$engineArchive"
    exit 0
}

if (-not (Test-Path -LiteralPath $msiPath)) {
    Write-Host '正在从 LibreOffice 官方镜像下载安装包……'
    Invoke-WebRequest -Uri $downloadUrl -OutFile $msiPath
}

$actualHash = (Get-FileHash -LiteralPath $msiPath -Algorithm SHA256).Hash
if ($actualHash -ne $expectedMsiSha256) {
    throw "LibreOffice 安装包校验失败。期望 $expectedMsiSha256，实际 $actualHash"
}

$extractRoot = Join-Path ([IO.Path]::GetTempPath()) ("WordPdfStudio-LO-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $extractRoot | Out-Null

try {
    Write-Host '正在提取 LibreOffice 便携运行文件……'
    $msiArguments = @('/a', "`"$msiPath`"", '/qn', "TARGETDIR=`"$extractRoot`"")
    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArguments -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "LibreOffice 提取失败，msiexec 返回代码 $($process.ExitCode)"
    }

    $soffice = Get-ChildItem -LiteralPath $extractRoot -Filter 'soffice.com' -File -Recurse | Select-Object -First 1
    if (-not $soffice) {
        throw '提取结果中没有找到 program\soffice.com。'
    }

    $installRoot = $soffice.Directory.Parent.FullName
    $programDirectory = $soffice.Directory.FullName
    $system64 = Join-Path $installRoot 'System64'
    if (Test-Path -LiteralPath $system64) {
        Get-ChildItem -LiteralPath $system64 -Filter '*.dll' -File | Copy-Item -Destination $programDirectory -Force
    }

    if (Test-Path -LiteralPath $engineArchive) {
        Remove-Item -LiteralPath $engineArchive -Force
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($engineArchive, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $rootWithSeparator = $installRoot.TrimEnd('\') + '\'
        foreach ($file in Get-ChildItem -LiteralPath $installRoot -File -Recurse) {
            $relativePath = $file.FullName.Substring($rootWithSeparator.Length).Replace('\', '/')
            if ($relativePath -match '^(help|System|System64)/' -or $relativePath -match '\.msi$') {
                continue
            }

            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $file.FullName,
                $relativePath,
                [IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
        }
    }
    finally {
        $archive.Dispose()
    }

    Write-Host "转换内核已生成：$engineArchive"
}
finally {
    if (Test-Path -LiteralPath $extractRoot) {
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }
}
