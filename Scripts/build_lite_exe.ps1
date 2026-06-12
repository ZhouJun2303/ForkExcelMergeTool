param(
    [string]$OutputPath = "dist\ExcelMergeFork-lite.exe",
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Read-AppVersion {
    $content = Get-Content -LiteralPath "Scripts\version.py" -Raw -Encoding UTF8
    $major = [regex]::Match($content, 'VERSION_MAJOR\s*=\s*(\d+)').Groups[1].Value
    $minor = [regex]::Match($content, 'VERSION_MINOR\s*=\s*(\d+)').Groups[1].Value
    if ($major -and $minor) {
        return "$major.$minor"
    }
    $direct = [regex]::Match($content, '__version__\s*=\s*["'']([^"'']+)["'']')
    if ($direct.Success) {
        return $direct.Groups[1].Value
    }
    throw "无法从 Scripts\version.py 读取版本号"
}

if (-not (Get-Command go -ErrorAction SilentlyContinue)) {
    throw "未找到 Go 编译器。请先安装 Go，或只构建完整单文件版。"
}

$cacheRoot = Join-Path $root ".cache"
if (-not (Test-Path -LiteralPath $cacheRoot)) {
    New-Item -ItemType Directory -Path $cacheRoot | Out-Null
}
if (-not $env:GOCACHE) {
    $env:GOCACHE = Join-Path $cacheRoot "go-build"
}

$launcherDir = Join-Path $root "Tools\lite_launcher"
if (-not (Test-Path -LiteralPath (Join-Path $launcherDir "main.go"))) {
    throw "未找到轻量启动器源码: $launcherDir"
}

$payloadRel = "Tools\lite_launcher\payload.zip"
Write-Host "Creating lite payload..."
& powershell -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\make_python_zip.ps1" -OutputPath $payloadRel
if ($LASTEXITCODE -ne 0) {
    throw "轻量版 payload 生成失败"
}

$version = Read-AppVersion
$out = Join-Path $root $OutputPath
$outDir = Split-Path -Parent $out
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

Write-Host "Building ExcelMergeFork-lite.exe (v$version)..."
Push-Location $launcherDir
try {
    & go build -trimpath -ldflags "-s -w -H=windowsgui -X main.appVersion=$version" -o $out .
    if ($LASTEXITCODE -ne 0) {
        throw "Go build failed"
    }
}
finally {
    Pop-Location
}

if (-not (Test-Path -LiteralPath $out)) {
    throw "未找到轻量版产物: $out"
}

$rootCopy = Join-Path $root "ExcelMergeFork-lite.exe"
Copy-Item -LiteralPath $out -Destination $rootCopy -Force
Write-Host "Generated $rootCopy"

if ($SelfTest) {
    Write-Host "Running lite self-test..."
    $oldLocalAppData = $env:LOCALAPPDATA
    $selfTestLocalAppData = Join-Path $cacheRoot "lite-localappdata"
    if (-not (Test-Path -LiteralPath $selfTestLocalAppData)) {
        New-Item -ItemType Directory -Path $selfTestLocalAppData | Out-Null
    }
    $env:LOCALAPPDATA = $selfTestLocalAppData
    try {
        & $rootCopy --lite-self-test
        if ($LASTEXITCODE -ne 0) {
            throw "轻量版自检失败"
        }
    }
    finally {
        $env:LOCALAPPDATA = $oldLocalAppData
    }
}
