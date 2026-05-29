param(
    [switch]$Build,
    [switch]$Test,
    [switch]$Prerelease,
    [string]$Repo = "ZhouJun2303/ForkExcelMergeTool"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Read-AppVersion {
    $content = Get-Content -LiteralPath "Scripts\version.py" -Raw -Encoding UTF8
    $match = [regex]::Match($content, '__version__\s*=\s*"%d\.%d"\s*%\s*\(VERSION_MAJOR,\s*VERSION_MINOR\)')
    if ($match.Success) {
        $major = [regex]::Match($content, 'VERSION_MAJOR\s*=\s*(\d+)').Groups[1].Value
        $minor = [regex]::Match($content, 'VERSION_MINOR\s*=\s*(\d+)').Groups[1].Value
        return "$major.$minor"
    }
    $direct = [regex]::Match($content, '__version__\s*=\s*["'']([^"'']+)["'']')
    if ($direct.Success) {
        return $direct.Groups[1].Value
    }
    throw "无法从 Scripts\version.py 读取版本号"
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "未找到 GitHub CLI: gh。请先安装并执行 gh auth login。"
}

if ($Build) {
    Write-Host "Building ExcelMergeFork.exe..."
    $buildArgs = @("--gui")
    & ".\build_exe.bat" @buildArgs
    if ($LASTEXITCODE -ne 0) {
        throw "打包失败"
    }
}

if ($Test) {
    Write-Host "Running merge mode tests..."
    & ".\run_merge_mode_tests.bat"
    if ($LASTEXITCODE -ne 0) {
        throw "测试失败"
    }
}

$version = Read-AppVersion
$tag = "v$version"
$exe = Join-Path $root "ExcelMergeFork.exe"
$shaFile = Join-Path $root "ExcelMergeFork.exe.sha256"

if (-not (Test-Path -LiteralPath $exe)) {
    throw "未找到产物: $exe。请先运行 .\package.bat 或使用 -Build。"
}

$sha = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant()
"$sha  ExcelMergeFork.exe" | Set-Content -LiteralPath $shaFile -Encoding ASCII

Write-Host "Version: $version"
Write-Host "Tag: $tag"
Write-Host "SHA256: $sha"

$existing = $null
$viewExitCode = 1
try {
    $existing = & gh release view $tag --repo $Repo --json tagName 2>$null
    $viewExitCode = $LASTEXITCODE
} catch {
    $existing = $null
    $viewExitCode = 1
}
if ($viewExitCode -eq 0 -and $existing) {
    throw "Release $tag 已存在。请递增版本号或手动处理已有 Release。"
}

$notes = @"
ExcelMergeFork $tag

- Windows 单文件版: ExcelMergeFork.exe
- SHA256: 见 ExcelMergeFork.exe.sha256
"@

$target = (& git rev-parse HEAD 2>$null).Trim()
if (-not $target) {
    $target = "HEAD"
}

$ghArgs = @(
    "release", "create", $tag,
    $exe,
    $shaFile,
    "--repo", $Repo,
    "--title", "ExcelMergeFork $tag",
    "--notes", $notes,
    "--target", $target
)
if ($Prerelease) {
    $ghArgs += "--prerelease"
}

& gh @ghArgs
if ($LASTEXITCODE -ne 0) {
    throw "GitHub Release 发布失败"
}

Write-Host "Release published: https://github.com/$Repo/releases/tag/$tag"
