$SkipTest = $false
$PackageZip = $false
$NoPause = $false
foreach ($a in $args) {
    switch -Regex ($a) {
        "^(?i)(-SkipTest|--dist|--no-test|no-test|skip)$" { $SkipTest = $true }
        "^(?i)(-PackageZip|--package-zip)$" { $PackageZip = $true }
        "^(?i)(-NoPause|--no-pause)$" { $NoPause = $true }
    }
}

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Fail([string] $Message) {
    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    if (-not $NoPause) { Pause }
    exit 1
}

Write-Host "========== ExcelMergeFork Package =========="
Write-Host ""

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    Fail "dotnet not found. Install .NET 8 SDK from https://dotnet.microsoft.com/download"
}

if (-not $SkipTest) {
    Write-Host "[1/3] Running tests..."
    & dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release --nologo
    if ($LASTEXITCODE -ne 0) { Fail "tests failed" }
} else {
    Write-Host "[1/3] Tests skipped"
}

Write-Host ""
Write-Host "[2/3] Publishing self-contained ExcelMergeFork.exe..."
New-Item -ItemType Directory -Force -Path dist\csharp | Out-Null
& dotnet publish src\ExcelMergeFork.App\ExcelMergeFork.App.csproj `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:DebugType=none `
    -o dist\csharp
if ($LASTEXITCODE -ne 0) { Fail "publish failed" }

$published = Join-Path $PSScriptRoot "dist\csharp\ExcelMergeFork.exe"
if (-not (Test-Path $published)) { Fail "published exe missing: $published" }

try {
    Copy-Item -Force $published (Join-Path $PSScriptRoot "ExcelMergeFork.exe")
} catch {
    Fail "cannot replace ExcelMergeFork.exe. Close the running app and retry. $($_.Exception.Message)"
}

$hash = (Get-FileHash -Algorithm SHA256 -Path .\ExcelMergeFork.exe).Hash.ToLowerInvariant()
Set-Content -Path .\ExcelMergeFork.exe.sha256 -Value "$hash  ExcelMergeFork.exe" -Encoding ascii

Write-Host ""
Write-Host "[3/3] Done. Output: $PSScriptRoot\ExcelMergeFork.exe"
Write-Host "SHA256: $hash"

if ($PackageZip) {
    $zip = Join-Path $PSScriptRoot "dist\ExcelMergeFork-package.zip"
    if (Test-Path $zip) { Remove-Item -Force $zip }
    Compress-Archive -Force -Path @(
        ".\ExcelMergeFork.exe",
        ".\README.md",
        ".\install_fork_integration.bat",
        ".\uninstall_fork_integration.bat",
        ".\install_git_integration.bat",
        ".\uninstall_git_integration.bat"
    ) -DestinationPath $zip
    Write-Host "Package zip: $zip"
}

if (-not $NoPause) {
    Write-Host ""
    Pause
}
exit 0
