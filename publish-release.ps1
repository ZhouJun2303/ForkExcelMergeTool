$SkipTest = $false
$Prerelease = $false
$NoPause = $false
foreach ($a in $args) {
    switch -Regex ($a) {
        "^(?i)(-SkipTest|--dist|--no-test)$" { $SkipTest = $true }
        "^(?i)(-Prerelease|--prerelease)$" { $Prerelease = $true }
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

Write-Host "========== ExcelMergeFork Release =========="
Write-Host ""

$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    Fail "GitHub CLI not found. Run: winget install --id GitHub.cli -e   then   gh auth login"
}

& gh auth status 2>$null
if ($LASTEXITCODE -ne 0) {
    Fail "GitHub CLI is not logged in. Run: gh auth login"
}

$versionFile = Join-Path $PSScriptRoot "src\ExcelMergeFork.Core\AppVersion.cs"
$text = Get-Content -Raw $versionFile
if ($text -notmatch 'Display = "([^"]+)"') {
    Fail "cannot read version from AppVersion.cs"
}
$version = $Matches[1]
$tag = "v$version"
Write-Host "Version: $tag"
Write-Host ""

$packageArgs = @("-NoPause", "-PackageZip")
if ($SkipTest) { $packageArgs += "-SkipTest" }
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "package.ps1") @packageArgs
if ($LASTEXITCODE -ne 0) { Fail "package failed" }

$exe = Join-Path $PSScriptRoot "ExcelMergeFork.exe"
$sha = Join-Path $PSScriptRoot "ExcelMergeFork.exe.sha256"
$zip = Join-Path $PSScriptRoot "dist\ExcelMergeFork-package.zip"
foreach ($f in @($exe, $sha, $zip)) {
    if (-not (Test-Path $f)) { Fail "missing artifact: $f" }
}

$notes = @"
ExcelMergeFork $tag

Windows self-contained desktop build (.NET 8 / WPF Fluent). Python is no longer required.

Download:
- ExcelMergeFork-package.zip (exe + Fork/Git install scripts)
- or ExcelMergeFork.exe alone

Fork setup:
- Merge/Diff Tool Path: full path to ExcelMergeFork.exe
- Merge Arguments: `$LOCAL,`$BASE,`$REMOTE,`$MERGED
- Diff Arguments: `"`$REMOTE`" `"`$LOCAL`"

Or run install_fork_integration.bat after unzip.

Checksum: ExcelMergeFork.exe.sha256
"@

$target = ""
try {
    $target = (& git -C $PSScriptRoot rev-parse HEAD).Trim()
} catch {
    Fail "cannot read git HEAD for --target"
}
if (-not $target) { Fail "empty git HEAD" }
Write-Host "Target commit: $target"

function Invoke-Gh([string[]] $GhArgs) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & gh @GhArgs
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
}

$existsCode = Invoke-Gh @("release", "view", $tag)
if ($existsCode -eq 0) {
    Write-Host "Release $tag already exists. Uploading assets with --clobber..."
    $uploadCode = Invoke-Gh @("release", "upload", $tag, $exe, $sha, $zip, "--clobber")
    if ($uploadCode -ne 0) { Fail "gh release upload failed" }
} else {
    Write-Host "Creating release $tag..."
    $createArgs = @(
        "release", "create", $tag,
        "--title", "ExcelMergeFork $tag",
        "--notes", $notes,
        "--target", $target,
        "--latest",
        $exe, $sha, $zip
    )
    if ($Prerelease) { $createArgs += "--prerelease" }
    $createCode = Invoke-Gh $createArgs
    if ($createCode -ne 0) { Fail "gh release create failed" }
}

Write-Host ""
Write-Host "Release published: $tag" -ForegroundColor Green
if (-not $NoPause) {
    Write-Host ""
    Pause
}
exit 0
