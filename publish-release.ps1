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

function Update-AppVersion {
    $versionFile = Join-Path $PSScriptRoot "src\ExcelMergeFork.Core\AppVersion.cs"
    $versionText = Get-Content -Raw $versionFile
    $majorMatch = [regex]::Match($versionText, 'public const int Major = (\d+);')
    $minorMatch = [regex]::Match($versionText, 'public const int Minor = (\d+);')
    $displayMatch = [regex]::Match($versionText, 'public const string Display = "([^"]+)";')
    if (-not $majorMatch.Success -or -not $minorMatch.Success -or -not $displayMatch.Success) {
        Fail "cannot read version from AppVersion.cs"
    }

    $major = [int]$majorMatch.Groups[1].Value
    $minor = [int]$minorMatch.Groups[1].Value + 1
    $display = "$major.$minor"
    $updatedVersionText = [regex]::Replace(
        $versionText,
        'public const int Minor = \d+;',
        "public const int Minor = $minor;"
    )
    $updatedVersionText = [regex]::Replace(
        $updatedVersionText,
        'public const string Display = "[^"]+";',
        "public const string Display = `"$display`";"
    )
    Set-Content -Path $versionFile -Value $updatedVersionText -Encoding utf8

    $projectFile = Join-Path $PSScriptRoot "src\ExcelMergeFork.App\ExcelMergeFork.App.csproj"
    $projectText = Get-Content -Raw $projectFile
    $updatedProjectText = [regex]::Replace($projectText, '<Version>[^<]+</Version>', "<Version>$display.0</Version>")
    $updatedProjectText = [regex]::Replace($updatedProjectText, '<InformationalVersion>[^<]+</InformationalVersion>', "<InformationalVersion>$display</InformationalVersion>")
    Set-Content -Path $projectFile -Value $updatedProjectText -Encoding utf8

    $manifestFile = Join-Path $PSScriptRoot "src\ExcelMergeFork.App\app.manifest"
    $manifestText = Get-Content -Raw $manifestFile
    $updatedManifestText = [regex]::Replace(
        $manifestText,
        '(<assemblyIdentity version=")[^"]+(" name="ExcelMergeFork\.app")',
        [System.Text.RegularExpressions.MatchEvaluator] {
            param($match)
            return "$($match.Groups[1].Value)$display.0.0$($match.Groups[2].Value)"
        }
    )
    Set-Content -Path $manifestFile -Value $updatedManifestText -Encoding utf8

    Write-Host "Version bumped: v$display"
}

Write-Host "========== ExcelMergeFork Release =========="
Write-Host ""

$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    Fail "GitHub CLI not found. Run: winget install --id GitHub.cli -e   then   gh auth login"
}

$oldEap = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"
& gh auth status 2>$null
$authCode = [int]$LASTEXITCODE
$ErrorActionPreference = $oldEap
if ($authCode -ne 0) {
    Fail "GitHub CLI is not logged in. Run: gh auth login"
}

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    Fail "dotnet not found. Install .NET 8 SDK from https://dotnet.microsoft.com/download"
}

Update-AppVersion

$packageArgs = @("-NoPause", "-PackageZip")
if ($SkipTest) { $packageArgs += "-SkipTest" }
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "package.ps1") @packageArgs
if ($LASTEXITCODE -ne 0) { Fail "package failed" }

$versionFile = Join-Path $PSScriptRoot "src\ExcelMergeFork.Core\AppVersion.cs"
$text = Get-Content -Raw $versionFile
if ($text -notmatch 'Display = "([^"]+)"') {
    Fail "cannot read version from AppVersion.cs"
}
$version = $Matches[1]
$tag = "v$version"
Write-Host "Version: $tag"
Write-Host ""

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
        return [int]$LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
}

function Test-GhRelease([string] $ReleaseTag) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    try {
        $null = & gh release view $ReleaseTag --json tagName 2>$null
        return ($LASTEXITCODE -eq 0)
    } finally {
        $ErrorActionPreference = $old
    }
}

$notesFile = Join-Path $env:TEMP "ExcelMergeFork-release-notes.md"
Set-Content -Path $notesFile -Value $notes -Encoding utf8

if (Test-GhRelease $tag) {
    Write-Host "Release $tag already exists. Uploading assets with --clobber..."
    $uploadCode = Invoke-Gh @("release", "upload", $tag, $exe, $sha, $zip, "--clobber")
    if ($uploadCode -ne 0) { Fail "gh release upload failed" }
} else {
    Write-Host "Creating release $tag..."
    $createArgs = @(
        "release", "create", $tag,
        "--title", "ExcelMergeFork $tag",
        "--notes-file", $notesFile,
        "--target", $target,
        "--latest",
        $exe, $sha, $zip
    )
    if ($Prerelease) { $createArgs += "--prerelease" }
    $createCode = Invoke-Gh $createArgs
    if ($createCode -ne 0 -and -not (Test-GhRelease $tag)) {
        Fail "gh release create failed"
    }
}

Write-Host ""
Write-Host "Release published: $tag" -ForegroundColor Green
if (-not $NoPause) {
    Write-Host ""
    Pause
}
exit 0
