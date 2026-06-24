param(
    [ValidateSet("full", "lite")]
    [string]$Kind = "full",
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path ".").Path

if (-not $OutputPath) {
    if ($Kind -eq "lite") {
        $OutputPath = "dist\ExcelMergeFork-lite-package.zip"
    } else {
        $OutputPath = "dist\ExcelMergeFork-package.zip"
    }
}

$out = Join-Path $root $OutputPath
$outDir = Split-Path -Parent $out
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
if (Test-Path -LiteralPath $out) {
    Remove-Item -LiteralPath $out -Force
}

$exeName = if ($Kind -eq "lite") { "ExcelMergeFork-lite.exe" } else { "ExcelMergeFork.exe" }
$exePath = Join-Path $root $exeName
if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw "未找到产物: $exePath。请先运行 package.bat。"
}

$stage = Join-Path $env:TEMP ("ExcelMergeFork-package-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $stage | Out-Null

$topLevel = @(
    $exeName,
    "README.md",
    "install_fork_integration.bat",
    "uninstall_fork_integration.bat",
    "install_git_integration.bat",
    "uninstall_git_integration.bat"
)

try {
    foreach ($item in $topLevel) {
        $src = Join-Path $root $item
        if (Test-Path -LiteralPath $src) {
            Copy-Item -LiteralPath $src -Destination $stage -Force
        }
    }

    Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $out -Force
}
finally {
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
}

$file = Get-Item -LiteralPath $out
Write-Host ("Created {0} ({1} bytes)" -f $file.FullName, $file.Length)
