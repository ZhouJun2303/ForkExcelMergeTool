param(
    [string]$OutputPath = "dist\ExcelMergeFork-python.zip"
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path ".").Path
$out = Join-Path $root $OutputPath
$outDir = Split-Path -Parent $out
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
if (Test-Path -LiteralPath $out) {
    Remove-Item -LiteralPath $out -Force
}

$stage = Join-Path $env:TEMP ("ExcelMergeFork-python-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $stage | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage "Scripts") | Out-Null

$topLevel = @(
    "MergeExcelFork.py",
    "ExcelMergeFork-python.cmd",
    "requirements.txt",
    "README.md"
)

foreach ($item in $topLevel) {
    $src = Join-Path $root $item
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination $stage -Force
    }
}

$scriptNames = @(
    "backup_util.py",
    "compare_core.py",
    "config.py",
    "conflict.py",
    "diff_gui.py",
    "ExcelMergeGUI.py",
    "excel_io.py",
    "git_util.py",
    "git_merge_driver.py",
    "gui_common.py",
    "log_util.py",
    "MergeExcelFork.py",
    "merge_core.py",
    "merge_gui.py",
    "preview_core.py",
    "update_manager.py",
    "version.py"
)

foreach ($name in $scriptNames) {
    $src = Join-Path (Join-Path $root "Scripts") $name
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination (Join-Path $stage "Scripts") -Force
    }
}

$assetSrc = Join-Path $root "Assets"
if (Test-Path -LiteralPath $assetSrc) {
    Copy-Item -LiteralPath $assetSrc -Destination (Join-Path $stage "Assets") -Recurse -Force
}

try {
    Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $out -Force
}
finally {
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
}

$file = Get-Item -LiteralPath $out
Write-Host ("Created {0} ({1} bytes)" -f $file.FullName, $file.Length)
