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
    "install_fork_integration.bat",
    "uninstall_fork_integration.bat",
    "install_git_integration.bat",
    "uninstall_git_integration.bat",
    "requirements.txt",
    "README.md"
)

foreach ($item in $topLevel) {
    $src = Join-Path $root $item
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination $stage -Force
    }
}

$scriptsSrc = Join-Path $root "Scripts"
Get-ChildItem -LiteralPath $scriptsSrc -Filter "*.py" -File | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $stage "Scripts") -Force
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
