param(
    [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"

function To-GitPath([string]$Path) {
    return ($Path -replace '\\', '/')
}

function Get-AttributesFile {
    $configured = (& git config --global --get core.attributesFile) 2>$null
    if ($LASTEXITCODE -eq 0 -and $configured) {
        return [Environment]::ExpandEnvironmentVariables($configured.Trim())
    }
    return Join-Path $env:USERPROFILE ".config\git\attributes"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
if (-not $ExePath) {
    $ExePath = Join-Path $root "ExcelMergeFork.exe"
}
$resolvedExe = (Resolve-Path -LiteralPath $ExePath -ErrorAction Stop).Path
if (-not (Test-Path -LiteralPath $resolvedExe -PathType Leaf)) {
    throw "ExcelMergeFork.exe 不存在: $resolvedExe"
}

$gitExe = To-GitPath $resolvedExe
$driver = '"' + $gitExe + '" --git-merge-driver "%O" "%A" "%B" "%P"'

Write-Host "Configuring global Git merge driver..."
& git config --global merge.excelmergefork.name "ExcelMergeFork workbook merge driver"
& git config --global merge.excelmergefork.driver $driver
& git config --global merge.excelmergefork.recursive binary

$attributesFile = Get-AttributesFile
$attributesDir = Split-Path -Parent $attributesFile
if ($attributesDir -and -not (Test-Path -LiteralPath $attributesDir)) {
    New-Item -ItemType Directory -Path $attributesDir | Out-Null
}
if (-not (Test-Path -LiteralPath $attributesFile)) {
    New-Item -ItemType File -Path $attributesFile | Out-Null
}

$content = Get-Content -LiteralPath $attributesFile -ErrorAction SilentlyContinue
$managed = @(
    "# ExcelMergeFork managed entry",
    "*.xlsx merge=excelmergefork",
    "*.XLSX merge=excelmergefork"
)
$changed = $false
foreach ($line in $managed) {
    if ($content -notcontains $line) {
        Add-Content -LiteralPath $attributesFile -Value $line
        $changed = $true
    }
}

Write-Host "Driver:"
& git config --global --get merge.excelmergefork.driver
Write-Host "Attributes file: $attributesFile"
if ($changed) {
    Write-Host "Added ExcelMergeFork attributes entries."
} else {
    Write-Host "Attributes entries already exist."
}
Write-Host "Done. New Git merges can use excelmergefork for .xlsx files."
