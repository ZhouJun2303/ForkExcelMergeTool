$ErrorActionPreference = "Stop"

function Get-AttributesFile {
    $configured = (& git config --global --get core.attributesFile) 2>$null
    if ($LASTEXITCODE -eq 0 -and $configured) {
        return [Environment]::ExpandEnvironmentVariables($configured.Trim())
    }
    return Join-Path $env:USERPROFILE ".config\git\attributes"
}

Write-Host "Removing ExcelMergeFork global Git merge driver..."
& git config --global --unset-all merge.excelmergefork.name 2>$null
& git config --global --unset-all merge.excelmergefork.driver 2>$null
& git config --global --unset-all merge.excelmergefork.recursive 2>$null

$attributesFile = Get-AttributesFile
if (Test-Path -LiteralPath $attributesFile) {
    $remove = @(
        "# ExcelMergeFork managed entry",
        "*.[xX][lL][sS] merge=excelmergefork",
        "*.[xX][lL][sS][xX] merge=excelmergefork",
        "*.[xX][lL][sS][mM] merge=excelmergefork",
        "*.[xX][lL][sS][bB] merge=excelmergefork",
        "*.[xX][lL][tT] merge=excelmergefork",
        "*.[xX][lL][tT][xX] merge=excelmergefork",
        "*.[xX][lL][tT][mM] merge=excelmergefork",
        "*.[xX][lL][aA] merge=excelmergefork",
        "*.[xX][lL][aA][mM] merge=excelmergefork",
        "*.[xX][lL][wW] merge=excelmergefork",
        "*.xlsx merge=excelmergefork",
        "*.XLSX merge=excelmergefork"
    )
    $content = Get-Content -LiteralPath $attributesFile
    $kept = @()
    foreach ($line in $content) {
        if ($remove -notcontains $line) {
            $kept += $line
        }
    }
    Set-Content -LiteralPath $attributesFile -Value $kept -Encoding UTF8
    Write-Host "Removed managed entries from: $attributesFile"
}

Write-Host "Done. Existing Fork configuration was not changed."
