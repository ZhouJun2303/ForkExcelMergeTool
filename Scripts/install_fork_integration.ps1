param(
    [string]$ExePath = "",
    [string]$SettingsPath = "",
    [string]$ToolName = "ExcelMergeFork",
    [switch]$NoBackup,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

function Resolve-ToolPath {
    param([string]$RequestedPath)

    if ($RequestedPath) {
        return (Resolve-Path -LiteralPath $RequestedPath -ErrorAction Stop).Path
    }

    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $root = Split-Path -Parent $scriptDir
    $candidates = @(
        (Join-Path $root "ExcelMergeFork.exe"),
        (Join-Path $root "ExcelMergeFork-lite.exe"),
        (Join-Path $root "ExcelMergeFork-python.cmd")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        }
    }

    throw "未找到 ExcelMergeFork.exe / ExcelMergeFork-lite.exe / ExcelMergeFork-python.cmd，请用 -ExePath 指定工具路径。"
}

function Resolve-ForkSettingsPath {
    param([string]$RequestedPath)

    if ($RequestedPath) {
        return $RequestedPath
    }

    if (-not $env:LOCALAPPDATA) {
        throw "无法读取 LOCALAPPDATA 环境变量，不能定位 Fork settings.json。"
    }

    return Join-Path $env:LOCALAPPDATA "Fork\settings.json"
}

function Ensure-JsonArray {
    param(
        [pscustomobject]$Settings,
        [string]$PropertyName
    )

    if (-not ($Settings.PSObject.Properties.Name -contains $PropertyName) -or $null -eq $Settings.$PropertyName) {
        Add-Member -InputObject $Settings -MemberType NoteProperty -Name $PropertyName -Value @()
    }
}

function Set-ObjectProperty {
    param(
        [pscustomobject]$Object,
        [string]$Name,
        $Value
    )

    if ($Object.PSObject.Properties.Name -contains $Name) {
        $Object.$Name = $Value
    } else {
        Add-Member -InputObject $Object -MemberType NoteProperty -Name $Name -Value $Value
    }
}

function Upsert-Tool {
    param(
        [object[]]$Tools,
        [string]$Name,
        [string]$Path,
        [string]$Arguments,
        [bool]$SetPrimary
    )

    $result = @()
    $found = $false

    foreach ($tool in @($Tools)) {
        if ($null -eq $tool) {
            continue
        }

        if ($tool.Type -eq "Custom" -and ($tool.Name -eq $Name -or $tool.Path -eq $Path)) {
            if ($found) {
                continue
            }
            Set-ObjectProperty $tool "Type" "Custom"
            Set-ObjectProperty $tool "Name" $Name
            Set-ObjectProperty $tool "Path" $Path
            Set-ObjectProperty $tool "Arguments" $Arguments
            if ($SetPrimary) {
                Set-ObjectProperty $tool "IsPrimary" $true
            }
            $found = $true
        } elseif ($SetPrimary -and ($tool.PSObject.Properties.Name -contains "IsPrimary")) {
            $tool.IsPrimary = $false
        }

        $result += $tool
    }

    if (-not $found) {
        $newTool = [ordered]@{
            Type = "Custom"
            Name = $Name
            Path = $Path
            Arguments = $Arguments
        }
        if ($SetPrimary) {
            $newTool.IsPrimary = $true
        }
        $result += [pscustomobject]$newTool
    }

    return $result
}

function Test-ManagedTool {
    param(
        [pscustomobject]$Tool,
        [string]$Name,
        [string]$Path
    )

    if ($null -eq $Tool) {
        return $false
    }
    return (
        $Tool.Type -eq "Custom" -and
        ($Tool.Name -eq $Name -or $Tool.Path -eq $Path)
    )
}

function Test-SelectedTool {
    param(
        [pscustomobject]$Tool,
        [string]$Path,
        [string]$Arguments
    )

    if ($null -eq $Tool) {
        return $false
    }
    return (
        $Tool.Type -eq "Custom" -and
        $Tool.ApplicationPath -eq $Path -and
        $Tool.Arguments -eq $Arguments
    )
}

$resolvedToolPath = Resolve-ToolPath $ExePath
if (-not (Test-Path -LiteralPath $resolvedToolPath -PathType Leaf)) {
    throw "工具文件不存在: $resolvedToolPath"
}

$resolvedSettingsPath = Resolve-ForkSettingsPath $SettingsPath
if (-not (Test-Path -LiteralPath $resolvedSettingsPath -PathType Leaf)) {
    throw "未找到 Fork 配置文件: $resolvedSettingsPath。请先启动并退出一次 Fork，或用 -SettingsPath 指定 settings.json。"
}

if ($Uninstall) {
    Write-Host "Removing ExcelMergeFork from Fork..."
} else {
    Write-Host "Injecting ExcelMergeFork into Fork..."
}
Write-Host "Tool: $resolvedToolPath"
Write-Host "Fork settings: $resolvedSettingsPath"
Write-Host "请在执行前关闭 Fork；如果 Fork 正在运行，退出时可能会覆盖本次写入。"

if (-not $NoBackup) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = "$resolvedSettingsPath.ExcelMergeForkBackup.$timestamp"
    Copy-Item -LiteralPath $resolvedSettingsPath -Destination $backupPath -Force
    Write-Host "Backup: $backupPath"
}

$raw = Get-Content -LiteralPath $resolvedSettingsPath -Raw -Encoding UTF8
$settings = $raw | ConvertFrom-Json

Ensure-JsonArray $settings "ExternalMergeTools"
Ensure-JsonArray $settings "ExternalDiffTools"

$mergeArgs = '$LOCAL,$BASE,$REMOTE,$MERGED'
$diffArgs = '"$REMOTE" "$LOCAL"'

if ($Uninstall) {
    $settings.ExternalMergeTools = @($settings.ExternalMergeTools | Where-Object { -not (Test-ManagedTool $_ $ToolName $resolvedToolPath) })
    $settings.ExternalDiffTools = @($settings.ExternalDiffTools | Where-Object { -not (Test-ManagedTool $_ $ToolName $resolvedToolPath) })
    if (Test-SelectedTool $settings.MergeTool $resolvedToolPath $mergeArgs) {
        $settings.MergeTool = $null
    }
    if (Test-SelectedTool $settings.ExternalDiffTool $resolvedToolPath $diffArgs) {
        $settings.ExternalDiffTool = $null
    }
} else {
    $settings.ExternalMergeTools = @(
        Upsert-Tool -Tools @($settings.ExternalMergeTools) -Name $ToolName -Path $resolvedToolPath -Arguments $mergeArgs -SetPrimary $true
    )
    $settings.ExternalDiffTools = @(
        Upsert-Tool -Tools @($settings.ExternalDiffTools) -Name $ToolName -Path $resolvedToolPath -Arguments $diffArgs -SetPrimary $false
    )

    $settings.MergeTool = [pscustomobject][ordered]@{
        Type = "Custom"
        ApplicationPath = $resolvedToolPath
        Arguments = $mergeArgs
    }
    $settings.ExternalDiffTool = [pscustomobject][ordered]@{
        Type = "Custom"
        ApplicationPath = $resolvedToolPath
        Arguments = $diffArgs
    }
}

$json = $settings | ConvertTo-Json -Depth 100
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($resolvedSettingsPath, $json + [Environment]::NewLine, $utf8NoBom)

if ($Uninstall) {
    Write-Host "Done. ExcelMergeFork Fork integration was removed."
} else {
    Write-Host "Done. Fork Merge Tool / External Diff Tool now point to ExcelMergeFork."
    Write-Host "Merge arguments: $mergeArgs"
    Write-Host "Diff arguments: $diffArgs"
}
