param(
    [switch]$Build,
    [switch]$Test,
    [switch]$Prerelease,
    [switch]$PrintNotes,
    [string]$NotesPath,
    [string]$Target,
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

function New-Sha256File {
    param(
        [string]$Path,
        [string]$AssetName
    )
    $shaPath = "$Path.sha256"
    $sha = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    "$sha  $AssetName" | Set-Content -LiteralPath $shaPath -Encoding ASCII
    return @{ Path = $shaPath; Hash = $sha }
}

function New-ReleaseNotes {
    param(
        [string]$Tag,
        [string]$NotesPath
    )
    $date = Get-Date -Format "yyyy-MM-dd"
    $notes = @"
ExcelMergeFork $Tag

发布日期：$date

下载建议：
- 推荐普通用户下载 ExcelMergeFork-package.zip：包含单文件版、README 和一键安装/移除 Fork 注入脚本。
- 只想手动配置的用户也可以单独下载 ExcelMergeFork.exe：单文件版，无需安装 Python。
- 已有 Python 环境、想要更小体积的用户下载 ExcelMergeFork-lite-package.zip 或 ExcelMergeFork-lite.exe：不内置 Python，需 Python 3.7+ 和 openpyxl。

两个版本：
- ExcelMergeFork.exe：完整独立版，Fork 里直接配置此 exe 路径。
- ExcelMergeFork-lite.exe：轻量启动器 exe，运行时检查 Python 环境；缺少环境会弹窗提示安装方法。

Fork 配置：
- 最方便：关闭 Fork 后双击 install_fork_integration.bat；移除时双击 uninstall_fork_integration.bat。
- 单独下载 exe 的用户可双击 exe 打开设置中心，点击「Fork 一键注入」里的安装/移除。
- Merge Tool Path / Diff Tool Path 填所下载 exe 的完整路径。
- Merge Arguments：`$LOCAL,`$BASE,`$REMOTE,`$MERGED
- Diff Arguments："`$REMOTE" "`$LOCAL"

更新说明：
- 完整版会更新为新的 ExcelMergeFork.exe。
- 轻量版会更新为新的 ExcelMergeFork-lite.exe，不会自动切换成完整版。

校验文件：
- ExcelMergeFork.exe.sha256
- ExcelMergeFork-lite.exe.sha256

验证：
- 已运行 run_merge_mode_tests.bat

上传资产：
- ExcelMergeFork-package.zip
- ExcelMergeFork-lite-package.zip
- ExcelMergeFork.exe
- ExcelMergeFork.exe.sha256
- ExcelMergeFork-lite.exe
- ExcelMergeFork-lite.exe.sha256
"@
    if ($NotesPath) {
        $extraPath = Resolve-Path -LiteralPath $NotesPath -ErrorAction Stop
        $extra = (Get-Content -LiteralPath $extraPath -Raw -Encoding UTF8).Trim()
        if ($extra) {
            $notes += "`n`n本次变更：`n$extra"
        }
    }
    return $notes
}

function Get-DefaultBranch {
    param([string]$Repo)
    try {
        $name = (& gh repo view $Repo --json defaultBranchRef --jq ".defaultBranchRef.name" 2>$null).Trim()
        if ($LASTEXITCODE -eq 0 -and $name) {
            return $name
        }
    } catch {
    }
    return $null
}

function Test-RemoteBranch {
    param(
        [string]$Repo,
        [string]$Branch
    )
    if (-not $Branch) {
        return $false
    }
    try {
        $encodedBranch = [System.Uri]::EscapeDataString($Branch)
        $name = (& gh api "repos/$Repo/branches/$encodedBranch" --jq ".name" 2>$null).Trim()
        return ($LASTEXITCODE -eq 0 -and $name)
    } catch {
        return $false
    }
}

function Resolve-ReleaseTarget {
    param(
        [string]$Repo,
        [string]$Target
    )
    if ($Target) {
        return $Target
    }

    $currentBranch = ""
    try {
        $currentBranch = (& git branch --show-current 2>$null).Trim()
    } catch {
        $currentBranch = ""
    }
    if ($currentBranch -and (Test-RemoteBranch -Repo $Repo -Branch $currentBranch)) {
        return $currentBranch
    }

    $defaultBranch = Get-DefaultBranch -Repo $Repo
    if ($defaultBranch) {
        if ($currentBranch) {
            Write-Host "当前分支 '$currentBranch' 在 GitHub 仓库中不可用，Release target 将使用默认分支 '$defaultBranch'。"
        }
        return $defaultBranch
    }

    return $null
}

$version = Read-AppVersion
$tag = "v$version"
$notes = New-ReleaseNotes -Tag $tag -NotesPath $NotesPath

if ($PrintNotes) {
    Write-Host "========== Release Notes Preview =========="
    Write-Host $notes
    exit 0
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "未找到 GitHub CLI: gh。请先安装并执行 gh auth login。"
}

if ($Build) {
    Write-Host "Building full and lite exe..."
    & ".\package.bat" --gui --dist --no-pause
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

$exe = Join-Path $root "ExcelMergeFork.exe"
$liteExe = Join-Path $root "ExcelMergeFork-lite.exe"

if (-not (Test-Path -LiteralPath $exe)) {
    throw "未找到产物: $exe。请先运行 .\package.bat 或使用 -Build。"
}
if (-not (Test-Path -LiteralPath $liteExe)) {
    throw "未找到产物: $liteExe。请先运行 .\package.bat 或使用 -Build。"
}

$fullSha = New-Sha256File -Path $exe -AssetName "ExcelMergeFork.exe"
$liteSha = New-Sha256File -Path $liteExe -AssetName "ExcelMergeFork-lite.exe"

& powershell -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\make_distribution_zip.ps1" -Kind full
if ($LASTEXITCODE -ne 0) {
    throw "完整分发包生成失败"
}
& powershell -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\make_distribution_zip.ps1" -Kind lite
if ($LASTEXITCODE -ne 0) {
    throw "轻量分发包生成失败"
}
$fullPackage = Join-Path $root "dist\ExcelMergeFork-package.zip"
$litePackage = Join-Path $root "dist\ExcelMergeFork-lite-package.zip"

Write-Host "Version: $version"
Write-Host "Tag: $tag"
Write-Host "ExcelMergeFork.exe SHA256: $($fullSha.Hash)"
Write-Host "ExcelMergeFork-lite.exe SHA256: $($liteSha.Hash)"

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

$target = Resolve-ReleaseTarget -Repo $Repo -Target $Target
if ($target) {
    Write-Host "Release target: $target"
} else {
    Write-Host "Release target: 未指定，交给 GitHub 使用默认分支"
}

$ghArgs = @(
    "release", "create", $tag,
    $fullPackage,
    $litePackage,
    $exe,
    $fullSha.Path,
    $liteExe,
    $liteSha.Path,
    "--repo", $Repo,
    "--title", "ExcelMergeFork $tag",
    "--notes", $notes
)
if ($target) {
    $ghArgs += @("--target", $target)
}
if ($Prerelease) {
    $ghArgs += "--prerelease"
}

& gh @ghArgs
if ($LASTEXITCODE -ne 0) {
    throw "GitHub Release 发布失败"
}

Write-Host "Release published: https://github.com/$Repo/releases/tag/$tag"
