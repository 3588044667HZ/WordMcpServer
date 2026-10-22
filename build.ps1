<#
.SYNOPSIS
    WordMcpServer 构建脚本：还原依赖 → 构建 → 注册插件 → 启动 Word → 健康检查 → 可选回归测试。

.DESCRIPTION
    注意：构建会写入 WordMcpServer\bin\<Configuration>，Word 运行时会锁定 DLL。
    如果 Word 正在运行，请先保存文档并关闭，或使用 -StopWord 强制结束（未保存内容会丢失）。

.PARAMETER Configuration
    Debug（默认）或 Release。

.PARAMETER Clean
    构建前清理 bin/<Configuration> 与 obj/<Configuration>。

.PARAMETER NoRestore
    跳过 NuGet 包还原（packages 目录已存在时可用）。

.PARAMETER Register
    注册/启用插件（首次安装、切换 Debug/Release、或注册表被清理时使用）。

.PARAMETER StopWord
    构建前强制结束 WINWORD.EXE（会丢失未保存文档）。

.PARAMETER StartWord
    构建完成后启动 Word，并等待 MCP /health 就绪。

.PARAMETER Test
    启动 Word 后运行 tests\run_tests.ps1 端到端回归测试。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build.ps1 -StopWord -Register -StartWord

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build.ps1 -StopWord -Register -StartWord -Test
#>
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [switch]$Clean,
    [switch]$NoRestore,
    [switch]$Register,
    [switch]$StopWord,
    [switch]$StartWord,
    [switch]$Test
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$projectDir = Join-Path $root "WordMcpServer"
$solution = Join-Path $root "WordMcpServer.sln"
$addinName = "WordMcpServer"
$outDir = Join-Path $projectDir ("bin\" + $Configuration)
$vstoPath = Join-Path $outDir ($addinName + ".vsto")
$vstoUrl = "file:///" + ($vstoPath -replace "\\", "/")

function Write-Step([string]$Text) {
    Write-Host ""
    Write-Host "=== $Text ===" -ForegroundColor Cyan
}

function Write-Ok([string]$Text) { Write-Host "  [OK] $Text" -ForegroundColor Green }
function Write-Warn2([string]$Text) { Write-Host "  [WARN] $Text" -ForegroundColor Yellow }
function Write-Err([string]$Text) { Write-Host "  [ERROR] $Text" -ForegroundColor Red }

function Find-MSBuild {
    $candidates = @()

    # 优先用 vswhere 定位最新 VS 的 MSBuild
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        try {
            $found = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null |
                Select-Object -First 1
            if ($found) { $candidates += $found }
        } catch {
            # 忽略 vswhere 失败，继续尝试固定路径
        }
    }

    $candidates += @(
        "D:\VisualStudioCommunity\MSBuild\Current\Bin\MSBuild.exe",
        (Join-Path $env:ProgramFiles "Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"),
        (Join-Path $env:ProgramFiles "Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"),
        (Join-Path $env:ProgramFiles "Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Find-WordExe {
    $candidates = @()

    $appPath = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe" -ErrorAction SilentlyContinue
    if ($appPath) {
        $default = $appPath.GetValue("")
        if ($default) { $candidates += $default }
        $path = $appPath.GetValue("Path")
        if ($path) { $candidates += (Join-Path $path "WINWORD.EXE") }
    }

    $candidates += @(
        "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Get-McpPort {
    $configPath = Join-Path $outDir "mcp.config.json"
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($config.Port -and [int]$config.Port -ge 1 -and [int]$config.Port -le 65535) {
                return [int]$config.Port
            }
        } catch {
            # 配置损坏时回退默认端口
        }
    }

    return 5101
}

# ---------------------------------------------------------------------------
Write-Step "WordMcpServer 构建（$Configuration）"
Write-Host "  根目录：$root"

# 1) Word 进程检查（DLL 占用）
$wordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($wordProcesses.Count -gt 0) {
    if ($StopWord) {
        Write-Warn2 "检测到 WINWORD.EXE 正在运行，强制结束（未保存的文档会丢失）..."
        Stop-Process -Name WINWORD -Force
        Start-Sleep -Seconds 2
        Write-Ok "Word 已结束"
    }
    else {
        Write-Err "WINWORD.EXE 正在运行，插件 DLL 被占用，无法构建。"
        Write-Host "        请先保存文档并关闭 Word，或加 -StopWord 参数。" -ForegroundColor Yellow
        exit 1
    }
}

# 2) 清理
if ($Clean) {
    Write-Step "清理 $Configuration 输出"
    Remove-Item (Join-Path $projectDir ("bin\" + $Configuration)) -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $projectDir ("obj\" + $Configuration)) -Recurse -Force -ErrorAction SilentlyContinue
    Write-Ok "已清理"
}

# 3) NuGet 还原
if (-not $NoRestore) {
    Write-Step "还原 NuGet 包"
    $nuget = Join-Path $root "tools\nuget.exe"
    if (-not (Test-Path $nuget)) {
        Write-Err "找不到 $nuget"
        exit 1
    }

    $restoreOutput = & $nuget restore $solution -NonInteractive 2>&1
    if ($LASTEXITCODE -ne 0) {
        if (Test-Path (Join-Path $root "packages")) {
            Write-Warn2 "nuget restore 失败（可能无网络），使用已有 packages 目录继续"
        }
        else {
            $restoreOutput | Write-Host
            Write-Err "NuGet 还原失败，且 packages 目录不存在"
            exit 1
        }
    }
    else {
        Write-Ok "依赖已还原"
    }
}

# 4) 构建
Write-Step "构建 $Configuration"
$msbuild = Find-MSBuild
if (-not $msbuild) {
    Write-Err "找不到 MSBuild.exe，请安装 Visual Studio（含 Office/SharePoint 开发工作负载）"
    exit 1
}

Write-Host "  MSBuild：$msbuild"
& $msbuild $solution ("/p:Configuration=" + $Configuration) /v:minimal /nologo
if ($LASTEXITCODE -ne 0) {
    Write-Err "构建失败"
    exit 1
}

if (-not (Test-Path $vstoPath)) {
    Write-Err "构建成功但未找到清单：$vstoPath"
    exit 1
}

Write-Ok "构建成功：$vstoPath"

# 5) 注册插件
if ($Register) {
    Write-Step "注册插件"
    $regKey = "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\$addinName"
    $currentManifest = $null
    if (Test-Path $regKey) {
        $current = Get-ItemProperty $regKey -ErrorAction SilentlyContinue
        if ($current -and $current.Manifest) {
            $currentManifest = ($current.Manifest -replace "\|vstolocal$", "")
        }
    }

    if ($currentManifest -eq $vstoUrl) {
        Write-Ok "注册表已指向当前清单，跳过安装"
    }
    else {
        $installer = "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
        $installed = $false
        if (Test-Path $installer) {
            Write-Host "  使用 VSTOInstaller 安装清单..."
            $install = Start-Process -FilePath $installer `
                -ArgumentList "/Install", $vstoUrl, "/Silent" -Wait -PassThru -NoNewWindow
            if ($install.ExitCode -eq 0) {
                $installed = $true
                Write-Ok "VSTOInstaller 安装成功"
            }
            else {
                Write-Warn2 "VSTOInstaller 返回 $($install.ExitCode)，尝试手工注册（|vstolocal）"
            }
        }

        if (-not $installed) {
            New-Item -Path $regKey -Force | Out-Null
            Set-ItemProperty $regKey -Name "Description" -Value $addinName
            Set-ItemProperty $regKey -Name "FriendlyName" -Value $addinName
            Set-ItemProperty $regKey -Name "LoadBehavior" -Value 3 -Type DWord
            Set-ItemProperty $regKey -Name "Manifest" -Value ($vstoUrl + "|vstolocal")
            Write-Ok "已写入注册表（|vstolocal）"
        }
    }

    # 确保启用并清除 Word 的禁用记录
    New-Item -Path $regKey -Force | Out-Null
    Set-ItemProperty $regKey -Name "LoadBehavior" -Value 3 -Type DWord
    if (-not (Get-ItemProperty $regKey -Name "Manifest" -ErrorAction SilentlyContinue)) {
        Set-ItemProperty $regKey -Name "Manifest" -Value ($vstoUrl + "|vstolocal")
    }

    Remove-Item "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Resiliency\DisabledItems" -Force -ErrorAction SilentlyContinue
    Write-Ok "LoadBehavior=3，已清除 Word 禁用记录"
}

# 6) 启动 Word 并等待 /health
$port = Get-McpPort
if ($StartWord) {
    Write-Step "启动 Word"
    $wordExe = Find-WordExe
    if (-not $wordExe) {
        Write-Err "找不到 WINWORD.EXE"
        exit 1
    }

    Start-Process $wordExe
    Write-Host "  已启动：$wordExe"

    if (-not $Register) {
        Write-Warn2 "未指定 -Register：如果插件尚未注册，MCP 服务不会启动"
    }

    $deadline = (Get-Date).AddSeconds(30)
    $healthy = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        try {
            $response = Invoke-WebRequest -Uri ("http://127.0.0.1:" + $port + "/health") `
                -TimeoutSec 2 -UseBasicParsing
            if ($response.StatusCode -eq 200 -and $response.Content -eq "ok") {
                $healthy = $true
                break
            }
        }
        catch {
            # 继续等待
        }
    }

    if ($healthy) {
        Write-Ok "MCP 服务器已就绪：http://127.0.0.1:$port/mcp"
    }
    else {
        Write-Warn2 "30 秒内未检测到 /health，请查看日志：%APPDATA%\WordMcpServer\Logs\wordmcp.log"
    }
}

# 7) 端到端测试
if ($Test) {
    Write-Step "运行端到端回归测试"
    $testScript = Join-Path $root "tests\run_tests.ps1"
    if (-not (Test-Path $testScript)) {
        Write-Err "找不到 $testScript"
        exit 1
    }

    & powershell -NoProfile -ExecutionPolicy Bypass -File $testScript `
        -BaseUrl ("http://127.0.0.1:" + $port + "/mcp")
    if ($LASTEXITCODE -ne 0) {
        Write-Err "回归测试失败"
        exit 1
    }

    Write-Ok "回归测试通过"
}

# 8) 汇总
Write-Step "完成"
Write-Host "  配置：$Configuration"
Write-Host "  输出：$outDir"
Write-Host "  清单：$vstoPath"
Write-Host "  端点：http://127.0.0.1:$port/mcp"
Write-Host ""
Write-Host "提示：更新插件后必须重启 Word（VSTO 无法热加载）。" -ForegroundColor DarkGray
