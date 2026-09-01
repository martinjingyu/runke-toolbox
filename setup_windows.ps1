﻿# 全新 Windows 电脑上的一键环境搭建：找不到 Python 就自动下载安装（装给当前用户，不需要管理员权限），
# 建一个项目专属的虚拟环境（.venv），装好依赖，然后启动软件。
#
# 用法：双击 setup_windows.bat（它会调用这个脚本）。
# 以后不用再装环境了，双击 run_windows.bat 直接启动就行。

$ErrorActionPreference = "Stop"

# 让中文正常显示在控制台里（不然哪怕脚本本身没问题，输出也可能是乱码）
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# 有些老 Windows 默认没启用 TLS 1.2，访问 python.org 这类 HTTPS 地址会直接失败，先强制打开
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$PythonVersion = "3.11.9"
$PythonInstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$VenvDir = Join-Path $ProjectRoot ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

function Get-PythonExe {
    # 依次试 py launcher 和 python 命令，返回一个真的能跑起来的 python.exe 完整路径；
    # 都找不到（或者只是 Windows 商店那个假的 python.exe 占位符）就返回 $null。
    foreach ($cmd in @("py", "python")) {
        try {
            if ($cmd -eq "py") {
                $exePath = & py -3 -c "import sys; print(sys.executable)" 2>$null
            } else {
                $exePath = & python -c "import sys; print(sys.executable)" 2>$null
            }
            if ($LASTEXITCODE -eq 0 -and $exePath) {
                return $exePath.Trim()
            }
        } catch {}
    }

    # 上面全靠 PATH 找命令——但改了注册表里的 PATH 之后，已经在运行的 explorer.exe 不一定会
    # 马上感知到，双击这个脚本重新弹出来的窗口，环境变量可能还是旧的，要等重新登录/重启资源
    # 管理器才会刷新。不能让这种情况被误判成"没装 Python"又重新下载一遍，所以就算上面的 PATH
    # 检测失败，也直接去 python.org per-user 安装的固定位置探一下，装过的话大概率就在这。
    $pythonRoot = Join-Path $env:LOCALAPPDATA "Programs\Python"
    $candidate = Get-ChildItem -Path $pythonRoot -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($candidate) {
        return $candidate.FullName
    }

    return $null
}

function Add-ToUserPath {
    # 把目录持久化地加进"当前用户"的 PATH（写注册表，不是只改当前进程），这样以后新开的
    # 命令行窗口、双击 run_windows.bat 都能直接找到 python，不用每次都靠这个脚本兜底。
    #
    # 之前踩过的坑：python.org 安装器的 PrependPath=1 理论上会自动做这件事，但实测不总是
    # 生效（跟 Windows 版本、安装器版本都有关系），所以这里不依赖它，自己再显式写一遍，
    # 双重保险。
    param([string[]]$Dirs)

    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    if ($null -eq $userPath) { $userPath = "" }

    $toAdd = $Dirs | Where-Object { $_ -and ($userPath -notlike "*$_*") }
    if ($toAdd.Count -eq 0) { return }

    $newUserPath = ($userPath.TrimEnd(";") + ";" + ($toAdd -join ";")).Trim(";")
    [System.Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")

    # 同时更新当前这个 PowerShell 进程的 PATH，这次运行马上就能用，不用等重开窗口
    $env:Path = ($toAdd -join ";") + ";" + $env:Path

    Write-Host "已经把 Python 永久加到你账户的 PATH 里，以后新开的命令行窗口也能直接用 python 命令。"
}

Write-Host "== 第 1 步：检查 Python =="
$systemPython = Get-PythonExe

if ($null -eq $systemPython) {
    Write-Host "没找到可用的 Python，从官网下载安装 $PythonVersion（只装给当前用户，不用管理员权限）..."
    $installerPath = Join-Path $env:TEMP "python-$PythonVersion-amd64.exe"
    # -UseBasicParsing：全新电脑上 IE 引擎可能从没初始化过，Invoke-WebRequest 默认解析方式会报错，
    # 这里用不到网页解析，加上这个参数绕开那个依赖
    Invoke-WebRequest -Uri $PythonInstallerUrl -OutFile $installerPath -UseBasicParsing
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_test=0" -Wait
    Remove-Item $installerPath -ErrorAction SilentlyContinue

    # 不依赖 PATH 有没有刷新成功——python.org per-user 安装是固定装到这个目录，直接去这里找，
    # 比"装完再搜 PATH 里有没有"更可靠
    $installDir = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311"
    $installedPython = Join-Path $installDir "python.exe"

    if (-not (Test-Path $installedPython)) {
        # 版本号目录名以后可能变（比如升级到 3.12），保险起见搜一下
        $pythonRoot = Join-Path $env:LOCALAPPDATA "Programs\Python"
        $candidate = Get-ChildItem -Path $pythonRoot -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($candidate) {
            $installedPython = $candidate.FullName
            $installDir = $candidate.DirectoryName
        }
    }

    if (-not (Test-Path $installedPython)) {
        Write-Host ""
        Write-Host "Python 装完了但没找到 python.exe，安装本身可能失败了。请手动检查一下，或者重新运行一次这个脚本。"
        exit 1
    }

    $systemPython = $installedPython
    Add-ToUserPath -Dirs @($installDir, (Join-Path $installDir "Scripts"))

    Write-Host "Python 装好了：$systemPython"
} else {
    Write-Host "已有可用的 Python：$systemPython"
}

Write-Host ""
Write-Host "== 第 2 步：建虚拟环境（.venv）=="
if (-not (Test-Path $VenvPython)) {
    & $systemPython -m venv $VenvDir
} else {
    Write-Host ".venv 已经存在，跳过"
}

Write-Host ""
Write-Host "== 第 3 步：安装依赖包 =="
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $ProjectRoot "requirements.txt")

Write-Host ""
Write-Host "== 环境搭建完成，启动程序 =="
& $VenvPython (Join-Path $ProjectRoot "main.py")
