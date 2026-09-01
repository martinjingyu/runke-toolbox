# 全新 Windows 电脑上的一键环境搭建：找不到 Python 就自动下载安装（装给当前用户，不需要管理员权限），
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
    return $null
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

    # 刚装完，当前这个 PowerShell 进程里的 PATH 还是旧的，从注册表重新读一遍
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "User") + ";" + `
                [System.Environment]::GetEnvironmentVariable("Path", "Machine")

    $systemPython = Get-PythonExe
    if ($null -eq $systemPython) {
        # 保险起见，per-user 安装的默认路径也试一下（不依赖 PATH 有没有刷新成功）
        $fallback = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"
        if (Test-Path $fallback) {
            $systemPython = $fallback
        }
    }
    if ($null -eq $systemPython) {
        Write-Host ""
        Write-Host "Python 应该是装好了，但这个窗口还没认到。请关掉这个窗口，重新打开一个，再双击一次 setup_windows.bat。"
        exit 1
    }
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
