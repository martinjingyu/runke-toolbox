
# 全新 Windows 电脑上的一键环境搭建：找不到 Python 就自动下载安装（装给当前用户，不需要管理员权限），
# 建一个项目专属的虚拟环境（.venv），装好依赖，然后启动软件。
#
# 用法：双击 setup_windows.bat（它会调用这个脚本）。
# 以后不用再装环境了，双击 run_windows.bat 直接启动就行。
#
# 上面故意留了一个空行——文件是 UTF-8 BOM 编码（中文注释要靠这个 Windows PowerShell 5.1
# 才能正确显示，不然会乱码甚至解析出错），但实测在某些机器上，如果 BOM 后面紧跟着的第一个
# 字符就是"#"，PowerShell 会把 BOM 和"#"粘在一起识别成一个命令名去执行，报"无法将"#"项识别
# 为 cmdlet"，脚本第一行就直接跑不起来。中间空一行能避开这个问题。

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

function Test-PythonWorks {
    # 光看 python.exe 这个文件存不存在不够——测试装/卸载装环境的时候，很容易在目录里留下
    # 卸载不干净的残留（比如 python.exe 还在，但旁边的 DLL/标准库已经被删了一部分），这种
    # "看着装了、其实跑不起来"的情况必须真的跑一下才能发现，不然后面 venv/pip 会莫名其妙
    # 失败，还不容易看出来是这个原因。
    param([string]$PythonExe)
    if (-not (Test-Path $PythonExe)) { return $false }
    try {
        $out = & $PythonExe -c "print('ok')" 2>$null
        return ($LASTEXITCODE -eq 0 -and $out -eq "ok")
    } catch {
        return $false
    }
}

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
    $candidates = Get-ChildItem -Path $pythonRoot -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue
    foreach ($c in $candidates) {
        if (Test-PythonWorks -PythonExe $c.FullName) {
            return $c.FullName
        }
    }

    return $null
}

function Invoke-ExternalCommand {
    # 跑 venv / pip 这些外部命令，同时把完整输出录下来再打印——不能指望它们的输出会正常显示：
    # 在 $ErrorActionPreference = "Stop" 底下，外部程序往 stderr 写一行东西，PowerShell 有时候
    # 会当成"终止性错误"直接把命令打断、甚至把真正的报错内容盖住看不清楚。这里显式地把
    # $ErrorActionPreference 临时放宽、完整收集输出，失败的时候原样打印出来，保证报错看得全。
    param([string]$Exe, [string[]]$CmdArgs)

    $previous = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & $Exe @CmdArgs 2>&1
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }

    foreach ($line in $output) {
        Write-Host $line
    }

    return $exitCode
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
    $exitCode = Invoke-ExternalCommand -Exe $systemPython -CmdArgs @("-m", "venv", $VenvDir)
    if ($exitCode -ne 0 -or -not (Test-Path (Join-Path $VenvDir "pyvenv.cfg"))) {
        Write-Host ""
        Write-Host "建虚拟环境失败了（完整报错见上面）。如果最近用 uninstall_windows.bat 卸载/重装过"
        Write-Host "Python，可能是卸载没干净、留了点残留文件——建议再运行一次 uninstall_windows.bat"
        Write-Host "彻底清理，然后重新打开一个窗口再运行这个脚本。"
        exit 1
    }
} else {
    Write-Host ".venv 已经存在，跳过"
}

Write-Host ""
Write-Host "== 第 3 步：安装依赖包 =="
# 升级 pip 这一步是锦上添花，不是必须的——Python 3.11.9 自带的 pip 已经完全够用（装
# PySide6/PyYAML 这种纯 wheel 包不需要多新的 pip），失败了没必要卡住整个安装，提醒一下、
# 继续往下走就行。
$exitCode = Invoke-ExternalCommand -Exe $VenvPython -CmdArgs @("-m", "pip", "install", "--upgrade", "pip")
if ($exitCode -ne 0) {
    Write-Host ""
    Write-Host "升级 pip 没成功（退出码 $exitCode），不影响继续安装，跳过这一步。"
}

$exitCode = Invoke-ExternalCommand -Exe $VenvPython -CmdArgs @("-m", "pip", "install", "-r", (Join-Path $ProjectRoot "requirements.txt"))
if ($exitCode -ne 0) {
    Write-Host ""
    Write-Host "安装依赖包失败了（退出码 $exitCode，完整报错见上面），停在这一步，不往下走了。"
    exit 1
}

Write-Host ""
Write-Host "== 环境搭建完成，启动程序 =="
& $VenvPython (Join-Path $ProjectRoot "main.py")
