# 把 setup_windows.ps1（以及软件运行时按需装的 Tesseract）装的东西清掉，方便重新测试
# "全新 Windows 电脑"那一套安装流程，不用真的找一台新电脑。
#
# 用法：双击 uninstall_windows.bat。
#
# 会做的事：
#   1. 删除项目自己的虚拟环境 (.venv)
#   2. 卸载本机装的 Python（如果是 setup_windows.ps1 装的那个 per-user 版本）
#   3. 卸载本机装的 Tesseract OCR（如果点开过"发货数量核对"、装过的话）
#   4. 从你账户的 PATH 里去掉上面这些东西装的时候加进去的目录
#
# 不会动的东西：项目代码本身、config.local.yaml、data/ 目录——这些跟"环境有没有装"无关，
# 删了没意义。

$ErrorActionPreference = "Continue"  # 清理脚本，某一步失败不该让后面几步也不做了
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Write-Host "这个脚本会：删除 .venv、卸载本机的 Python 和 Tesseract、清理 PATH。"
Write-Host "只清理这个项目相关的东西，不会动系统里其它软件。"
Write-Host ""
$confirm = Read-Host "确定要继续吗？输入 y 确认"
if ($confirm -ne "y") {
    Write-Host "取消了，什么都没动。"
    exit 0
}

function Uninstall-ByDisplayNamePattern {
    param([string]$Pattern, [string]$ExtraArgs = "")

    $found = $false
    foreach ($root in @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                         "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                         "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")) {
        $keys = Get-ChildItem $root -ErrorAction SilentlyContinue
        foreach ($key in $keys) {
            $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
            if ($props.DisplayName -like $Pattern) {
                Write-Host "找到已安装的：$($props.DisplayName)，正在卸载..."
                $uninstallString = $props.QuietUninstallString
                if (-not $uninstallString) { $uninstallString = $props.UninstallString }
                if ($uninstallString) {
                    try {
                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "$uninstallString $ExtraArgs" -Wait
                        $found = $true
                    } catch {
                        Write-Host "卸载 $($props.DisplayName) 的时候出错：$($_.Exception.Message)"
                    }
                }
            }
        }
    }
    return $found
}

Write-Host ""
Write-Host "== 第 1 步：删除 .venv =="
try {
    $VenvDir = Join-Path $ProjectRoot ".venv"
    if (Test-Path $VenvDir) {
        Remove-Item -Recurse -Force $VenvDir
        Write-Host "已删除：$VenvDir"
    } else {
        Write-Host "本来就没有，跳过"
    }
} catch {
    Write-Host "删除 .venv 出错：$($_.Exception.Message)"
}

Write-Host ""
Write-Host "== 第 2 步：卸载 Python =="
# Python 的卸载可能会跳出系统自己的确认/进度窗口，跳出来正常点掉就行
if (-not (Uninstall-ByDisplayNamePattern -Pattern "Python 3.1*")) {
    Write-Host "没找到本机装过 Python（或者不是 setup_windows.ps1 那种 per-user 安装方式），跳过"
}

Write-Host ""
Write-Host "== 第 3 步：卸载 Tesseract OCR =="
if (-not (Uninstall-ByDisplayNamePattern -Pattern "Tesseract-OCR*" -ExtraArgs "/S")) {
    Write-Host "没找到本机装过 Tesseract，跳过"
}

Write-Host ""
Write-Host "== 第 4 步：清理 PATH =="
try {
    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    if ($userPath) {
        $parts = $userPath -split ";" | Where-Object {
            $_ -and ($_ -notlike "*\Programs\Python\*") -and ($_ -notlike "*Tesseract-OCR*")
        }
        [System.Environment]::SetEnvironmentVariable("Path", ($parts -join ";"), "User")
        Write-Host "已经把 Python / Tesseract 相关的目录从 PATH 里去掉"
    }
} catch {
    Write-Host "清理 PATH 出错：$($_.Exception.Message)"
}

Write-Host ""
Write-Host "== 清理完成 =="
Write-Host "现在可以重新打开一个命令行窗口（保证 PATH 是刷新过的），再双击 setup_windows.bat 测试全新安装流程。"
