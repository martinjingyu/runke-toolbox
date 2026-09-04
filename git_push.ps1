$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

# --- 同步项目到服务器 ---
$NAS_UNC = "\\RK\公共文件"
$NAS_TARGET_DIR = Join-Path $NAS_UNC "个人\黄靖禺\runke-toolbox"
$LOCAL_DIR = Get-Location

# 不需要同步的目录/文件（虚拟环境、缓存等）
$ExcludeDirNames = @('.git', '.venv', 'venv', 'env', '__pycache__', '.pytest_cache', '.mypy_cache', '.ruff_cache', 'node_modules', '.idea')
$ExcludeFilePatterns = @('*.pyc')

function Test-Excluded($relPath) {
    $parts = $relPath -split '\\'
    foreach ($p in $parts) {
        if ($ExcludeDirNames -contains $p) { return $true }
    }
    foreach ($pattern in $ExcludeFilePatterns) {
        if ($relPath -like $pattern) { return $true }
    }
    return $false
}

if (-not (Test-Path $NAS_UNC)) {
    Write-Error "错误：无法访问共享文件夹 $NAS_UNC。请先在资源管理器地址栏输入 $NAS_UNC 确认可以连接后重试。"
    exit 1
}

New-Item -ItemType Directory -Force -Path $NAS_TARGET_DIR | Out-Null

Write-Host "同步项目文件到服务器：$NAS_TARGET_DIR"

$localFiles = Get-ChildItem -Path $LOCAL_DIR -Recurse -File | Where-Object {
    $rel = $_.FullName.Substring($LOCAL_DIR.Path.Length + 1)
    -not (Test-Excluded $rel)
}
$localRelSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($f in $localFiles) {
    $rel = $f.FullName.Substring($LOCAL_DIR.Path.Length + 1)
    [void]$localRelSet.Add($rel)
}

# 删除服务器上不再存在于本地的文件
Get-ChildItem -Path $NAS_TARGET_DIR -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($NAS_TARGET_DIR.Length + 1)
    if (-not $localRelSet.Contains($rel)) {
        Remove-Item -Force $_.FullName
    }
}

# 复制本地文件到服务器
foreach ($f in $localFiles) {
    $rel = $f.FullName.Substring($LOCAL_DIR.Path.Length + 1)
    $dest = Join-Path $NAS_TARGET_DIR $rel
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    }
    Copy-Item -Path $f.FullName -Destination $dest -Force
}

# 清理同步后产生的空目录
Get-ChildItem -Path $NAS_TARGET_DIR -Recurse -Directory |
    Sort-Object { $_.FullName.Length } -Descending |
    Where-Object { (Get-ChildItem $_.FullName -Force | Measure-Object).Count -eq 0 } |
    Remove-Item -Force

Write-Host "完成：已同步到服务器。"

# --- 推送到 GitHub ---
# 工作区有没提交的改动的话，自动 add 全部改动、commit -m "quick fix"，再推送。
$gitStatus = git status --porcelain
if ($gitStatus) {
    Write-Host "工作区有未提交的改动，自动 commit……"
    git add -A
    git commit -m "quick fix"
}

Write-Host "推送到 GitHub…"
git push origin HEAD
Write-Host "完成：已推送到 GitHub。"
