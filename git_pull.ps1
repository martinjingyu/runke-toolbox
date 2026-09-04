$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

# --- 从服务器拉取项目 ---
$NAS_UNC = "\\RK\公共文件"
$NAS_SOURCE_DIR = Join-Path $NAS_UNC "个人\黄靖禺\runke-toolbox"
$LOCAL_DIR = Get-Location

# 不需要同步的目录/文件（虚拟环境、缓存等），保留本地已有的这些内容不被清理
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

if (-not (Test-Path $NAS_SOURCE_DIR)) {
    Write-Error "错误：无法访问共享文件夹 $NAS_SOURCE_DIR。请先在资源管理器地址栏输入 $NAS_UNC 确认可以连接后重试。"
    exit 1
}

Write-Host "从服务器拉取项目文件：$NAS_SOURCE_DIR -> $LOCAL_DIR"

$remoteFiles = Get-ChildItem -Path $NAS_SOURCE_DIR -Recurse -File
$remoteRelSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($f in $remoteFiles) {
    $rel = $f.FullName.Substring($NAS_SOURCE_DIR.Length + 1)
    [void]$remoteRelSet.Add($rel)
}

# 删除本地不再存在于服务器上的文件（排除 .git、venv、缓存等）
Get-ChildItem -Path $LOCAL_DIR -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($LOCAL_DIR.Path.Length + 1)
    if (Test-Excluded $rel) { return }
    if (-not $remoteRelSet.Contains($rel)) {
        Remove-Item -Force $_.FullName
    }
}

# 从服务器复制文件到本地
foreach ($f in $remoteFiles) {
    $rel = $f.FullName.Substring($NAS_SOURCE_DIR.Length + 1)
    $dest = Join-Path $LOCAL_DIR $rel
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    }
    Copy-Item -Path $f.FullName -Destination $dest -Force
}

# 清理同步后产生的空目录（排除虚拟环境、缓存等目录）
Get-ChildItem -Path $LOCAL_DIR -Recurse -Directory | ForEach-Object {
    $rel = $_.FullName.Substring($LOCAL_DIR.Path.Length + 1)
    if (-not (Test-Excluded $rel)) {
        [PSCustomObject]@{ Item = $_; Rel = $rel }
    }
} | Sort-Object { $_.Item.FullName.Length } -Descending | ForEach-Object {
    $dir = $_.Item
    if ((Get-ChildItem $dir.FullName -Force | Measure-Object).Count -eq 0) {
        Remove-Item -Force $dir.FullName
    }
}

Write-Host "完成：已从服务器拉取到本地。"
