#!/bin/bash
set -e

cd "$(dirname "$0")"

# --- 推送到 git ---
git add -A
if git diff --cached --quiet; then
    echo "没有改动，跳过 commit"
else
    git commit -m "quick fix"
fi
git push

# --- 同步项目到 NAS ---
NAS_SMB_URL="smb://RK._smb._tcp.local/公共文件"
NAS_MOUNT_POINT="/Volumes/公共文件"
NAS_TARGET_DIR="$NAS_MOUNT_POINT/个人/黄靖禺/runke-toolbox"

if [ ! -d "$NAS_MOUNT_POINT" ]; then
    echo "NAS 共享文件夹未挂载，尝试挂载 $NAS_SMB_URL ..."
    open "$NAS_SMB_URL"
    for i in $(seq 1 15); do
        [ -d "$NAS_MOUNT_POINT" ] && break
        sleep 1
    done
fi

if [ ! -d "$NAS_MOUNT_POINT" ]; then
    echo "错误：NAS 共享文件夹挂载失败。请先用 Finder（Cmd+K）连接 $NAS_SMB_URL 后重试。" >&2
    exit 1
fi

mkdir -p "$NAS_TARGET_DIR"

echo "同步项目文件到 NAS：$NAS_TARGET_DIR"
# 只同步 git 已追踪的文件，自动排除 .git/__pycache__/.venv/data 里的真实业务数据等无关文件
git ls-files -z | rsync -a --files-from=- --from0 --delete . "$NAS_TARGET_DIR/"

echo "完成：已推送到 git 并同步到 NAS。"
