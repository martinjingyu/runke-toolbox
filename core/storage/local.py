"""本地磁盘实现。当前所有数据都走这个后端；NAS 对接前，模块不用改代码。"""
from __future__ import annotations

from pathlib import Path

from .base import StorageBackend


class LocalStorage(StorageBackend):
    def __init__(self, root: str):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)

    def _resolve(self, path: str) -> Path:
        return self.root / path

    def read_bytes(self, path: str) -> bytes:
        return self._resolve(path).read_bytes()

    def write_bytes(self, path: str, data: bytes) -> None:
        target = self._resolve(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)

    def list_files(self, path: str = "") -> list[str]:
        target = self._resolve(path)
        if not target.exists():
            return []
        return sorted(p.name for p in target.iterdir() if p.is_file())

    def exists(self, path: str) -> bool:
        return self._resolve(path).exists()
