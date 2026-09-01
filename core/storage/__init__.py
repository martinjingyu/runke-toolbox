"""统一的数据读写入口：模块只认这个接口，不用关心数据实际存在本地还是 NAS。"""
from __future__ import annotations

from typing import Any

from .base import StorageBackend
from .local import LocalStorage


def get_storage(config: dict[str, Any]) -> StorageBackend:
    storage_config = config.get("storage", {})
    backend = storage_config.get("backend", "local")

    if backend == "local":
        return LocalStorage(root=storage_config.get("local_root", "data"))

    if backend == "nas":
        from .nas import NasStorage

        return NasStorage(storage_config.get("nas", {}))

    raise ValueError(f"未知的 storage.backend: {backend!r}")


__all__ = ["StorageBackend", "LocalStorage", "get_storage"]
