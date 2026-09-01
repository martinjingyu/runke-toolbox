"""NAS 后端预留接口——按需求先不做，等要接的时候在这里实现下面几个方法即可。

上层模块只通过 core.storage.get_storage() 拿 StorageBackend，不会直接 import 这个类，
所以接上 NAS 之后，modules/ 下的业务代码不需要改动。
"""
from __future__ import annotations

from typing import Any

from .base import StorageBackend


class NasStorage(StorageBackend):
    def __init__(self, nas_config: dict[str, Any]):
        self.nas_config = nas_config
        raise NotImplementedError(
            "NAS 存储后端尚未实现。接口已在 StorageBackend 定义好，"
            "对接时在本文件里实现 read_bytes / write_bytes / list_files / exists 即可。"
        )

    def read_bytes(self, path: str) -> bytes:
        raise NotImplementedError

    def write_bytes(self, path: str, data: bytes) -> None:
        raise NotImplementedError

    def list_files(self, path: str = "") -> list[str]:
        raise NotImplementedError

    def exists(self, path: str) -> bool:
        raise NotImplementedError
