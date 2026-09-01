"""存储后端的统一接口。本地/NAS/云端未来都实现这一套方法，模块代码不用区分。"""
from __future__ import annotations

from abc import ABC, abstractmethod


class StorageBackend(ABC):
    @abstractmethod
    def read_bytes(self, path: str) -> bytes:
        """按相对路径读取文件内容。"""

    @abstractmethod
    def write_bytes(self, path: str, data: bytes) -> None:
        """按相对路径写入文件内容，父目录不存在时自动创建。"""

    @abstractmethod
    def list_files(self, path: str = "") -> list[str]:
        """列出某个相对路径下的文件（不递归）。"""

    @abstractmethod
    def exists(self, path: str) -> bool:
        """相对路径对应的文件/目录是否存在。"""
