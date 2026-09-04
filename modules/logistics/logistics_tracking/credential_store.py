"""货代平台的登录账号密码——只存在本机（QSettings：Windows 是注册表、Mac 是本地 plist/ini），
不走 core/storage 那一套给"业务数据表格"用的、以后可能对接 NAS 共享给全组的存储层，也不提交
进 git。这是纯本机机密，道理跟 config.local.yaml 存 Amazon SP-API 密钥一样——只是这里账号密码
是用户在界面"账号管理"里直接输入的，不是改配置文件。

跟 shipment_plan_apply 记"上次选过的文件路径"用的是同一个 QSettings 机制，但这里单独包一个类，
因为要支持"一个平台配好几个账号，按顺序试"（见 platforms/registry.py 的 _lookup_multi_account，
比如海德嘉新旧账号是两个独立租户），还要支持增删改查、调整优先顺序。

存储格式：每个平台一个 QSettings key，值是这个平台的账号列表序列化成的 JSON 文本
（[[账号, 密码], ...]，顺序就是优先尝试顺序）——用 JSON 文本而不是 QSettings 原生的数组接口
（beginWriteArray/endArray），因为原生数组在不同平台的行为细节不完全一致，JSON 文本更省心，
出问题也好排查（QSettings 的图形化查看工具能直接看到这段文本）。
"""
from __future__ import annotations

import json
from dataclasses import dataclass

from PySide6.QtCore import QSettings

from .platforms.registry import PLATFORM_LABELS, Account

_SETTINGS_KEY_PREFIX = "logistics_tracking/credentials/"


@dataclass
class StoredAccount:
    username: str
    password: str

    def as_tuple(self) -> Account:
        return (self.username, self.password)


class CredentialStore:
    def __init__(self, settings: QSettings | None = None):
        # settings 参数留给测试注入一个指向临时文件的 QSettings（IniFormat），不用碰用户
        # 真实的本机存储；正常使用不传，用默认的 QSettings()（读 core/app.py 设置好的
        # organization/application name）。
        self._settings = settings or QSettings()

    def accounts_for(self, platform_code: str) -> list[StoredAccount]:
        raw = self._settings.value(_SETTINGS_KEY_PREFIX + platform_code, "")
        if not raw:
            return []
        try:
            data = json.loads(raw)
        except (TypeError, ValueError):
            return []
        return [StoredAccount(username=u, password=p) for u, p in data]

    def set_accounts(self, platform_code: str, accounts: list[StoredAccount]) -> None:
        key = _SETTINGS_KEY_PREFIX + platform_code
        if not accounts:
            self._settings.remove(key)
            return
        data = [[a.username, a.password] for a in accounts]
        self._settings.setValue(key, json.dumps(data, ensure_ascii=False))

    def accounts_by_platform(self) -> dict[str, list[Account]]:
        """给 tracking_pipeline.TrackingSheet.build_preview() 用：{平台代码: [(账号,密码), ...]}，
        只包含真的配了账号的平台。
        """
        result: dict[str, list[Account]] = {}
        for code in PLATFORM_LABELS:
            accounts = self.accounts_for(code)
            if accounts:
                result[code] = [a.as_tuple() for a in accounts]
        return result
