"""供应商全称 -> 短代码 的映射（比如"东莞市盛鑫灯饰有限公司" -> "SX"）——采购订单文件里
供应商写的是全称，但采购汇总表「供应商名称」列、发货计划汇总表「工厂」列历史上一直填的都是
这种短代码。这份映射只存在本机（QSettings），道理跟 logistics_tracking/credential_store.py
存货代账号一样：是用户在界面「供应商映射」设置里自己维护的对照表，不是业务数据表格本身的内容。
"""
from __future__ import annotations

import json

from PySide6.QtCore import QSettings

_SETTINGS_KEY = "purchase_order_import/supplier_codes"


class SupplierCodeStore:
    def __init__(self, settings: QSettings | None = None):
        self._settings = settings or QSettings()

    def mapping(self) -> dict[str, str]:
        raw = self._settings.value(_SETTINGS_KEY, "")
        if not raw:
            return {}
        try:
            data = json.loads(raw)
        except (TypeError, ValueError):
            return {}
        if not isinstance(data, dict):
            return {}
        return {str(k): str(v) for k, v in data.items()}

    def set_mapping(self, mapping: dict[str, str]) -> None:
        if not mapping:
            self._settings.remove(_SETTINGS_KEY)
            return
        self._settings.setValue(_SETTINGS_KEY, json.dumps(mapping, ensure_ascii=False))

    def resolve(self, supplier_name: str | None) -> str | None:
        if not supplier_name:
            return None
        return self.mapping().get(supplier_name.strip())
