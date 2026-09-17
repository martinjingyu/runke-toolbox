"""ERP 导出、CastleGate(CG) CSV 导出里的字段，怎么对应到汇总表的"仓库"和"平台"。

这几条规则不是猜的，是业务确认过的：
- ERP「发货仓库」字段：含 IL 归 IL 表，含 TX1 归 TX1 表，含 CA1 归 CA1 表，
  剩下的（比如 TS-WH13/TS-WH25 这种）一律归 CG。
- ERP「店铺」字段：Overstock/Wayfair US/Wayfair3/Wayfair RQ 直接对应汇总表里
  OS/WF-RX/WF-TS/WF-RQ 这 4 个平台区块。
- CG 的 3 份 CSV 本来就是 CastleGate（CG 仓库）自己导出的订单，仓库固定是 CG，
  平台由业务人员在界面里为每份文件选（文件名一般是 CG-RQ/CG-RX/CG-TS 这种，
  只做默认猜测用，不强制，选错了人工可以改）。
"""
from __future__ import annotations

WAREHOUSE_TX1 = "TX1"
WAREHOUSE_IL = "IL"
WAREHOUSE_CA1 = "CA1"
WAREHOUSE_CG = "CG"

PLATFORM_OS = "OS"
PLATFORM_WF_RX = "WF-RX"
PLATFORM_WF_RQ = "WF-RQ"
PLATFORM_WF_TS = "WF-TS"

# 这 4 个是《输入说明》里明确要导入的平台；Lowe's/HD/Walmart 这次不涉及。
IMPORTABLE_PLATFORMS = (PLATFORM_OS, PLATFORM_WF_RX, PLATFORM_WF_RQ, PLATFORM_WF_TS)

ERP_SHOP_TO_PLATFORM = {
    "Overstock": PLATFORM_OS,
    "Wayfair US": PLATFORM_WF_RX,
    "Wayfair3": PLATFORM_WF_TS,
    "Wayfair RQ": PLATFORM_WF_RQ,
}


class UnknownShopError(Exception):
    pass


def classify_erp_warehouse(ship_warehouse: str) -> str:
    text = (ship_warehouse or "")
    if "IL" in text:
        return WAREHOUSE_IL
    if "TX1" in text:
        return WAREHOUSE_TX1
    if "CA1" in text:
        return WAREHOUSE_CA1
    return WAREHOUSE_CG


def classify_erp_platform(shop: str) -> str:
    platform = ERP_SHOP_TO_PLATFORM.get((shop or "").strip())
    if platform is None:
        raise UnknownShopError(
            f"ERP 数据里的「店铺」字段 {shop!r} 不在已知的平台对照表里"
            f"（已知的有：{', '.join(ERP_SHOP_TO_PLATFORM)}），需要人工确认这是哪个平台"
        )
    return platform


# CSV 文件名前缀 -> 默认猜测的平台（仅供界面预选，不强制）。
CSV_FILENAME_PLATFORM_HINTS = {
    "CG-RQ": PLATFORM_WF_RQ,
    "CG-RX": PLATFORM_WF_RX,
    "CG-TS": PLATFORM_WF_TS,
}


def guess_platform_from_filename(filename: str) -> str | None:
    upper = filename.upper()
    for prefix, platform in CSV_FILENAME_PLATFORM_HINTS.items():
        if prefix in upper:
            return platform
    return None
