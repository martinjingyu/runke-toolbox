"""物流商代码(表格"物流商"列里的缩写) -> 展示名 / 查询函数 的登记表。

跟原来那个独立脚本(logistics_tracking/platforms/registry.py)不一样：账号密码不再写死在代码里
——账号密码现在是用户在界面"账号管理"里输入、保存在本机的(见 ../credential_store.py)，运行时
由调用方（tracking_pipeline.py）传进来。这个模块只负责两件事：告诉界面"有哪些物流商代码、
中文名叫什么"（填下拉框用），以及"给定一个物流商代码 + 一批账号 + 一批运单号，怎么去查"。

同一个物流商可能配了不止一个账号(比如海德嘉新旧账号是两个独立租户，各自能查到的运单不重叠)——
_lookup_multi_account() 按用户在"账号管理"里排的顺序依次试，一个运单只要在某个账号下查到了就
不用再试下一个账号，所有账号都没查到才真正算"未找到"。

还没接入自动查询的物流商，用 _not_implemented() 占位：不会报错崩掉整条流水线，而是让这个物流商
名下的每个运单号都返回一个"平台尚未接入自动查询"的 RouteResult，这样最终写回表格的"是否有更新"
列会如实显示这个原因，而不是空着或者假装查过了。
"""
from __future__ import annotations

from typing import Callable

from .base import RouteResult
from .nextsls import NextslsClient
from .niuku import NiukuClient
from .ylyn import YlynClient
from .zhongbao import ZhongbaoClient
from .kqgyl import KqgylClient

# 物流商代码 -> 中文展示名。代码就是出货跟踪表"物流商"列里的缩写，账号管理界面的下拉框、
# 预览表格里的"平台尚未接入自动查询"提示都用这份名单，加新物流商在这里加一行就行。
#
# ZB 曾经错标成"至美通"(zipto.cn)——跟业务确认过，表格里"物流商"列写 ZB 的实际指"众包"
# (ops.zbao56.com)，两家是完全不同的公司。
PLATFORM_LABELS: dict[str, str] = {
    "HDJ": "海德嘉",
    "HX": "恒信",
    "ZB": "众包",
    "KQ": "凯琦",
    "XQH": "新企航",
    "YH": "盈和",
    "CH": "长河",
    "NK": "纽酷",
    "YLYN": "壹鹿有你",
    "ZY": "众壹",
}

# nextsls(ParcelOS 白标系统)这几个物流商，域名不一样但接口/字段完全一样，见 nextsls.py 顶部说明。
_NEXTSLS_BASE_URLS: dict[str, str] = {
    "HDJ": "http://haidej.nextsls.com",
    "HX": "http://hengxe.nextsls.com",
}

# 壹鹿有你(YLYN)那套若依风格后台，"众壹"(ZY)是同一个系统的另一个租户/品牌——域名不同、
# RSA公钥/clientId都一样，见 ylyn.py 顶部说明。
_YLYN_DOMAINS: dict[str, str] = {
    "YLYN": "yl.noms.logistics-tms.com",
    "ZY": "zy.noms.logistics-tms.com",
}

Account = tuple[str, str]  # (账号, 密码)，按优先尝试顺序排列


def _lookup_multi_account(
    client_factory: Callable[[str, str], object],
    accounts: list[Account],
    waybill_numbers: list[str],
) -> dict[str, RouteResult]:
    remaining = list(waybill_numbers)
    results: dict[str, RouteResult] = {}
    login_errors: list[str] = []

    for username, password in accounts:
        if not remaining:
            break
        try:
            client = client_factory(username, password)
        except Exception as exc:  # noqa: BLE001 - 记录下来，继续试下一个账号
            login_errors.append(f"账号[{username}]登录失败: {exc}")
            continue
        batch = client.get_last_routes(remaining)
        still_missing = []
        for wb in remaining:
            r = batch.get(wb)
            if r and r.found:
                results[wb] = r
            else:
                still_missing.append(wb)
        remaining = still_missing

    reason = "未找到该运单(已尝试全部账号)"
    if login_errors and not results:
        reason = "; ".join(login_errors)
    for wb in remaining:
        results[wb] = RouteResult(waybill=wb, error=reason)
    return results


def _not_implemented(name: str, waybill_numbers: list[str]) -> dict[str, RouteResult]:
    return {
        wb: RouteResult(waybill=wb, error=f"{name}: 平台尚未接入自动查询")
        for wb in waybill_numbers
    }


def _no_accounts(name: str, waybill_numbers: list[str]) -> dict[str, RouteResult]:
    return {
        wb: RouteResult(waybill=wb, error=f"{name}: 还没在「账号管理」里配置这个平台的账号密码")
        for wb in waybill_numbers
    }


def get_last_routes_for_carrier(
    carrier_code: str, accounts: list[Account], waybill_numbers: list[str]
) -> dict[str, RouteResult]:
    """给定物流商代码 + 这个物流商配置的账号列表(按优先顺序) + 一批运单号，返回查询结果。

    accounts 从 credential_store.py 里读出来，为空说明用户还没给这个物流商配账号——即使这个
    物流商本身已经接入了自动查询，没账号也查不了，跟"平台尚未接入自动查询"分开报不同的原因，
    方便人一眼看出是该去接入代码还是该去填账号。
    """
    label = PLATFORM_LABELS.get(carrier_code, carrier_code)

    if carrier_code in _NEXTSLS_BASE_URLS:
        if not accounts:
            return _no_accounts(label, waybill_numbers)
        base_url = _NEXTSLS_BASE_URLS[carrier_code]

        def factory(username: str, password: str) -> NextslsClient:
            return NextslsClient(username=username, password=password, base_url=base_url)

        return _lookup_multi_account(factory, accounts, waybill_numbers)

    if carrier_code == "NK":
        if not accounts:
            return _no_accounts(label, waybill_numbers)

        def factory(account: str, password: str) -> NiukuClient:
            return NiukuClient(account=account, password=password)

        return _lookup_multi_account(factory, accounts, waybill_numbers)

    if carrier_code in _YLYN_DOMAINS:
        if not accounts:
            return _no_accounts(label, waybill_numbers)
        domain = _YLYN_DOMAINS[carrier_code]

        def factory(account: str, password: str) -> YlynClient:
            return YlynClient(account=account, password=password, base_url=f"https://{domain}/prod-api", domain=domain)

        return _lookup_multi_account(factory, accounts, waybill_numbers)

    if carrier_code == "ZB":
        if not accounts:
            return _no_accounts(label, waybill_numbers)

        # 众包这家"账号管理"里存的其实是 appKey/appToken(一对应用级密钥，不是网页登录密码)，
        # 借用同一套账号密码 UI 的两个输入框存，见 zhongbao.py 顶部说明。
        def factory(app_key: str, app_token: str) -> ZhongbaoClient:
            return ZhongbaoClient(app_key=app_key, app_token=app_token)

        return _lookup_multi_account(factory, accounts, waybill_numbers)

    if carrier_code == "KQ":
        if not accounts:
            return _no_accounts(label, waybill_numbers)

        def factory(username: str, password: str) -> KqgylClient:
            return KqgylClient(username=username, password=password)

        return _lookup_multi_account(factory, accounts, waybill_numbers)

    if carrier_code in PLATFORM_LABELS:
        return _not_implemented(label, waybill_numbers)

    return {
        wb: RouteResult(waybill=wb, error=f"未知物流商代码: {carrier_code}")
        for wb in waybill_numbers
    }
