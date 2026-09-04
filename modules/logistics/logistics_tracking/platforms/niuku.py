"""纽酷(usniuku.com)运单最后路由查询。

跟其它几个平台不一样，纽酷提供了正式的开放API文档(用户直接给的PDF)，不用逆向：
    1. POST /portal/api/1.0/openApi/login          登录，body={account, password}，
       返回 data.token，后续请求放在请求头 "token" 里，24小时有效
    2. GET  /portal/api/1.0/openApi/findLogisticsTrack?clientNo=<运单号>   查这个运单的完整轨迹

findLogisticsTrack 返回的是这个运单从下单到签收的完整事件列表，每条 {date, content, ...}。
实测发现顺序是按时间**正序**(从最早到最新)，跟接口文档里给的示例(两条记录、看起来是倒序)刚好
相反——不确定文档示例是不是巧合，所以这里不依赖任何顺序假设，直接按 date 字段解析比大小，
取时间最晚的一条当"最后路由"，这样不管服务端实际返回顺序是什么都不会取错。

查不到的运单号，接口返回 {"code":"order_not_exist","message":"订单不存在","success":false}
(不是 HTTP 错误状态码，200 也会带这个 body)。

跟纽酷不一样的是这个接口是逐个运单号查的(没有批量导出)，get_last_routes() 因此按运单号数量
线性发请求——真的要查的运单号很多的话会比较慢，但目前实测账号名下的量还不大，暂时不做额外的
并发优化，等以后量级明显增长了再看要不要在这一个 client 内部也拆线程。
"""
from __future__ import annotations

from datetime import datetime

import requests

from .base import RouteResult

BASE_URL = "https://api.usniuku.com"


class NiukuClient:
    def __init__(self, account: str, password: str, base_url: str = BASE_URL, session=None):
        # session 参数留给测试用假会话注入，不联网也能测——正常使用不用传，自己建一个。
        self.base_url = base_url
        self.session = session or requests.Session()
        self.token = self._login(account, password)

    def _login(self, account: str, password: str) -> str:
        r = self.session.post(
            f"{self.base_url}/portal/api/1.0/openApi/login",
            json={"account": account, "password": password},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        if data.get("code") != "200":
            raise RuntimeError(f"登录失败: {data.get('message')}")
        return data["data"]["token"]

    def _find_logistics_track(self, waybill: str) -> list[dict]:
        r = self.session.get(
            f"{self.base_url}/portal/api/1.0/openApi/findLogisticsTrack",
            params={"clientNo": waybill},
            headers={"token": self.token},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        if data.get("code") != "200":
            return []
        return data.get("data") or []

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            events = self._find_logistics_track(wb)
            if not events:
                results[wb] = RouteResult(waybill=wb, error="未找到该运单")
                continue
            latest = max(events, key=_parse_date)
            results[wb] = RouteResult(
                waybill=wb,
                found=True,
                last_route=f"{latest.get('date', '')} {latest.get('content', '')}".strip(),
                raw_events=events,
            )
        return results


def _parse_date(event: dict) -> datetime:
    try:
        return datetime.strptime(event.get("date", ""), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return datetime.min
