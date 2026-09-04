"""众包(ops.zbao56.com，"AU-OPS/乐代云智能操作系统")运单跟踪查询。

这家不是查网页登录后的会话接口——网页账号密码（"账号管理"里配的那两个字段）在这家权作
appKey/appToken 用，走的是它们自己开放的 EDI API（文档在 https://ops.zbao56.com/#/docs，
Swagger 生成的）：请求头带 appKey + appToken 这对应用级密钥（要登进网页后台的"开发者中心"
才能生成/看到，不是登网页用的账号密码——这两者刚好都是"账号管理"里的两个输入框，所以还是
复用同一套 UI，只是这里存的内容含义不一样）。

GET /edi/web-services/v5/tracking?trackingRef=<运单号>：trackingRef 可以是工作号/主单号/
快递跟踪号/PO号/SO号/Shipment ID/客户入仓号，运单号列填的是哪种格式这个接口都认，不用
额外转换。没有批量接口，一次只能查一个运单号。

返回结构里 dataList 是带时间的轨迹事件列表（{time, context, node, nodeTime}），文档没保证
顺序，跟 niuku.py 一样按时间比大小取最新一条当"最后路由"；code 字段只在失败时才是错误码
（4xx），成功时是"200"，跟 HTTP 状态码语义一致，不是这条记录本身的业务状态。
"""
from __future__ import annotations

from datetime import datetime

import requests

from .base import RouteResult

BASE_URL = "https://ops.zbao56.com"


class ZhongbaoClient:
    def __init__(self, app_key: str, app_token: str, base_url: str = BASE_URL, session=None):
        self.base_url = base_url
        self.session = session or requests.Session()
        self._headers = {"appKey": app_key, "appToken": app_token}

    def _track_one(self, tracking_ref: str) -> dict:
        r = self.session.get(
            f"{self.base_url}/edi/web-services/v5/tracking",
            params={"trackingRef": tracking_ref},
            headers=self._headers,
            timeout=15,
        )
        r.raise_for_status()
        return r.json()

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            data = self._track_one(wb)
            code = str(data.get("code", ""))
            if code.startswith("4"):
                results[wb] = RouteResult(waybill=wb, error=data.get("description") or "未找到该运单")
                continue

            events = data.get("dataList") or []
            if not events:
                results[wb] = RouteResult(waybill=wb, error="暂无路由信息")
                continue
            latest = max(events, key=_parse_time)
            results[wb] = RouteResult(
                waybill=wb,
                found=True,
                last_route=f"{latest.get('time', '')} {latest.get('context', '')}".strip(),
                raw_events=events,
            )
        return results


def _parse_time(event: dict) -> datetime:
    try:
        return datetime.strptime(event.get("time", ""), "%Y-%m-%d %H:%M")
    except ValueError:
        return datetime.min
