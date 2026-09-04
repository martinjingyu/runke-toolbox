"""凯琦(KQ, sys.kqgyl.com)运单最后路由查询。

登录：POST /v1/customerLogin，body={username,password}，返回 data.token；后续请求要带请求头
"x-token"（前端是从 localStorage["Authorization"] 读出来塞进这个头的，是这套系统自己的命名，
跟标准的 Authorization 头无关）。这条登录接口之前已经用真实账号验证过。

运单列表：POST /v1/waybill/manage/selectList——用户在"运单管理"页面（/operations/waybill-manager）
上实际调的接口，是浏览器 Network 面板里截的真实请求（不是逆向猜的）。这个接口不是按单号一个个
查，是分页拉列表，所以跟 ylyn.py/nextsls.py 一样：get_last_routes() 一次性把账号名下全部运单
分页拉全（pageSize 开到 200），本地按 waybillNumber 建字典匹配。接口本身支持一堆筛选字段
（selectField/selectMethod/trackNumberSelectField 那些），但语义没摸清楚，所以这里全部留空
不筛选、直接拉全量——账号名下运单量不大的话没问题，量大了再考虑要不要啃筛选参数省流量。

"最后路由"：内容优先用 trackInfo（有实际轨迹事件才会有值，比如"2026年9月3日 已签收"，内容
本身有时候会带日期，但格式不统一——有的只有一个日期，有的像预报信息一样一段话里带好几个
日期），没有的话（新单还没轨迹事件）退回用 nodeName（当前节点名，比如"已预报"）。时间统一用
trackOperTime（这条 trackInfo/nodeName 对应的更新时间，实测这个字段可靠、trackTime 反而
一直是 null 没用上），不依赖 trackInfo 文本里可能有可能没有、格式还不统一的日期，保证每条都
带上一个干净、跟其它货代格式一致的时间前缀。
"""
from __future__ import annotations

import requests

from .base import RouteResult

BASE_URL = "https://sys.kqgyl.com"
_PAGE_SIZE = 200

# 浏览器 Network 面板里截到的真实请求体，字段名和默认值原样照抄——大部分是筛选条件，这里
# 全部留空/None 表示不筛选，只有 currentPage/pageSize 会被 _fetch_all_waybills() 覆盖。
_LIST_PAYLOAD_DEFAULTS: dict = {
    "shipStartDate": None,
    "shipEndDate": None,
    "deliveryTimeStartDate": "",
    "deliveryTimeEndDate": "",
    "allotStatus": None,
    "channelIds": [],
    "deliveryId": [],
    "deliveryType": [],
    "postcodeField": "",
    "postcodeStr": "",
    "selectBookingField": "",
    "selectBookingFieldStr": "",
    "selectBookingMethod": "",
    "selectField": "",
    "selectFieldStr": "",
    "selectMethod": "",
    "selectPostcodeMethod": "",
    "status": None,
    "trackNumberSelectField": "",
    "trackNumberSelectFieldStr": "",
    "trackNumberSelectMethod": "",
    "trackOperatorIds": [],
    "trackTimeEnd": "",
    "trackTimeStart": "",
}


class KqgylClient:
    def __init__(self, username: str, password: str, base_url: str = BASE_URL, session=None):
        self.base_url = base_url
        self.session = session or requests.Session()
        self.token = self._login(username, password)

    def _login(self, username: str, password: str) -> str:
        r = self.session.post(
            f"{self.base_url}/v1/customerLogin",
            json={"username": username, "password": password},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        if data.get("code") != 200:
            raise RuntimeError(f"登录失败: {data.get('msg')}")
        return data["data"]["token"]

    def _headers(self) -> dict:
        return {"x-token": self.token}

    def _fetch_all_waybills(self) -> list[dict]:
        rows: list[dict] = []
        page = 1
        while True:
            payload = {**_LIST_PAYLOAD_DEFAULTS, "currentPage": page, "pageSize": _PAGE_SIZE}
            r = self.session.post(
                f"{self.base_url}/v1/waybill/manage/selectList",
                json=payload,
                headers=self._headers(),
                timeout=30,
            )
            r.raise_for_status()
            data = r.json()
            if data.get("code") != 200:
                raise RuntimeError(f"查询运单列表失败: {data.get('msg')}")
            page_data = data.get("data") or {}
            batch = page_data.get("data") or []
            rows.extend(batch)
            total = page_data.get("totalCount", 0)
            if not batch or len(rows) >= total:
                break
            page += 1
        return rows

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        by_waybill = {
            row["waybillNumber"]: row for row in self._fetch_all_waybills() if row.get("waybillNumber")
        }

        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            row = by_waybill.get(wb)
            if row is None:
                results[wb] = RouteResult(waybill=wb, error="未找到该运单")
                continue

            content = row.get("trackInfo") or row.get("nodeName")
            oper_time = row.get("trackOperTime")
            last_route = f"{oper_time} {content}".strip() if oper_time else content

            if not last_route:
                results[wb] = RouteResult(waybill=wb, error="暂无路由信息")
                continue
            results[wb] = RouteResult(waybill=wb, found=True, last_route=last_route, raw_events=[row])
        return results
