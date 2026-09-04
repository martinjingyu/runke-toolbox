"""ParcelOS 白标 TMS 系统(nextsls.com)的运单最后路由查询。

海德嘉(haidej.nextsls.com)和恒信(hengxe.nextsls.com)用的是同一套白标系统(登录后台都写着
"Supported by ParcelOS")，接口路径、字段结构完全一样，只有域名/账号密码不同，所以共用这一个
client，靠 base_url 区分租户。以后遇到别的租户如果也是这套系统(域名形如 *.nextsls.com)，
直接复用这个 client 即可，不用重新逆向。

## 数据来源: 批量导出，不是逐个查询

运单列表页(/tms/wos/shipment)有个"导出"按钮，点开后选模板"运单"、点导出，会直接同步返回一个
表格(不是异步后台任务，接口是 POST /rest/tms/wos/shipment/more_export_save)。这个表格里本来就有
一列叫"最后路由"，跟平台自己的产品设计对上了，不用再自己去拼路由文字。

一开始想的是"搜索运单号拿内部id，再查详情页里的路由信息面板"，这条路子实测走不通：
1. 搜索接口 /rest/tms/wos/shipment/lists 的 keywords 如果一次塞多个运单号(逗号分隔)，服务端会
   稳定丢掉一部分匹配——不是分页/展示上限，是服务端返回的 total 本身就比真实匹配数少，同样的输入
   重复查丢的还是同一批，不是随机抖动。改成一个运单号发一次搜索请求虽然能绕开，但对这个账号下
   105 个未完成运单实测查全了以后，命中率也只有 19/105——很多单查不到不是因为号错了，是因为它在
   "另一个账号"(海德嘉新旧账号是两个独立租户)下，两个账号一起试也才勉强够用，还很慢(每个运单号
   两次请求)。
2. 改用批量导出以后，同样这 105 个运单号，两个账号各导出一次(不到 10 个请求)，本地按运单号精确
   匹配，105/105 全部命中——因为导出接口拿到的是账号名下全量数据，不受"关键词搜索"那套模糊匹配
   逻辑的影响。

所以这个 client 只做一件事：导出账号下全部运单数据，在内存里按运单号建一个 dict，查询就是本地查
这个 dict，不再对某个运单号单独发请求。多个运单号一起查跟查一个运单号，请求次数是一样的(都是
一次导出)，get_last_routes() 因此不需要限制批量大小。

## 已知限制

- 导出接口一次拿到的是这个账号下"当前能看到的全部运单"，如果运单量非常大(比如账号下有几万条)，
  这个导出可能会变慢或者被服务端加限制——目前测试的两个账号(1319/334条)响应都在几秒内，暂时没
  遇到这个问题，量级如果显著增长需要重新评估。
- "最后路由"是平台自己产品逻辑拼出来的文字(不是原始 tracking 事件日志)，格式不统一(有的是"预计
  X/X送仓"，有的是"已签收，待回传POD"，有的只是" 已签收")，这是平台本身的设计，不是这份代码的
  问题，直接透传即可。
"""
from __future__ import annotations

import requests

from .base import RouteResult

_EXPORT_TEMPLATE = "export_shipment"
_COL_WAYBILL = "运单号"
_COL_LAST_ROUTE = "最后路由"


class NextslsClient:
    def __init__(self, username: str, password: str, base_url: str, session=None):
        # session 参数留给测试用假会话注入，不联网也能测——正常使用不用传，自己建一个。
        self.base_url = base_url
        self.session = session or requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest",
        })
        self._login(username, password)

    def _login(self, username: str, password: str) -> None:
        r = self.session.post(
            f"{self.base_url}/rest/tms/wos/auth/login",
            params={"redirect_url": "/tms/wos"},
            json={"username": username, "password": password},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        if not data.get("success"):
            raise RuntimeError(f"登录失败: {data.get('info')}")

    def _export_all_shipments(self) -> dict[str, str]:
        """导出这个账号名下全部运单，返回 {运单号: 最后路由文字}。"""
        r = self.session.post(
            f"{self.base_url}/rest/tms/wos/shipment/more_export_save",
            json={
                "data_type": "all",
                "template_name": _EXPORT_TEMPLATE,
                "parentId": "",
                "change_fields": [{"template_name": ""}],
                "data": {"page": 1, "pageSize": 30, "activeTab": "all", "ids": []},
                "search": {},
            },
            timeout=60,
        )
        r.raise_for_status()
        data = r.json()
        if not data.get("success"):
            raise RuntimeError(f"导出失败: {data.get('info')}")

        table = data["data"]["data"]
        if not table:
            return {}
        header = table[0]
        try:
            idx_waybill = header.index(_COL_WAYBILL)
            idx_route = header.index(_COL_LAST_ROUTE)
        except ValueError as exc:
            raise RuntimeError(f"导出表头缺少预期的列: {exc}") from None

        result: dict[str, str] = {}
        for row in table[1:]:
            if not row:  # 导出结果里偶尔会有空行
                continue
            result[row[idx_waybill]] = row[idx_route]
        return result

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        """输入一批运单号，返回 {运单号: RouteResult}。"""
        export_map = self._export_all_shipments()
        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            if wb not in export_map:
                results[wb] = RouteResult(waybill=wb, error="未找到该运单")
                continue
            route_text = (export_map[wb] or "").strip()
            if not route_text:
                results[wb] = RouteResult(waybill=wb, found=True, error="暂无路由信息")
                continue
            results[wb] = RouteResult(waybill=wb, found=True, last_route=route_text)
        return results
