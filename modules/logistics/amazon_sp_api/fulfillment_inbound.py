"""FBA 入库货件（Fulfillment Inbound）相关的数据获取——目前只有"按更新时间查货件列表"这一个
用途：给之后要做的"FBA 货件追踪"之类的工具提供 ShipmentId（货件编号）和 AmazonReferenceId
（亚马逊内部参考编号）。

分页：真实账号一次查询超过 50 条时确认会被截断（第一次实现漏掉了这个，是凭记忆猜"这个接口不
支持分页"，后来查官方文档 https://developer-docs.amazon/sp-api/reference/getshipments
证实是错的——response payload 里有 NextToken 字段，QueryType 还有第三种取值 NEXT_TOKEN
专门用来翻页）。这里按官方约定实现：第一页用 DATE_RANGE 查，后续页只带 MarketplaceId +
NextToken（其它过滤条件已经编码在 NextToken 里，不能重复传）。
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass

from .client import SPApiClient

# 翻页安全上限，防止 NextToken 因为亚马逊那边的 bug 或者代码逻辑问题一直不清空，陷入死循环。
# 一页最多见过 50 条，200 页对应 1 万条货件，正常用户不可能一次查询范围内有这么多，触发这个
# 上限基本可以断定是哪里错了，不是真的还有更多数据。
_MAX_PAGES = 200

# DATE_RANGE 查询模式除了日期范围，还要求必须额外指定 ShipmentStatusList 或 ShipmentIdList
# 之一（同样是真实调用踩出来的，不是文档写清楚的）——这里列出货件的全部状态（官方文档给出的
# 11 个合法值），默认覆盖整个生命周期，不遗漏任何状态的货件。
ALL_SHIPMENT_STATUSES = [
    "WORKING",
    "READY_TO_SHIP",
    "SHIPPED",
    "IN_TRANSIT",
    "DELIVERED",
    "CHECKED_IN",
    "RECEIVING",
    "CLOSED",
    "CANCELLED",
    "DELETED",
    "ERROR",
]


@dataclass
class InboundShipment:
    shipment_id: str  # FBA 货件编号，形如 FBA15G4XXXXX
    amazon_reference_id: str | None  # 亚马逊内部参考编号
    shipment_name: str | None
    status: str | None


def list_inbound_shipments(client: SPApiClient, days_back: int = 90) -> list[InboundShipment]:
    """查最近 days_back 天内有更新的 FBA 入库货件。

    DATE_RANGE 查询模式下 LastUpdatedAfter 和 LastUpdatedBefore 必须同时传，只传前者会被
    亚马逊拒绝（HTTP 400 "LastUpdatedBefore value is empty"，真实调用踩出来的，不是文档
    写清楚的）——后者固定传"现在"就行。
    """
    now = datetime.datetime.now(datetime.timezone.utc)
    last_updated_after = (now - datetime.timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    last_updated_before = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    params: dict[str, str] = {
        "QueryType": "DATE_RANGE",
        "MarketplaceId": client.marketplace_id,
        "LastUpdatedAfter": last_updated_after,
        "LastUpdatedBefore": last_updated_before,
        "ShipmentStatusList": ",".join(ALL_SHIPMENT_STATUSES),
    }

    raw_shipments: list[dict] = []
    for _ in range(_MAX_PAGES):
        payload = client.get("/fba/inbound/v0/shipments", params=params)
        page = payload.get("payload", {}) if isinstance(payload.get("payload"), dict) else {}
        raw_shipments.extend(page.get("ShipmentData", []))

        next_token = page.get("NextToken")
        if not next_token:
            break
        # 翻页请求只带 MarketplaceId + NextToken，其它过滤条件已经编码在 token 里，
        # 重复传会被拒绝。
        params = {
            "QueryType": "NEXT_TOKEN",
            "MarketplaceId": client.marketplace_id,
            "NextToken": next_token,
        }
    else:
        raise RuntimeError(
            f"翻页超过 {_MAX_PAGES} 页仍未结束，可能是 NextToken 没有正常清空——先别信任这批数据，排查一下。"
        )

    return [
        InboundShipment(
            shipment_id=raw["ShipmentId"],
            amazon_reference_id=raw.get("AmazonReferenceId"),
            shipment_name=raw.get("ShipmentName"),
            status=raw.get("ShipmentStatus"),
        )
        for raw in raw_shipments
    ]
