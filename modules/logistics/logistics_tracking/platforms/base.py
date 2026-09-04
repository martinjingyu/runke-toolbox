"""所有平台 client 共用的返回结构。"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class RouteResult:
    waybill: str
    found: bool = False
    last_route: str | None = None
    error: str | None = None
    raw_events: list[dict] = field(default_factory=list)
