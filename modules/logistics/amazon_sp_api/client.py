"""SP-API 的通用调用封装：LWA 换取 Access Token（自动缓存/续期）+ 发请求的公共逻辑。

具体某个业务接口（比如查 FBA 货件）不写在这个文件里，见同目录下 fulfillment_inbound.py
这种按业务拆开的文件——这里只负责"怎么认证、怎么发请求、错了怎么报错"这些跟具体业务无关的部分。
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

import requests

from core.config import load_config

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"

# 常用站点的 marketplace id，业务代码按需引用，不用每次去记那串编号
MARKETPLACE_ID_US = "ATVPDKIKX0DER"
MARKETPLACE_ID_CA = "A2EUQ1WTGCTBG2"


class SPApiError(RuntimeError):
    """SP-API 调用失败，带上 HTTP 状态码和原始响应内容方便排查。"""

    def __init__(self, message: str, status_code: int | None = None, body: str | None = None):
        super().__init__(message)
        self.status_code = status_code
        self.body = body


@dataclass
class SPApiConfig:
    client_id: str
    client_secret: str
    refresh_token: str
    marketplace_id: str
    base_url: str


def load_sp_api_config() -> SPApiConfig:
    """从 config.yaml / config.local.yaml 的 amazon_sp_api 一节读凭证。

    真实密钥只应该出现在 config.local.yaml（gitignore 掉的那份），config.yaml 里对应字段
    留空模板即可。这里按"涉及真实业务数据，宁可报错不要猜"的约定，缺字段直接抛错，
    不悄悄用空字符串继续跑（那样错误会在真正发请求时才炸，报错信息反而不清楚）。
    """
    raw = load_config().get("amazon_sp_api") or {}

    missing = [
        key
        for key in ("client_id", "client_secret", "refresh_token", "marketplace_id", "base_url")
        if not raw.get(key)
    ]
    if missing:
        raise SPApiError(
            "config.local.yaml 里 amazon_sp_api 缺少必填项："
            + "、".join(missing)
            + "。参考 config.yaml 里 amazon_sp_api 一节的模板，把真实凭证填到 config.local.yaml。"
        )

    return SPApiConfig(
        client_id=raw["client_id"],
        client_secret=raw["client_secret"],
        refresh_token=raw["refresh_token"],
        marketplace_id=raw["marketplace_id"],
        base_url=raw["base_url"],
    )


class SPApiClient:
    """一个 SP-API 客户端实例对应一个卖家账户的一组凭证。

    Access Token 有效期约 1 小时，内部自动缓存并在快过期时自动用 Refresh Token 换新的，
    调用方不需要关心 Token 的生命周期，只管调 get()。
    """

    # 提前多久（秒）就当作快过期，避免请求发出去的路上刚好过期
    _EXPIRY_SAFETY_MARGIN_SECONDS = 60

    def __init__(self, config: SPApiConfig | None = None, session: requests.Session | None = None):
        self._config = config or load_sp_api_config()
        self._session = session or requests.Session()
        self._access_token: str | None = None
        self._access_token_expires_at: float = 0.0

    @property
    def marketplace_id(self) -> str:
        return self._config.marketplace_id

    def _ensure_access_token(self) -> str:
        if self._access_token and time.monotonic() < self._access_token_expires_at:
            return self._access_token

        resp = self._session.post(
            LWA_TOKEN_URL,
            data={
                "grant_type": "refresh_token",
                "refresh_token": self._config.refresh_token,
                "client_id": self._config.client_id,
                "client_secret": self._config.client_secret,
            },
            timeout=15,
        )
        if resp.status_code != 200:
            raise SPApiError(
                f"获取 Access Token 失败（HTTP {resp.status_code}）",
                status_code=resp.status_code,
                body=resp.text,
            )

        payload = resp.json()
        self._access_token = payload["access_token"]
        expires_in = payload.get("expires_in", 3600)
        self._access_token_expires_at = time.monotonic() + expires_in - self._EXPIRY_SAFETY_MARGIN_SECONDS
        return self._access_token

    def get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        """发一个 GET 请求，path 是不带域名的接口路径，比如 "/fba/inbound/v0/shipments"。

        返回解析好的 JSON（dict）。非 200 状态码直接抛 SPApiError，调用方不需要自己判断
        status_code——按"宁可报错，不要猜"的约定，请求失败就是失败，不猜测/不静默吞掉。
        """
        access_token = self._ensure_access_token()
        resp = self._session.get(
            f"{self._config.base_url}{path}",
            headers={
                "x-amz-access-token": access_token,
                "content-type": "application/json",
            },
            params=params or {},
            timeout=30,
        )
        if resp.status_code != 200:
            raise SPApiError(
                f"SP-API 请求失败：GET {path}（HTTP {resp.status_code}）",
                status_code=resp.status_code,
                body=resp.text,
            )
        return resp.json()
