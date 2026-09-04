"""壹鹿有你(YLYN, yl.noms.logistics-tms.com)运单最后路由查询。

这家没有像纽酷那样的开放API文档，是照着网站前端（Vben Admin + 若依风格的后台，"跨境物流OMS"）
逆向摸出来的：

    1. 登录 POST /prod-api/auth/login，请求体要走一层 AES+RSA 混合加密（前端拦截器对所有
       POST/PUT 都这么处理）：随机造一个 32 位字符串当 AES 密钥，AES-ECB-PKCS7 加密请求体，
       AES 密钥本身（先 base64 一次）用 RSA 公钥加密放进 encrypt-key 请求头——RSA 公钥写死在
       /_app.config.js 里。前端代码里响应体理论上也能反过来加密（同一个 config.js 里还给了
       一把"私钥"），但实测这个部署的所有响应（包括登录）都没带 encrypt-key 响应头、就是
       明文 JSON，那把"私钥"实际根本没用上（拿它自己验证加解密也对不上，大概率是随便配的
       占位值，不是真的公私钥配对）——所以这里只实现请求加密，响应直接当明文 JSON 解析。
    2. 网站上"运单查询"页面（界面 URL 是 order-manage/manage）背后查的是 GET /oms/BILL/list
       （不加密，GET 不在加密范围内）——这家没有 NK 那种"一个运单号查一条完整轨迹事件列表"
       的接口，返回的是"提单/货件"记录，但每条记录自带 trackRuleName（最新轨迹描述）+
       trackRuleDate（对应时间），实测这正好就是"详细轨迹里最后一条"，不用再另外调接口去挑
       最新的。

网站本身就有"导出"功能、不用一个运单号一个运单号去查——这里照样思路：get_last_routes() 一次性
把这个账号名下所有 BILL 记录分页拉全（GET /oms/BILL/list 的 total/rows 分页），在本地按
billCode 建一个字典，用户表格里要查的运单号直接从这个字典里取结果，不会对每个运单号单独发
一次请求。

"众壹"(ZY, zy.noms.logistics-tms.com) 是同一套系统的另一个租户/品牌——/_app.config.js 里的
RSA 公钥、clientId 跟这边完全一样，实测就是同一个后端，只是域名不同，所以 registry.py 里 ZY
直接复用这个 YlynClient、传不同的 base_url 就行，不用单独再写一份。
"""
from __future__ import annotations

import base64
import json
import random
import string

import requests
from Crypto.Cipher import AES, PKCS1_v1_5
from Crypto.PublicKey import RSA
from Crypto.Util.Padding import pad

from .base import RouteResult

BASE_URL = "https://yl.noms.logistics-tms.com/prod-api"
DOMAIN = "yl.noms.logistics-tms.com"
TENANT_ID = "000000"  # 登录页面上这个字段是自动填的（取「/auth/tenant/list」返回的第一个租户），
# 不是用户手选的下拉框，实测就是这个值，跟登录用的具体账号无关。
CLIENT_ID = "e5cd7e4891bf95d1d19206ce24a7b32e"

# 前端 /_app.config.js 里写死的 RSA 公钥，只用来加密请求（见模块顶部说明——响应不加密，
# 配套的"私钥"没有实际用上，这里也就不需要它）。
_PUBLIC_KEY_B64 = (
    "MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKoR8mX0rGKLqzcWmOzbfj64K8ZIgOdHnzkXSOVOZbFu/"
    "TJhZ7rFAN+eaGkl3C4buccQd/EjEsj9ir7ijT7h96MCAwEAAQ=="
)
_PUBLIC_KEY = RSA.import_key(base64.b64decode(_PUBLIC_KEY_B64))

_PAGE_SIZE = 100


def _random_key(length: int = 32) -> str:
    charset = string.digits + string.ascii_lowercase + string.ascii_uppercase
    return "".join(random.choice(charset) for _ in range(length))


def _rsa_encrypt(data: bytes) -> str:
    return base64.b64encode(PKCS1_v1_5.new(_PUBLIC_KEY).encrypt(data)).decode()


def _aes_encrypt(plaintext: bytes, key: bytes) -> str:
    return base64.b64encode(AES.new(key, AES.MODE_ECB).encrypt(pad(plaintext, 16))).decode()


class YlynClient:
    def __init__(self, account: str, password: str, base_url: str = BASE_URL, domain: str = DOMAIN, session=None):
        self.base_url = base_url
        self.domain = domain
        self.session = session or requests.Session()
        self.token = self._login(account, password)

    # ---- 登录 + 加解密 ----

    def _login(self, account: str, password: str) -> str:
        payload = {
            "username": account,
            "password": password,
            "grantType": "password",
            "tenantId": TENANT_ID,
            "domain": self.domain,
            "clientId": CLIENT_ID,
        }
        data = self._encrypted_post("/auth/login", payload)
        if data.get("code") != 200:
            raise RuntimeError(f"登录失败: {data.get('msg')}")
        return data["data"]["access_token"]

    def _auth_headers(self) -> dict:
        return {"Authorization": f"Bearer {self.token}", "ClientID": CLIENT_ID}

    def _encrypted_post(self, path: str, payload: dict) -> dict:
        key = _random_key(32)
        key_b64 = base64.b64encode(key.encode()).decode()
        body = _aes_encrypt(json.dumps(payload, ensure_ascii=False).encode("utf-8"), key.encode())
        r = self.session.post(
            f"{self.base_url}{path}",
            data=body.encode("utf-8"),
            headers={
                "Content-Type": "application/json;charset=UTF-8",
                "encrypt-key": _rsa_encrypt(key_b64.encode()),
            },
            timeout=15,
        )
        r.raise_for_status()
        return r.json()

    # ---- 查询 ----

    def _fetch_all_bills(self) -> list[dict]:
        rows: list[dict] = []
        page = 1
        while True:
            r = self.session.get(
                f"{self.base_url}/oms/BILL/list",
                params={"pageNum": page, "pageSize": _PAGE_SIZE},
                headers=self._auth_headers(),
                timeout=20,
            )
            r.raise_for_status()
            data = r.json()
            if data.get("code") != 200:
                raise RuntimeError(f"查询运单列表失败: {data.get('msg')}")
            batch = data.get("data", {}).get("rows") or data.get("rows") or []
            rows.extend(batch)
            total = (data.get("data") or data).get("total", 0)
            if not batch or len(rows) >= total:
                break
            page += 1
        return rows

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        bills_by_code = {bill["billCode"]: bill for bill in self._fetch_all_bills() if bill.get("billCode")}

        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            bill = bills_by_code.get(wb)
            if bill is None:
                results[wb] = RouteResult(waybill=wb, error="未找到该运单")
                continue

            track_name = bill.get("trackRuleName")
            track_date = bill.get("trackRuleDate")
            if track_name:
                last_route = f"{track_date} {track_name}".strip() if track_date else track_name
            else:
                # 还没有轨迹事件的新单，退回用整体状态（比如"已出仓"），别留空。
                last_route = bill.get("status")

            if not last_route:
                results[wb] = RouteResult(waybill=wb, error="暂无路由信息")
                continue
            results[wb] = RouteResult(waybill=wb, found=True, last_route=last_route, raw_events=[bill])
        return results
