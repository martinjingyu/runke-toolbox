"""众包(ops.zbao56.com，AngularJS + Spring 后台，系统里显示"电商业务系统")运单最后路由查询。

## 之前那版(appKey/appToken 走开放 EDI API)为什么废弃了

一开始以为"账号管理"里存的账号密码是给 GET /edi/web-services/v5/tracking 这个开放 EDI 接口用的
appKey/appToken（这接口本身是真的、文档里也真有）。但实测这俩字段其实是**网页登录**用的账号密码，
appKey/appToken 是另一码事——要登进网页后台、在"接口授权管理"(/#/interface-authorization，
菜单里叫"API MGMT")里手动生成，而且这个页面还要求账号有 `AU_API` 权限（要"老板"账号在"员工管理"
里单独给员工账号勾选开通，默认没有）。这条路对咱们用的这个员工账号走不通，所以换了下面这条路：
直接复用网页登录本身的账号密码，把网页"物流订单"列表页背后真正调的接口摸出来直接调。

## 登录：POST /api/authentication，密码要 RSA 加密

这是 Spring Security 经典的表单登录(不是 JSON/JWT 那套)，body 是
`application/x-www-form-urlencoded`：`username=<账号>&password=<RSA加密后的密码>&remember-me=false&submit=Login`。

密码不能明文传——前端登录框提交前会用 `EncryptionService`(AngularJS service，底层用的是开源库
JSEncrypt)把密码用一把写死在前端 JS 包里的 RSA 公钥加密(标准 PKCS1v1.5 padding)，加密结果再
base64 一下才是 password 字段的值。这里用 pycryptodome 照抄同一个公钥、同一种 padding 复现，
写法跟 ylyn.py 登录用的 RSA+AES 混合加密是一个路数(这家更简单，只有 RSA，没有 AES 那层)。
(公开发布的前端 JS 里翻出来的，属于任何人打开这个网站都能直接下载到的公开信息，不是破解什么私有
秘密——公钥本来就是设计成可以公开的。)

登录接口还要求带 CSRF cookie：先 GET 一次首页/任意页面，服务端会种一个 `XSRF-TOKEN` cookie，
登录请求要把这个值原样放进 `X-XSRF-TOKEN` 请求头，不然会被 Spring Security 的 CSRF 过滤器拦。
登录成功后服务端会种一个 `SESSION` cookie，后续接口调用只要带这个 cookie 就认（`requests.Session`
自动处理，不用手动管）。

## 运单列表：PUT /api/bookings/getFilterPage

这是网页"物流订单(ALL)"列表页(/#/all-shipments)背后真正调的接口，不是逐个运单号查、是分页拉
列表：`?page=0&size=<n>&sort=id,desc`，body 是 `{"fmsType":"all","isShipper":<bool>}`。
`isShipper` 对应列表页顶部"客户身份"/"发货人身份"这两个 tab——实测"发货人身份"(isShipper=true)
是"客户身份"(isShipper=false)的超集(多一条)，所以这里两个都拉一遍再按 jobNum 去重合并，不只
依赖一个 tab，防止漏单。account 名下运单量不大(实测两三百条)，size 直接开到 5000 一次拉全，不用
真分页。

单号字段用 `jobNum`(网页"订单号"列，形如 `ZBSZ26072657`，业务侧确认这个就是运单跟踪表里 ZB 那栏
填的号)。响应里还有 `soNum`/`hblNum`/`mblNum`/`amsNum` 这几个其它单号字段，目前没用上，如果以后
发现运单表里 ZB 号码对不上 jobNum、其实填的是这几个里的某个，再来改匹配字段。

"最后路由"用 `lastestTkStatus` 字段(网页列表里"最新货物动态"那一列)，前面拼一个 `lastModifiedTime`
当时间前缀，跟其它货代模块格式保持一致——虽然 `lastestTkStatus` 文本末尾自己往往也带一段时间，但
格式不统一(有的有、有的没有)，不能依赖它。还没有轨迹更新的新单 `lastestTkStatus` 是 null，退回用
`status` 这个业务状态码兜底，不留空。
"""
from __future__ import annotations

import base64

import requests
from Crypto.Cipher import PKCS1_v1_5
from Crypto.PublicKey import RSA

from .base import RouteResult

BASE_URL = "https://ops.zbao56.com"

# 前端 JS 包(app-*.js)里 EncryptionService 写死的 RSA 公钥，标准 PKCS1v1.5 padding，用来加密
# 登录密码。公钥本身是给所有访问者用的公开信息，浏览器一样能直接下载到这段 JS。
_PUBLIC_KEY_B64 = (
    "MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCGB9JGiTWEr4wiWYO7VwWf6HfcK52wjNYLg9/x9UMf"
    "SdvOmzxi1ryHTjY2SY5ru03Qm7E+feiszB+K+ns4G803mgneI7ej6b4lmsoe4PlD05fsmx9ut2QyWL13"
    "rLuLkUW2/zmVrvf+qZn2SQSiTPlLYrBQAwcKMaXYlIrEV6/A6wIDAQAB"
)
_PUBLIC_KEY = RSA.import_key(base64.b64decode(_PUBLIC_KEY_B64))

_PAGE_SIZE = 5000


def _rsa_encrypt_password(plaintext: str) -> str:
    cipher = PKCS1_v1_5.new(_PUBLIC_KEY)
    return base64.b64encode(cipher.encrypt(plaintext.encode("utf-8"))).decode()


class ZhongbaoClient:
    def __init__(self, username: str, password: str, base_url: str = BASE_URL, session=None):
        self.base_url = base_url
        self.session = session or requests.Session()
        self._login(username, password)

    def _xsrf_header(self) -> dict:
        token = self.session.cookies.get("XSRF-TOKEN")
        return {"X-XSRF-TOKEN": token} if token else {}

    def _login(self, username: str, password: str) -> None:
        # 先随便 GET 一次，拿服务端种下来的 XSRF-TOKEN cookie，登录请求要带这个。
        self.session.get(f"{self.base_url}/", timeout=15)

        enc_password = _rsa_encrypt_password(password)
        body = (
            f"username={requests.utils.quote(username, safe='')}"
            f"&password={requests.utils.quote(enc_password, safe='')}"
            f"&remember-me=false&submit=Login"
        )
        r = self.session.post(
            f"{self.base_url}/api/authentication",
            data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded", **self._xsrf_header()},
            timeout=15,
        )
        if r.status_code != 200:
            raise RuntimeError(f"登录失败(账号或密码错误，HTTP {r.status_code})")

    def _fetch_bookings(self, is_shipper: bool) -> list[dict]:
        r = self.session.put(
            f"{self.base_url}/api/bookings/getFilterPage",
            params={"page": 0, "size": _PAGE_SIZE, "sort": "id,desc"},
            json={"fmsType": "all", "isShipper": is_shipper},
            headers=self._xsrf_header(),
            timeout=30,
        )
        r.raise_for_status()
        return r.json() or []

    def _fetch_all_bookings(self) -> list[dict]:
        # "客户身份"/"发货人身份"两个 tab 实测不完全是同一批运单(后者是前者的超集，但不放心
        # 保证以后一直是超集关系)，两边都拉一遍按 jobNum 合并，防止漏单。
        by_job_num: dict[str, dict] = {}
        for is_shipper in (False, True):
            for row in self._fetch_bookings(is_shipper):
                job_num = row.get("jobNum")
                if job_num:
                    by_job_num[job_num] = row
        return list(by_job_num.values())

    def get_last_routes(self, waybill_numbers: list[str]) -> dict[str, RouteResult]:
        by_job_num = {row["jobNum"]: row for row in self._fetch_all_bookings()}

        results: dict[str, RouteResult] = {}
        for wb in waybill_numbers:
            row = by_job_num.get(wb)
            if row is None:
                results[wb] = RouteResult(waybill=wb, error="未找到该运单")
                continue

            content = row.get("lastestTkStatus") or row.get("status")
            modified_time = row.get("lastModifiedTime")
            last_route = f"{modified_time} {content}".strip() if modified_time else content

            if not last_route:
                results[wb] = RouteResult(waybill=wb, error="暂无路由信息")
                continue
            results[wb] = RouteResult(waybill=wb, found=True, last_route=last_route, raw_events=[row])
        return results
