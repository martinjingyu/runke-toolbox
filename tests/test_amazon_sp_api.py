"""modules/logistics/amazon_sp_api 的测试。

分两类：
- 合成数据的单测（假的 requests.Session，不联网，跑得快）——覆盖 token 缓存/续期、
  错误处理、货件数据解析这些逻辑本身对不对。
- 一个真实联网的集成测试，只有本机 config.local.yaml 填了真实凭证才会跑，没填就跳过
  （不能因为拿不到真实凭证就让测试直接失败），标了 @pytest.mark.slow。
"""
from __future__ import annotations

import pytest

import modules.logistics.amazon_sp_api.client as sp_api_client_module
from modules.logistics.amazon_sp_api.client import SPApiClient, SPApiConfig, SPApiError, load_sp_api_config
from modules.logistics.amazon_sp_api.fulfillment_inbound import list_inbound_shipments


class _FakeResponse:
    def __init__(self, status_code: int, json_body: dict | None = None, text: str = ""):
        self.status_code = status_code
        self._json_body = json_body or {}
        self.text = text or str(json_body)

    def json(self):
        return self._json_body


class _FakeSession:
    """按测试用例需要，事先塞好 token 请求和后续 GET 请求分别该返回什么。"""

    def __init__(self, token_response: _FakeResponse, get_responses: list[_FakeResponse] | None = None):
        self.token_response = token_response
        self.get_responses = list(get_responses or [])
        self.post_call_count = 0
        self.get_call_count = 0

    def post(self, url, data=None, timeout=None):
        self.post_call_count += 1
        return self.token_response

    def get(self, url, headers=None, params=None, timeout=None):
        self.get_call_count += 1
        return self.get_responses[self.get_call_count - 1]


def _make_config(**overrides) -> SPApiConfig:
    base = dict(
        client_id="amzn1.application-oa2-client.test",
        client_secret="amzn1.oa2-cs.v1.test",
        refresh_token="Atzr|test",
        marketplace_id="ATVPDKIKX0DER",
        base_url="https://sellingpartnerapi-na.amazon.com",
    )
    base.update(overrides)
    return SPApiConfig(**base)


def test_get_fetches_and_reuses_cached_access_token():
    session = _FakeSession(
        token_response=_FakeResponse(200, {"access_token": "tok-1", "expires_in": 3600}),
        get_responses=[
            _FakeResponse(200, {"payload": {"ShipmentData": []}}),
            _FakeResponse(200, {"payload": {"ShipmentData": []}}),
        ],
    )
    client = SPApiClient(config=_make_config(), session=session)

    client.get("/fake/path")
    client.get("/fake/path")

    assert session.post_call_count == 1  # 第二次调用复用缓存的 token，没有重新换
    assert session.get_call_count == 2


def test_get_raises_sp_api_error_on_token_failure():
    session = _FakeSession(token_response=_FakeResponse(400, text="invalid_grant"))
    client = SPApiClient(config=_make_config(), session=session)

    with pytest.raises(SPApiError) as exc_info:
        client.get("/fake/path")

    assert exc_info.value.status_code == 400


def test_get_raises_sp_api_error_on_api_failure():
    session = _FakeSession(
        token_response=_FakeResponse(200, {"access_token": "tok-1", "expires_in": 3600}),
        get_responses=[_FakeResponse(403, text="Forbidden: missing role")],
    )
    client = SPApiClient(config=_make_config(), session=session)

    with pytest.raises(SPApiError) as exc_info:
        client.get("/fake/path")

    assert exc_info.value.status_code == 403


def test_list_inbound_shipments_parses_shipment_fields():
    session = _FakeSession(
        token_response=_FakeResponse(200, {"access_token": "tok-1", "expires_in": 3600}),
        get_responses=[
            _FakeResponse(
                200,
                {
                    "payload": {
                        "ShipmentData": [
                            {
                                "ShipmentId": "FBA15G4XXXXX",
                                "AmazonReferenceId": "REF-123",
                                "ShipmentName": "9/2 补货",
                                "ShipmentStatus": "WORKING",
                            },
                            {
                                # AmazonReferenceId 缺失的情况也要能正常解析，不报错
                                "ShipmentId": "FBA15G4YYYYY",
                                "ShipmentStatus": "SHIPPED",
                            },
                        ]
                    }
                },
            )
        ],
    )
    client = SPApiClient(config=_make_config(), session=session)

    shipments = list_inbound_shipments(client, days_back=30)

    assert len(shipments) == 2
    assert shipments[0].shipment_id == "FBA15G4XXXXX"
    assert shipments[0].amazon_reference_id == "REF-123"
    assert shipments[1].amazon_reference_id is None


def test_list_inbound_shipments_follows_next_token_across_pages():
    session = _FakeSession(
        token_response=_FakeResponse(200, {"access_token": "tok-1", "expires_in": 3600}),
        get_responses=[
            _FakeResponse(
                200,
                {
                    "payload": {
                        "ShipmentData": [{"ShipmentId": "FBA1", "ShipmentStatus": "WORKING"}],
                        "NextToken": "page-2-token",
                    }
                },
            ),
            _FakeResponse(
                200,
                {"payload": {"ShipmentData": [{"ShipmentId": "FBA2", "ShipmentStatus": "SHIPPED"}]}},
            ),
        ],
    )
    client = SPApiClient(config=_make_config(), session=session)

    shipments = list_inbound_shipments(client, days_back=30)

    assert session.get_call_count == 2
    assert [s.shipment_id for s in shipments] == ["FBA1", "FBA2"]


def test_load_sp_api_config_raises_clear_error_when_incomplete(monkeypatch):
    monkeypatch.setattr(
        sp_api_client_module,
        "load_config",
        lambda: {"amazon_sp_api": {"client_id": "", "client_secret": "x"}},
    )

    with pytest.raises(SPApiError) as exc_info:
        load_sp_api_config()

    assert "client_id" in str(exc_info.value)


def _real_credentials_configured() -> bool:
    try:
        config = load_sp_api_config()
    except SPApiError:
        return False
    return not config.client_id.startswith("填入真实的")


@pytest.mark.slow
@pytest.mark.skipif(not _real_credentials_configured(), reason="config.local.yaml 里还没填真实的 SP-API 凭证")
def test_list_inbound_shipments_against_real_account():
    client = SPApiClient()
    shipments = list_inbound_shipments(client, days_back=365)
    # 不断言一定有货件（账号可能确实一年内没发过货），只验证调用链路本身走得通、
    # 拿到的数据结构没有报错。
    for shipment in shipments:
        assert shipment.shipment_id
