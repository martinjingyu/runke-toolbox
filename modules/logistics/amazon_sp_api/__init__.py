"""亚马逊 Selling Partner API 的底层数据源，供物流仓库部门下的各个工具复用。

不是一个带界面的"工具"，没有 panel.py，也不在 hub.py 的 ToolInfo 列表里出现——纯粹是
一层"调用 SP-API、把结果整理成 Python 数据结构"的业务逻辑，不 import PySide6，方便单测，
也方便以后被多个不同的工具（比如"FBA 货件追踪"界面、定时同步脚本）共用。

用法：

    from modules.logistics.amazon_sp_api.client import SPApiClient
    from modules.logistics.amazon_sp_api.fulfillment_inbound import list_inbound_shipments

    client = SPApiClient()  # 自动从 config.yaml / config.local.yaml 读凭证
    shipments = list_inbound_shipments(client, days_back=90)

凭证配置见 config.local.yaml 里的 amazon_sp_api 一节（真实密钥不进 git，只在本机）。
"""
