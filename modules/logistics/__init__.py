"""物流仓库模块。

目前有一个功能：发货数量核对（Walmart），见 walmart_shipment_reconcile/ 和 panel.py。
其他功能要等对应的《需求申请表》/《开发确认回执单》确认完再加，见
/Users/jingyuhuang/Documents/Work/闰科/软件开发SOP。
"""
from .panel import build_panel

MODULE_INFO = {
    "id": "logistics",
    "name": "物流仓库",
    "description": "仓管部门相关的自动化功能",
    "build_panel": build_panel,
}
