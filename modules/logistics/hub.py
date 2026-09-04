"""物流仓库部门的入口——列出这个部门有哪些工具，点了哪个才进那个工具的界面。

这个文件本身只 import PySide6 和 core 里的通用组件，不 import 具体工具用到的 openpyxl/
pymupdf/pylibdmtx/pytesseract 这些——那些是"发货数量核对"这一个工具自己的事，只有真的点开
它才会被检查/安装/import（见 core/hub_widget.py 和 walmart_shipment_reconcile/panel.py）。
以后物流仓库加新工具，在下面的列表里加一个 ToolInfo 就行。
"""
from __future__ import annotations

from PySide6.QtWidgets import QWidget

from core.dependency import pip_package
from core.hub_widget import HubWidget, ToolInfo

from .walmart_shipment_reconcile.tesseract_dependency import tesseract_ocr


def _build_walmart_reconcile_panel() -> QWidget:
    from .walmart_shipment_reconcile.panel import WalmartReconcilePanel

    return WalmartReconcilePanel()


def _build_fba_label_redact_panel() -> QWidget:
    from .fba_label_redact.panel import FbaLabelRedactPanel

    return FbaLabelRedactPanel()


def _build_shipment_plan_apply_panel() -> QWidget:
    from .shipment_plan_apply.panel import ShipmentPlanApplyPanel

    return ShipmentPlanApplyPanel()


def _build_logistics_tracking_panel() -> QWidget:
    from .logistics_tracking.panel import LogisticsTrackingPanel

    return LogisticsTrackingPanel()


def _build_purchase_order_import_panel() -> QWidget:
    from .purchase_order_import.panel import PurchaseOrderImportPanel

    return PurchaseOrderImportPanel()


def build_panel() -> QWidget:
    tools = [
        ToolInfo(
            id="walmart_shipment_reconcile",
            name="发货数量核对（Walmart）",
            description="核对箱唛实际发货数量和发货计划表是否一致，并按 SKU 拆分箱唛 PDF",
            build_panel=_build_walmart_reconcile_panel,
            dependencies=[
                pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
                pip_package("pymupdf", import_name="fitz", display_name="PyMuPDF（读取 PDF）"),
                pip_package("Pillow", import_name="PIL", display_name="Pillow（图片处理）"),
                pip_package("pylibdmtx", display_name="pylibdmtx（解析箱唛条码）"),
                pip_package("pytesseract", display_name="pytesseract（OCR 接口）"),
                tesseract_ocr(),
            ],
        ),
        ToolInfo(
            id="fba_label_redact",
            name="FBA 标签发货人信息脱敏",
            description="批量去掉一个目录下箱唛 PDF 的发货人信息，加拿大目的地额外整个删掉发货地",
            build_panel=_build_fba_label_redact_panel,
            dependencies=[
                pip_package("pymupdf", import_name="fitz", display_name="PyMuPDF（读取/编辑 PDF）"),
            ],
        ),
        ToolInfo(
            id="shipment_plan_apply",
            name="发货计划自动更新",
            description="把运营提交的发货计划表导入，自动更新采购订单汇总表和发货计划汇总表",
            build_panel=_build_shipment_plan_apply_panel,
            dependencies=[
                pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
            ],
        ),
        ToolInfo(
            id="logistics_tracking",
            name="物流跟踪自动更新",
            description="按物流商分组，并行去各货代平台查询运单最后路由，自动更新物流跟踪表格",
            build_panel=_build_logistics_tracking_panel,
            dependencies=[
                pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
                pip_package("requests", display_name="requests（调用货代平台接口）"),
                pip_package("pycryptodome", import_name="Crypto", display_name="pycryptodome（壹鹿有你/众壹登录用的 AES/RSA 加密）"),
            ],
        ),
        ToolInfo(
            id="purchase_order_import",
            name="采购订单批量导入",
            description="批量读取一个文件夹里的采购订单文件，自动追加进采购订单汇总表和发货计划汇总表",
            build_panel=_build_purchase_order_import_panel,
            dependencies=[
                pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
            ],
        ),
    ]
    return HubWidget(tools)
