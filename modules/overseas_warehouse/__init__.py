"""海外仓运营部门模块。

目前有一个功能：海外仓销量汇总表导入，见 hub.py（部门入口，轻量）和
sales_import/panel.py（工具自己的界面，点开才会被 import）。其它功能要等对应的
《需求申请表》/《开发确认回执单》确认完再加，见 /Users/jingyuhuang/Documents/Work/闰科/软件开发SOP。
"""
from .hub import build_panel

MODULE_INFO = {
    "id": "overseas_warehouse",
    "name": "海外仓运营",
    "description": "海外仓运营部门相关的自动化功能",
    "build_panel": build_panel,
}
