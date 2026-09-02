"""可复用的"改动预览"组件：红色显示变化前，绿色显示变化后，按一组"识别字段"（比如订单号+
型号）把同一条记录的变化前/变化后放在一起，自动隐藏两边完全没变化的列，可以指定一个字段
允许直接在表格里编辑（改了立刻写回对应的 openpyxl worksheet，不用等外面"确认写入"再收集）。

给"读一批 Excel 数据、算出一批要写的改动、写之前先给人看一眼确认"这种流程用的——
`modules/logistics/shipment_plan_apply` 就是这么用的（见该模块的 panel.py），以后哪个模块
有类似"改真实业务数据之前先预览"的需求，直接复用这个组件，不用重新写一遍表格渲染、分组、
隐藏没变化的列、可编辑字段、筛选这些逻辑——这些跟"发货计划"这个具体业务完全无关，纯粹是
"两份数据摆在一起对比着看"的通用需要。

典型用法：

    from core.diff_preview import DiffPreviewGroup, DiffTable

    group = DiffPreviewGroup("采购订单汇总表 · 改动对比", key_fields=["订单号", "型号"],
                              editable_field="备注")
    layout.addWidget(group)
    ...
    diff = DiffTable(headers=[...], before_rows=[{...}, ...], after_rows=[{...}, ...])
    group.fill(diff, ws=purchase_ws, col_index=remark_col_index)

行字典（before_rows/after_rows 里的每个 dict）如果想让某一行支持编辑，要在里面塞一个
`ROW_INDEX_KEY` 键，值是这一行在真实工作表里的行号——只有"变化后"的行、且这一行带了这个
key，对应的 editable_field 那一格才会真的允许编辑。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

REMOVED_COLOR = QColor("#F8CBAD")  # 变化前
ADDED_COLOR = QColor("#C6E0B4")  # 变化后

ROW_INDEX_KEY = "__row__"

_MAX_VISIBLE_ROWS = 30
_GROUP_ROLE = Qt.ItemDataRole.UserRole  # 这一行属于第几组（一组=一条变化前+它对应的几条变化后）
_EDIT_TARGET_ROLE = Qt.ItemDataRole.UserRole + 1  # 可编辑格子存 (worksheet, 行号, 列号)，改完直接写回


def format_cell(value) -> str:
    """把单元格的原始值（None/数字/字符串/datetime）转成给人看的字符串。"""
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        # 本地算出来的公式结果经常带一长串浮点误差位数，这里只是给人看，四舍五入到 4 位小数
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)


@dataclass
class DiffTable:
    headers: list
    before_rows: list[dict] = field(default_factory=list)
    after_rows: list[dict] = field(default_factory=list)


class DiffPreviewGroup(QGroupBox):
    """一个"标题 + 筛选框 + 变化前/变化后对比表格"的完整组件，本身就是一个 QGroupBox，
    直接 addWidget 进布局就行。
    """

    itemEdited = Signal()  # 可编辑字段被人改了一次就发一次，方便外面标记"有未保存的手工修改"

    def __init__(
        self,
        title: str,
        key_fields: list[str],
        editable_field: str | None = None,
        parent=None,
    ):
        super().__init__(title, parent)
        self._key_fields = key_fields
        self._editable_field = editable_field

        layout = QVBoxLayout(self)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("筛选"))
        self._filter_edit = QLineEdit()
        self._filter_edit.setPlaceholderText("输入关键字，只看匹配的那几条")
        filter_row.addWidget(self._filter_edit)
        layout.addLayout(filter_row)

        self._table = QTableWidget(0, 0)
        # 表格整体允许编辑，但具体哪个格子能不能改，由每个格子自己的 ItemIsEditable 标志决定
        # （见 fill()）——只有"变化后"那些行的 editable_field 格会带这个标志。
        self._table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.EditKeyPressed
        )
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._table.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self._table)

        self._filter_edit.textChanged.connect(self._apply_filter)

    @property
    def table(self) -> QTableWidget:
        """极少数场景需要直接摸底层 QTableWidget（比如测试里查行数）才用，正常用 fill()/clear() 就够。"""
        return self._table

    def clear(self) -> None:
        self._table.setRowCount(0)
        self._table.setColumnCount(0)
        self._filter_edit.clear()

    def fill(self, diff: DiffTable, ws=None, col_index: int | None = None) -> None:
        """按 key_fields 分组渲染 diff；ws/col_index 是 editable_field 要写回的实际工作表和列号，
        不给的话这一批就不能编辑（比如这张表本身没有对应字段）。

        列会先筛一遍：一列如果在所有组里"变化前"和"变化后"的显示值都完全一样（比如查不出结果、
        两边都是空的外部链接公式），就不显示，只留下真的有变化的列，加上 key_fields、
        editable_field ——这两类字段就算"没变化"也要留着，不然认不出是哪一条、也没法编辑。

        headers 里可能混了非字符串的值（比如表头本身就是日期），直接丢给
        setHorizontalHeaderLabels 会在 PySide6 里报"_pythonToCppCopy: Cannot copy-convert"
        的错误，所以表头文字要单独转一次字符串；查具体某一格的值还是要用原始（没转字符串的）
        header 去 row.get(h) 找。
        """
        key_fields = self._key_fields
        editable_field = self._editable_field

        def key_of(row: dict) -> tuple:
            return tuple(row.get(f) for f in key_fields)

        after_by_key: dict[tuple, list[dict]] = {}
        for row in diff.after_rows:
            after_by_key.setdefault(key_of(row), []).append(row)

        groups: list[tuple[dict, list[dict]]] = [
            (before_row, after_by_key.get(key_of(before_row), [])) for before_row in diff.before_rows
        ]

        changed_columns: set = set()
        for before_row, after_rows in groups:
            for h in diff.headers:
                before_text = format_cell(before_row.get(h))
                if any(format_cell(r.get(h)) != before_text for r in after_rows):
                    changed_columns.add(h)

        always_visible = set(key_fields)
        if editable_field:
            always_visible.add(editable_field)
        visible_headers = [h for h in diff.headers if h in always_visible or h in changed_columns]

        table = self._table
        display_headers = ["", *[format_cell(h) for h in visible_headers]]
        table.setColumnCount(len(display_headers))
        table.setHorizontalHeaderLabels(display_headers)

        rendered: list[tuple[int, str, dict, QColor, bool]] = []
        for group_id, (before_row, after_rows) in enumerate(groups):
            rendered.append((group_id, "－ 变化前", before_row, REMOVED_COLOR, False))
            for after_row in after_rows:
                rendered.append((group_id, "＋ 变化后", after_row, ADDED_COLOR, True))

        table.setRowCount(len(rendered))
        table.blockSignals(True)  # 填表这几十上百次 setItem 不是"用户编辑"，别触发写回
        try:
            for r, (group_id, label, row, color, is_after) in enumerate(rendered):
                marker = QTableWidgetItem(label)
                marker.setBackground(color)
                marker.setData(_GROUP_ROLE, group_id)
                marker.setFlags(marker.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(r, 0, marker)
                for c, h in enumerate(visible_headers, start=1):
                    item = QTableWidgetItem(format_cell(row.get(h)))
                    item.setBackground(color)
                    if (
                        editable_field
                        and is_after
                        and h == editable_field
                        and ws is not None
                        and col_index
                    ):
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                        item.setData(_EDIT_TARGET_ROLE, (ws, row.get(ROW_INDEX_KEY), col_index))
                    else:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(r, c, item)
        finally:
            table.blockSignals(False)

        table.resizeColumnsToContents()
        self._auto_size_height()

    def _auto_size_height(self) -> None:
        # 外层调用方一般会把整个页面放进一个纵向可滚动的容器里，表格自己不该再单独滚动——
        # 所以要让表格的高度正好放得下所有行；但改动一多（比如一次导入几十个 SKU），表格会
        # 变得很长，撑得后面的按钮要划很久才能看到，所以封顶最多显示 30 行，超过的部分表格
        # 自己内部上下滚动，不再无限撑高。
        table = self._table
        table.resizeRowsToContents()
        header_h = table.horizontalHeader().height()
        visible_rows = min(table.rowCount(), _MAX_VISIBLE_ROWS)
        rows_h = sum(table.rowHeight(r) for r in range(visible_rows))
        frame = table.frameWidth() * 2
        needs_scroll = table.rowCount() > _MAX_VISIBLE_ROWS
        table.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded if needs_scroll else Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        table.setFixedHeight(header_h + rows_h + frame + 24)  # 24：给横向滚动条（列多的表会有）留点空间

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        target = item.data(_EDIT_TARGET_ROLE)
        if target is None:
            return
        ws, row_index, col_index = target
        ws.cell(row=row_index, column=col_index, value=item.text())
        self.itemEdited.emit()

    def _apply_filter(self, text: str) -> None:
        # 按组过滤：一组（一条变化前 + 它对应的几条变化后）里只要有一格命中关键字，整组都留着，
        # 不会出现"变化前显示、对应的变化后被筛掉了"这种看不完整的情况。
        table = self._table
        text = text.strip().lower()
        row_count = table.rowCount()
        if not text:
            for r in range(row_count):
                table.setRowHidden(r, False)
            return

        matching_groups: set = set()
        for r in range(row_count):
            marker = table.item(r, 0)
            group_id = marker.data(_GROUP_ROLE) if marker else None
            for c in range(table.columnCount()):
                cell = table.item(r, c)
                if cell is not None and text in cell.text().lower():
                    matching_groups.add(group_id)
                    break

        for r in range(row_count):
            marker = table.item(r, 0)
            group_id = marker.data(_GROUP_ROLE) if marker else None
            table.setRowHidden(r, group_id not in matching_groups)
