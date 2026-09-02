# runke-toolbox 项目说明（给接手开发的 code agent 看）

闰科（跨境电商卖家）内部工具箱，桌面软件，核心壳 + 按部门拆分的可插拔模块。**接下来会一个模块
一个模块地交给不同的 agent 开发**，这份文档的目的是让每个 agent 不用重新摸索一遍架构、少踩已经
踩过的坑、界面风格保持统一。

开发前务必看一遍这份文档，尤其是"设计原则"和"踩过的坑"两节——不是可读可不读的背景资料，是
真实踩过、代价不小的教训。

## 项目结构

```
core/                    核心壳，跟具体业务无关
  app.py                 启动窗口（MainWindow）、main()
  module_registry.py     发现 modules/ 下的部门模块
  hub_widget.py           部门内"工具列表 → 点进去具体功能"的导航组件（HubWidget/ToolInfo）
  dependency.py           工具按需装依赖的机制（Dependency/pip_package）
  config.py               读 config.yaml / config.local.yaml
  storage/                统一的数据读写接口（local/nas 两种后端）
  diff_preview.py         可复用组件：预览一批改动（红=变化前/绿=变化后），见下文"模板组件"
  backup.py               可复用：写入真实文件前先备份

modules/
  <department>/
    __init__.py            MODULE_INFO 字典，核心壳靠这个发现模块
    hub.py                  部门入口，列出这个部门有哪些工具（ToolInfo 列表）
    <tool_name>/
      panel.py              这个工具的界面
      *.py                  业务逻辑（不 import PySide6，纯 Python，方便单测）

modules/logistics/ 是目前唯一的部门（物流仓库），下面三个工具可以直接当参考实现：
  walmart_shipment_reconcile/  最早写的，模式最简单：单一后台线程 + 进度条 + 结果表格
  fba_label_redact/            PDF 批处理，模式类似
  shipment_plan_apply/         最复杂也最新，预览+确认写入两阶段、用到了 core/diff_preview.py
                                这个模板组件——新工具如果涉及"改真实业务数据"，照着这个抄

tests/                    pytest，跟 modules/ 结构对应，一个模块一个 test_xxx.py
```

## 设计原则（新模块必须遵守）

### 1. 核心壳极简，工具自己的依赖自己声明、按需装

`requirements.txt`（软件启动就要装的东西）只有 `PySide6` + `PyYAML`，装起来要快。**不要**往这个
文件里加任何具体工具用到的库。

一个工具需要 `openpyxl`/`pymupdf`/`pytesseract` 之类的东西，在部门的 `hub.py` 里，通过
`ToolInfo.dependencies` 声明：

```python
from core.dependency import pip_package

ToolInfo(
    id="xxx",
    name="...",
    description="...",
    build_panel=_build_xxx_panel,   # 见下一条：必须是懒加载
    dependencies=[
        pip_package("openpyxl", display_name="openpyxl（读写 Excel）"),
        pip_package("pymupdf", import_name="fitz", display_name="PyMuPDF（读取 PDF）"),
    ],
)
```

`core/hub_widget.py` 只有在用户**第一次点开这个工具**时才检查/安装这些依赖（装好了就不会再问）。
不装的话工具打不开，不会影响软件里其它工具正常用。

`environment.yml`（Mac 开发/测试用的 conda 环境）可以装得比较全，跟 `requirements.txt` 故意不对称
——这是给开发机准备的，不代表最终用户也要装这么多。

### 2. 部门 hub.py 必须懒加载，模块之间互不拖累

`hub.py` 本身只能 `import PySide6` 和 `core` 里的东西，**不能**在文件顶部 import 具体工具用到的
重量级库（`fitz`/`openpyxl`/`pylibdmtx` 等）。每个工具的 `build_panel` 要写成这样：

```python
def _build_xxx_panel() -> QWidget:
    from .xxx_tool.panel import XxxPanel   # import 放函数体内，真的点开才会执行
    return XxxPanel()
```

这样一个工具的依赖没装好，不会导致这个部门里其它工具，或者其它部门，一起打不开——
`core/module_registry.py` 也专门做了这层隔离：一个部门模块加载失败（比如它自己的 `__init__.py`
import 时炸了），只影响这一个部门显示"加载失败"，其它部门照常用。

### 3. 界面风格统一——照抄已有模式，不要另起炉灶

**后台任务**：任何耗时操作（读大表格、跑 OCR、批量处理 PDF）都要放 `QThread` 子类里跑，
不能卡住界面。约定的信号命名：

```python
class _XxxWorker(QThread):
    succeeded = Signal(object)   # 结果对象，具体类型工具自己定
    failed = Signal(str)         # 错误信息（字符串，直接进 QMessageBox）
    progress = Signal(int, int)  # done, total —— 没有这个信号就说明是不可拆分进度的任务
```

**进度条**：优先给真实的 done/total 百分比，不要图省事用不可拆分进度的"一直转圈"——
`walmart_shipment_reconcile` 按 PDF 页数报进度，`shipment_plan_apply` 按"已经处理了几笔分摊"
报进度，都是从一开始就不确定总数、拿到第一个回调才把进度条从"忙碌样式"切成百分比样式，
参考这两个的 `_on_progress`。

**取消**：用协作式取消（`threading.Event`），不要用 `QThread.terminate()`——实测线程卡在
`concurrent.futures`/多进程等待里的时候，`terminate()` 会直接卡死拿不回来。

**关闭软件**：工具的 Panel 如果有后台线程，要实现 `stop_running_tasks()` 方法（`MainWindow.closeEvent`
和 `HubWidget.stop_running_tasks()` 会自动调用所有已经打开过的工具的这个方法，包括已经点了
"返回"但线程还没跑完的）。

**文件选择**：单文件用 `QFileDialog.getOpenFileName`，多文件 `getOpenFileNames`，见各工具 panel.py
里的 `_file_picker_row` 辅助函数（每个文件里各抄一份，没有提到共享，因为足够小、不值得再抽一层）。
长期不变的路径（比如汇总表这种"每次都用同一份文件"的场景）用 `QSettings` 记住上次选的路径，
见 `shipment_plan_apply/panel.py` 的 `_SETTINGS_KEY_*` 那几行，`core/app.py` 的
`main()` 里已经设置了 `setOrganizationName`/`setApplicationName`，新模块不用再设一遍。

**颜色约定**：结果有"对/错"或者"变化前/变化后"这种二元状态时，统一用这两个颜色（跟
`core/diff_preview.py` 里保持一致）：

```python
REMOVED_COLOR / MISMATCH_COLOR ≈ QColor("#F8CBAD")  # 偏红：不一致/变化前/有问题
ADDED_COLOR / MATCH_COLOR ≈ QColor("#C6E0B4")        # 偏绿：一致/变化后/没问题
```

### 4. 业务逻辑和界面彻底分离

`panel.py` 只管 Qt 控件和信号槽，实际的读表格/算逻辑/写文件都在同目录下别的 `.py` 文件里，
**不 import PySide6**——这样业务逻辑可以直接用 pytest 测，不需要起 Qt 应用、不需要
`QT_QPA_PLATFORM=offscreen`。`walmart_shipment_reconcile/reconcile.py`、
`shipment_plan_apply/planner.py` 都是这个模式。

### 5. 涉及真实业务数据：宁可报错，不要猜

这条是从 `shipment_plan_apply`（改真实的采购/发货数据）踩出来的血泪教训，任何新模块只要涉及
**修改用户已有的业务数据文件**（不是"生成一份新报告"这种只增不改的场景），都要遵守：

- **表头匹配用精确匹配，不用模糊/关键词匹配**。SKU 识别、姓名脱敏这种"允许有轻微书写差异"
  的场景才用模糊匹配（`normalize_sku` 那种），业务数据的字段名不该有歧义。
- **能找到多个候选、又没法唯一确定选哪个的时候，直接报错，不要选"看起来最像"的那个**——
  `shipment_summary.py` 的 `AmbiguousPendingRowError` 是真实教训：曾经因为"先找到哪行就用哪行"
  的逻辑，把数据写错了行。
- **算出来的结果超出预期范围（比如要写入的数量比容量还大）要报错，不要静默地当成极限情况处理**——
  同一个文件里 `InconsistentQuantityError` 也是真实踩过的坑。
- **改动之前一定要有"预览 → 人工确认 → 才真的写盘"这个环节，写盘前自动备份原文件**——
  见下面"模板组件"一节，这个流程已经封装成可复用组件了，不用每个模块重新写一遍。
- 遇到"猜不出用户意图"的地方（比如运营给的表格该用哪个 sheet、多个候选选哪个），交给界面
  让人工选，不要自动猜——`shipment_plan_apply` 的 sheet 选择、模板类型识别都是"自动识别 +
  人工可改"，不是全自动。

### 6. 真实数据驱动开发，不要只凭读代码/猜测下结论

写涉及 Excel 读写的逻辑时，先用真实样例数据跑一遍、打印中间结果看看跟猜的是否一致，再决定怎么写，
不要靠"我觉得公式应该是这样"就直接实现。这个项目里好几个严重 bug（openpyxl 插入行不会自动调整
公式引用、`setHorizontalHeaderLabels` 不接受非字符串表头、read_only 模式下访问超出范围的行会
直接抛 `IndexError`）都是拿真实数据跑出来才发现的，纯读代码/读文档发现不了。测试真实数据的路径
约定见"测试约定"一节。

## 模板组件（新模块优先复用，不要重新实现）

### `core/diff_preview.py` —— "改动预览"组件

任何"读一批 Excel 数据、算出一批要写的改动、写之前先给人看一眼确认"的场景都能直接用：

```python
from core.diff_preview import DiffPreviewGroup, DiffTable

group = DiffPreviewGroup(
    "采购订单汇总表 · 改动对比",
    key_fields=["订单号", "型号"],       # 用来把同一条记录的"变化前/变化后"配对分组
    editable_field="备注",              # 可选：允许在表格里直接改这一列，改了立刻写回 worksheet
)
layout.addWidget(group)   # 它本身就是 QGroupBox，直接扔进布局

# 业务逻辑那边算出 before_rows / after_rows（每行是 {表头: 值} 字典）之后：
diff = DiffTable(headers=[...], before_rows=[...], after_rows=[...])
group.fill(diff, ws=some_openpyxl_worksheet, col_index=备注列号)   # ws/col_index 不给就不能编辑
```

自带的能力：红/绿两色区分变化前后、按 key_fields 分组配对、**自动隐藏两边完全没变化的列**（只留
下真正变了的 + key_fields + editable_field）、表格内筛选框、最多显示 30 行超出部分内部滚动、
可编辑字段改了直接写回 worksheet 对象。`shipment_plan_apply/panel.py` 是唯一的调用方，也是最好
的参考代码。

行字典如果要支持编辑，要带上 `core.diff_preview.ROW_INDEX_KEY` 这个 key，值是这一行在真实
worksheet 里的行号——业务逻辑那边生成快照的时候顺手塞进去，参考
`shipment_plan_apply/diff.py` 的 `_snapshot_purchase`/`_snapshot_summary`。

### `core/backup.py` —— 写入前备份

```python
from core.backup import backup_file

backup_path = backup_file(target_path)   # 在原文件旁边生成带时间戳的备份，失败直接抛异常
```

调用方要保证：备份失败就整个中止写入，不能带着"这次没备份成"的状态继续写。

### 还没抽出来但以后可能值得抽的

- 三段式"选模板类型 → 选 sheet → 校验表头"的文件导入流程（`shipment_plan_apply/shipment_templates.py`
  目前是工具自己实现的，如果第二个工具也要处理"运营给的、格式不完全统一的 Excel 表格"，可以
  考虑抽出来）。
- Excel 公式安全求值（`shipment_plan_apply/column_utils.py` 的 `resolve_cell_value`，只支持
  "纯本行四则运算/字符串拼接"这种公式，跨表引用/函数调用会放弃返回 None）——如果以后别的模块
  也要在预览界面里把公式列显示成算出来的结果而不是公式原文，这个可以直接复用或者抽到 core 里。

## 新增一个部门模块

在 `modules/` 下新建一个包，`__init__.py` 参考 `modules/logistics/__init__.py`：

```python
from .hub import build_panel

MODULE_INFO = {
    "id": "xxx",           # 唯一标识
    "name": "部门中文名",
    "description": "一句话描述",
    "build_panel": build_panel,
}
```

`core/module_registry.py` 会自动发现，不用改 `core/` 任何代码。`hub.py` 参考
`modules/logistics/hub.py`，是一个 `build_panel() -> QWidget` 函数，内部用
`core.hub_widget.HubWidget(tools)` 返回一个"工具列表 → 点进去具体功能"的导航界面（`tools` 是
`ToolInfo` 列表）。

## 新增一个部门内的工具

在部门目录下新建一个子包（工具自己的所有代码，业务逻辑+panel.py），在部门的 `hub.py` 里加一个
`ToolInfo`：懒加载的 `build_panel`、声明好 `dependencies`。参考本文档"设计原则"第 2 条的代码示例。

开发顺序建议：先写业务逻辑（不碰 PySide6），拿真实数据反复验证，写 pytest 测试；业务逻辑稳定了
再写 `panel.py` 把它接到界面上，UI 层能复用的组件（`core/diff_preview.py` 等）优先复用。

## 测试约定

- `tests/test_<module>.py`，一一对应 `modules/<dept>/<tool>/` 里的业务逻辑文件。
- 合成数据（用 `openpyxl.Workbook()` 现造一个最小化的测试文件）优先，跑得快、不依赖本机
  有没有真实业务文件，CI/别的机器上也能跑。
- 真实数据集成测试用 `REAL_DATA_DIR = Path("/Users/jingyuhuang/Documents/Work/闰科/...")` 这种
  写死的绝对路径 + `@pytest.mark.skipif(not REAL_DATA_DIR.exists(), reason="...")`，本机有数据就
  跑、没有就跳过，不能因为拿不到真实数据就让测试直接失败。真的很慢的测试（要跑几分钟那种）额外
  加 `@pytest.mark.slow`，`pytest.ini` 里已经注册了这个 marker，日常开发用 `-m "not slow"` 跳过。
- `pytest tests/` 跑全部；涉及 PySide6 的地方（构造 Panel、smoke test）在没有显示器的环境要加
  `QT_QPA_PLATFORM=offscreen` 环境变量。

## 部署

- Windows 用户用 `setup_windows.bat`（内部调 `setup_windows.ps1`）一键装 Python + 建虚拟环境 +
  装 `requirements.txt`；`uninstall_windows.bat` 清空环境方便重新测试安装流程。日常启动用
  `run_windows.vbs`（隐藏控制台窗口）。
- 这几个脚本改动很容易踩 Windows 特有的坑（见下一节），改之前最好想清楚要解决的问题，改完最好
  能找一台真实 Windows 机器（或者至少用 `pwsh`，Mac 上 `brew install powershell` 能装）实际跑一下
  再确认。

## 踩过的坑（不要重蹈覆辙）

- **业务人员口头描述的字段含义，跟这个字段在真实数据里实际的含义，可能对不上——拿真实数据
  验证过才能信，不能只凭对话里的一句话就去实现。** `shipment_plan_apply` 早期开发时，用户
  说过"有些货码本质上是同一个货，只是打了不同的货码在卖，货号没货的时候可以用另一个货号的
  库存"，据此实现了一个"货号缺货时，查这个货号在售产品信息总表里的'变体'字段，换那个货号
  再查一次库存"的兜底逻辑。后来用真实数据验证到一个具体案例才发现：'变体'字段实际标的是
  "同一个造型、不同颜色"的一组货号（比如"变体=TD-CK-206"这一个标签下，挂着白色/灰色/亮棕色/
  泥土色等 20 多个不同颜色的货号）——颜色不同的货根本不能互相顶替发货，这个兜底逻辑一旦真的
  跑到生产环境，会导致"白色缺货，系统偷偷把库存里的灰色当白色发出去"这种实际发错货的严重
  后果。最后拿掉了这个兜底：一个货号自己名下所有采购订单都不够，就是真的缺货，直接报错，
  不去动任何"变体"关联货号的库存。教训是：哪怕业务人员的描述听起来逻辑通顺、哪怕这个描述
  就是他们自己说的，只要要实现的是"自动拿一份数据去顶替/影响另一份数据"这种有实际后果的
  逻辑，动手写代码之前也要先挑几个真实的具体案例，把这个字段在数据里到底是怎么用的、跟哪些
  别的字段的实际值对照着看一遍，不能只凭对话内容里的一句抽象描述。见
  `modules/logistics/shipment_plan_apply/product_lookup.py` 和 `purchase_book.py` 现在的
  实现（不做任何跨货号库存替代）。
- **openpyxl 插入行/列不会像 Excel 那样自动调整公式里的单元格引用**。这个库只是读写文件底层
  格式，没有公式引擎。往表格中间插入一行，插入点以下所有行（不只是你直接操作的那一行）里，
  公式文本中引用的行号都要手动改，否则打开 Excel 之后这些公式会读到错位的数据，还不容易发现。
  参考 `shipment_summary.py` 的 `_reindex_shifted_rows`/`_shift_formula_refs`（通用规则：公式里
  任何 `>= 插入点` 的行号都要 +1，不只是"自己引用自己这一行"这一种情况，区间公式如
  `SUM(E6:E100)` 的两个边界都要分别判断）。插入列同理，见 `purchase_book.py` 的
  `find_or_create_date_column`。—— 已经问过用户要不要换成 `xlwings`（遥控真实 Excel 来做，
  公式调整交给 Excel 自己处理）代替这套手写逻辑，用户选择继续用 openpyxl + 手写修正（更快、
  不依赖 Excel 安装），如果以后这套手写逻辑又踩到没覆盖的公式写法，可以重新考虑这个选项。
- **`insert_rows`/`insert_cols` 同样不会像 Excel 那样自动处理"格式"**，这条差点造成大问题：
  用户反馈"写入了新表但排版都没了"，一查是两个叠加的坑——① 新插入的那一整行/整列，格子本身
  是完全没有任何样式（字体、填充色、边框、数字格式）的空白格子，需要手动从旁边的行/列把样式
  抄过去（`copy()` 每个 `cell.font`/`.fill`/`.border`/`.alignment`/`.number_format`/
  `.protection`，不能只抄 `.value`）；② 行高/列宽这种"整行/整列"级别的设置（`row_dimensions`/
  `column_dimensions`，用行号/列字母当 key）**不会跟着插入操作一起往下/往右挪**——插入点以下
  原来的行/列，格子内容和样式会正确挪过去（这个 openpyxl 做对了），但行高列宽的设置留在原来
  的行号/列号上不动，插入之后要么新插入的行/列意外"继承"了不该属于它的旧设置，要么原来那一行/
  列的设置丢了、变成默认值——必须手动把 `row_dimensions`/`column_dimensions` 从插入点开始
  一个个往后挪一位（从最后一行/列开始往前处理，不然会覆盖还没读出来的旧值）。**这个坑用之前
  那种"只复制 cell.value、不保留格式"生成的测试数据是发现不了的**——本项目早期为了测试方便
  用 `ws.append(values)` 拆分出来的几份"（测试用）"参考文件本身就没有任何格式，所以这个问题
  在开发阶段的测试里完全不会暴露，是用户拿真实的、有格式的生产文件测试才发现的。教训：凡是
  会调用 `insert_rows`/`insert_cols` 的代码，测试数据必须是带真实格式的（哪怕是最简单的填色/
  边框/行高），不能图省事只测"值对不对"；现在 `tests/test_shipment_plan_apply.py` 里
  `test_shipment_summary_insert_above_preserves_formatting`/
  `test_purchase_book_insert_date_column_preserves_formatting` 这两个测试专门锁定这条，
  三份"（测试用）"参考文件也已经改成用"删除其它 sheet 再另存"的方式重新生成（保留完整格式），
  不再是只复制值的版本。
- **`QTableWidget.setHorizontalHeaderLabels()` 只接受字符串**，传非字符串（比如 Excel 表头本身
  就是 `datetime`，或者裸数字）会在 PySide6 底层报一堆
  `_pythonToCppCopy: Cannot copy-convert ... to C++` 的错误（不会崩溃，但表现出来就是各种
  数据显示不对/界面异常，不容易第一时间联想到是这个原因）。所有要显示给人看的表头/文本都要
  先转成 `str`，取具体单元格的值再用原始（没转字符串的）值当 key 去查，因为业务数据字典是拿
  原始值存的。见 `core/diff_preview.py` 的 `format_cell`。
- **openpyxl `read_only` 模式下，`ws[row_idx]` 或类似按行号取值的写法，如果 `row_idx` 超过表格
  实际行数，会直接抛 `IndexError`**（普通模式不会，会返回空行）——扫描表头等场景要用
  `min(想扫的行数, ws.max_row)` 封顶，不能无脑扫固定的行数。见 `column_utils.py` 的
  `find_header_row`。
- **公式安全求值时，字符白名单检查要放在"公式原文"上，不能放在"替换完单元格引用之后的表达式"
  上**——业务数据里的字符串值（比如订单号"GH-2501009"、型号"TD-RZ-419"）本身就带字母，如果对
  替换后的内容做字符白名单检查，会把这些合法值误判成"不安全"直接放弃求值。应该检查"公式原文
  挖掉所有单元格引用之后剩下的字符"是不是只有运算符——这样公式结构本身的安全性和引用的实际内容
  是分开检查的。见 `column_utils.py` 的 `resolve_cell_value`。同一个函数里还有一个相关教训：
  Excel 里 `=+B2` 这种开头多余的加号是合法写法（等价于 `=B2`），但 Python 对字符串做一元 `+`
  会直接报错，替换成实际值求值之前要把这种没意义的开头 `+` 去掉。
- **"同一个字段组合对应多条记录"不一定是异常，先弄清楚业务上这些记录之间是什么关系，
  不要凭直觉假设"肯定要唯一确定是哪一条"**——`shipment_summary.py` 这里踩了两次坑，是一次
  对同一个问题的过度纠正：
  1. 一开始的实现是"采购单号+型号"查到第一条待定行就用，结果同一个采购单号+型号同时有多条
     待定行时（比如一条剩 36、一条历史遗留的 0），曾经把货错误地写进了那条该是 0 的行。
  2. 第一次的"修复"矫枉过正：改成要求待定行的数量必须跟某一笔具体采购记录的未出货数量
     精确匹配，查到不唯一就报错。结果发现"同一个采购单号+型号同时有好几行待定"其实是
     **正常状态**——这些行本来就是同一批"还没决定发去哪"的库存，互相之间没有区别，
     它们的数量总和才是跟采购汇总表对应的东西，不是某一行单独对应。
  最终的正确模型：把这类"同一个 key 下有多条记录"的情况当成一个可以按顺序消耗的库存池——
  校验的是**总量**够不够，不够就报错；消耗的时候按某个稳定顺序（这里是表里的行顺序）依次扣，
  一行不够扣下一行，跳过数量本来就是 0 的行。不要假设一定能唯一定位到"这一条"，也不要假设
  "多条"就是数据错误——两种极端都可能是错的，具体哪种要靠问业务、看真实数据里这些记录之间
  到底是什么关系来确定。见 `shipment_summary.py` 的 `apply_shipment`/`pending_rows`/
  `total_pending_quantity`，以及 `purchase_book.py` 里从多笔采购订单依次分摊的同款逻辑
  （这个问题最早其实是在那边正确解决的，`shipment_summary.py` 后来才补齐同一套思路）。
- **PySide6 的 `QThread.terminate()` 在线程卡在 `concurrent.futures`/多进程等待时会直接卡死**，
  取消长任务要用协作式取消（`threading.Event`，工作线程自己检查、自己收尾），不要指望强制终止。
- Windows PowerShell 脚本：文件要用 UTF-8 BOM 编码中文注释才能正常显示，但 BOM 后面紧跟 `#`
  在某些机器上会解析失败（报"无法将#项识别为cmdlet"），文件开头留一个空行再写注释能绕开；
  `$ErrorActionPreference = "Stop"` 不会让外部命令的非零退出码变成"错误"，要显式检查
  `$LASTEXITCODE`；已装的 Python 是否能用不能只看文件存不存在，要真的执行一下（`Test-PythonWorks`
  这种模式），残留的不完整安装会导致后续步骤诡异地失败。

## 其它参考

- 需求确认 SOP、术语规范：`/Users/jingyuhuang/Documents/Work/闰科/软件开发SOP`
- `README.md`：给人看的简介
- `SETUP.md`：开发环境搭建 + Windows 部署细节
- `modules/logistics/shipment_plan_apply/` 下每个 `.py` 文件顶部的模块级 docstring 写得比较详细，
  是理解"预览+确认写入"这一整套流程最快的入口。
