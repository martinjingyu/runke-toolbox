# runke-toolbox

闰科内部业务工具，核心壳 + 按部门模块的结构：

- `core/` 通用部分——启动窗口、加载配置、统一的数据存取接口（`core/storage`）、模块加载（`core/module_registry.py`）、
  部门内"工具列表 → 点进去具体功能"的导航（`core/hub_widget.py`）、工具按需装依赖的机制（`core/dependency.py`）
- `modules/` 各部门模块，每个部门一个子目录，`hub.py` 是部门入口（轻量，只列出工具、声明各工具需要的依赖，
  不 import 具体工具用到的重量级库）。当前只有 `modules/logistics`（物流仓库），第一个工具
  `walmart_shipment_reconcile`（Walmart 发货数量核对）已经实现，界面在
  `modules/logistics/walmart_shipment_reconcile/panel.py`，业务人员**第一次点开这个工具**时才会
  检查/安装它自己的依赖（openpyxl/pymupdf/pylibdmtx/pytesseract + Tesseract），不是软件一启动就
  装好——见 [SETUP.md](SETUP.md) "各工具自己的依赖" 一节。运行 `python main.py` 就能在软件里直接用；
  命令行版本 `scripts/run_walmart_reconcile.py` 还留着，方便批处理或调试。三个输入：箱唛 PDF、翻译表
  （WM-SKU→货号）、发货计划表——匹配链路和几个容易踩的坑（发货计划表里同一 SKU+仓库横跨很多历史批次、
  要靠 SHIPMENT ID 才能定位到具体是哪一批）写在 `modules/logistics/walmart_shipment_reconcile/reconcile.py`
  顶部注释里

Windows 上想直接用（不装开发工具），双击 `setup_windows.bat` 一键装环境启动，见 [SETUP.md](SETUP.md)。
- `config.yaml` 全局配置；机器专属的覆盖项放 `config.local.yaml`（不提交 git）

数据存取现在只走本地磁盘（`core/storage/local.py`），NAS 对接接口已经在 `core/storage/base.py` / `core/storage/nas.py` 留好，
模块代码通过 `core.storage.get_storage()` 拿存储对象，以后切到 NAS 不需要改模块代码。

新增部门模块的方式：在 `modules/` 下新建一个包，`__init__.py` 里写一个 `MODULE_INFO` 字典（参考
`modules/logistics/__init__.py`），核心壳会自动识别出来，不用改 `core/` 里的代码。

需求确认流程和术语规范见 `/Users/jingyuhuang/Documents/Work/闰科/软件开发SOP`。

开发环境搭建见 [SETUP.md](SETUP.md)。
