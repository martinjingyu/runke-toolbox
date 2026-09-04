# runke-toolbox

闰科内部业务工具，核心壳 + 按部门模块的结构。**接手开发的 code agent 先看 [CLAUDE.md](CLAUDE.md)**
——架构、设计原则、可复用模板组件、踩过的坑都写在那里，比这份 README 详细得多。

- `core/` 通用部分——启动窗口、加载配置、统一的数据存取接口（`core/storage`）、模块加载（`core/module_registry.py`）、
  部门内"工具列表 → 点进去具体功能"的导航（`core/hub_widget.py`）、工具按需装依赖的机制（`core/dependency.py`）、
  "预览改动再确认写入"的可复用组件（`core/diff_preview.py`）、写入前自动备份（`core/backup.py`）
- `modules/` 各部门模块，每个部门一个子目录，`hub.py` 是部门入口（轻量，只列出工具、声明各工具需要的依赖，
  不 import 具体工具用到的重量级库）。当前只有 `modules/logistics`（物流仓库），四个工具：
  - `walmart_shipment_reconcile`（发货数量核对）：核对箱唛实际发货数量和发货计划表是否一致
  - `fba_label_redact`（FBA 标签发货人信息脱敏）：批量去掉箱唛 PDF 的发货人信息
  - `shipment_plan_apply`（发货计划自动更新）：把运营提交的发货计划表导入，自动更新采购订单汇总表
    和发货计划汇总表——涉及真实业务数据的修改，预览+人工确认+自动备份，是目前最复杂的一个工具，
    也是 `core/diff_preview.py` 这个模板组件的来源
  - `logistics_tracking`（物流跟踪自动更新）：按物流商分组，并行去各货代平台查询运单最后路由，
    自动更新物流跟踪表格；登录账号密码在界面里维护，只存本机（`QSettings`），不进 `core/storage`/git

  每个工具的依赖都是**业务人员第一次点开这个工具**时才检查/安装，不是软件一启动就全装好——见
  [SETUP.md](SETUP.md) "各工具自己的依赖" 一节。运行 `python main.py` 就能在软件里直接用。

Windows 上想直接用（不装开发工具），双击 `setup_windows.bat` 一键装环境启动，见 [SETUP.md](SETUP.md)。
- `config.yaml` 全局配置；机器专属的覆盖项放 `config.local.yaml`（不提交 git）

数据存取现在只走本地磁盘（`core/storage/local.py`），NAS 对接接口已经在 `core/storage/base.py` / `core/storage/nas.py` 留好，
模块代码通过 `core.storage.get_storage()` 拿存储对象，以后切到 NAS 不需要改模块代码。

新增部门模块的方式：在 `modules/` 下新建一个包，`__init__.py` 里写一个 `MODULE_INFO` 字典（参考
`modules/logistics/__init__.py`），核心壳会自动识别出来，不用改 `core/` 里的代码。

需求确认流程和术语规范见 `/Users/jingyuhuang/Documents/Work/闰科/软件开发SOP`。

开发环境搭建见 [SETUP.md](SETUP.md)。
