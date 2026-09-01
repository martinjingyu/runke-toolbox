# 开发环境搭建

## Windows（全新电脑，什么都没装）

双击 **`setup_windows.bat`**：

- 没装 Python 会自动从官网下载安装 3.11.9（只装给当前用户，不需要管理员权限）
- 建一个项目专属的虚拟环境 `.venv`（不影响系统里其它 Python 项目）
- 自动装好 `requirements.txt` 里的依赖——**只有核心壳需要的东西**（PySide6、PyYAML），几秒钟装完
- 装完直接启动软件

装好之后，以后要打开软件：
- 双击 **`run_windows.vbs`** —— 日常推荐用这个，不会弹出黑色命令行窗口
- 双击 **`run_windows.bat`** —— 会弹一下命令行窗口，调试的时候想看报错信息可以用这个

两个效果一样，只是要不要看到那个黑窗口的区别。`.bat` 之所以会弹窗口，是因为双击 `.bat`
这件事本身就是 Windows 的 Explorer 拉起 `cmd.exe` 去解释这个批处理文件，跟脚本内容无关，
没法从 `.bat` 内部把这一下"闪一下"完全消掉；`.vbs` 是通过 `wscript.exe`（本身没有窗口）
以隐藏模式拉起程序，才能做到全程没有任何窗口。

## 想重新测试一遍"全新电脑"的安装流程

双击 **`uninstall_windows.bat`**：删掉 `.venv`、卸载本机装的 Python 和 Tesseract（如果是这个
项目的脚本装的）、把对应的目录从 PATH 里清掉——相当于把这台机器恢复到 `setup_windows.bat`
第一次运行之前的状态，不用真的找一台新电脑就能反复测试安装流程。会弹出一个确认提示，输入
`y` 才会真的动手；Python/Tesseract 卸载这两步可能会跳出系统自己的确认/进度窗口，正常点掉就行。

## 各工具自己的依赖，是点开那个工具才装的，不是 setup 时候一次装齐

`setup_windows.ps1` 只管核心壳（启动窗口、部门列表这些）。具体某个部门的某个工具要用到什么
（比如"发货数量核对"要读 Excel、要做 OCR），是在业务人员**第一次点开那个工具**的时候，软件
自己检查装了没有，没装会弹窗问"现在安装吗"，同意了就在后台装（带进度提示），装完才真正打开
界面——不会因为软件里加了新工具，所有人的电脑上都被迫多装一堆自己用不上的东西。

这套机制在 `core/dependency.py`（工具怎么声明"我需要什么"）和 `core/hub_widget.py`（点开工具
之前先检查/装）里；某个工具具体需要什么，写在这个工具自己的模块里（比如
`modules/logistics/hub.py` 里"发货数量核对"这一项的 `dependencies` 列表）。

## Mac / Linux（开发用）

用 conda 建一个专属环境，不要装进 base，避免以后别的项目互相打架。这个环境跟上面 Windows
用户装的东西不一样——开发/测试要能跑通所有工具，所以 `environment.yml` 比 `requirements.txt`
多装了各工具自己的依赖（详见 `environment.yml` 里的注释）。

```bash
conda env create -f environment.yml
conda activate runke-toolbox
```

以后核心壳加了新依赖：改完 `environment.yml`，同时同步一份到 `requirements.txt`；如果是某个
工具自己的新依赖，只改 `environment.yml`（开发要用）和那个工具自己声明依赖的地方（比如
`modules/logistics/hub.py`），不要加进 `requirements.txt`。然后：

```bash
conda env update -f environment.yml --prune
```

运行框架：

```bash
python main.py
```

跑测试：

```bash
pytest
```

## 条码解码需要的系统依赖（Mac 开发环境）

箱唛上的 GTIN / QUANTITY / SHIPMENT ID 是解 Data Matrix 二维码拿的，不用 OCR，更准。`pylibdmtx`
这个 pip 包本身在 Mac 上还依赖一个系统层面的库：

```bash
brew install libdmtx
```

装完之后 `environment.yml` 里的 `pylibdmtx` 才能真正解出码，不然 import 就会报
`Unable to find dmtx shared library`。Windows 上 pylibdmtx 自带 dmtx.dll，不需要这一步。

## OCR（Mac 开发环境；Windows 端是点开"发货数量核对"时自动装）

条码里没有编码 SKU 文字（只有 GTIN），而这次核对用的几份表都不含 GTIN，只能靠 SKU 文字互相
对应，所以要从箱唛印刷文字里 OCR 出 SKU（每个 GTIN 只 OCR 一次代表页，不是每页都读）。这个功能
离不开 Tesseract：

```bash
brew install tesseract
```

没装的话，`run()` 一开始就会报错说清楚原因（不会跑到一半才发现所有 SKU 都是空的——这个坑
真实踩过一次：pytesseract 这个 Python 包装了，但 Tesseract 引擎本身没装，OCR 每次静默失败，
整批核对结果全部"查不到"，界面看起来像是正常跑完了，其实数据是空的）。

Windows 上不用手动装——业务人员第一次点开"发货数量核对"时，软件会自动检测、下载、安装
（用的是 [UB-Mannheim 提供的安装包](https://github.com/UB-Mannheim/tesseract/wiki)，可能会
跳出 Windows 的"是否允许此应用对你的设备进行更改"提示，点"是"）。

## Windows 上给同事用

现在这套是给开发用的环境，同事那边最终应该是拿一个用 PyInstaller 打包好的 exe，双击就能跑，不需要装 Python 或
conda——这个打包流程还没做，等框架和第一个模块都跑通了再弄。
