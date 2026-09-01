# 开发环境搭建

## Windows（全新电脑，什么都没装）

双击 **`setup_windows.bat`**：

- 没装 Python 会自动从官网下载安装 3.11.9（只装给当前用户，不需要管理员权限）
- 建一个项目专属的虚拟环境 `.venv`（不影响系统里其它 Python 项目）
- 自动装好 `requirements.txt` 里的依赖
- 装完直接启动软件

第一次运行如果卡在"Python 应该是装好了，但这个窗口还没认到"——关掉窗口重新打开一个，再双击一次
`setup_windows.bat` 就行（Windows 装完新程序后，已经打开的旧窗口有时候刷新不到最新的环境变量）。

装好之后，以后要打开软件，双击 **`run_windows.bat`** 就行，不用每次都重新装环境。

条码解码用到的 `pylibdmtx` 在 Windows 上不需要额外装东西——PyPI 上有专门给 Windows 打包的版本，
把需要的 dll 直接放在包里了（跟 Mac 不一样，Mac 上才需要额外 `brew install libdmtx`，见下面）。

## Mac / Linux（开发用）

用 conda 建一个专属环境，不要装进 base，避免以后别的项目互相打架。

```bash
conda env create -f environment.yml
conda activate runke-toolbox
```

以后加了新依赖：改完 `environment.yml`（同时同步一份到 `requirements.txt`，给 Windows 上打包用），然后：

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
`Unable to find dmtx shared library`。Windows 上打包时 pylibdmtx 会自带 dmtx.dll，不需要这一步，
到时候在打包环节确认一下就行。

## OCR（Walmart 发货数量核对这个功能必须装，不是可选的）

条码里没有编码 SKU 文字（只有 GTIN），而这次核对用的几份表都不含 GTIN，只能靠 SKU 文字互相
对应，所以要从箱唛印刷文字里 OCR 出 SKU（每个 GTIN 只 OCR 一次代表页，不是每页都读）。这个功能
离不开 Tesseract：

```bash
brew install tesseract
```

没装的话，`run()` 一开始就会报错说清楚原因（不会跑到一半才发现所有 SKU 都是空的）。Windows 上装
[Tesseract 的 Windows 安装包](https://github.com/UB-Mannheim/tesseract/wiki)；`setup_windows.ps1`
目前没有自动装这个，先手动装，之后有需要再把这一步也自动化。

## Windows 上给同事用

现在这套是给开发用的环境，同事那边最终应该是拿一个用 PyInstaller 打包好的 exe，双击就能跑，不需要装 Python 或
conda——这个打包流程还没做，等框架和第一个模块都跑通了再弄。
