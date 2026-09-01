import multiprocessing

from core.app import main

if __name__ == "__main__":
    # 以后用 PyInstaller 打包成 exe 后，物流模块用到的 ProcessPoolExecutor 必须有这一行，
    # 不然打包后的程序一启动就会不停地把自己重新拉起来。开发阶段没有这一行也不会出问题，
    # 提前加上是为了不在打包那天才踩到这个坑。
    multiprocessing.freeze_support()
    main()
