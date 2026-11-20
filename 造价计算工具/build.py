# -*- coding: utf-8 -*-
"""打包成单文件 exe：python build.py  ->  dist/造价计算工具.exe"""

import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
NAME = "造价计算工具"


def main():
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile", "--windowed",
        "--name", NAME,
        "--paths", os.path.join(BASE, "src"),
        "--distpath", os.path.join(BASE, "dist"),
        "--workpath", os.path.join(BASE, "build"),
        "--specpath", os.path.join(BASE, "build"),
        os.path.join(BASE, "src", "main.py"),
    ]
    print(" ".join(args))
    try:
        subprocess.check_call(args, cwd=BASE)
    except subprocess.CalledProcessError:
        print("\n打包失败：如果提示拒绝访问，请先关闭正在运行的 %s.exe 再重试。" % NAME)
        return 1
    out = os.path.join(BASE, "dist", NAME + ".exe")
    print("完成：%s" % out)
    if os.path.exists(out):
        print("大小：%.1f MB" % (os.path.getsize(out) / 1024.0 / 1024.0))


if __name__ == "__main__":
    main()
