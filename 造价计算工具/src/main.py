# -*- coding: utf-8 -*-
"""程序入口"""

import os
import sys
import traceback

import tkinter as tk


def dpi_scale():
    """开启 DPI 感知并返回屏幕缩放系数（96dpi = 1.0）。"""
    if sys.platform != "win32":
        return 1.0
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            return 1.0
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return max(1.0, min(3.0, float(dpi) / 96.0))
    except Exception:
        return 1.0


def center(root, width, height):
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    w = min(width, sw - 40)
    h = min(height, sh - 80)
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2 - 20)
    root.geometry("%dx%d+%d+%d" % (int(w), int(h), int(x), int(y)))


def main():
    scale = dpi_scale()
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 96.0 / 72.0 * scale)
    except Exception:
        pass
    root.title("造价计算工具 - 材料调差表 / 量增量减费用调整表")
    root.minsize(int(900 * scale), int(560 * scale))
    center(root, int(1440 * scale), int(800 * scale))

    import app
    app.run(root, scale=scale)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        try:
            base = os.path.dirname(os.path.abspath(sys.argv[0]))
            with open(os.path.join(base, "error.log"), "w", encoding="utf-8") as fh:
                fh.write(err)
        except Exception:
            pass
        try:
            import tkinter.messagebox as mb
            mb.showerror("程序出错", err[-1500:])
        except Exception:
            pass
