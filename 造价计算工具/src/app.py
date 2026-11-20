# -*- coding: utf-8 -*-
"""主界面：两个页签（量增量减费用调整表 / 材料调差表）"""

import json
import os
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import sheet
from page import AdjustPage, MaterialPage
from sheet import C_APP, C_MUTED, C_TEXT, FONT

APP_TITLE = "造价计算工具"
APP_VERSION = "2.0"
DEVELOPER = "开元咨询-三部-王友强"

C_TOP = "#1F3A5F"
C_TOP_SUB = "#9FB8D4"


def pick_font(root):
    global FONT
    try:
        fams = set(tkfont.families(root))
    except Exception:
        return FONT
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimHei", "宋体"):
        if name in fams:
            FONT = name
            break
    sheet.FONT = FONT
    return FONT


class App:
    def __init__(self, root, scale=1.0):
        self.root = root
        self.s = scale
        self.file_path = None
        self.last_dir = os.path.join(os.path.expanduser("~"), "Documents")
        if not os.path.isdir(self.last_dir):
            self.last_dir = os.path.expanduser("~")
        self._style()
        self._build()

    # ------------------------------------------------------------ 构建 --
    def _build(self):
        s = self.s
        self.root.configure(bg=C_APP)

        top = tk.Frame(self.root, bg=C_TOP, height=int(56 * s))
        top.pack(side="top", fill="x")
        top.pack_propagate(False)
        tk.Label(top, text="造价计算工具", bg=C_TOP, fg="#FFFFFF",
                 font=(FONT, 15, "bold")).pack(side="left", padx=(int(18 * s), int(12 * s)))
        tk.Label(top, text="材料调差表　·　量增量减费用调整表　（黄色为参数，绿色为结果）",
                 bg=C_TOP, fg=C_TOP_SUB, font=(FONT, 9)).pack(side="left",
                                                            pady=(int(4 * s), 0))
        tk.Label(top, text="v" + APP_VERSION, bg=C_TOP, fg=C_TOP_SUB,
                 font=(FONT, 8)).pack(side="right", padx=(int(6 * s), int(18 * s)))
        tk.Label(top, text=DEVELOPER, bg=C_TOP, fg="#C8D9EC",
                 font=(FONT, 9)).pack(side="right", padx=(0, int(10 * s)))

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(side="top", fill="both", expand=True)
        self.adjust_page = AdjustPage(self.nb, self, s)
        self.material_page = MaterialPage(self.nb, self, s)
        # 页签顺序：材料调差表在前，量增量减费用调整表在后
        self.pages = [self.material_page, self.adjust_page]
        self.nb.add(self.material_page, text="  " + self.material_page.TAB_TITLE + "  ")
        self.nb.add(self.adjust_page, text="  " + self.adjust_page.TAB_TITLE + "  ")
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab_changed())

        self.root.bind("<Control-s>", lambda e: self.do_save())
        self.root.bind("<Control-S>", lambda e: self.do_save())
        self.root.bind("<Control-o>", lambda e: self.do_open())
        self.root.bind("<Control-O>", lambda e: self.do_open())
        self.root.bind("<Control-e>", lambda e: self._dispatch("do_export"))
        self.root.bind("<Control-E>", lambda e: self._dispatch("do_export"))
        self.root.bind("<Control-i>", lambda e: self._dispatch("do_import"))
        self.root.bind("<Control-I>", lambda e: self._dispatch("do_import"))
        self.root.bind("<Control-z>", lambda e: self._dispatch("do_undo"))
        self.root.bind("<Control-Z>", lambda e: self._dispatch("do_undo"))
        self.root.bind("<Control-y>", lambda e: self._dispatch("do_redo"))
        self.root.bind("<Control-Y>", lambda e: self._dispatch("do_redo"))

    def _style(self):
        s = self.s
        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        pad = (int(11 * s), int(5 * s))
        st.configure("TButton", font=(FONT, 9), padding=pad)
        st.configure("Soft.TButton", font=(FONT, 9), padding=pad,
                     background="#FFFFFF", foreground="#26405E", borderwidth=1,
                     focuscolor="#FFFFFF")
        st.map("Soft.TButton", background=[("active", "#EAF2FB")],
               foreground=[("disabled", "#A9B4C2")])
        st.configure("Accent.TButton", font=(FONT, 9, "bold"), padding=pad,
                     background="#1A73E8", foreground="#FFFFFF", borderwidth=0,
                     focuscolor="#1A73E8")
        st.map("Accent.TButton", background=[("active", "#1667D0")],
               foreground=[("active", "#FFFFFF")])
        st.configure("TScrollbar", background="#D6E0EC", troughcolor="#EDF1F7",
                     borderwidth=0, arrowsize=int(12 * s))
        st.configure("TNotebook", background=C_APP, borderwidth=0, tabmargins=(int(10 * s), int(6 * s), 0, 0))
        # 选中的页签更大、绿色高亮；未选中的小而浅
        st.configure("TNotebook.Tab", font=(FONT, 10), padding=(int(14 * s), int(5 * s)),
                     background="#E6EBF2", foreground="#6B7F95", borderwidth=0)
        st.map("TNotebook.Tab",
               background=[("selected", "#17A05E"), ("active", "#D9E6F5"),
                           ("!selected", "#E6EBF2")],
               foreground=[("selected", "#FFFFFF"), ("active", "#24405E")],
               font=[("selected", (FONT, 11, "bold"))],
               padding=[("selected", (int(24 * s), int(9 * s)))],
               expand=[("selected", (0, 0, 0, 0))])

    # ------------------------------------------------------------ 工具 --
    def current_page(self):
        try:
            idx = self.nb.index(self.nb.select())
            return self.pages[idx]
        except Exception:
            return self.pages[0]

    def _dispatch(self, method):
        page = self.current_page()
        fn = getattr(page, method, None)
        if fn:
            fn()
        return "break"

    def _on_tab_changed(self):
        page = self.current_page()
        try:
            page.compose_status_refresh()
        except AttributeError:
            pass

    # ------------------------------------------------------------ 工程 --
    def project_data(self):
        for p in self.pages:
            p.flush_edit()
        return {
            "version": 2,
            "pages": {
                "adjust": self.adjust_page.state(),
                "material": self.material_page.state(),
            },
        }

    def do_save(self):
        """有工程文件时直接保存；还没有工程文件时弹窗另存为（选择路径）。"""
        path = self.file_path
        if not path:
            cur = self.current_page()
            name = (cur.var_proj.get().strip() or "造价工程") + ".json"
            path = filedialog.asksaveasfilename(
                title="保存工程 - 请选择保存位置", defaultextension=".json",
                initialfile=name, initialdir=self.last_dir,
                filetypes=[("工程文件", "*.json")])
            if not path:
                return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(self.project_data(), fh, ensure_ascii=False, indent=1)
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self.file_path = path
        folder = os.path.dirname(path)
        if folder:
            self.last_dir = folder
        self.set_title(path)
        try:
            self.current_page().flash("已保存：%s" % os.path.basename(path), ok=True)
        except Exception:
            pass

    def do_save_as(self):
        """另存为：总是弹窗选择新的保存路径。"""
        self.file_path = None
        self.do_save()

    def do_open(self):
        path = filedialog.askopenfilename(
            title="打开工程", initialdir=self.last_dir,
            filetypes=[("工程文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))
            return
        pages = data.get("pages")
        try:
            if pages:
                self.adjust_page.load_state(pages.get("adjust", {}))
                self.material_page.load_state(pages.get("material", {}))
            else:
                # 兼容旧版工程（只有量增量减表）
                self.adjust_page.load_state(data)
        except Exception as exc:
            messagebox.showerror("打开失败", "工程文件格式不正确：\n%s" % exc)
            return
        self.file_path = path
        folder = os.path.dirname(path)
        if folder:
            self.last_dir = folder
        self.set_title(path)
        for p in self.pages:
            p.refresh()

    def set_title(self, path):
        self.root.title("%s - %s" % (APP_TITLE, os.path.basename(path)))


def run(root, scale=1.0):
    pick_font(root)
    App(root, scale=scale)
    root.mainloop()
