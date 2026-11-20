# -*- coding: utf-8 -*-
"""AI 识图窗口：选图纸（PDF/图片）→ 识别 → 预览结果 → 追加到雨污水表。"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ai_model
import ai_vision
from sheet import C_APP, C_TEXT, FONT

C_TOP = "#1F3A5F"
C_OK = "#17A05E"
C_ERR = "#C0392B"
C_HINT = "#5B6B7C"

FILE_TYPES = [
    ("图纸文件", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff"),
    ("PDF 图纸", "*.pdf"),
    ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"),
    ("所有文件", "*.*"),
]


class AIVisionDialog(tk.Toplevel):
    def __init__(self, master, app, on_apply=None, scale=1.0):
        tk.Toplevel.__init__(self, master)
        self.s = scale or 1.0
        self.app = app
        self.on_apply = on_apply
        self.rows = []
        self.path = ""
        self._busy = False
        self._photo = None

        self.title("AI 识图")
        self.configure(bg=C_APP)
        try:
            self.transient(master)
        except Exception:
            pass
        self.resizable(False, False)
        self._build()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ 布局 --
    def _center(self):
        """固定 1215×510，屏幕居中。"""
        width, height = 1215, 510
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)
            self.geometry("%dx%d+%d+%d" % (width, height, x, y))
        except Exception:
            self.geometry("%dx%d" % (width, height))

    def _build(self):
        s = self.s
        head = tk.Frame(self, bg=C_TOP, height=int(52 * s))
        head.pack(side="top", fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="AI 识图", bg=C_TOP, fg="#FFFFFF",
                 font=(FONT, 12, "bold")).pack(side="left", padx=int(16 * s),
                                               pady=int(12 * s))
        tk.Label(head, text="识别平面图 / 纵断面图，自动提取桩号、井号、高程、管径等",
                 bg=C_TOP, fg="#9FB8D4", font=(FONT, 9)).pack(side="left",
                                                              pady=(int(4 * s), 0))

        top = tk.Frame(self, bg=C_APP)
        top.pack(side="top", fill="x", padx=int(16 * s), pady=(int(10 * s), int(6 * s)))
        ttk.Button(top, text="选择图纸…", style="Accent.TButton",
                   command=self._pick_file).pack(side="left")
        self.var_file = tk.StringVar(value="尚未选择文件")
        tk.Label(top, textvariable=self.var_file, bg=C_APP, fg=C_TEXT, font=(FONT, 9),
                 anchor="w").pack(side="left", padx=(int(10 * s), 0), fill="x", expand=True)

        body = tk.Frame(self, bg=C_APP)
        body.pack(side="top", fill="both", expand=True,
                  padx=int(16 * s), pady=(0, int(6 * s)))

        left = tk.Frame(body, bg=C_APP, width=250)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        tk.Label(left, text="图纸预览", bg=C_APP, fg=C_HINT,
                 font=(FONT, 8), anchor="w").pack(side="top", fill="x")
        self.lbl_preview = tk.Label(left, text="（选择图纸后显示第一页）", bg="#FFFFFF",
                                    fg=C_HINT, font=(FONT, 9), relief="solid", bd=1,
                                    anchor="center", justify="center")
        self.lbl_preview.pack(side="top", fill="both", expand=True, pady=(int(4 * s), 0))

        right = tk.Frame(body, bg=C_APP)
        right.pack(side="left", fill="both", expand=True, padx=(int(12 * s), 0))
        tk.Label(right, text="识别结果（可先核对，再追加到表格）", bg=C_APP, fg=C_HINT,
                 font=(FONT, 8), anchor="w").pack(side="top", fill="x")

        wrap = tk.Frame(right, bg=C_APP)
        wrap.pack(side="top", fill="both", expand=True, pady=(int(4 * s), 0))
        keys = [key for key, _title in ai_vision.FIELDS]
        self.tree = ttk.Treeview(wrap, columns=keys, show="headings", height=9)
        for key, title in ai_vision.FIELDS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=92, anchor="center", stretch=True)
        bar = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")

        hint_box = tk.Frame(self, bg=C_APP)
        hint_box.pack(side="top", fill="x", padx=int(16 * s))
        tk.Label(hint_box, text="补充要求", bg=C_APP, fg=C_TEXT,
                 font=(FONT, 9)).pack(side="left")
        self.var_hint = tk.StringVar()
        tk.Entry(hint_box, textvariable=self.var_hint, font=(FONT, 9),
                 relief="solid", bd=1).pack(side="left", fill="x", expand=True,
                                            padx=(int(8 * s), 0))

        foot = tk.Frame(self, bg=C_APP)
        foot.pack(side="bottom", fill="x", padx=int(16 * s), pady=int(12 * s))
        self.btn_run = ttk.Button(foot, text="开始识别", style="Accent.TButton",
                                  command=self._start)
        self.btn_run.pack(side="left")
        self.btn_apply = ttk.Button(foot, text="追加到表格", command=self._apply,
                                    state="disabled")
        self.btn_apply.pack(side="left", padx=(int(8 * s), 0))
        ttk.Button(foot, text="关闭", command=self._on_close).pack(side="right")

        self.lbl_status = tk.Label(self, text="选择图纸后点「开始识别」；识别结果不会自动改表，确认后再追加。",
                                   bg=C_APP, fg=C_HINT, font=(FONT, 9), anchor="w",
                                   justify="left")
        self.lbl_status.pack(side="bottom", fill="x", padx=int(16 * s),
                             pady=(0, int(6 * s)))

        self.tree.bind("<<TreeviewSelect>>", lambda e: None)

    # ------------------------------------------------------------ 交互 --
    def _set_status(self, text, color=C_HINT):
        try:
            self.lbl_status.configure(text=text, fg=color)
        except Exception:
            pass

    def _pick_file(self):
        path = filedialog.askopenfilename(title="选择图纸", filetypes=FILE_TYPES)
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in ai_vision.SUPPORTED_EXTS:
            messagebox.showwarning("格式不支持",
                                   "目前支持 PDF 和常见图片（png/jpg/bmp/tif）。\n"
                                   "DWG/DXF 请先打印或导出为 PDF。")
            return
        self.path = path
        self.var_file.set(path)
        self._show_preview(path)
        self._set_status("已选择：%s，点「开始识别」" % os.path.basename(path))

    def _show_preview(self, path):
        try:
            from PIL import Image, ImageTk
        except Exception:
            self.lbl_preview.configure(text="（缺少 Pillow，无法预览）")
            return
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".pdf":
                import fitz
                doc = fitz.open(path)
                try:
                    if len(doc) == 0:
                        raise RuntimeError("PDF 没有页面")
                    pix = doc[0].get_pixmap(dpi=90)
                    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                finally:
                    doc.close()
            else:
                image = Image.open(path).convert("RGB")
            image.thumbnail((230, 300))
            self._photo = ImageTk.PhotoImage(image)
            self.lbl_preview.configure(image=self._photo, text="")
        except Exception as exc:
            self.lbl_preview.configure(text="预览失败：%s" % exc)

    def _start(self):
        if self._busy:
            return
        if not self.path:
            messagebox.showinfo("请先选择图纸", "点「选择图纸…」挑一个 PDF 或图片。")
            return
        profile = ai_model.active_profile(ai_model.load_config())
        if not ai_model.is_vision_model(profile.get("model")):
            messagebox.showwarning(
                "当前模型不能看图",
                "请先到顶部「AI模型」里换成带 VL / vision 的模型，例如 qwen-vl-max、glm-4v-plus、gpt-4o。")
            return

        self._busy = True
        self.btn_run.configure(state="disabled")
        self.btn_apply.configure(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.rows = []
        hint = self.var_hint.get().strip()
        self._set_status("正在识别…")

        def progress(page, total):
            self._safe_after(lambda: self._set_status(
                "正在识别第 %d / %d 页…（大图通常需要 30～120 秒，请勿关闭窗口）"
                % (page, total)))

        def worker():
            try:
                rows = ai_vision.ocr_drawing(profile, self.path, hint=hint,
                                             on_progress=progress)
            except Exception as exc:
                self._safe_after(lambda: self._finish_error(str(exc)))
            else:
                self._safe_after(lambda: self._finish_ok(rows))

        threading.Thread(target=worker, daemon=True).start()

    def _safe_after(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass

    def _finish_error(self, message):
        self._busy = False
        try:
            self.btn_run.configure(state="normal")
        except Exception:
            pass
        self._set_status("识别失败：%s" % self._friendly_error(message), C_ERR)

    @staticmethod
    def _friendly_error(message):
        """把接口原始报错翻译成能照着做的提示。"""
        text = message or ""
        low = text.lower()
        if ("model disabled" in low or "model not found" in low
                or "does not exist" in low or "invalid model" in low):
            return ("当前模型在平台上不可用（已下线或未开通）。请到顶部「AI模型」换可识图的模型，"
                    "如 Qwen/Qwen3-VL-30B-A3B-Instruct、qwen3-vl-plus、glm-4v-plus、gpt-4o")
        if "401" in text or "invalid api key" in low or "unauthorized" in low:
            return "API Key 无效或已过期，请到「AI模型」重新填写并保存"
        if "403" in text:
            return "无权限调用该模型（未开通 / 未实名 / 额度用尽）：%s" % text
        if "timed out" in low or "timeout" in low or "超时" in text:
            return ("请求超时（大图识别慢）。建议换更快的模型，如 Qwen/Qwen3-VL-30B-A3B-Instruct，"
                    "或把图纸按桩号裁成小块再识别")
        return text

    def _finish_ok(self, rows):
        self._busy = False
        try:
            self.btn_run.configure(state="normal")
        except Exception:
            pass
        self.rows = rows
        for row in rows:
            self.tree.insert("", "end", values=[row.get(k, "") for k, _t in ai_vision.FIELDS])
        if not rows:
            self._set_status("没有识别到表格行；可换清晰的图纸或加补充要求再试", C_ERR)
            return
        try:
            self.btn_apply.configure(state="normal")
        except Exception:
            pass
        self._set_status("识别到 %d 行，核对无误后点「追加到表格」" % len(rows), C_OK)

    def _apply(self):
        if not self.rows:
            return
        if not self.on_apply:
            return
        try:
            self.on_apply(self.rows)
        except Exception as exc:
            messagebox.showerror("追加失败", str(exc))
            return
        self.destroy()

    def _on_close(self):
        self.destroy()
