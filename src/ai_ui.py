# -*- coding: utf-8 -*-
"""AI 模型配置窗口：选择引擎、填接口/密钥/模型，支持读取模型列表与测试连接。"""

import threading
import tkinter as tk
from tkinter import ttk

import ai_model
from sheet import C_APP, C_TEXT, FONT

C_TOP = "#1F3A5F"
C_OK = "#17A05E"
C_ERR = "#C0392B"
C_HINT = "#5B6B7C"


class AIModelDialog(tk.Toplevel):
    def __init__(self, master, app, scale=1.0, on_saved=None):
        tk.Toplevel.__init__(self, master)
        self.s = scale or 1.0
        self.app = app
        self.on_saved = on_saved
        self.cfg = ai_model.load_config()
        self._engine = self.cfg["engine"]
        self._busy = False

        self.title("AI 模型")
        self.configure(bg=C_APP)
        try:
            self.transient(master)
        except Exception:
            pass
        self.resizable(False, False)
        self._center(int(560 * self.s), int(470 * self.s))

        self._build()
        self._load_engine(self._engine)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    # ------------------------------------------------------------ 布局 --
    def _center(self, width, height):
        try:
            self.update_idletasks()
            x = self.master.winfo_rootx() + (self.master.winfo_width() - width) // 2
            y = self.master.winfo_rooty() + (self.master.winfo_height() - height) // 3
            self.geometry("%dx%d+%d+%d" % (max(0, x), max(0, y), width, height))
        except Exception:
            self.geometry("%dx%d" % (width, height))

    def _build(self):
        s = self.s
        head = tk.Frame(self, bg=C_TOP, height=int(52 * s))
        head.pack(side="top", fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="AI 模型", bg=C_TOP, fg="#FFFFFF",
                 font=(FONT, 12, "bold")).pack(side="left", padx=int(16 * s), pady=int(12 * s))
        tk.Label(head, text="供 AI 服务调用，密钥只存本机", bg=C_TOP, fg="#9FB8D4",
                 font=(FONT, 9)).pack(side="left", pady=(int(4 * s), 0))

        body = tk.Frame(self, bg=C_APP)
        body.pack(side="top", fill="both", expand=True, padx=int(18 * s), pady=int(14 * s))

        self.var_engine = tk.StringVar(value=self._engine)
        self.var_model = tk.StringVar()
        self.var_endpoint = tk.StringVar()
        self.var_key = tk.StringVar()
        self.var_temp = tk.DoubleVar(value=ai_model.DEFAULT_TEMPERATURE)
        self.var_timeout = tk.StringVar(value=str(ai_model.DEFAULT_TIMEOUT))
        self.var_show = tk.BooleanVar(value=False)

        row = 0
        self._label(body, "引擎").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.cb_engine = ttk.Combobox(body, textvariable=self.var_engine, state="readonly",
                                      values=ai_model.ENGINES, width=18,
                                      font=(FONT, 9))
        self.cb_engine.grid(row=row, column=1, columnspan=2, sticky="we", pady=int(6 * s))
        self.cb_engine.bind("<<ComboboxSelected>>", self._on_engine_change)

        row += 1
        self._label(body, "模型").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.cb_model = ttk.Combobox(body, textvariable=self.var_model, width=28,
                                     font=(FONT, 9))
        self.cb_model.grid(row=row, column=1, sticky="we", pady=int(6 * s))
        self.btn_models = ttk.Button(body, text="读取列表", width=9,
                                     command=self._on_list_models)
        self.btn_models.grid(row=row, column=2, sticky="w", padx=(int(8 * s), 0))

        row += 1
        self._label(body, "地址").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        tk.Entry(body, textvariable=self.var_endpoint, font=(FONT, 9),
                 relief="solid", bd=1).grid(row=row, column=1, columnspan=2,
                                            sticky="we", pady=int(6 * s))

        row += 1
        self._label(body, "密钥").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.ent_key = tk.Entry(body, textvariable=self.var_key, font=(FONT, 9),
                                relief="solid", bd=1, show="*")
        self.ent_key.grid(row=row, column=1, sticky="we", pady=int(6 * s))
        tk.Checkbutton(body, text="显示", variable=self.var_show, command=self._toggle_key,
                       bg=C_APP, fg=C_TEXT, activebackground=C_APP,
                       font=(FONT, 8), bd=0, highlightthickness=0).grid(
            row=row, column=2, sticky="w", padx=(int(8 * s), 0))

        row += 1
        self._label(body, "温度").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.sc_temp = ttk.Scale(body, from_=0.0, to=2.0, variable=self.var_temp,
                                 command=self._on_temp, orient="horizontal")
        self.sc_temp.grid(row=row, column=1, sticky="we", pady=int(6 * s))
        self.lbl_temp = tk.Label(body, text="0.3", bg=C_APP, fg=C_HINT, font=(FONT, 8),
                                 width=4, anchor="w")
        self.lbl_temp.grid(row=row, column=2, sticky="w", padx=(int(8 * s), 0))

        row += 1
        self._label(body, "超时").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        tk.Spinbox(body, from_=5, to=300, increment=5, textvariable=self.var_timeout,
                   width=6, font=(FONT, 9), relief="solid", bd=1).grid(
            row=row, column=1, sticky="w", pady=int(6 * s))
        tk.Label(body, text="秒", bg=C_APP, fg=C_HINT, font=(FONT, 8)).grid(
            row=row, column=2, sticky="w", padx=(int(8 * s), 0))

        row += 1
        self.lbl_status = tk.Label(body, text="", bg=C_APP, fg=C_HINT, font=(FONT, 9),
                                   anchor="w", justify="left", wraplength=int(470 * s))
        self.lbl_status.grid(row=row, column=0, columnspan=3, sticky="we",
                             pady=(int(8 * s), 0))

        foot = tk.Frame(self, bg=C_APP)
        foot.pack(side="bottom", fill="x", padx=int(18 * s), pady=(0, int(14 * s)))
        self.btn_test = ttk.Button(foot, text="测试连接", width=10,
                                   command=self._on_test)
        self.btn_test.pack(side="left")
        ttk.Button(foot, text="保存", style="Accent.TButton", width=10,
                   command=self._on_save).pack(side="right")
        ttk.Button(foot, text="取消", width=8,
                   command=self.destroy).pack(side="right", padx=(0, int(8 * s)))

        body.columnconfigure(1, weight=1)
        self.var_temp.trace_add("write", lambda *a: self._sync_temp_label())
        self._sync_temp_label()

    def _label(self, parent, text):
        return tk.Label(parent, text=text, bg=C_APP, fg=C_TEXT, font=(FONT, 9))

    # ------------------------------------------------------------ 交互 --
    def _sync_temp_label(self):
        try:
            self.lbl_temp.configure(text="%.1f" % float(self.var_temp.get()))
        except Exception:
            pass

    def _on_temp(self, value):
        try:
            self.lbl_temp.configure(text="%.1f" % float(value))
        except Exception:
            pass

    def _toggle_key(self):
        self.ent_key.configure(show="" if self.var_show.get() else "*")

    def _set_status(self, text, color=C_HINT):
        try:
            self.lbl_status.configure(text=text, fg=color)
        except Exception:
            pass

    def _collect_profile(self):
        try:
            timeout = int(float(self.var_timeout.get()))
        except Exception:
            timeout = ai_model.DEFAULT_TIMEOUT
        timeout = max(5, min(300, timeout))
        return {
            "endpoint": ai_model.normalize_endpoint(self.var_endpoint.get()),
            "model": self.var_model.get().strip(),
            "key": self.var_key.get().strip(),
            "temperature": round(float(self.var_temp.get()), 2),
            "timeout": timeout,
        }

    def _load_engine(self, engine):
        profile = self.cfg["profiles"].get(engine) or ai_model.default_profile(engine)
        self._engine = engine
        self.var_engine.set(engine)
        self.var_endpoint.set(profile.get("endpoint") or "")
        self.var_model.set(profile.get("model") or "")
        self.var_key.set(profile.get("key") or "")
        _name, _endpoint, models = ai_model.provider_preset(engine)
        self.cb_model.configure(values=models)
        self._set_status("")

    def _on_engine_change(self, _event=None):
        new = self.var_engine.get()
        if new == self._engine:
            return
        self.cfg["profiles"][self._engine] = self._collect_profile()
        self._load_engine(new)

    # ------------------------------------------------------------ 网络 --
    def _run_async(self, work, on_ok, busy_text):
        if self._busy:
            return
        self._busy = True
        self._set_status(busy_text)
        self.btn_models.configure(state="disabled")
        self.btn_test.configure(state="disabled")

        def worker():
            try:
                result = work()
            except Exception as exc:
                self._safe_after(lambda: self._finish(None, str(exc), C_ERR))
            else:
                self._safe_after(lambda: self._finish(on_ok, result, C_OK))

        threading.Thread(target=worker, daemon=True).start()

    def _safe_after(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass

    def _finish(self, handler, value, color):
        self._busy = False
        try:
            self.btn_models.configure(state="normal")
            self.btn_test.configure(state="normal")
        except Exception:
            pass
        if color == C_ERR:
            self._set_status(value, C_ERR)
        else:
            self._set_status(handler(value) if handler else value, C_OK)

    def _on_list_models(self):
        profile = self._collect_profile()

        def work():
            return ai_model.list_models(profile)

        def ok(names):
            if not names:
                return "接口未返回模型列表，可手动填写"
            self.cb_model.configure(values=names)
            if self.var_model.get().strip() not in names:
                self.var_model.set(names[0])
            return "读取到 %d 个模型" % len(names)

        self._run_async(work, ok, "正在读取模型列表…")

    def _on_test(self):
        profile = self._collect_profile()

        def ok(text):
            return "连接成功：%s" % text

        self._run_async(lambda: ai_model.test_connection(profile), ok, "正在测试连接…")

    # ------------------------------------------------------------ 保存 --
    def _on_save(self):
        self.cfg["profiles"][self._engine] = self._collect_profile()
        self.cfg["engine"] = self._engine
        self.cfg["temperature"] = round(float(self.var_temp.get()), 2)
        try:
            self.cfg["timeout"] = max(5, min(300, int(float(self.var_timeout.get()))))
        except Exception:
            self.cfg["timeout"] = ai_model.DEFAULT_TIMEOUT
        if not ai_model.save_config(self.cfg):
            self._set_status("保存失败：无法写入配置文件", C_ERR)
            return
        if self.on_saved:
            try:
                self.on_saved(ai_model.load_config())
            except Exception:
                pass
        self.destroy()
