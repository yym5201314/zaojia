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
C_WARN = "#B26A00"

# 回答风格（对应请求里的 temperature，不直接暴露数字）
STYLE_OPTIONS = [
    ("strict", "严谨", 0.2, "严谨：按资料如实作答，不加工结论，适合核对数据、算量取值"),
    ("balanced", "均衡", 0.7, "均衡：正常问答、适度归纳，适合一般咨询和写说明"),
    ("creative", "灵活", 1.2, "灵活：允许扩写、润色、多方案比较，引用具体数字要自己核对"),
]


def style_key_of(value):
    """把已保存的 temperature 吸附到最接近的风格档位。"""
    try:
        number = float(value)
    except Exception:
        number = 0.3
    best = STYLE_OPTIONS[0]
    for item in STYLE_OPTIONS:
        if abs(item[2] - number) < abs(best[2] - number):
            best = item
    return best[0]


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

        self._build()
        self._load_engine(self._engine)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())

    # ------------------------------------------------------------ 布局 --
    def _center(self):
        try:
            self.update_idletasks()
            width = max(int(560 * self.s), self.winfo_reqwidth())
            height = self.winfo_reqheight() + int(6 * self.s)
            x = self.master.winfo_rootx() + (self.master.winfo_width() - width) // 2
            y = self.master.winfo_rooty() + (self.master.winfo_height() - height) // 3
            self.geometry("%dx%d+%d+%d" % (width, height, max(0, x), max(0, y)))
        except Exception:
            pass

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
        self.var_style = tk.StringVar(value=STYLE_OPTIONS[0][0])
        self.var_timeout = tk.StringVar(value=str(ai_model.DEFAULT_TIMEOUT))
        self.var_show = tk.BooleanVar(value=False)

        # 顺序：引擎 → 地址 → 密钥 → 读取模型并选择 → 温度 → 超时
        row = 0
        self._label(body, "引擎").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.cb_engine = ttk.Combobox(body, textvariable=self.var_engine, state="readonly",
                                      values=ai_model.ENGINES, width=18,
                                      font=(FONT, 9))
        self.cb_engine.grid(row=row, column=1, columnspan=2, sticky="we", pady=int(6 * s))
        self.cb_engine.bind("<<ComboboxSelected>>", self._on_engine_change)

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
        self._label(body, "模型").grid(row=row, column=0, sticky="e",
                                      padx=(0, int(10 * s)), pady=int(6 * s))
        self.cb_model = ttk.Combobox(body, textvariable=self.var_model, width=28,
                                     font=(FONT, 9))
        self.cb_model.grid(row=row, column=1, sticky="we", pady=int(6 * s))
        self.btn_models = ttk.Button(body, text="读取列表", width=9,
                                     command=self._on_list_models)
        self.btn_models.grid(row=row, column=2, sticky="w", padx=(int(8 * s), 0))

        row += 1
        self.lbl_vision = tk.Label(body, text="", bg=C_APP, fg=C_HINT, font=(FONT, 8),
                                   anchor="w", justify="left", wraplength=int(500 * s))
        self.lbl_vision.grid(row=row, column=0, columnspan=3, sticky="we")

        row += 1
        self._label(body, "回答风格").grid(row=row, column=0, sticky="e",
                                        padx=(0, int(10 * s)), pady=int(6 * s))
        style_box = tk.Frame(body, bg=C_APP)
        style_box.grid(row=row, column=1, columnspan=2, sticky="w", pady=int(4 * s))
        for key, label, _value, _desc in STYLE_OPTIONS:
            tk.Radiobutton(style_box, text=label, value=key, variable=self.var_style,
                           command=self._on_style, bg=C_APP, fg=C_TEXT,
                           activebackground=C_APP, selectcolor="#FFFFFF",
                           font=(FONT, 9), bd=0, highlightthickness=0).pack(
                side="left", padx=(0, int(16 * s)))

        row += 1
        self.lbl_style = tk.Label(body, text="", bg=C_APP, fg=C_HINT, font=(FONT, 8),
                                  anchor="w", justify="left", wraplength=int(500 * s))
        self.lbl_style.grid(row=row, column=0, columnspan=3, sticky="we")

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
                                   anchor="w", justify="left", wraplength=int(500 * s))
        self.lbl_status.grid(row=row, column=0, columnspan=3, sticky="we",
                             pady=(int(8 * s), 0))

        foot = tk.Frame(self, bg=C_APP)
        foot.pack(side="bottom", fill="x", padx=int(18 * s), pady=(0, int(14 * s)))
        self.btn_test = ttk.Button(foot, text="测试连接", width=10,
                                   command=self._on_test)
        self.btn_test.pack(side="left")
        ttk.Button(foot, text="保存", style="Accent.TButton", width=10,
                   command=self._on_save).pack(side="right")
        ttk.Button(foot, text="关闭", width=8,
                   command=self._on_close).pack(side="right", padx=(0, int(8 * s)))

        body.columnconfigure(1, weight=1)
        self._on_style()
        self.var_model.trace_add("write", lambda *a: self._update_vision_hint())
        self._update_vision_hint()

    def _label(self, parent, text):
        return tk.Label(parent, text=text, bg=C_APP, fg=C_TEXT, font=(FONT, 9))

    # ------------------------------------------------------------ 交互 --
    def _on_style(self):
        key = self.var_style.get()
        desc = ""
        for item in STYLE_OPTIONS:
            if item[0] == key:
                desc = item[3]
        self.lbl_style.configure(text=desc)

    def _style_value(self):
        key = self.var_style.get()
        for item in STYLE_OPTIONS:
            if item[0] == key:
                return item[2]
        return STYLE_OPTIONS[0][2]

    def _update_vision_hint(self):
        model = self.var_model.get().strip()
        if not model:
            self.lbl_vision.configure(text="识别图纸需要选择带 VL / vision 的模型", fg=C_HINT)
        elif ai_model.is_vision_model(model):
            self.lbl_vision.configure(text="当前模型可识别图纸（图片压缩后随问题一起发送）", fg=C_OK)
        else:
            self.lbl_vision.configure(
                text="当前模型只能处理文字；识别图纸请换带 VL / vision 的模型"
                     "（如 qwen-vl-max、glm-4v-plus、gpt-4o）", fg=C_WARN)

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
            "temperature": self._style_value(),
            "timeout": timeout,
        }

    def _load_engine(self, engine):
        profile = self.cfg["profiles"].get(engine) or ai_model.default_profile(engine)
        self._engine = engine
        self.var_engine.set(engine)
        self.var_endpoint.set(profile.get("endpoint") or "")
        self.var_model.set(profile.get("model") or "")
        self.var_key.set(profile.get("key") or "")
        self.var_style.set(style_key_of(self.cfg.get("temperature")))
        self._on_style()
        _name, preset_endpoint = ai_model.provider_preset(engine)[:2]
        self.cb_model.configure(values=ai_model.model_choices(engine))
        self._update_vision_hint()
        endpoint = profile.get("endpoint") or ""
        configured = bool(profile.get("key")) or bool(endpoint and endpoint != preset_endpoint)
        if configured:
            self._set_status("已载入上次保存的配置：%s / %s"
                             % (engine, profile.get("model") or "未选模型"), C_OK)
        else:
            self._set_status("请填写地址和密钥，再点「读取列表」选择模型")

    def _on_engine_change(self, _event=None):
        new = self.var_engine.get()
        if new == self._engine:
            return
        self.cfg["profiles"][self._engine] = self._collect_profile()
        self._load_engine(new)
        self._persist(notify=False)

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
            self._persist(notify=False)
            return "读取到 %d 个模型" % len(names)

        self._run_async(work, ok, "正在读取模型列表…")

    def _on_test(self):
        profile = self._collect_profile()

        def ok(text):
            return "连接成功：%s" % text

        self._run_async(lambda: ai_model.test_connection(profile), ok, "正在测试连接…")

    # ------------------------------------------------------------ 保存 --
    def _persist(self, notify=True):
        """写盘；关闭窗口、切换引擎、读取模型后都会自动调用。"""
        self.cfg["profiles"][self._engine] = self._collect_profile()
        self.cfg["engine"] = self._engine
        self.cfg["temperature"] = self._style_value()
        try:
            self.cfg["timeout"] = max(5, min(300, int(float(self.var_timeout.get()))))
        except Exception:
            self.cfg["timeout"] = ai_model.DEFAULT_TIMEOUT
        if not ai_model.save_config(self.cfg):
            self._set_status("保存失败：无法写入配置文件", C_ERR)
            return False
        if self.on_saved:
            try:
                self.on_saved(ai_model.load_config())
            except Exception:
                pass
        if notify:
            profile = self.cfg["profiles"].get(self._engine) or {}
            self._set_status("已保存：%s / %s（下次打开自动带出）"
                             % (self._engine, profile.get("model") or "未选模型"), C_OK)
        return True

    def _on_save(self):
        self._persist(True)

    def _on_close(self):
        self._persist(notify=False)
        self.destroy()
