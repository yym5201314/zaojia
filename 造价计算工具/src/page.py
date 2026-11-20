# -*- coding: utf-8 -*-
"""页签基类 + 两张业务表页签"""

import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import calc
import excel_io
import fmt
import material
from sheet import (C_APP, C_LINE, C_MUTED, C_TEXT, FONT, SheetGrid)

C_TOPBAR = "#E8EDF4"


class BasePage(tk.Frame):
    TAB_TITLE = "表"
    HINT = ""
    TITLE = ""

    def __init__(self, master, app, scale=1.0):
        tk.Frame.__init__(self, master, bg=C_APP)
        self.app = app
        self.s = scale
        self.rows = []
        self.var_proj = tk.StringVar()
        self.var_stat = tk.StringVar()
        self._undo = []
        self._redo = []
        self._suppress_undo = False
        self.btn_undo = None
        self.btn_redo = None
        self.hint_label = None
        self.grid = None
        self._build()
        self.reset_data()

    # ------------------------------------------------------------ 钩子 --
    def columns(self):
        return []

    def new_row(self):
        return {}

    def recompute(self):
        pass

    def status_text(self):
        return ""

    def export_file(self, path):
        raise NotImplementedError

    def import_rows(self, items):
        raise NotImplementedError

    def help_text(self):
        return ""

    def import_full_rows(self, grid):
        """表格带有本表完整列时返回完整行数据，否则返回 None。"""
        return None

    def build_params(self, bar):
        """在项目名称右侧添加项目级参数控件。"""
        pass

    def apply_all(self):
        pass

    def extra_buttons(self):
        return []

    def state(self):
        return {"project": self.var_proj.get(), "rows": [dict(r) for r in self.rows],
                "params": self.params()}

    def params(self):
        return {}

    def load_state(self, data):
        rows = []
        for item in data.get("rows", []):
            row = self.new_row()
            for k, v in item.items():
                row[k] = v
            rows.append(row)
        self.rows = rows or [self.new_row() for _ in range(8)]
        self.var_proj.set(data.get("project", ""))
        self.load_params(data.get("params", {}))
        self.sync_columns()
        self.set_rows(self.rows)

    def load_params(self, params):
        pass

    def reset_data(self):
        self.rows = [self.new_row() for _ in range(8)]
        self.set_rows(self.rows)

    # ------------------------------------------------------------ 构建 --
    def _build(self):
        s = self.s
        bar1 = tk.Frame(self, bg=C_APP)
        bar1.pack(side="top", fill="x", padx=int(14 * s), pady=(int(10 * s), 0))
        tk.Label(bar1, text="项目名称", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")
        tk.Entry(bar1, textvariable=self.var_proj, font=(FONT, 10), width=26,
                 relief="solid", bd=1).pack(side="left", padx=(int(6 * s), int(16 * s)))
        self.build_params(bar1)
        ttk.Button(bar1, text="应用到全部行", style="Soft.TButton",
                   command=self.do_apply_all).pack(side="left", padx=(int(10 * s), 0))

        # 导入 / 导出按钮靠右，与项目名称同一行
        io_box = tk.Frame(bar1, bg=C_APP)
        io_box.pack(side="right")
        ttk.Button(io_box, text="导出 Excel 结果表", command=self.do_export,
                   style="Soft.TButton").pack(side="right", padx=(int(6 * s), 0))
        ttk.Button(io_box, text="导入 Excel（序号/名称/单位）", command=self.do_import,
                   style="Accent.TButton").pack(side="right")

        bar2 = tk.Frame(self, bg=C_APP)
        bar2.pack(side="top", fill="x", padx=int(14 * s), pady=(int(8 * s), int(8 * s)))
        btns = [
            ("打开工程", self.app.do_open, "Soft.TButton"),
            ("保存工程", self.app.do_save, "Soft.TButton"),
            ("另存为", self.app.do_save_as, "Soft.TButton"),
            ("插入行", self.insert_row, "Soft.TButton"),
            ("删除行", self.del_row, "Soft.TButton"),
        ]
        btns.extend(self.extra_buttons())
        btns.append(("使用说明", self.show_help, "Soft.TButton"))
        for text, cmd, style in btns:
            ttk.Button(bar2, text=text, command=cmd, style=style).pack(
                side="left", padx=(0, int(6 * s)))

        right_box = tk.Frame(bar2, bg=C_APP)
        right_box.pack(side="right")
        self.btn_redo = ttk.Button(right_box, text="重做 →", command=self.do_redo,
                                   style="Soft.TButton", width=7)
        self.btn_redo.pack(side="right", padx=(int(4 * s), 0))
        self.btn_undo = ttk.Button(right_box, text="← 撤销", command=self.do_undo,
                                   style="Soft.TButton", width=7)
        self.btn_undo.pack(side="right")

        self.grid = SheetGrid(self, self.columns(), scale=s,
                              on_change=self._on_grid_change,
                              before_change=self.push_undo)
        self.grid.set_row_factory(self.new_row, self.is_empty_row)
        self.grid.on_before_reload = self.recompute
        self.grid.pack(side="top", fill="both", expand=True,
                       padx=int(14 * s), pady=(0, int(6 * s)))

        status = tk.Frame(self, bg=C_TOPBAR, height=int(28 * s))
        status.pack(side="bottom", fill="x")
        status.pack_propagate(False)
        tk.Label(status, textvariable=self.var_stat, bg=C_TOPBAR, fg="#31465F",
                 font=(FONT, 9), anchor="w").pack(side="left", padx=int(14 * s))
        self.hint_label = tk.Label(status, text=self.HINT, bg=C_TOPBAR, fg=C_MUTED,
                                   font=(FONT, 8))
        self.hint_label.pack(side="right", padx=int(14 * s))
        self._update_history_buttons()

    # ------------------------------------------------------------ 数据 --
    def is_empty_row(self, row):
        return not any(str(v or "").strip() for k, v in row.items()
                       if not k in ("f", "p1", "amount", "bid", "avg", "diff",
                                    "price", "lo", "hi", "flt"))

    def set_rows(self, rows):
        self.rows = rows
        if self.grid is not None:
            self.grid.set_rows(rows)
        self._on_grid_change()

    def refresh(self):
        self.recompute()
        if self.grid is not None:
            self.grid.reload()
        self._on_grid_change()

    def _on_grid_change(self):
        # 单元格改动后立即重算结果列，再刷新底栏
        try:
            self.recompute()
        except Exception:
            pass
        self.var_stat.set(self.compose_status())

    def compose_status_refresh(self):
        self.var_stat.set(self.compose_status())

    def compose_status(self):
        text = self.status_text()
        if self.grid is not None:
            try:
                cells, nums, total = self.grid.selection_stats()
            except Exception:
                cells, nums, total = 0, 0, 0.0
            if cells > 1:
                text += "　|　选中 %d 格" % cells
                if nums:
                    text += "，数值合计 %s" % fmt.fmt_money(total)
        return text

    def count_rows(self):
        return len([r for r in self.rows if not self.is_empty_row(r)])

    # -------------------------------------------------------- 撤销重做 --
    def snapshot(self):
        return {"state": self.state(),
                "sel": (self.grid.sel_row, self.grid.sel_col) if self.grid else (0, 0)}

    def push_undo(self):
        if self._suppress_undo:
            return
        self._undo.append(self.snapshot())
        if len(self._undo) > 100:
            self._undo.pop(0)
        self._redo = []
        self._update_history_buttons()

    def flush_edit(self):
        self._suppress_undo = True
        try:
            if self.grid is not None:
                self.grid.commit_pending()
        finally:
            self._suppress_undo = False

    def restore(self, snap):
        self.load_state(snap["state"])
        self.sync_columns()
        r, c = snap.get("sel", (0, 1))
        if self.grid is not None:
            self.grid.sel_row = min(max(0, r), max(0, len(self.rows) - 1))
            self.grid.sel_col = min(max(0, c), len(self.columns()) - 1)
        self.set_rows(self.rows)

    def sync_columns(self):
        """子类可在此刷新动态列（材料调差的月份列）。"""

    def do_undo(self):
        self.flush_edit()
        if not self._undo:
            self.flash("已经回到最早一步，没有可撤销的操作")
            return
        snap = self._undo.pop()
        self._redo.append(self.snapshot())
        self.restore(snap)
        self._update_history_buttons()

    def do_redo(self):
        self.flush_edit()
        if not self._redo:
            self.flash("没有可重做的操作")
            return
        snap = self._redo.pop()
        self._undo.append(self.snapshot())
        self.restore(snap)
        self._update_history_buttons()

    def _update_history_buttons(self):
        for btn, stack in ((self.btn_undo, self._undo), (self.btn_redo, self._redo)):
            if btn is None:
                continue
            try:
                btn.state(["!disabled"] if stack else ["disabled"])
            except Exception:
                pass

    def flash(self, msg, ok=False):
        if self.hint_label is None:
            return
        self.hint_label.config(text=msg, fg="#1E7F4F" if ok else "#C0392B")
        self.root_after(2200, lambda: self.hint_label.config(text=self.HINT, fg=C_MUTED))

    def root_after(self, ms, fn):
        try:
            self.winfo_toplevel().after(ms, fn)
        except Exception:
            pass

    # ---------------------------------------------------------- 行操作 --
    def insert_row(self):
        self.flush_edit()
        self.push_undo()
        idx = min(self.grid.sel_row, len(self.rows) - 1)
        self.rows.insert(idx + 1, self.new_row())
        self.grid.sel_row = idx + 1
        self.grid.sel_col = 1
        self.set_rows(self.rows)

    def del_row(self):
        self.flush_edit()
        if not self.rows:
            return
        self.push_undo()
        idx = self.grid.sel_row
        del self.rows[idx]
        if not self.rows:
            self.rows.append(self.new_row())
        self.grid.sel_row = min(idx, len(self.rows) - 1)
        self.set_rows(self.rows)

    def do_apply_all(self):
        self.flush_edit()
        self.push_undo()
        self.apply_all()
        self.refresh()

    # ------------------------------------------------------------ 文件 --
    def do_import(self):
        self.flush_edit()
        path = filedialog.askopenfilename(
            title="选择要导入的 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("CSV 文件", "*.csv"),
                       ("所有文件", "*.*")])
        if not path:
            return
        try:
            table = excel_io.read_table(path)
        except Exception as exc:
            messagebox.showerror("导入失败", "读取文件出错：\n%s" % exc)
            return
        project, items = excel_io.extract_items(table)
        full = None
        try:
            full = self.import_full_rows(table)
        except Exception:
            full = None
        if full is None and not items:
            messagebox.showwarning(
                "导入失败",
                "没有识别到清单行。\n请确认表格中包含「序号 / 名称 / 单位」等列。")
            return
        self.push_undo()
        if project:
            self.var_proj.set(project)
        if full is not None:
            self.rows = full
            self.sync_columns()
            self.grid.sel_row = 0
            self.grid.sel_col = 1
            self.set_rows(self.rows)
            messagebox.showinfo(
                "导入完成",
                "已导入 %d 行完整数据（含数量、价格、参数）。" % len(full))
            return
        self.import_rows(items)
        self.grid.sel_row = 0
        self.grid.sel_col = 1
        self.set_rows(self.rows)
        messagebox.showinfo("导入完成", "已导入 %d 行。" % len(items))

    def do_export(self):
        self.flush_edit()
        name = (self.var_proj.get().strip() or self.TITLE or "结果表")
        path = filedialog.asksaveasfilename(
            title="导出 Excel", defaultextension=".xlsx",
            initialfile=name + ".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            self.export_file(path)
        except Exception as exc:
            messagebox.showerror("导出失败", "写入文件出错：\n%s" % exc)
            return
        messagebox.showinfo("导出完成", "已保存：\n%s" % path)

    def show_help(self):
        win = tk.Toplevel(self.winfo_toplevel())
        win.title("%s - 使用说明" % (self.TITLE or "使用说明"))
        win.configure(bg="#FFFFFF")
        body = tk.Frame(win, bg="#FFFFFF")
        body.pack(fill="both", expand=True)
        bar = ttk.Scrollbar(body, orient="vertical")
        txt = tk.Text(body, font=(FONT, 10), wrap="word", bd=0, padx=16, pady=12,
                      bg="#FFFFFF", fg=C_TEXT, yscrollcommand=bar.set)
        bar.configure(command=txt.yview)
        txt.insert("1.0", self.help_text())
        txt.configure(state="disabled")
        bar.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=(6, 10))
        w = int(860 * self.s)
        h = int(760 * self.s)
        sh = win.winfo_screenheight()
        if h > sh - 120:
            h = max(int(480 * self.s), sh - 120)
        win.geometry("%dx%d" % (w, h))
        win.transient(self.winfo_toplevel())


# ==================================================== 表一：量增量减 ====

class AdjustPage(BasePage):
    TAB_TITLE = "量增量减费用调整表"
    TITLE = "量增量减费用调整表"
    HINT = "拖动可框选多格看合计　·　Ctrl+V 粘贴　·　Ctrl+Z 撤销"

    def __init__(self, master, app, scale=1.0):
        self.var_dev = tk.StringVar(value="15")
        self.var_flt = tk.StringVar(value="5")
        BasePage.__init__(self, master, app, scale)

    def columns(self):
        return calc.COLUMNS

    def new_row(self):
        return calc.new_row()

    def recompute(self):
        calc.compute_all(self.rows)

    def build_params(self, bar):
        s = self.s
        tk.Label(bar, text="清单规定偏差", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")
        tk.Entry(bar, textvariable=self.var_dev, font=(FONT, 10), width=6, justify="right",
                 relief="solid", bd=1).pack(side="left", padx=(int(6 * s), int(2 * s)))
        tk.Label(bar, text="%", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")
        tk.Label(bar, text="   报价浮动率", bg=C_APP, fg=C_TEXT,
                 font=(FONT, 9)).pack(side="left")
        tk.Entry(bar, textvariable=self.var_flt, font=(FONT, 10), width=6, justify="right",
                 relief="solid", bd=1).pack(side="left", padx=(int(6 * s), int(2 * s)))
        tk.Label(bar, text="%", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")

    def apply_all(self):
        dev = self.var_dev.get().strip()
        flt = self.var_flt.get().strip()
        if dev and fmt.parse_pct(dev) is None:
            self.flash("清单规定偏差格式不正确")
            return
        if flt and fmt.parse_pct(flt) is None:
            self.flash("报价浮动率格式不正确")
            return
        for row in self.rows:
            if self.is_empty_row(row):
                continue
            row["dev"] = dev
            row["flt"] = flt

    def status_text(self):
        total_a, total_b = calc.totals(self.rows)
        diff = total_a - total_b
        return ("清单 %d 项　·　合计调整合价 %s　·　合计中标价 %s　·　差额 %s"
                % (self.count_rows(), fmt.fmt_money(total_a),
                   fmt.fmt_money(total_b), fmt.fmt_money(diff)))

    def params(self):
        return {"dev": self.var_dev.get(), "flt": self.var_flt.get()}

    def load_params(self, params):
        self.var_dev.set(params.get("dev", "15"))
        self.var_flt.set(params.get("flt", "5"))

    def import_rows(self, items):
        self.rows = [calc.new_row(name=it["name"], unit=it["unit"]) for it in items]

    def export_file(self, path):
        calc.export_table(path, self.var_proj.get().strip(), self.rows)

    IMPORT_ALIASES = {
        "name": ("名称", "清单名称", "项目名称及特征", "分项工程名称", "工程名称"),
        "unit": ("单位", "计量单位"),
        "q1": ("结算量q1", "结算量"),
        "q0": ("招标清单量q0", "招标清单量"),
        "p2": ("招标单价p2", "招标单价"),
        "p0": ("中标单价p0", "中标单价"),
        "dev": ("清单规定偏差j", "清单规定偏差", "偏差"),
        "flt": ("报价浮动率l", "报价浮动率", "浮动率"),
        "note": ("备注",),
    }

    def import_full_rows(self, grid):
        """读回本表（或结构相同的表）的完整数据：数量、单价、偏差、浮动率。"""
        titles, start = excel_io.read_header(grid)
        if not titles:
            return None
        norm = [excel_io.norm_title(t) for t in titles]
        idx = {}
        for field, aliases in self.IMPORT_ALIASES.items():
            for j, t in enumerate(norm):
                if t in aliases:
                    idx[field] = j
                    break
        if "name" not in idx or "q1" not in idx:
            return None
        if "p0" not in idx and "p2" not in idx:
            return None

        def cell(line, field):
            j = idx.get(field)
            if j is None or j >= len(line):
                return None
            return line[j]

        rows = []
        for line in grid[start:]:
            if not line:
                continue
            name = str(cell(line, "name") or "").strip()
            if not name or name in ("合计", "小计", "总计", "本页小计"):
                continue
            row = calc.new_row(name=name,
                               unit=str(cell(line, "unit") or "").strip())
            for field in ("q1", "q0"):
                v = fmt.parse_num(cell(line, field))
                if v is not None:
                    row[field] = fmt.fmt_qty(v)
            for field in ("p2", "p0"):
                v = fmt.parse_num(cell(line, field))
                if v is not None:
                    row[field] = fmt.fmt_price(v)
            for field in ("dev", "flt"):
                v = fmt.parse_pct(cell(line, field))
                if v is not None:
                    row[field] = fmt.fmt_pct(v)
            row["note"] = str(cell(line, "note") or "").strip()
            rows.append(row)
        return rows or None

    def help_text(self):
        return HELP_ADJUST


# ==================================================== 表二：材料调差 ====

class MaterialPage(BasePage):
    TAB_TITLE = "材料调差表"
    TITLE = "材料调差表"
    HINT = "月份列可增删　·　拖动框选看合计　·　Ctrl+Z 撤销"

    def __init__(self, master, app, scale=1.0):
        self.var_flt = tk.StringVar(value="5")
        self.var_start = tk.StringVar(value="2026-01")
        self.month_count = 4
        BasePage.__init__(self, master, app, scale)

    # --- 结构 ---
    def month_titles(self):
        return material.month_titles(self.month_count, self.var_start.get())

    def month_keys(self):
        return material.month_keys(self.month_count)

    def columns(self):
        return material.build_columns(self.month_titles())

    def new_row(self):
        return material.new_row(self.month_keys())

    def recompute(self):
        material.ensure_month_keys(self.rows, self.month_keys())
        material.compute_all(self.rows, self.month_keys(), self.var_flt.get())

    # --- 参数 ---
    def build_params(self, bar):
        s = self.s
        tk.Label(bar, text="统一风险幅度", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")
        tk.Entry(bar, textvariable=self.var_flt, font=(FONT, 10), width=6, justify="right",
                 relief="solid", bd=1).pack(side="left", padx=(int(6 * s), int(2 * s)))
        tk.Label(bar, text="%", bg=C_APP, fg=C_TEXT, font=(FONT, 9)).pack(side="left")
        tk.Label(bar, text="   施工期起始", bg=C_APP, fg=C_TEXT,
                 font=(FONT, 9)).pack(side="left")
        ent = tk.Entry(bar, textvariable=self.var_start, font=(FONT, 10), width=10,
                       justify="center", relief="solid", bd=1)
        ent.pack(side="left", padx=(int(6 * s), int(4 * s)))
        ent.bind("<Return>", lambda e: self.on_start_changed())
        ent.bind("<FocusOut>", lambda e: self.on_start_changed())
        self.var_flt.trace_add("write", lambda *a: self.on_param_changed())

    def on_start_changed(self):
        self.sync_columns()
        self.refresh()

    def on_param_changed(self):
        if self.grid is None:
            return
        self.refresh()

    def apply_all(self):
        self.recompute()

    def extra_buttons(self):
        return [("＋ 月份列", self.add_month, "Soft.TButton"),
                ("－ 月份列", self.del_month, "Soft.TButton")]

    def add_month(self):
        self.flush_edit()
        self.push_undo()
        self.month_count = min(36, self.month_count + 1)
        self.sync_columns()
        self.refresh()

    def del_month(self):
        self.flush_edit()
        if self.month_count <= 1:
            self.flash("至少保留一个月")
            return
        self.push_undo()
        self.month_count -= 1
        self.sync_columns()
        self.refresh()

    # --- 状态 / 存档 ---
    def status_text(self):
        ta = material.totals(self.rows)
        return ("材料 %d 项　·　合计本期调整金额 %s"
                % (self.count_rows(), fmt.fmt_money(ta)))

    def is_empty_row(self, row):
        keys = ["name", "unit", "qty", "base", "bid", "note"] + self.month_keys()
        return not any(str(row.get(k, "") or "").strip() for k in keys)

    def params(self):
        return {"flt": self.var_flt.get(), "months": self.month_count,
                "start": self.var_start.get()}

    def load_params(self, params):
        self.var_flt.set(params.get("flt", "5"))
        self.month_count = int(params.get("months", 4) or 4)
        self.var_start.set(params.get("start", "2026-01"))

    def sync_columns(self):
        material.ensure_month_keys(self.rows, self.month_keys())
        if self.grid is not None:
            self.grid.set_columns(self.columns())

    def import_rows(self, items):
        keys = self.month_keys()
        rows = []
        for it in items:
            qty = it.get("qty")
            rows.append(material.new_row(
                keys, name=it["name"], unit=it["unit"],
                qty=("" if qty is None else fmt.fmt_qty(qty))))
        self.rows = rows

    def export_file(self, path):
        material.export_table(path, self.var_proj.get().strip(), self.rows,
                              self.month_keys(), self.month_titles(),
                              self.var_flt.get())

    IMPORT_ALIASES = {
        "name": ("材料名称", "名称"),
        "unit": ("单位", "计量单位"),
        "qty": ("本期数量", "数量"),
        "flt": ("风险幅度",),
        "base": ("基准价",),
        "bid": ("投标价/中标价", "投标价", "中标价"),
        "note": ("备注",),
    }
    _MONTH_RE = re.compile(r"^\d{4}年\d{1,2}月$")

    def import_full_rows(self, grid):
        """读回完整材料调差数据：数量、风险幅度、基准价/中标价、各月信息价。"""
        titles, start = excel_io.read_header(grid)
        if not titles:
            return None
        norm = [excel_io.norm_title(t) for t in titles]
        idx = {}
        for field, aliases in self.IMPORT_ALIASES.items():
            for j, t in enumerate(norm):
                if t in aliases:
                    idx[field] = j
                    break
        months = [(j, str(t).strip()) for j, t in enumerate(titles)
                  if self._MONTH_RE.match(str(t).strip())]
        if not months or "base" not in idx or "bid" not in idx:
            return None

        # 按导入表的月份列重建月份列
        self.month_count = len(months)
        self.var_start.set(months[0][1])
        keys = self.month_keys()

        def cell(line, field):
            j = idx.get(field)
            if j is None or j >= len(line):
                return None
            return line[j]

        rows = []
        flt_vals = []
        for line in grid[start:]:
            if not line:
                continue
            name = str(cell(line, "name") or "").strip()
            if not name or name in ("合计", "小计", "总计", "本页小计"):
                continue
            qty = fmt.parse_num(cell(line, "qty"))
            base = fmt.parse_num(cell(line, "base"))
            bid = fmt.parse_num(cell(line, "bid"))
            row = material.new_row(
                keys, name=name, unit=str(cell(line, "unit") or "").strip(),
                qty=("" if qty is None else fmt.fmt_qty(qty)),
                base=("" if base is None else fmt.fmt_price(base)),
                bid=("" if bid is None else fmt.fmt_price(bid)))
            for key, (j, _title) in zip(keys, months):
                v = fmt.parse_num(line[j]) if j < len(line) else None
                row[key] = "" if v is None else fmt.fmt_price(v)
            row["note"] = str(cell(line, "note") or "").strip()
            rows.append(row)
            flt_vals.append(fmt.parse_pct(cell(line, "flt")))
        if not rows:
            return None

        # 出现最多的风险幅度作为顶部统一值，与之不同的行单独记
        counts = {}
        for v in flt_vals:
            if v is None:
                continue
            k = round(v, 6)
            counts[k] = counts.get(k, 0) + 1
        mode = max(counts, key=counts.get) if counts else None
        if mode is not None:
            self.var_flt.set(fmt.fmt_pct(mode))
        for row, v in zip(rows, flt_vals):
            if v is None or (mode is not None and round(v, 6) == mode):
                row["flt_own"] = ""
            else:
                row["flt_own"] = fmt.fmt_pct(v)
        return rows

    def help_text(self):
        return HELP_MATERIAL


HELP_ADJUST = """【一、条款约定】（表格按此计算）
一、Q1 > 1.15*Q0 时，超出部分按 P2*1.15 计算：
      P0 > P2*1.15  → 按 P2*1.15 执行
      P0 ≤ P2*1.15  → 按中标价执行
二、Q1 < 0.85*Q0 时，按 P2*0.85*(1-L) 计算：
      P0 < P2*0.85*(1-L) → 按 P2*0.85*(1-L) 执行
      P0 ≥ P2*0.85*(1-L) → 按中标价执行

【二、计算公式】（与导出的 Excel 一致）
F 需调价量 = IF(Q1>(1+J)*Q0, Q1-(1+J)*Q0, IF(Q1<Q0*(1-J), Q1, 0))
I 结算单价 = IF(AND(Q1>Q0*(1+J), P0<P2*(1+J)), P0,
        IF(AND(Q1<Q0*(1-J), P0>P2*(1-J)*(1-L)), P0,
        IF(P0<P2*(1-L)*(1-J), P2*(1-L)*(1-J),
        IF(P0>P2*(1+J), P2*(1+J), P0))))
L 调整合价 = IF(Q1>Q0, (Q1-F)*P0+F*I, IF(Q1<Q0, F*I, 0))
M 中标价 = Q0*P0

【三、录入】
Q1 结算量　Q0 招标清单量　P2 招标单价　P0 中标单价
J 清单规定偏差（默认 15）　L 报价浮动率（默认 5），填 15 或 15% 均可
黄色格手工填写，绿色格自动计算；序号/名称/单位 可直接填或从 Excel 导入

【四、快捷操作】
双击或直接输入 = 编辑；回车 / Tab = 确认并移动；拖动 = 框选看合计
Ctrl+Z 撤销　Ctrl+Y 重做　Ctrl+C 复制　Delete 清空选中格
Ctrl+V 粘贴：多行多列从选区左上角铺开，剪贴板只有一个值时填满选区
Ctrl+S 保存工程　Ctrl+O 打开工程　Ctrl+E 导出 Excel
"""

HELP_MATERIAL = """【一、调整规则】
F = 风险幅度　G = 基准价　H = 投标价/中标价　I = 信息价平均　N = 调整值

一、中标价 < 基准价
  1）I > G*(1+F)　→ N = I − G*(1+F)　上调
  2）I < H*(1−F)　→ N = I − H*(1−F)　下调
  3）H*(1−F) ≤ I ≤ G*(1+F)　→ N = 0　不调整
二、中标价 > 基准价
  4）I < G*(1−F)　→ N = I − G*(1−F)　下调
  5）I > H*(1+F)　→ N = I − H*(1+F)　上调
  6）G*(1−F) ≤ I ≤ H*(1+F)　→ N = 0　不调整
三、中标价 = 基准价
  7）I < G*(1−F)　→ N = I − G*(1−F)　下调
  8）I > G*(1+F)　→ N = I − G*(1+F)　上调
  9）G*(1−F) ≤ I ≤ G*(1+F)　→ N = 0　不调整

合成一条公式（与导出的 Excel 一致）：
N = IF(I < MIN(H*(1-F), G*(1-F)), I - MIN(H*(1-F), G*(1-F)),
    IF(I > MAX(H*(1+F), G*(1+F)), I - MAX(H*(1+F), G*(1+F)), 0))
结算价 P = H + N　　本期调整金额 = E * N
信息价平均 I = 各月信息价之和 ÷ 有价月份数
  （空白或填 0 的月份不计入，未实施的月份留空即可）

【二、录入】
E 本期数量、G 基准价、H 投标价/中标价、各月信息价 手工填（黄色）
F 风险幅度：默认取顶部「统一风险幅度」（灰色字），个别材料不同时
  直接在该格填 3%，清空（Delete）即恢复统一值
I / N / O / P 自动计算（绿色）
月份列用「＋ 月份列 / － 月份列」增减，标题按「施工期起始」自动排月

【三、快捷操作】
双击或直接输入 = 编辑；回车 / Tab = 确认并移动；拖动 = 框选看合计
Ctrl+Z 撤销　Ctrl+Y 重做　Ctrl+C 复制　Delete 清空
Ctrl+V 粘贴：多行多列从选区左上角铺开，剪贴板只有一个值时填满选区
Ctrl+S 保存工程　Ctrl+O 打开工程　Ctrl+E 导出 Excel
"""
