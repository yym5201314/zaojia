# -*- coding: utf-8 -*-
"""通用表格控件（自绘 Canvas：可编辑、可拖选、隔行底色、黄参数/绿结果）"""

import re
import tkinter as tk
from tkinter import ttk

import fmt

C_APP = "#F2F5F9"
C_LINE = "#C7D6E8"
C_HEAD = "#DDEBF7"
C_HEAD_TX = "#1B3A5C"
C_PARAM = "#FFF6D8"
C_RESULT = "#E3F6E6"
C_SEQ = "#EEF3F9"
C_SEL = "#1A73E8"
C_TEXT = "#1F2933"
C_ERR = "#C00000"
C_MUTED = "#8A97A6"

FONT = "Microsoft YaHei UI"

_ROW_BUF = 10      # 可视区上下预绘的行数（滚动缓冲，防闪白）

_STATION_RE = re.compile(r"^([A-Za-z\u4e00-\u9fff]*)\s*[Kk]\s*(\d+)\s*\+\s*(\d+(?:\.\d+)?)$")
_PLAIN_NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def fmt_plain_num(value):
    """3.0 → 3；3.5 → 3.5（文本列填充用）。"""
    txt = ("%.6f" % value).rstrip("0").rstrip(".")
    return "0" if txt in ("", "-0") else txt


def station_parts(text):
    """拆桩号：'AK0+020' → ('A', 20.0)、'K1+020' → ('', 1020.0)；不是桩号返回 None。"""
    m = _STATION_RE.match(str(text or "").strip())
    if not m:
        return None
    return m.group(1), int(m.group(2)) * 1000.0 + float(m.group(3))


def parse_station(text):
    """K0+020 / AK0+020 → 20.0（米）；不是桩号返回 None。"""
    parts = station_parts(text)
    return None if parts is None else parts[1]


def fmt_station(value, prefix=""):
    """20.0 → K0+020；prefix='A' → AK0+020（填充拖动时用）。"""
    km = int(value // 1000)
    meter = value - km * 1000
    txt = ("%.3f" % meter).rstrip("0").rstrip(".")
    if txt in ("", "-0"):
        txt = "0"
    if "." in txt:
        head, _, dec = txt.partition(".")
        txt = head.zfill(3) + "." + dec
    else:
        txt = txt.zfill(3)
    return "%sK%d+%s" % (prefix or "", km, txt)


class SheetGrid(tk.Frame):
    def __init__(self, master, columns, scale=1.0, on_change=None, before_change=None,
                 font_family=None):
        tk.Frame.__init__(self, master, bg=C_LINE, bd=0, highlightthickness=0)
        self.cols = list(columns)
        self.rows = []
        self.on_change = on_change
        self.before_change = before_change
        self.s = scale
        self.row_h = max(22, int(26 * scale))
        self.group_h = max(18, int(22 * scale))
        self.col_h = max(22, int(26 * scale))
        self.head_h = self.group_h + self.col_h
        self.family = font_family or FONT
        fpx = max(11, int(13 * scale))
        self.font = (self.family, -fpx)
        self.font_b = (self.family, -fpx, "bold")
        self.sel_row = 0
        self.sel_col = 1
        self.range = None
        self.freeze_cols = 0      # 冻结在左侧的列数
        self.xoff = 0             # 横向滚动偏移（像素，仅作用于未冻结列）
        self._shift_anchor = (0, 1)   # Shift+左键框选起点
        self._width_override = {}
        self._resize = None
        self._anchor = None
        self._press_sel = None
        self.editor = None
        self.edit_win = None
        self.edit_pos = None
        self.on_before_reload = None
        self._new_row = lambda: {}
        self._row_empty = lambda row: False
        self._drawn_rows = set()   # 已绘制的行号（滚动增量绘制用）
        self._fill_handle = None   # 填充柄命中区域（画布坐标）
        self._filling = None       # 正在拖动填充柄时的状态
        self._hover_fill = False   # 鼠标是否悬停在填充柄上
        self._calc_layout()
        self._build()

    # ---------------------------------------------------------- 布局 --
    def _calc_layout(self):
        s = self.s
        self.row_h = max(22, int(26 * s))
        self.group_h = max(18, int(22 * s))
        self.col_h = max(22, int(26 * s))
        self.head_h = self.group_h + self.col_h
        widths = []
        for c in self.cols:
            w = self._width_override.get(c.get("key"))
            widths.append(max(36, int(w if w is not None else c["width"] * s)))
        self.widths = widths
        self.xs = []
        acc = 0
        for w in self.widths:
            self.xs.append(acc)
            acc += w
        self.total_w = acc

    def set_columns(self, columns):
        self.cols = list(columns)
        self._calc_layout()
        self.range = None
        self.sel_col = min(self.sel_col, len(self.cols) - 1)
        self.head.configure(height=self.head_h)
        self.redraw()

    # ------------------------------------------ 冻结列 / 横向滚动 --
    def frozen_n(self):
        return max(0, min(self.freeze_cols, len(self.cols)))

    def frozen_w(self):
        n = self.frozen_n()
        return self.xs[n] if n < len(self.xs) else self.total_w

    def view_w(self):
        """未冻结区域的可视宽度。"""
        try:
            w = self.body.winfo_width()
        except Exception:
            w = 0
        return max(1, w - self.frozen_w())

    def max_xoff(self):
        return max(0, self.total_w - self.frozen_w() - self.view_w())

    def col_dx(self, j):
        """列 j 在画布上的左边界（已考虑冻结与横向偏移）。"""
        if j < self.frozen_n():
            return self.xs[j]
        return self.xs[j] - self.xoff

    def virtual_x(self, x):
        """画布局部 x → 列坐标（0..total_w）。"""
        fw = self.frozen_w()
        if x < fw:
            return x
        return x + self.xoff

    def set_freeze(self, cols):
        """冻结最左侧 cols 列（0 = 取消）。"""
        self.freeze_cols = max(0, min(int(cols), len(self.cols)))
        self.xoff = 0
        self.redraw()
        self._sync_hbar()

    def _set_xoff(self, v):
        v = int(min(max(0, v), self.max_xoff()))
        if v == self.xoff:
            return
        self.xoff = v
        self._sync_hbar()
        self._draw_head()
        self._draw_body()

    def _sync_hbar(self):
        try:
            content = max(1, self.total_w - self.frozen_w())
            vw = self.view_w()
            if content <= vw:
                self.hbar.set(0.0, 1.0)
            else:
                self.hbar.set(self.xoff / float(content),
                              min(1.0, (self.xoff + vw) / float(content)))
        except Exception:
            pass

    def _on_body_configure(self, event=None):
        try:
            self.body.configure(
                scrollregion=(0, 0, max(1, self.body.winfo_width()),
                              max(1, len(self.rows) * self.row_h)))
        except Exception:
            pass
        if self.xoff > self.max_xoff():
            self.xoff = self.max_xoff()
        self._sync_hbar()
        self._draw_body()

    def _on_head_configure(self, event=None):
        try:
            self.head.configure(
                scrollregion=(0, 0, max(1, self.head.winfo_width()), self.head_h))
        except Exception:
            pass
        self._draw_head()

    # ---------------------------------------------------------- 构建 --
    def _build(self):
        self.head = tk.Canvas(self, height=self.head_h, bg=C_HEAD, highlightthickness=0,
                              bd=0, scrollregion=(0, 0, self.total_w, self.head_h))
        self.body = tk.Canvas(self, bg="#FFFFFF", highlightthickness=0, bd=0, takefocus=1)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.head.grid(row=0, column=0, sticky="ew")
        self.body.grid(row=1, column=0, sticky="nsew")
        self.vbar.grid(row=1, column=1, sticky="ns")
        self.hbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)
        self.body.configure(yscrollcommand=self._on_yview)

        self.body.bind("<Button-1>", self._on_press)
        self.body.bind("<B1-Motion>", self._on_drag)
        self.body.bind("<ButtonRelease-1>", self._on_release)
        self.body.bind("<Double-Button-1>", self._on_dblclick)
        self.body.bind("<Configure>", self._on_body_configure)
        self.body.bind("<Motion>", self._on_motion)
        self.body.bind("<Leave>", lambda e: self.body.configure(cursor=""))
        for w in (self.body, self.head):
            w.bind("<MouseWheel>", self._on_wheel)
            w.bind("<Shift-MouseWheel>", self._on_wheel)
        self.head.bind("<Configure>", self._on_head_configure)
        # 表头分隔线：拖动调整列宽
        self.head.bind("<Motion>", self._on_head_motion)
        self.head.bind("<Button-1>", self._on_head_press)
        self.head.bind("<B1-Motion>", self._on_head_drag)
        self.head.bind("<ButtonRelease-1>", self._on_head_release)
        self.head.bind("<Leave>", lambda e: self.head.configure(cursor=""))

        self.body.bind("<Up>", lambda e: self._nav(-1, 0))
        self.body.bind("<Down>", lambda e: self._nav(1, 0))
        self.body.bind("<Left>", lambda e: self._nav(0, -1))
        self.body.bind("<Right>", lambda e: self._nav(0, 1))
        self.body.bind("<Prior>", lambda e: self._nav(-10, 0))
        self.body.bind("<Next>", lambda e: self._nav(10, 0))
        self.body.bind("<Return>", lambda e: self._edit_key())
        self.body.bind("<F2>", lambda e: self._edit_key())
        self.body.bind("<Tab>", lambda e: self._nav(0, 1))
        self.body.bind("<Shift-Tab>", lambda e: self._nav(0, -1))
        self.body.bind("<Delete>", lambda e: self._clear_cell())
        self.body.bind("<Control-v>", lambda e: self._paste())
        self.body.bind("<Control-V>", lambda e: self._paste())
        self.body.bind("<Control-c>", lambda e: self._copy())
        self.body.bind("<Control-C>", lambda e: self._copy())
        self.body.bind("<Key>", self._on_typing)

    # ---------------------------------------------------- 列宽拖动 --
    def _edge_at(self, x):
        """命中的是哪一列的右边界（用于拖动列宽）。"""
        for j in range(len(self.widths)):
            if abs(x - (self.xs[j] + self.widths[j])) <= max(4, int(4 * self.s)):
                return j
        return None

    def _on_head_motion(self, event):
        if self._resize is not None:
            self.head.configure(cursor="sb_h_double_arrow")
            return "break"
        j = self._edge_at(self.virtual_x(event.x))
        try:
            self.head.configure(cursor="sb_h_double_arrow" if j is not None else "")
        except Exception:
            pass
        return "break"

    def _on_head_press(self, event):
        x = self.virtual_x(event.x)
        j = self._edge_at(x)
        self._resize = (j, x, self.widths[j]) if j is not None else None
        return "break"

    def _on_head_drag(self, event):
        if self._resize is None:
            return "break"
        j, x0, w0 = self._resize
        dx = self.virtual_x(event.x) - x0
        w = int(max(36, min(1200, w0 + dx)))
        if w == self.widths[j]:
            return "break"
        key = self.cols[j].get("key")
        if key:
            self._width_override[key] = w
        self._calc_layout()
        self.redraw()
        return "break"

    def _on_head_release(self, event):
        self._resize = None
        try:
            self.head.configure(cursor="")
        except Exception:
            pass
        return "break"

    # ------------------------------------------------------ 滚动同步 --
    def _yview(self, *args):
        self.body.yview(*args)

    def _xview(self, *args):
        """横向滚动：自己维护偏移量，冻结列固定在左边不动。"""
        if not args:
            return
        if args[0] == "moveto":
            try:
                frac = float(args[1])
            except Exception:
                return
            self._set_xoff(int(frac * (self.total_w - self.frozen_w())))
            return
        if args[0] == "scroll":
            try:
                num = int(args[1])
            except Exception:
                return
            if len(args) > 2 and args[2] == "pages":
                step = self.view_w()
            else:
                step = max(20, int(40 * self.s))
            self._set_xoff(self.xoff + num * step)

    def _on_yview(self, first, last):
        self.vbar.set(first, last)
        self._draw_body(force=False)

    def _on_wheel(self, event):
        step = -1 * int(event.delta / 120) if event.delta else 0
        if not step:
            step = -1 if event.delta > 0 else 1
        if event.state & 0x0001:
            self._set_xoff(self.xoff + step * max(20, int(40 * self.s)))
        else:
            self.body.yview_scroll(step, "units")
        return "break"

    # ---------------------------------------------------------- 数据 --
    def set_rows(self, rows):
        self.rows = rows
        self.range = None
        self._anchor = None
        if self.sel_row >= len(self.rows):
            self.sel_row = max(0, len(self.rows) - 1)
        self.reload()
        self._selection_changed()

    def reload(self):
        if self.on_before_reload:
            self.on_before_reload()
        self.redraw()

    def redraw(self):
        h = max(1, len(self.rows)) * self.row_h
        try:
            self.body.configure(scrollregion=(0, 0, max(1, self.body.winfo_width()), h),
                                yscrollincrement=self.row_h)
            self.head.configure(scrollregion=(0, 0, max(1, self.head.winfo_width()),
                                              self.head_h))
        except Exception:
            pass
        if self.xoff > self.max_xoff():
            self.xoff = self.max_xoff()
        self._sync_hbar()
        self._draw_head()
        self._draw_body()

    def _split_freeze(self, c0, c1):
        """把列区间按冻结边界拆成若干段。"""
        fz = self.frozen_n()
        if fz <= c0 or fz > c1:
            return [(c0, c1)]
        return [(c0, fz - 1), (fz, c1)]

    # ---------------------------------------------------------- 绘制 --
    def _draw_head(self):
        c = self.head
        c.delete("all")
        n = len(self.cols)
        if not n:
            return
        fz = self.frozen_n()
        vw = max(1, c.winfo_width())
        segs = []
        i = 0
        while i < n:
            group = self.cols[i].get("group") or ""
            j = i
            while group and j + 1 < n and (self.cols[j + 1].get("group") or "") == group:
                j += 1
            if fz and i < fz:
                segs.append((i, min(j, fz - 1), group, True))
            if j >= fz:
                segs.append((max(i, fz), j, group, False))
            i = j + 1
        # 先画可滚动部分，再画冻结列（覆盖在上面）
        for is_frozen in (False, True):
            for a, b, group, fz_seg in segs:
                if fz_seg != is_frozen:
                    continue
                x0 = self.col_dx(a)
                x1 = self.col_dx(b) + self.widths[b]
                if x1 < -20 or x0 > vw + 20:
                    continue
                self._draw_head_seg(c, a, b, group, x0, x1)
        if fz:
            c.create_line(self.frozen_w(), 0, self.frozen_w(), self.head_h,
                          fill="#7E9CBC", width=2)

    def _draw_head_seg(self, c, a, b, group, x0, x1):
        if group:
            c.create_rectangle(x0, 0, x1, self.group_h, fill="#CFE2F5", outline=C_LINE)
            c.create_text((x0 + x1) / 2, self.group_h / 2, text=group,
                          fill=C_HEAD_TX, font=self.font_b)
            for k in range(a, b + 1):
                cx0 = self.col_dx(k)
                cx1 = cx0 + self.widths[k]
                c.create_rectangle(cx0, self.group_h, cx1, self.head_h,
                                   fill=C_HEAD, outline=C_LINE)
                c.create_text((cx0 + cx1) / 2, (self.group_h + self.head_h) / 2,
                              text=self.cols[k]["title"], fill=C_HEAD_TX,
                              font=self.font_b)
        else:
            c.create_rectangle(x0, 0, x1, self.head_h, fill=C_HEAD, outline=C_LINE)
            c.create_text((x0 + x1) / 2, self.head_h / 2, text=self.cols[a]["title"],
                          fill=C_HEAD_TX, font=self.font_b)

    def _col_order(self):
        fz = self.frozen_n()
        return list(range(fz, len(self.cols))) + list(range(0, fz))

    def _draw_row(self, c, i, vw, order):
        y0 = i * self.row_h
        y1 = y0 + self.row_h
        row = self.rows[i]
        tags = ("cell", "row_%d" % i)
        for j in order:
            col = self.cols[j]
            x0 = self.col_dx(j)
            x1 = x0 + self.widths[j]
            if x1 < -40 or x0 > vw + 40:
                continue
            c.create_rectangle(x0, y0, x1, y1, fill=self._bg(i, col),
                               outline=C_LINE, tags=tags)
            text, color, anchor = self._text(i, row, col)
            if not text:
                continue
            if anchor == "e":
                tx = x1 - 5 * self.s
            elif anchor == "center":
                tx = (x0 + x1) / 2
            else:
                tx = x0 + 6 * self.s
            c.create_text(tx, (y0 + y1) / 2, text=text, anchor=anchor,
                          fill=color, font=self.font, tags=tags)

    def _draw_body(self, force=True):
        """绘制可见区。force=False 只补齐/回收滚出视口的行，滚动时不整屏重建。"""
        c = self.body
        n = len(self.rows)
        try:
            top = c.canvasy(0)
            vh = c.winfo_height()
            vw = c.winfo_width()
        except Exception:
            top, vh, vw = 0, 600, 1000
        # 上下各预绘若干行，滚动时新暴露区域已有内容，不会闪白
        buf = _ROW_BUF
        start = max(0, int(top // self.row_h) - 1 - buf)
        end = min(n, int((top + vh) // self.row_h) + 2 + buf)
        want = set(range(start, end))
        drawn = getattr(self, "_drawn_rows", None)
        if force or drawn is None:
            c.delete("cell")
            drawn = set()
        else:
            for i in drawn - want:
                c.delete("row_%d" % i)
            drawn = drawn & want
        order = self._col_order()
        for i in sorted(want - drawn):
            self._draw_row(c, i, vw, order)
            drawn.add(i)
        self._drawn_rows = drawn
        c.delete("overlay")
        fz = self.frozen_n()
        if fz:
            c.create_line(self.frozen_w(), 0, self.frozen_w(), vh,
                          fill="#7E9CBC", width=2, tags="overlay")
        if self.range and n:
            r0, c0, r1, c1 = self.range
            r1 = min(r1, n - 1)
            for a, b in self._split_freeze(c0, c1):
                c.create_rectangle(self.col_dx(a), r0 * self.row_h,
                                   self.col_dx(b) + self.widths[b],
                                   (r1 + 1) * self.row_h,
                                   outline=C_SEL, width=1, fill=C_SEL,
                                   stipple="gray25", tags="overlay")
        if n and 0 <= self.sel_row < n:
            x0 = self.col_dx(self.sel_col)
            x1 = x0 + self.widths[self.sel_col]
            y0 = self.sel_row * self.row_h
            c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y0 + self.row_h - 1,
                               outline=C_SEL, width=2, tags="overlay")
        # 填充柄（Excel 风格）：选区右下角的小方块，按住拖动可复制/递增填充
        self._fill_handle = None
        if n:
            if self.range:
                br = min(self.range[2], n - 1)
                bc = min(self.range[3], len(self.cols) - 1)
            else:
                br, bc = self.sel_row, self.sel_col
            if 0 <= br < n and 0 <= bc < len(self.cols):
                hx = self.col_dx(bc) + self.widths[bc]
                hy = (br + 1) * self.row_h
                if -1 <= hy - top <= vh + 1:
                    c.create_rectangle(hx - 4, hy - 4, hx + 4, hy + 4,
                                       fill=C_SEL, outline="#FFFFFF", tags="overlay")
                    self._fill_handle = (hx - 5, hy - 5, hx + 5, hy + 5)
        if self.edit_win is not None:
            try:
                if self.edit_pos is not None:
                    er, ec = self.edit_pos
                    c.coords(self.edit_win, self.col_dx(ec), er * self.row_h)
                c.tag_raise(self.edit_win)
            except Exception:
                pass

    def _bg(self, i, col):
        role = col.get("role")
        if role == "result":
            return C_RESULT
        if role == "param":
            return C_PARAM
        if col["kind"] == "seq":
            return C_SEQ
        return "#FFFFFF" if i % 2 == 0 else "#F8FBFF"

    def seq_of(self, i):
        """序号列显示的行号：不把「断开」行算进去。"""
        return sum(1 for r in self.rows[:i] if r.get("kind") != "break") + 1

    def _text(self, i, row, col):
        """数据一律居中显示。"""
        text, color, _anchor = self._text_impl(i, row, col)
        return text, color, "center"

    def _text_impl(self, i, row, col):
        key = col["key"]
        kind = col["kind"]
        if kind == "seq":
            if row.get("kind") == "break":
                return "断开", C_TEXT, "center"
            return str(self.seq_of(i)), C_TEXT, "center"
        if kind == "text_hint":
            anchor = col.get("anchor", "w")
            val = str(row.get(key, "") or "").strip()
            if val:
                if col.get("pct"):
                    v = fmt.parse_pct(val)
                    if v is None:
                        return val, C_ERR, anchor
                    return fmt.fmt_pct(v), C_TEXT, anchor
                if col.get("numeric") and fmt.parse_num(val) is None:
                    return val, C_ERR, anchor
                return val, C_TEXT, anchor
            hint = col.get("hint")
            if hint is not None:
                try:
                    txt = hint(row)
                except Exception:
                    txt = ""
                if txt:
                    return txt, C_MUTED, anchor
            return "", C_TEXT, anchor
        if kind == "text":
            val = str(row.get(key, "") or "")
            return val, C_TEXT, ("center" if col.get("center") else "w")
        raw = row.get(key)
        if col.get("role") == "result":
            if raw is None:
                return "", C_TEXT, "e"
            return fmt.fmt_by_kind(kind, raw), C_TEXT, "e"
        s = str(raw or "").strip()
        if not s:
            return "", C_TEXT, "e"
        if kind == "percent":
            v = fmt.parse_pct(s)
            if v is None:
                return s, C_ERR, "e"
            return fmt.fmt_pct(v), C_TEXT, "e"
        v = fmt.parse_num(s)
        if v is None:
            return s, C_ERR, "e"
        return fmt.fmt_by_kind(kind, v), C_TEXT, "e"

    # ---------------------------------------------------------- 交互 --
    def _hit(self, x, y):
        cx = self.virtual_x(x)
        cy = self.body.canvasy(y)
        r = int(cy // self.row_h)
        if r < 0 or r >= len(self.rows):
            return None
        for j in range(len(self.cols)):
            if self.xs[j] <= cx < self.xs[j] + self.widths[j]:
                return r, j
        return None

    def _hit_clamped(self, x, y):
        if not self.rows:
            return None
        cx = self.virtual_x(x)
        cy = self.body.canvasy(y)
        r = min(max(0, int(cy // self.row_h)), len(self.rows) - 1)
        j = 0 if cx < 0 else len(self.cols) - 1
        for k in range(len(self.cols)):
            if self.xs[k] <= cx < self.xs[k] + self.widths[k]:
                j = k
                break
        return r, j

    # ------------------------------------------------------ 选中区域 --
    def _selection_changed(self):
        if self.on_change:
            self.on_change()

    def selected_cells(self):
        if self.range and self.rows:
            r0, c0, r1, c1 = self.range
            r1 = min(r1, len(self.rows) - 1)
            return [(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)]
        if self.rows:
            return [(self.sel_row, self.sel_col)]
        return []

    def cell_number(self, r, c):
        if not (0 <= r < len(self.rows)) or not (0 <= c < len(self.cols)):
            return None
        col = self.cols[c]
        if col["kind"] in ("text", "text_hint", "seq"):
            return None
        raw = self.rows[r].get(col["key"])
        if col.get("role") == "result":
            return raw if isinstance(raw, (int, float)) else None
        if col["kind"] == "percent":
            return fmt.parse_pct(raw)
        return fmt.parse_num(raw)

    def selection_stats(self):
        cells = self.selected_cells()
        values = [v for v in (self.cell_number(r, c) for r, c in cells) if v is not None]
        return len(cells), len(values), (sum(values) if values else 0.0)

    def plain_text(self, r, c):
        col = self.cols[c]
        row = self.rows[r]
        if col["kind"] == "seq":
            if row.get("kind") == "break":
                return "断开"
            return str(self.seq_of(r))
        if col["kind"] in ("text", "text_hint"):
            return str(row.get(col["key"], "") or "")
        if col.get("role") == "result":
            return fmt.fmt_by_kind(col["kind"], row.get(col["key"]))
        return str(row.get(col["key"], "") or "")

    def _notify_before(self):
        if self.before_change:
            self.before_change()

    def _anchor_cell(self):
        """框选起点：上次单击/方向键定位的单元格。"""
        a = getattr(self, "_shift_anchor", None)
        if a and 0 <= a[0] < max(1, len(self.rows)) and 0 <= a[1] < len(self.cols):
            return a
        return (self.sel_row, self.sel_col)

    def _on_press(self, event):
        was_editing = self.editor is not None
        if was_editing:
            self._commit_edit(silent=True)
        self.body.focus_set()
        if self._fill_hit(event.x, event.y):
            # 按住填充柄：进入填充拖动
            r0, c0, r1, c1 = self._fill_source()
            self._filling = {"r0": r0, "c0": c0, "r1": r1, "c1": c1, "target": None}
            self._anchor = None
            self._press_sel = None
            return "break"
        self._press_sel = None if was_editing else (self.sel_row, self.sel_col)
        hit = self._hit(event.x, event.y)
        self._anchor = hit
        if not hit:
            return "break"
        r, j = hit
        if (event.state & 0x0001) and self.rows:
            # Shift + 左键：从上次定位的单元格框选到本次点击的单元格
            ar, ac = self._anchor_cell()
            self.sel_row, self.sel_col = r, j
            if (r, j) == (ar, ac):
                self.range = None
            else:
                self.range = (min(ar, r), min(ac, j), max(ar, r), max(ac, j))
            self._anchor = None
            self._press_sel = None
            self._draw_body()
            self._selection_changed()
            return "break"
        self._shift_anchor = (r, j)
        self.sel_row, self.sel_col = r, j
        self.range = None
        self._draw_body()
        self._selection_changed()
        return "break"

    def _fill_source(self):
        """填充源区域（选区或当前单元格）。"""
        n = len(self.rows)
        if self.range:
            r0, c0, r1, c1 = self.range
            return r0, c0, min(r1, n - 1), min(c1, len(self.cols) - 1)
        return self.sel_row, self.sel_col, self.sel_row, self.sel_col

    def _fill_hit(self, x, y):
        h = self._fill_handle
        if not h or not self.rows:
            return False
        try:
            cx = self.virtual_x(x)
            cy = self.body.canvasy(y)
        except Exception:
            return False
        return h[0] <= cx <= h[2] and h[1] <= cy <= h[3]

    def _on_motion(self, event):
        if self._filling is not None:
            return "break"
        hit = self._fill_hit(event.x, event.y)
        if hit != self._hover_fill:
            self._hover_fill = hit
            try:
                self.body.configure(cursor="crosshair" if hit else "")
            except Exception:
                pass
        return "break"

    def _draw_fill_preview(self):
        c = self.body
        c.delete("fpreview")
        f = self._filling
        if not f or f.get("target") is None:
            return
        r0, r1, c0, c1, t = f["r0"], f["r1"], f["c0"], f["c1"], f["target"]
        if t > r1:
            y0, y1 = (r1 + 1) * self.row_h, (t + 1) * self.row_h
        else:
            y0, y1 = t * self.row_h, r0 * self.row_h
        for a, b in self._split_freeze(c0, c1):
            c.create_rectangle(self.col_dx(a), y0,
                               self.col_dx(b) + self.widths[b], y1,
                               outline=C_SEL, dash=(4, 2), width=1,
                               tags="fpreview")

    @staticmethod
    def _is_num_col(col):
        return (col.get("kind") in ("num", "price", "percent", "money")
                or bool(col.get("pct")) or bool(col.get("numeric")))

    def _cell_num(self, col, text):
        """单元格文本 → 数值（桩号按米计）；无法解析返回 None。"""
        s = str(text or "").strip()
        if not s:
            return None
        v = parse_station(s)
        if v is not None:
            return v
        if col.get("kind") == "percent" or col.get("pct"):
            return fmt.parse_pct(s)
        if self._is_num_col(col):
            return fmt.parse_num(s)
        # 文本列里的纯数字（如 1、2）也按等差填充
        if _PLAIN_NUM_RE.match(s):
            return float(s)
        return None

    def _cell_out_func(self, col, src):
        """按源值的写法决定回写格式。"""
        last = str(src[-1] if src else "").strip()
        parts = station_parts(last)
        if parts is not None:
            prefix = parts[0]
            return lambda v: fmt_station(v, prefix)
        if self._is_num_col(col):
            if col.get("kind") == "percent" or col.get("pct"):
                return fmt.fmt_pct
            kind = col.get("kind")
            return lambda v: fmt.fmt_by_kind(kind, v)
        return fmt_plain_num

    def _fill_series(self, col, src):
        """返回 f(k)：第 k 个填充值。数值/桩号按等差叠加，其它循环复制。"""
        nums = [self._cell_num(col, v) for v in src]
        if nums and all(v is not None for v in nums):
            out = self._cell_out_func(col, src)
            if len(nums) >= 2:
                step = nums[-1] - nums[-2]
                last = nums[-1]
                return lambda k: out(last + step * (k + 1))
            only = nums[0]
            return lambda k: out(only)
        n = len(src) or 1
        return lambda k: src[k % n]

    def _apply_fill(self, f):
        """把源区域按拖动方向填充到目标行。"""
        if not self.rows:
            return
        n = len(self.rows)
        r0, c0, r1, c1 = f["r0"], f["c0"], f["r1"], f["c1"]
        t = min(max(0, f["target"]), n - 1)
        down = t > r1
        idxs = list(range(r1 + 1, t + 1)) if down else list(range(r0 - 1, t - 1, -1))
        if not idxs:
            return
        cols = [j for j in range(c0, c1 + 1)
                if self.cols[j].get("edit") and self.cols[j].get("kind") != "seq"
                and self.cols[j].get("role") != "result"]
        if not cols:
            return
        self._notify_before()
        for j in cols:
            col = self.cols[j]
            key = col["key"]
            if down:
                src = [self.rows[i].get(key, "") for i in range(r0, r1 + 1)]
            else:
                src = [self.rows[i].get(key, "") for i in range(r1, r0 - 1, -1)]
            fn = self._fill_series(col, src)
            for k, i in enumerate(idxs):
                self.rows[i][key] = fn(k)
        if self.on_change:
            self.on_change()
        self._draw_body()
        self._selection_changed()

    def _on_drag(self, event):
        if self._filling is not None:
            f = self._filling
            if event.y < 0:
                self.body.yview_scroll(-1, "units")
            elif event.y > self.body.winfo_height():
                self.body.yview_scroll(1, "units")
            if self.rows:
                row = int(self.body.canvasy(event.y) // self.row_h)
                row = min(max(0, row), len(self.rows) - 1)
                f["target"] = None if f["r0"] <= row <= f["r1"] else row
                self._draw_fill_preview()
            return "break"
        if self._anchor is None:
            return "break"
        if event.y < 0:
            self.body.yview_scroll(-1, "units")
        elif event.y > self.body.winfo_height():
            self.body.yview_scroll(1, "units")
        if event.x < 0:
            self._set_xoff(self.xoff - max(20, int(40 * self.s)))
        elif event.x > self.body.winfo_width():
            self._set_xoff(self.xoff + max(20, int(40 * self.s)))
        hit = self._hit_clamped(event.x, event.y)
        if not hit:
            return "break"
        r, j = hit
        self.sel_row, self.sel_col = r, j
        ar, ac = self._anchor
        if (r, j) == (ar, ac):
            self.range = None
        else:
            self.range = (min(ar, r), min(ac, j), max(ar, r), max(ac, j))
        self._draw_body()
        self._selection_changed()
        return "break"

    def _on_release(self, event):
        if self._filling is not None:
            f = self._filling
            self._filling = None
            try:
                self.body.delete("fpreview")
                self.body.configure(cursor="")
            except Exception:
                pass
            if f.get("target") is not None:
                self._apply_fill(f)
            return "break"
        anchor = self._anchor
        self._anchor = None
        if anchor and self.range is None and self._press_sel == anchor:
            self.begin_edit()
        self._press_sel = None
        return "break"

    def _on_dblclick(self, event):
        if self.editor is not None:
            self._commit_edit(silent=True)
        hit = self._hit(event.x, event.y)
        if hit:
            self.sel_row, self.sel_col = hit
            self.range = None
            self._draw_body()
            self.begin_edit()
        return "break"

    def _nav(self, dr, dc):
        if not self.rows:
            return "break"
        self._commit_edit(silent=True)
        if dr > 0 and self.sel_row >= len(self.rows) - 1 and not self._row_empty(self.rows[-1]):
            self.rows.append(self._new_row())
            self.redraw()
        r = min(max(0, self.sel_row + dr), len(self.rows) - 1)
        c = min(max(0, self.sel_col + dc), len(self.cols) - 1)
        self.sel_row, self.sel_col = r, c
        self._shift_anchor = (r, c)
        self.range = None
        self.ensure_visible(r, c)
        self._draw_body()
        self.body.focus_set()
        return "break"

    def set_row_factory(self, new_row, is_empty_row):
        self._new_row = new_row
        self._row_empty = is_empty_row

    def _edit_key(self):
        self.begin_edit()
        return "break"

    def _on_typing(self, event):
        if self.editor is not None:
            return
        ch = event.char
        if ch and ch.isprintable() and event.keysym not in (
                "Up", "Down", "Left", "Right", "Return", "Tab", "Escape", "Delete",
                "Home", "End", "Prior", "Next"):
            if self.cols[self.sel_col].get("edit"):
                self.begin_edit(initial=ch)
                return "break"

    def _clear_cell(self):
        if not self.rows:
            return "break"
        targets = [(r, c) for r, c in self.selected_cells() if self.cols[c].get("edit")]
        if not targets:
            return "break"
        self._notify_before()
        for r, c in targets:
            self.rows[r][self.cols[c]["key"]] = ""
        self.reload()
        self._selection_changed()
        return "break"

    def _copy(self):
        try:
            cells = self.selected_cells()
            if not cells:
                return "break"
            rs = sorted(set(r for r, _ in cells))
            cs = sorted(set(c for _, c in cells))
            lines = ["\t".join(self.plain_text(r, c) for c in cs) for r in rs]
            self.clipboard_clear()
            self.clipboard_append("\n".join(lines))
        except Exception:
            pass
        return "break"

    def _paste(self):
        try:
            text = self.clipboard_get()
        except Exception:
            return "break"
        if not text:
            return "break"
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        if not lines:
            return "break"
        data = [ln.split("\t") for ln in lines]
        # 粘贴起点：有框选时用选区左上角，否则用当前单元格
        r1 = c1 = None
        if self.range and self.rows:
            r0, c0, r1, c1 = self.range
            r1 = min(r1, len(self.rows) - 1)
        else:
            r0, c0 = self.sel_row, self.sel_col
        # 剪贴板只有一个值时，填满整个选区（与 Excel 一致）
        if r1 is not None and (r1 > r0 or c1 > c0) \
                and len(data) == 1 and len(data[0]) <= 1:
            val = data[0][0].strip() if data[0] else ""
            self._notify_before()
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    col = self.cols[cc]
                    if col.get("edit"):
                        self.rows[rr][col["key"]] = val
            self.sel_row, self.sel_col = r0, c0
            self.reload()
            self._selection_changed()
            return "break"
        self._notify_before()
        need = r0 + len(data)
        while len(self.rows) < need:
            self.rows.append(self._new_row())
        max_c = c0
        for dr, cells in enumerate(data):
            rr = r0 + dr
            if rr >= len(self.rows):
                break
            for dc, val in enumerate(cells):
                cc = c0 + dc
                if cc >= len(self.cols):
                    break
                col = self.cols[cc]
                if not col.get("edit"):
                    continue
                self.rows[rr][col["key"]] = val.strip()
                max_c = max(max_c, cc)
        last_r = min(r0 + len(data) - 1, len(self.rows) - 1)
        self.sel_row, self.sel_col = r0, c0
        if last_r > r0 or max_c > c0:
            self.range = (r0, c0, last_r, max_c)
        else:
            self.range = None
        self.reload()
        self._selection_changed()
        return "break"

    def ensure_visible(self, r, c):
        n = max(1, len(self.rows))
        total_h = n * self.row_h
        top = self.body.canvasy(0)
        vh = self.body.winfo_height()
        y0 = r * self.row_h
        y1 = y0 + self.row_h
        if y0 < top:
            self.body.yview_moveto(max(0.0, y0 / total_h))
        elif y1 > top + vh:
            self.body.yview_moveto(max(0.0, (y1 - vh) / total_h))
        # 横向：冻结列始终可见，其余列调整偏移
        if c >= self.frozen_n():
            vw = self.view_w()
            x0 = self.xs[c]
            x1 = x0 + self.widths[c]
            if x0 < self.xoff:
                self._set_xoff(int(x0))
            elif x1 > self.xoff + vw:
                self._set_xoff(int(x1 - vw))

    # ---------------------------------------------------------- 编辑 --
    def begin_edit(self, initial=None):
        if not self.rows:
            return
        col = self.cols[self.sel_col]
        if not col.get("edit"):
            return
        self._commit_edit(silent=True)
        i, j = self.sel_row, self.sel_col
        raw = str(self.rows[i].get(col["key"], "") or "")
        if initial is None:
            text = raw
            select = True
        else:
            text = initial
            select = False
        numeric = (col["kind"] in ("num", "price", "percent", "money")
                   or bool(col.get("pct")) or bool(col.get("numeric")))
        ent = tk.Entry(self.body, font=self.font, bd=1, relief="solid",
                       bg="#FFFFFF", fg=C_TEXT, insertbackground=C_TEXT,
                       highlightthickness=0, justify="right" if numeric else "left")
        ent.insert(0, text)
        if select:
            ent.select_range(0, "end")
        else:
            ent.icursor("end")
        self.editor = ent
        self.edit_pos = (i, j)
        self.edit_win = self.body.create_window(self.col_dx(j), i * self.row_h,
                                                anchor="nw", window=ent,
                                                width=self.widths[j],
                                                height=self.row_h)
        ent.focus_set()
        ent.bind("<Return>", lambda e: self._commit_and_move(1, 0))
        ent.bind("<KP_Enter>", lambda e: self._commit_and_move(1, 0))
        ent.bind("<Tab>", lambda e: self._commit_and_move(0, 1))
        ent.bind("<Shift-Tab>", lambda e: self._commit_and_move(0, -1))
        ent.bind("<Escape>", lambda e: self._cancel_edit())
        ent.bind("<FocusOut>", lambda e: self._commit_edit(silent=True))

    def commit_pending(self):
        self._commit_edit(silent=True)

    def _commit_and_move(self, dr, dc):
        self._commit_edit(move=(dr, dc))
        return "break"

    def _cancel_edit(self):
        self._commit_edit(cancel=True)
        self.body.focus_set()
        return "break"

    def _commit_edit(self, move=None, cancel=False, silent=False):
        if self.editor is None:
            return
        ent = self.editor
        win = self.edit_win
        pos = self.edit_pos
        self.editor = None
        self.edit_win = None
        self.edit_pos = None
        try:
            value = ent.get()
        except Exception:
            value = None
        try:
            ent.destroy()
        except Exception:
            pass
        if win is not None:
            try:
                self.body.delete(win)
            except Exception:
                pass
        changed = False
        if not cancel and value is not None and pos is not None and self.rows:
            r, c = pos
            if 0 <= r < len(self.rows) and 0 <= c < len(self.cols):
                col = self.cols[c]
                row = self.rows[r]
                if col.get("edit") and str(row.get(col["key"], "") or "") != value:
                    self._notify_before()
                    row[col["key"]] = value
                    changed = True
                    if self.on_change:
                        self.on_change()
        if move is not None:
            n = len(self.rows)
            r = min(max(0, self.sel_row + move[0]), n - 1)
            c = min(max(0, self.sel_col + move[1]), len(self.cols) - 1)
            self.sel_row, self.sel_col = r, c
            self.ensure_visible(r, c)
            self._draw_body()
            self.body.focus_set()
        elif changed or not silent:
            self._draw_body()
