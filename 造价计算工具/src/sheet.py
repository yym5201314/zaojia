# -*- coding: utf-8 -*-
"""通用表格控件（自绘 Canvas：可编辑、可拖选、隔行底色、黄参数/绿结果）"""

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
        self.head.configure(height=self.head_h, scrollregion=(0, 0, self.total_w, self.head_h))
        self.redraw()

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
        self.body.configure(yscrollcommand=self._on_yview, xscrollcommand=self._on_xview)

        self.body.bind("<Button-1>", self._on_press)
        self.body.bind("<B1-Motion>", self._on_drag)
        self.body.bind("<ButtonRelease-1>", self._on_release)
        self.body.bind("<Double-Button-1>", self._on_dblclick)
        self.body.bind("<Configure>", lambda e: self._draw_body())
        for w in (self.body, self.head):
            w.bind("<MouseWheel>", self._on_wheel)
            w.bind("<Shift-MouseWheel>", self._on_wheel)
        self.head.bind("<Configure>", lambda e: self._draw_head())
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
        j = self._edge_at(self.head.canvasx(event.x))
        try:
            self.head.configure(cursor="sb_h_double_arrow" if j is not None else "")
        except Exception:
            pass
        return "break"

    def _on_head_press(self, event):
        x = self.head.canvasx(event.x)
        j = self._edge_at(x)
        self._resize = (j, x, self.widths[j]) if j is not None else None
        return "break"

    def _on_head_drag(self, event):
        if self._resize is None:
            return "break"
        j, x0, w0 = self._resize
        dx = self.head.canvasx(event.x) - x0
        w = int(max(36, min(1200, w0 + dx)))
        if w == self.widths[j]:
            return "break"
        key = self.cols[j].get("key")
        if key:
            self._width_override[key] = w
        self._calc_layout()
        self.head.configure(scrollregion=(0, 0, self.total_w, self.head_h))
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
        self.body.xview(*args)
        self.head.xview(*args)

    def _on_yview(self, first, last):
        self.vbar.set(first, last)
        self._draw_body()

    def _on_xview(self, first, last):
        self.hbar.set(first, last)
        self.head.xview_moveto(first)

    def _on_wheel(self, event):
        step = -1 * int(event.delta / 120) if event.delta else 0
        if not step:
            step = -1 if event.delta > 0 else 1
        if event.state & 0x0001:
            self.body.xview_scroll(step, "units")
            self.head.xview_scroll(step, "units")
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
        self.body.configure(scrollregion=(0, 0, self.total_w, h))
        self.head.configure(scrollregion=(0, 0, self.total_w, self.head_h))
        self._draw_head()
        self._draw_body()

    # ---------------------------------------------------------- 绘制 --
    def _draw_head(self):
        c = self.head
        c.delete("all")
        i = 0
        n = len(self.cols)
        try:
            vw = c.winfo_width()
            left = c.canvasx(0)
        except Exception:
            vw, left = 1000, 0
        right = left + max(vw, 200)
        while i < n:
            col = self.cols[i]
            group = col.get("group") or ""
            j = i
            while group and j + 1 < n and (self.cols[j + 1].get("group") or "") == group:
                j += 1
            x0 = self.xs[i]
            x1 = self.xs[j] + self.widths[j]
            if x1 < left - 40 or x0 > right + 40:
                i = j + 1
                continue
            if group:
                c.create_rectangle(x0, 0, x1, self.group_h, fill="#CFE2F5", outline=C_LINE)
                c.create_text((x0 + x1) / 2, self.group_h / 2, text=group,
                              fill=C_HEAD_TX, font=self.font_b)
                for k in range(i, j + 1):
                    cx0 = self.xs[k]
                    cx1 = cx0 + self.widths[k]
                    c.create_rectangle(cx0, self.group_h, cx1, self.head_h,
                                       fill=C_HEAD, outline=C_LINE)
                    c.create_text((cx0 + cx1) / 2, (self.group_h + self.head_h) / 2,
                                  text=self.cols[k]["title"], fill=C_HEAD_TX, font=self.font_b)
            else:
                c.create_rectangle(x0, 0, x1, self.head_h, fill=C_HEAD, outline=C_LINE)
                c.create_text((x0 + x1) / 2, self.head_h / 2, text=col["title"],
                              fill=C_HEAD_TX, font=self.font_b)
            i = j + 1

    def _draw_body(self):
        c = self.body
        c.delete("cell")
        n = len(self.rows)
        try:
            top = c.canvasy(0)
            vh = c.winfo_height()
        except Exception:
            top, vh = 0, 600
        start = max(0, int(top // self.row_h) - 1)
        end = min(n, int((top + vh) // self.row_h) + 2)
        for i in range(start, end):
            y0 = i * self.row_h
            y1 = y0 + self.row_h
            row = self.rows[i]
            for j, col in enumerate(self.cols):
                x0 = self.xs[j]
                x1 = x0 + self.widths[j]
                c.create_rectangle(x0, y0, x1, y1, fill=self._bg(i, col),
                                   outline=C_LINE, tags="cell")
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
                              fill=color, font=self.font, tags="cell")
        if self.range and n:
            r0, c0, r1, c1 = self.range
            r1 = min(r1, n - 1)
            c.create_rectangle(self.xs[c0], r0 * self.row_h,
                               self.xs[c1] + self.widths[c1], (r1 + 1) * self.row_h,
                               outline=C_SEL, width=1, fill=C_SEL,
                               stipple="gray25", tags="cell")
        if n and 0 <= self.sel_row < n:
            x0 = self.xs[self.sel_col]
            x1 = x0 + self.widths[self.sel_col]
            y0 = self.sel_row * self.row_h
            c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y0 + self.row_h - 1,
                               outline=C_SEL, width=2, tags="cell")
        if self.edit_win is not None:
            try:
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

    def _text(self, i, row, col):
        key = col["key"]
        kind = col["kind"]
        if kind == "seq":
            return str(i + 1), C_TEXT, "center"
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
        cx = self.body.canvasx(x)
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
        cx = self.body.canvasx(x)
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
            return str(r + 1)
        if col["kind"] in ("text", "text_hint"):
            return str(row.get(col["key"], "") or "")
        if col.get("role") == "result":
            return fmt.fmt_by_kind(col["kind"], row.get(col["key"]))
        return str(row.get(col["key"], "") or "")

    def _notify_before(self):
        if self.before_change:
            self.before_change()

    def _on_press(self, event):
        was_editing = self.editor is not None
        if was_editing:
            self._commit_edit(silent=True)
        self.body.focus_set()
        self._press_sel = None if was_editing else (self.sel_row, self.sel_col)
        hit = self._hit(event.x, event.y)
        self._anchor = hit
        if not hit:
            return "break"
        r, j = hit
        self.sel_row, self.sel_col = r, j
        self.range = None
        self._draw_body()
        self._selection_changed()
        return "break"

    def _on_drag(self, event):
        if self._anchor is None:
            return "break"
        if event.y < 0:
            self.body.yview_scroll(-1, "units")
        elif event.y > self.body.winfo_height():
            self.body.yview_scroll(1, "units")
        if event.x < 0:
            self.body.xview_scroll(-1, "units")
        elif event.x > self.body.winfo_width():
            self.body.xview_scroll(1, "units")
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
        left = self.body.canvasx(0)
        vw = self.body.winfo_width()
        x0 = self.xs[c]
        x1 = x0 + self.widths[c]
        if x0 < left:
            self.body.xview_moveto(max(0.0, x0 / self.total_w))
        elif x1 > left + vw:
            self.body.xview_moveto(max(0.0, (x1 - vw) / self.total_w))

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
        self.edit_win = self.body.create_window(self.xs[j], i * self.row_h, anchor="nw",
                                                window=ent, width=self.widths[j],
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
