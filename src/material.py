# -*- coding: utf-8 -*-
"""表二：材料调差表 —— 计算核心

Excel 列对照（B~Q 为打印区域）：

    B 序号     C 材料名称     D 单位      E 本期数量
    F 风险幅度值（项目统一，个别材料可在该行单独填写）
    G 基准价    H 投标价/中标价
    I 结算期信息价（= 各月信息价平均值）
    J..  施工期各月信息价（列数可变，可插入/删除）
    N 调整值   O 本期调整金额   P 结算价   Q 备注

公式：
    I = AVERAGE(各月信息价)
    N = IF(I < MIN(H*(1-F), G*(1-F)), I - MIN(H*(1-F), G*(1-F)),
        IF(I > MAX(H*(1+F), G*(1+F)), I - MAX(H*(1+F), G*(1+F)), 0))
    O = E * N
    P = H + N
"""

import re

import excel_io
import fmt


# ---------------------------------------------------------- 月份列 ----

def parse_start(text, default_year=2026):
    """解析起始年月，返回 (年, 月)。支持 2026-01 / 2026年1月 / 202601。"""
    s = (text or "").strip()
    m = re.match(r"^(\d{4})\s*[-/.年]\s*(\d{1,2})", s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo
    m = re.match(r"^(\d{4})(\d{2})$", s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo
    return default_year, 1


def month_titles(count, start_text):
    y, m = parse_start(start_text)
    out = []
    for _ in range(max(1, count)):
        out.append("%d年%d月" % (y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def month_keys(count):
    return ["m%d" % (i + 1) for i in range(max(1, count))]


# ---------------------------------------------------------- 列定义 ----

def build_columns(titles):
    cols = [
        {"key": "seq", "title": "序号", "group": "", "width": 44,
         "kind": "seq", "edit": False, "role": "plain"},
        {"key": "name", "title": "材料名称", "group": "", "width": 150,
         "kind": "text", "edit": True, "role": "plain"},
        {"key": "unit", "title": "单位", "group": "", "width": 50,
         "kind": "text", "edit": True, "role": "plain", "center": True},
        {"key": "qty", "title": "本期数量", "group": "", "width": 84,
         "kind": "num", "edit": True, "role": "param"},
        {"key": "flt_own", "title": "风险幅度", "group": "", "width": 78,
         "kind": "text_hint", "edit": True, "role": "param", "pct": True,
         "anchor": "e", "hint": lambda row: str(row.get("flt_hint") or "")},
        {"key": "base", "title": "基准价", "group": "", "width": 78,
         "kind": "price", "edit": True, "role": "param"},
        {"key": "bid", "title": "投标价/中标价", "group": "", "width": 104,
         "kind": "price", "edit": True, "role": "param"},
        {"key": "avg", "title": "信息价平均", "group": "结算期信息价", "width": 96,
         "kind": "price", "edit": False, "role": "result"},
    ]
    for i, title in enumerate(titles):
        # 宽度要容得下「2025年10月」这类表头，避免导出后文字挤压
        cols.append({"key": "m%d" % (i + 1), "title": title, "group": "结算期信息价",
                     "width": 90, "kind": "price", "edit": True, "role": "param"})
    cols.extend([
        {"key": "diff", "title": "调整值", "group": "", "width": 74,
         "kind": "price", "edit": False, "role": "result"},
        {"key": "amount", "title": "本期调整金额", "group": "", "width": 100,
         "kind": "money", "edit": False, "role": "result"},
        {"key": "price", "title": "结算价", "group": "", "width": 82,
         "kind": "price", "edit": False, "role": "result"},
        {"key": "note", "title": "备注", "group": "", "width": 150,
         "kind": "text", "edit": True, "role": "plain"},
    ])
    return cols


def new_row(monkey_keys=None, name="", unit="", qty="", base="", bid="", note=""):
    row = {"name": name, "unit": unit, "qty": qty,
           "base": base, "bid": bid, "note": note,
           "flt": "", "flt_own": "", "flt_hint": "",
           "avg": None, "diff": None, "amount": None, "price": None,
           "lo": None, "hi": None}
    for key in (monkey_keys or month_keys(4)):
        row[key] = ""
    return row


def ensure_month_keys(rows, keys):
    """月份列增减后，保证每行都有对应的键。"""
    for row in rows:
        for key in keys:
            row.setdefault(key, "")
        for key in list(row.keys()):
            if key.startswith("m") and key[1:].isdigit() and key not in keys:
                row.pop(key, None)


# -------------------------------------------------------------- 计算 ----

def compute_all(rows, keys, flt_text):
    f_proj = fmt.parse_pct(flt_text)
    f_proj = 0.0 if f_proj is None else f_proj
    for row in rows:
        # 行内单独填写了风险幅度就用行内的，否则沿用项目统一值
        own = str(row.get("flt_own", "") or "").strip()
        f_own = fmt.parse_pct(own) if own else None
        f = f_own if f_own is not None else f_proj
        row["flt"] = own if f_own is not None else flt_text
        proj = fmt.parse_pct(flt_text)
        # 只有填了本期数量的材料行才显示灰色风险幅度提示（分类行/插行不显示）
        if fmt.parse_num(row.get("qty")) is None:
            row["flt_hint"] = ""
        else:
            row["flt_hint"] = fmt.fmt_pct(f_proj) if proj is not None else flt_text
        # 空白或 0 的月份不计入平均（未实施的月份留空或填 0 即可）
        vals = [fmt.parse_num(row.get(k)) for k in keys]
        vals = [v for v in vals if v is not None and v != 0]
        avg = (sum(vals) / len(vals)) if vals else None
        row["avg"] = avg

        g = fmt.parse_num(row.get("base"))
        h = fmt.parse_num(row.get("bid"))
        diff = amount = price = None
        lo = hi = None
        if avg is not None and g is not None and h is not None:
            lo = min(h * (1 - f), g * (1 - f))
            hi = max(h * (1 + f), g * (1 + f))
            if avg < lo:
                diff = avg - lo
            elif avg > hi:
                diff = avg - hi
            else:
                diff = 0.0
            price = h + diff
            qty = fmt.parse_num(row.get("qty"))
            if qty is not None:
                amount = qty * diff
        row["lo"], row["hi"] = lo, hi
        row["diff"], row["amount"], row["price"] = diff, amount, price
    return rows


def totals(rows):
    return sum(r["amount"] for r in rows if r.get("amount") is not None)


def is_empty_row(row):
    keys = ["name", "unit", "qty", "base", "bid", "note"]
    keys += [k for k in row.keys() if k.startswith("m") and k[1:].isdigit()]
    return not any(str(row.get(k, "") or "").strip() for k in keys)


# -------------------------------------------------------------- 导出 ----

def export_table(path, project_name, rows, keys, titles, flt_text):
    """导出 xlsx：不含「计算规则」列（规则见使用说明）。"""
    columns = build_columns(titles)
    widths = [float(c["width"]) / 7.6 for c in columns]

    data = []
    idx = 0
    for row in rows:
        if is_empty_row(row):
            continue
        idx += 1
        note = str(row.get("note", "") or "").strip()
        line = [idx, row.get("name", ""), row.get("unit", ""),
                fmt.parse_num(row.get("qty")), fmt.parse_pct(row.get("flt")),
                fmt.parse_num(row.get("base")), fmt.parse_num(row.get("bid")),
                row.get("avg")]
        for key in keys:
            line.append(fmt.parse_num(row.get(key)))
        line.extend([row.get("diff"), row.get("amount"), row.get("price"), note])
        data.append(line)

    first_month = next((i for i, c in enumerate(columns) if c["key"] == keys[0]), 8)
    last_month = first_month + len(keys) - 1
    excel_io.write_sheet(
        path, "材料调差",
        "材料调差表",
        "项目名称：%s" % (project_name or ""),
        "单位：元",
        columns, data,
        col_widths=widths,
        print_area=("A", excel_io._col_letter(len(columns) - 1)),
        outline_cols=(first_month, last_month),
    )
