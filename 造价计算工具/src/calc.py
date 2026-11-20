# -*- coding: utf-8 -*-
"""表一：量增量减费用调整表 —— 计算核心

与 Excel 公式一一对应（以第 5 行为例）：

    F5 = IF(D5>(1+J5)*E5, D5-(1+J5)*E5, IF(D5<E5*(1-J5), D5, 0))
    I5 = IF(AND(D5>E5*(1+J5), H5<G5*(1+J5)), H5,
            IF(AND(D5<E5*(1-J5), H5>G5*(1-J5)*(1-K5)), H5,
            IF(H5<G5*(1-K5)*(1-J5), G5*(1-K5)*(1-J5),
            IF(H5>G5*(1+J5), G5*(1+J5), H5))))
    L5 = IF(D5>E5, (D5-F5)*H5+F5*I5, IF(D5<E5, F5*I5, 0))
    M5 = E5*H5

列含义：
    D 结算量 Q1      E 招标清单量 Q0    F 需调价量（结果）
    G 招标单价 P2    H 中标单价 P0      I 结算单价 P1（结果）
    J 清单规定偏差   K 报价浮动率       L 调整合价（结果）   M 中标价（结果）
"""

import excel_io
import fmt

# ---------------------------------------------------------------- 列定义 ----

COLUMNS = [
    {"key": "seq",    "title": "序号",       "group": "",     "width": 50,  "kind": "seq",     "edit": False, "role": "plain"},
    {"key": "name",   "title": "名称",       "group": "",     "width": 186, "kind": "text",    "edit": True,  "role": "plain"},
    {"key": "unit",   "title": "单位",       "group": "",     "width": 56,  "kind": "text",    "edit": True,  "role": "plain", "center": True},
    {"key": "q1",     "title": "结算量Q1",   "group": "工程量", "width": 90,  "kind": "num",     "edit": True,  "role": "param"},
    {"key": "q0",     "title": "招标清单量Q0", "group": "工程量", "width": 100, "kind": "num",     "edit": True,  "role": "param"},
    {"key": "f",      "title": "需调价量",   "group": "工程量", "width": 86,  "kind": "num",     "edit": False, "role": "result"},
    {"key": "p2",     "title": "招标单价P2", "group": "单价",   "width": 94,  "kind": "price",   "edit": True,  "role": "param"},
    {"key": "p0",     "title": "中标单价P0", "group": "单价",   "width": 94,  "kind": "price",   "edit": True,  "role": "param"},
    {"key": "p1",     "title": "结算单价P1", "group": "单价",   "width": 94,  "kind": "price",   "edit": False, "role": "result"},
    {"key": "dev",    "title": "清单规定偏差J", "group": "单价", "width": 96,  "kind": "percent", "edit": True,  "role": "param"},
    {"key": "flt",    "title": "报价浮动率L", "group": "单价",  "width": 88,  "kind": "percent", "edit": True,  "role": "param"},
    {"key": "amount", "title": "调整合价",   "group": "合价",   "width": 118, "kind": "money",   "edit": False, "role": "result"},
    {"key": "bid",    "title": "中标价",     "group": "合价",   "width": 108, "kind": "money",   "edit": False, "role": "result"},
    {"key": "note",   "title": "备注",       "group": "",     "width": 108, "kind": "text",    "edit": True,  "role": "plain"},
]

INPUT_KEYS = ["q1", "q0", "p2", "p0", "dev", "flt"]

COL_INDEX = {c["key"]: i for i, c in enumerate(COLUMNS)}

# 兼容旧接口
parse_num = fmt.parse_num
parse_pct = fmt.parse_pct


# -------------------------------------------------------------- 计算 ----

def new_row(name="", unit="", q1="", q0="", p2="", p0="", dev="", flt="", note=""):
    return {
        "name": name, "unit": unit,
        "q1": q1, "q0": q0, "p2": p2, "p0": p0,
        "dev": dev, "flt": flt, "note": note,
        "f": None, "p1": None, "amount": None, "bid": None,
    }


def compute_row(row):
    """按 Excel 公式计算 F / I / L / M 四列，返回 (f, p1, amount, bid)。"""
    q1 = fmt.parse_num(row.get("q1"))
    q0 = fmt.parse_num(row.get("q0"))
    p2 = fmt.parse_num(row.get("p2"))
    p0 = fmt.parse_num(row.get("p0"))
    j = fmt.parse_pct(row.get("dev"))
    k = fmt.parse_pct(row.get("flt"))
    j = 0.0 if j is None else j
    k = 0.0 if k is None else k

    f = p1 = amount = bid = None

    # F 需调价量
    if q1 is not None and q0 is not None:
        if q1 > (1 + j) * q0:
            f = q1 - (1 + j) * q0
        elif q1 < q0 * (1 - j):
            f = q1
        else:
            f = 0.0

    # I 结算单价
    if q1 is not None and q0 is not None and p2 is not None and p0 is not None:
        if q1 > q0 * (1 + j) and p0 < p2 * (1 + j):
            p1 = p0
        elif q1 < q0 * (1 - j) and p0 > p2 * (1 - j) * (1 - k):
            p1 = p0
        elif p0 < p2 * (1 - k) * (1 - j):
            p1 = p2 * (1 - k) * (1 - j)
        elif p0 > p2 * (1 + j):
            p1 = p2 * (1 + j)
        else:
            p1 = p0

    # L 调整合价
    if q1 is not None and q0 is not None and f is not None and p1 is not None and p0 is not None:
        if q1 > q0:
            amount = (q1 - f) * p0 + f * p1
        elif q1 < q0:
            amount = f * p1
        else:
            amount = 0.0

    # M 中标价
    if q0 is not None and p0 is not None:
        bid = q0 * p0

    return f, p1, amount, bid


def compute_all(rows, flt_override=None):
    for row in rows:
        if flt_override is not None:
            row["flt"] = flt_override
        row["f"], row["p1"], row["amount"], row["bid"] = compute_row(row)
    return rows


def totals(rows):
    ta = sum(r["amount"] for r in rows if r.get("amount") is not None)
    tb = sum(r["bid"] for r in rows if r.get("bid") is not None)
    return ta, tb


def is_empty_row(row):
    return not any(
        str(row.get(k, "") or "").strip()
        for k in ("name", "unit", "q1", "q0", "p2", "p0", "dev", "flt", "note")
    )


# -------------------------------------------------------------- 导出 ----

_EXPORT_WIDTHS = [6.5, 30.0, 7.0, 10.0, 12.0, 10.0, 11.0, 11.0, 11.0, 12.5, 11.5, 14.0, 13.0, 16.0]


def export_table(path, project_name, rows):
    """导出「量增量减费用调整表」xlsx。"""
    data = []
    idx = 0
    for row in rows:
        if is_empty_row(row):
            continue
        idx += 1
        data.append([
            idx,
            row.get("name", ""),
            row.get("unit", ""),
            fmt.parse_num(row.get("q1")),
            fmt.parse_num(row.get("q0")),
            row.get("f"),
            fmt.parse_num(row.get("p2")),
            fmt.parse_num(row.get("p0")),
            row.get("p1"),
            fmt.parse_pct(row.get("dev")),
            fmt.parse_pct(row.get("flt")),
            row.get("amount"),
            row.get("bid"),
            row.get("note", ""),
        ])
    excel_io.write_sheet(
        path, "调整表",
        "量增量减费用调整表",
        "项目名称：%s" % (project_name or ""),
        "单位：元",
        COLUMNS, data,
        col_widths=_EXPORT_WIDTHS,
        print_area=("A", _col_letter(len(COLUMNS) - 1)),
    )


def _col_letter(idx):
    s = ""
    n = idx + 1
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s
