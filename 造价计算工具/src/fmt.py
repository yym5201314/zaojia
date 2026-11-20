# -*- coding: utf-8 -*-
"""数值解析与格式化（两张表共用）"""


def _to_float(text):
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    s = str(text).strip()
    if not s:
        return None
    s = s.replace(",", "").replace("，", "").replace(" ", "").replace("　", "")
    s = s.replace("％", "%")
    if s.endswith("%"):
        s = s[:-1]
        try:
            return float(s) / 100.0
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_num(text):
    """普通数值解析（空白返回 None）。"""
    return _to_float(text)


def parse_pct(text):
    """百分比解析：15% / 15 / 0.15 都按 15% 处理；小于 1 的数按小数形式。"""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        v = float(text)
        return v if abs(v) < 1 else v / 100.0
    s = str(text).strip().replace("％", "%").replace(" ", "")
    if not s:
        return None
    has_pct = s.endswith("%")
    if has_pct:
        s = s[:-1]
    s = s.replace(",", "").replace("，", "")
    try:
        v = float(s)
    except ValueError:
        return None
    if has_pct:
        return v / 100.0
    return v if abs(v) < 1 else v / 100.0


def _trim(s, min_dp=2):
    if "." in s:
        head, dec = s.split(".", 1)
    else:
        head, dec = s, ""
    dec = dec.rstrip("0")
    if len(dec) < min_dp:
        dec = dec.ljust(min_dp, "0")
    return head + "." + dec


def fmt_qty(v):
    if v is None:
        return ""
    return _trim("%.3f" % v, 2)


def fmt_price(v):
    if v is None:
        return ""
    return _trim("%.4f" % v, 2)


def fmt_money(v):
    if v is None:
        return ""
    return "{:,.2f}".format(v)


def fmt_pct(v):
    if v is None:
        return ""
    s = ("%.4f" % (v * 100.0)).rstrip("0").rstrip(".")
    if s in ("", "-0"):
        s = "0"
    return s + "%"


def fmt_by_kind(kind, v):
    if kind == "money":
        return fmt_money(v)
    if kind == "percent":
        return fmt_pct(v)
    if kind == "price":
        return fmt_price(v)
    if kind == "num":
        return fmt_qty(v)
    return "" if v is None else str(v)
