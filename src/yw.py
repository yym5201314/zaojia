# -*- coding: utf-8 -*-
"""表三：雨污水工程（沟槽土方）计算核心

与原表《2021.4.12-雨污水工程计算式（坡比调整）》逐列对应：

    A 序号            B 桩号          C 井号         D 检查井型号
    E 判断挖填方段   F 是否车行道     G 原地面高程    H 换填厚度（本程序新增）
    I 竣工路基高程   J 清表后高程     K 设计管内底高程 L 沟槽抄测底高程
    M 设计沟槽底高程 N 设计路面高程   O 检查井高度    P 管顶覆土深度
    Q 路基至沟槽底高度 R 挖土深度     S 管线中心长度  T 扣减长度
    U 管道净长       V 管径          W 管道级别      X 管道内径
    Y 壁厚           Z C1            AA C2           AB 工作面a
    AC 底宽          AD 放坡系数      AE 挖沟槽断面积 AF 挖方
    AG 180度砂石基础面积 AH 管道面积  AI 砂石基础体积 AJ 管道体积
    AK 回填连砂石面积 AL 回填连砂石体积 AM 管顶50cm以上回填土量
    AN 井个数        AO 砼井周加强    AP 井室高度    AQ 井筒高度
    AR 满足最小井筒高度 AS 井室计算高度 AT 井底厚度h1

两条关键规则（原表靠人工保证，这里由程序判断）：
1. 主管：相邻两行（前一检查井与本井）断面积平均 × 管线中心长度；
2. 支管：段内每一行都与「段首行」平均（支管从主管井引出，起点断面相同）。
   R（管线中心长度）为空的行只是段首基准行，不出体积；
   「断开」行之后的第一个数据行重新起段，不跨断开做平均。
"""

import excel_io
import fmt

ROAD_THICK = 1.66          # 路面结构层厚 0.86+0.8
CLEAR_THICK = 0.5          # 清表厚度（清表后高程 = 原地面高程 - 该值）
BACK_CUT = 0.5             # 管顶反开挖深度（填方段挖土深度用，随项目不同）
MIN_SHAFT = 0.4            # 最小井筒高度
WELL_BOTTOM = 0.25         # 井底厚度 h1（AT 默认值）


def is_branch(row):
    """支管行：井号带 “-”（如 W凤1-1）。"""
    return "-" in str(row.get("well", "") or "")


def is_break(row):
    return row.get("kind") == "break"


# ------------------------------------------------------------ 列定义 ----

def _col(key, title, width, kind, edit, role):
    return {"key": key, "title": title, "group": "", "width": width,
            "kind": kind, "edit": edit, "role": role}


COLUMNS = [
    {"key": "seq", "title": "序号", "group": "", "width": 44,
     "kind": "seq", "edit": False, "role": "plain"},
    _col("station", "桩号", 84, "text", True, "plain"),
    _col("well", "井号", 84, "text", True, "plain"),
    _col("well_type", "检查井型号", 96, "text", True, "plain"),
    _col("cut_fill", "判断挖填方段", 86, "text", False, "result"),
    _col("road", "是否车行道", 78, "text", True, "plain"),
    _col("ground", "原地面高程", 88, "price", True, "param"),
    _col("fill_thick", "换填厚度", 78, "num", True, "param"),
    _col("road_base", "竣工路基高程", 92, "price", False, "result"),
    _col("clear", "清表后高程", 84, "price", False, "result"),
    _col("pipe_bottom", "设计管内底高程", 100, "price", True, "param"),
    {"key": "trench_bottom_in", "title": "沟槽抄测底高程", "group": "", "width": 104,
     "kind": "text_hint", "edit": True, "role": "param", "numeric": True,
     "anchor": "e", "hint": lambda row: fmt.fmt_price(row.get("design_bottom"))},
    _col("design_bottom", "设计沟槽底高程", 100, "price", False, "result"),
    _col("road_top", "设计路面高程", 92, "price", True, "param"),
    _col("well_h", "检查井高度", 86, "price", False, "result"),
    _col("cover", "管顶覆土深度", 92, "price", False, "result"),
    _col("base_gap", "路基至沟槽底高度", 108, "price", False, "result"),
    _col("depth", "挖土深度", 80, "price", False, "result"),
    _col("len_c", "管线中心长度", 96, "num", True, "param"),
    _col("len_cut", "扣减长度", 76, "num", True, "param"),
    _col("len_n", "管道净长", 80, "num", False, "result"),
    _col("dia", "管径", 66, "text", True, "plain"),
    _col("grade", "管道级别", 70, "text", True, "plain"),
    _col("inner", "管道内径", 78, "num", True, "param"),
    _col("wall", "壁厚t", 66, "num", True, "param"),
    _col("c1", "C1", 50, "num", True, "param"),
    _col("c2", "C2", 50, "num", True, "param"),
    _col("work", "工作面a", 70, "num", True, "param"),
    _col("width", "底宽", 62, "num", False, "result"),
    _col("slope", "放坡系数", 70, "num", False, "result"),
    _col("area_trench", "挖沟槽断面积", 96, "num", False, "result"),
    _col("vol_cut", "挖方", 86, "num", False, "result"),
    _col("area_base", "180度砂石基础面积", 116, "num", False, "result"),
    _col("area_pipe", "管道面积", 80, "num", False, "result"),
    _col("vol_base", "180度砂石基础体积", 116, "num", False, "result"),
    _col("vol_pipe", "管道体积", 86, "num", False, "result"),
    _col("area_back", "回填连砂石面积", 110, "num", False, "result"),
    _col("vol_back", "回填连砂石体积", 110, "num", False, "result"),
    _col("vol_soil", "管顶50cm以上回填土量", 140, "num", False, "result"),
    _col("well_cnt", "井个数", 70, "num", True, "param"),
    _col("well_ring", "砼井周加强", 100, "num", False, "result"),
    _col("well_room", "井室高度", 80, "num", False, "result"),
    _col("shaft_h", "井筒高度", 80, "num", False, "result"),
    _col("shaft_min", "满足最小井筒高度", 110, "num", False, "result"),
    _col("ring_h", "井室计算高度", 96, "num", False, "result"),
    _col("bot_thick", "井底厚度h1", 84, "num", True, "param"),
]

INPUT_KEYS = ("mark", "station", "well", "well_type", "road", "ground",
              "trench_bottom_in", "pipe_bottom", "road_top", "len_c", "len_cut",
              "dia", "grade", "inner", "wall", "c1", "c2", "work", "well_cnt",
              "bot_thick", "fill_thick")

RESULT_KEYS = tuple(c["key"] for c in COLUMNS if c.get("role") == "result")


def new_row(**kw):
    row = {"kind": "data", "trench_bottom": None}
    for c in COLUMNS:
        row[c["key"]] = None if c.get("role") == "result" else ""
    for k, v in kw.items():
        row[k] = v
    return row


def new_break():
    return {"kind": "break", "mark": "断开"}


def is_empty_row(row):
    if is_break(row):
        return False
    for key in ("well", "station", "ground", "pipe_bottom", "road_top", "len_c"):
        if str(row.get(key, "") or "").strip():
            return False
    return True


# -------------------------------------------------------------- 计算 ----

def _n(v):
    return fmt.parse_num(v)


def _clear_results(row):
    for key in RESULT_KEYS:
        row[key] = None


def _slope(q):
    if q is None:
        return None
    if 1.5 < q < 3:
        return 0.67
    if q > 3:
        return 1.0
    return 0.0


def _section(row, road_thick, well_bottom, clear_thick=CLEAR_THICK,
             back_cut=BACK_CUT):
    """按公式算本行的断面类结果（不涉及上下行平均）。"""
    g = _n(row.get("ground"))
    j = _n(row.get("pipe_bottom"))
    m = _n(row.get("road_top"))
    w = _n(row.get("inner"))
    x = _n(row.get("wall"))
    y = _n(row.get("c1"))
    z = _n(row.get("c2"))
    a = _n(row.get("work"))
    fill = _n(row.get("fill_thick")) or 0.0
    road = str(row.get("road") or "").strip()

    # J 清表后高程：原地面高程 - 项目级清表厚度
    clear = g - clear_thick if g is not None else None
    design = j - x - y if None not in (j, x, y) else None   # M 设计沟槽底
    # L 沟槽抄测底：手填（实测）优先，留空则跟随设计沟槽底
    trench = _n(row.get("trench_bottom_in"))
    if trench is None:
        trench = design

    # I 竣工路基高程：不在车行道上取设计路面高程，否则扣路面结构层厚，两者都扣换填厚度
    if m is None:
        road_base = None
    elif road == "否":
        road_base = m - fill
    else:
        road_base = m - road_thick - fill

    well_h = m - j if None not in (m, j) else None       # O 检查井高度
    cover = None                                        # P 管顶覆土深度
    if None not in (m, trench, y, w, x):
        cover = m - (trench + y + w + x * 2)

    base_gap = None                                     # Q 路基至沟槽底
    if None not in (road_base, clear, trench):
        base_gap = (clear - trench) if road_base > clear else (road_base - trench)

    cut_fill = None
    depth = None
    if None not in (road_base, clear, trench, w, x, j):
        # 填方段要考虑管顶反开挖：内径 + 壁厚 + 反开挖深度 + 管内底 - 抄测底
        edge = w + x + back_cut + j - trench
        if road_base > clear and (clear - trench) > edge:
            cut_fill, depth = "挖方段", clear - trench
        elif clear > road_base and (road_base - trench) > edge:
            cut_fill, depth = "挖方段", road_base - trench
        else:
            cut_fill, depth = "填方段", edge

    width = (w + x * 2 + a * 2) if None not in (w, x, a) else None
    slope = _slope(depth)
    area_trench = None
    if None not in (width, depth, slope):
        area_trench = (width + depth * slope) * depth

    area_pipe = None
    if None not in (w, x):
        area_pipe = 3.141592653589793 * (w + x * 2) ** 2 / 4.0

    area_base = None
    if None not in (width, y, z, slope, area_pipe):
        area_base = (width + (y + z) * slope) * (y + z) - area_pipe / 2.0

    area_back = None
    if None not in (width, y, x, w, slope, area_pipe):
        area_back = ((width + (y + x * 2 + w + 0.5) * slope)
                     * (y + x * 2 + w + 0.5) - area_pipe / 2.0)

    # 井周加强
    bot = _n(row.get("bot_thick"))
    bot = well_bottom if bot is None else bot
    well_room = (w + 1.8 + 0.1) if w is not None else None
    shaft_h = (well_h - well_room) if None not in (well_h, well_room) else None
    shaft_min = None
    if shaft_h is not None:
        shaft_min = shaft_h if shaft_h > MIN_SHAFT else MIN_SHAFT
    ring_h = None
    if None not in (well_h, shaft_min, well_room):
        if abs(shaft_min - MIN_SHAFT) < 1e-9:
            ring_h = well_h + bot - shaft_min
        else:
            ring_h = well_room + bot
    # 砼井周加强：只有车行道上的检查井才做（原表按「是否车行道」判断）
    well_ring = None
    if road == "是" and None not in (ring_h, shaft_min):
        well_ring = (2 * 3.141592653589793 * 1.2 * ring_h
                     + 2 * 3.141592653589793 * 0.95
                     * (shaft_min - 0.16 - WELL_BOTTOM / 2.0))

    row.update({
        "clear": clear, "design_bottom": design, "trench_bottom": trench,
        "road_base": road_base, "well_h": well_h, "cover": cover,
        "base_gap": base_gap, "cut_fill": cut_fill, "depth": depth,
        "width": width, "slope": slope, "area_trench": area_trench,
        "area_pipe": area_pipe, "area_base": area_base,
        "area_back": area_back, "well_room": well_room,
        "shaft_h": shaft_h, "shaft_min": shaft_min, "ring_h": ring_h,
        "well_ring": well_ring,
    })


def _volume(row, base):
    """按基准行做平均面积法，算挖方/基础/管道/回填。"""
    r = _n(row.get("len_c"))
    s = _n(row.get("len_cut")) or 0.0
    if r is None or base is None:
        row.update({"len_n": None, "vol_cut": None, "vol_base": None,
                    "vol_pipe": None, "vol_back": None, "vol_soil": None})
        return
    t = r - s
    row["len_n"] = t

    ad, b_ad = row.get("area_trench"), base.get("area_trench")
    af, b_af = row.get("area_base"), base.get("area_base")
    aj, b_aj = row.get("area_back"), base.get("area_back")
    ag = row.get("area_pipe")

    vol_cut = (b_ad + ad) / 2.0 * r if None not in (ad, b_ad) else None
    vol_base = (b_af + af) / 2.0 * t if None not in (af, b_af) else None
    vol_pipe = ag * t if ag is not None else None

    vol_back = None
    if row.get("cut_fill") == "填方段":
        if None not in (vol_cut, vol_base, vol_pipe):
            vol_back = vol_cut - vol_base - vol_pipe
    else:
        if None not in (aj, b_aj, vol_base):
            vol_back = (b_aj + aj) / 2.0 * t - vol_base

    vol_soil = None
    if None not in (vol_cut, vol_base, vol_pipe, vol_back):
        vol_soil = vol_cut - vol_base - vol_pipe - vol_back

    row.update({"len_n": t, "vol_cut": vol_cut, "vol_base": vol_base,
                "vol_pipe": vol_pipe, "vol_back": vol_back,
                "vol_soil": vol_soil})


def compute_all(rows, road_thick=ROAD_THICK, well_bottom=WELL_BOTTOM,
                clear_thick=CLEAR_THICK, back_cut=BACK_CUT):
    for row in rows:
        if is_break(row):
            _clear_results(row)
            continue
        _section(row, road_thick, well_bottom, clear_thick, back_cut)

    seg_i = None    # 当前段首行下标
    prev_i = None   # 上一个数据行下标
    for i, row in enumerate(rows):
        if is_break(row):
            seg_i = prev_i = None
            row["_base_i"] = None
            continue
        if _n(row.get("len_c")) is None:
            # 段首行：只提供起点断面，不出体积
            row.update({"len_n": None, "vol_cut": None, "vol_base": None,
                        "vol_pipe": None, "vol_back": None, "vol_soil": None})
            row["_base_i"] = None
            seg_i = prev_i = i
            continue
        base_i = seg_i if (is_branch(row) and seg_i is not None) else prev_i
        _volume(row, rows[base_i] if base_i is not None else None)
        row["_base_i"] = base_i
        prev_i = i
    return rows


# -------------------------------------------------------------- 汇总 ----

def segments(rows):
    """按段首（R 为空）切段：返回 [("break", None) 或 ("seg", [行...])]。

    段内出现带 “-” 的井号即为支管段（支管从主管井引出）。
    """
    segs = []
    cur = []
    for row in rows:
        if is_break(row):
            if cur:
                segs.append(("seg", cur))
                cur = []
            segs.append(("break", None))
            continue
        if is_empty_row(row):
            continue
        if not cur or _n(row.get("len_c")) is None:
            if cur:
                segs.append(("seg", cur))
            cur = [row]
        else:
            cur.append(row)
    if cur:
        segs.append(("seg", cur))
    return segs


def _seg_is_branch(seg):
    return any("-" in str(r.get("well") or "") for r in seg)


def _q(v):
    return "—" if v is None else fmt.fmt_qty(v)


def formula_text(row, key, rows=None, road_thick=ROAD_THICK,
                 clear_thick=CLEAR_THICK, back_cut=BACK_CUT):
    """结果列的计算式（带本行实际数值），用于选中单元格时显示。"""
    if row is None:
        return ""
    g = _n(row.get("ground"))
    j = _n(row.get("pipe_bottom"))
    m = _n(row.get("road_top"))
    w = _n(row.get("inner"))
    x = _n(row.get("wall"))
    y = _n(row.get("c1"))
    z = _n(row.get("c2"))
    a = _n(row.get("work"))
    fill = _n(row.get("fill_thick")) or 0.0
    road = str(row.get("road") or "").strip()
    clear = row.get("clear")
    design = row.get("design_bottom")
    trench = row.get("trench_bottom")
    road_base = row.get("road_base")
    well_h = row.get("well_h")
    depth = row.get("depth")
    width = row.get("width")
    slope = row.get("slope")
    area_pipe = row.get("area_pipe")
    base = None
    if rows and isinstance(row.get("_base_i"), int):
        bi = row["_base_i"]
        if 0 <= bi < len(rows):
            base = rows[bi]

    if key == "clear":
        return ("清表后高程 = 原地面高程 %s - 清表厚度 %s = %s"
                % (_q(g), _q(clear_thick), _q(clear)))
    if key == "design_bottom":
        return ("设计沟槽底 = 设计管内底 %s - C1 %s - C2 %s = %s"
                % (_q(j), _q(y), _q(z), _q(design)))
    if key == "well_h":
        return "检查井高度 = 设计路面 %s - 设计管内底 %s = %s" % (_q(m), _q(j), _q(well_h))
    if key == "road_base":
        if road == "否":
            return ("竣工路基高程 = 设计路面 %s - 换填厚度 %s = %s（是否车行道=否，不扣路面结构层）"
                    % (_q(m), _q(fill), _q(road_base)))
        return ("竣工路基高程 = 设计路面 %s - 路面结构层厚 %s - 换填厚度 %s = %s"
                % (_q(m), _q(road_thick), _q(fill), _q(road_base)))
    if key == "cover":
        return ("管顶覆土深度 = 设计路面 %s - (抄测底 %s + C2 %s + 内径 %s + 2×壁厚 %s) = %s"
                % (_q(m), _q(trench), _q(z), _q(w), _q(x), _q(row.get("cover"))))
    if key == "base_gap":
        return ("路基至沟槽底 = IF(路基 %s > 清表 %s, 清表-抄测底, 路基-抄测底) = %s"
                % (_q(road_base), _q(clear), _q(row.get("base_gap"))))
    if key == "depth":
        if row.get("cut_fill") == "填方段":
            return ("挖土深度 = 内径 %s + 壁厚 %s + 反开挖 %s + 管内底 %s - 抄测底 %s = %s（填方段，含管顶反开挖）"
                    % (_q(w), _q(x), _q(back_cut), _q(j), _q(trench), _q(depth)))
        return ("挖土深度 = 路基(或清表) %s - 抄测底 %s = %s（挖方段）"
                % (_q(road_base), _q(trench), _q(depth)))
    if key == "cut_fill":
        if None in (road_base, clear, trench, w, x, j):
            return "判断挖填方段：原地面/管内底/路面等数据不全，暂无法判断"
        need = w + x + back_cut + j - trench
        d_qb = clear - trench          # 清表后高程 - 沟槽抄测底
        d_lu = road_base - trench      # 竣工路基高程 - 沟槽抄测底
        base = ("判断挖填方段（考虑管顶反开挖）：清表−抄测底 = %s−%s = %s；"
                "路基−抄测底 = %s−%s = %s；填方所需深度 = 内径 %s+壁厚 %s+反开挖 %s"
                "+管内底 %s−抄测底 %s = %s。"
                % (_q(clear), _q(trench), _q(d_qb),
                   _q(road_base), _q(trench), _q(d_lu),
                   _q(w), _q(x), _q(back_cut), _q(j), _q(trench), _q(need)))
        if road_base > clear and d_qb > need:
            return base + ("路基高于清表、且清表−抄测底 %s > 填方所需 %s → 挖方段，挖深取清表−抄测底 %s"
                           % (_q(d_qb), _q(need), _q(d_qb)))
        if clear > road_base and d_lu > need:
            return base + ("清表高于路基、且路基−抄测底 %s > 填方所需 %s → 挖方段，挖深取路基−抄测底 %s"
                           % (_q(d_lu), _q(need), _q(d_lu)))
        return base + ("两个高差都不足以达到填方所需深度 %s → 填方段，挖深取 %s（含反开挖）"
                       % (_q(need), _q(need)))
    if key == "width":
        return ("底宽 = 内径 %s + 2×壁厚 %s + 2×工作面 %s = %s"
                % (_q(w), _q(x), _q(a), _q(width)))
    if key == "slope":
        return ("放坡系数 = IF(1.5<挖深 %s<3, 0.67, IF(挖深>3, 1.0, 0)) = %s"
                % (_q(depth), _q(slope)))
    if key == "area_pipe":
        return ("管道面积 = π×(内径 %s + 2×壁厚 %s)²/4 = %s"
                % (_q(w), _q(x), _q(area_pipe)))
    if key == "area_trench":
        return ("挖沟槽断面积 = (底宽 %s + 挖深 %s×放坡 %s)×挖深 %s = %s"
                % (_q(width), _q(depth), _q(slope), _q(depth), _q(row.get("area_trench"))))
    if key == "area_base":
        return ("180度砂石基础面积 = (底宽 %s + (C2 %s + C1 %s)×放坡 %s)×(C2+C1) - 管道面积 %s/2 = %s"
                % (_q(width), _q(z), _q(y), _q(slope), _q(area_pipe),
                   _q(row.get("area_base"))))
    if key == "area_back":
        return ("回填连砂石面积 = (底宽 %s + (C2+2×壁厚+内径+0.5)×放坡 %s)×(同上) - 管道面积 %s/2 = %s"
                % (_q(width), _q(slope), _q(area_pipe), _q(row.get("area_back"))))
    if key == "len_n":
        return ("管道净长 = 管线中心长度 %s - 扣减长度 %s = %s"
                % (_q(_n(row.get("len_c"))), _q(_n(row.get("len_cut")) or 0.0),
                   _q(row.get("len_n"))))
    if key in ("vol_cut", "vol_base", "vol_pipe", "vol_back", "vol_soil"):
        if row.get("_base_i") is None and key != "vol_soil":
            return "%s：本行是段首行（管线中心长度为空），不出体积" % _COL_TITLE.get(key, key)
        b_ad = base.get("area_trench") if base else None
        b_af = base.get("area_base") if base else None
        b_aj = base.get("area_back") if base else None
        r_len = _n(row.get("len_c"))
        ln = row.get("len_n")
        if key == "vol_cut":
            return ("挖方 = (基准行断面积 %s + 本行断面积 %s)/2 × 中心长 %s = %s"
                    % (_q(b_ad), _q(row.get("area_trench")), _q(r_len),
                       _q(row.get("vol_cut"))))
        if key == "vol_base":
            return ("180度砂石基础体积 = (基准行 %s + 本行 %s)/2 × 净长 %s = %s"
                    % (_q(b_af), _q(row.get("area_base")), _q(ln),
                       _q(row.get("vol_base"))))
        if key == "vol_pipe":
            return ("管道占位体积 = 管道面积 %s × 净长 %s = %s"
                    % (_q(row.get("area_pipe")), _q(ln), _q(row.get("vol_pipe"))))
        if key == "vol_back":
            if row.get("cut_fill") == "填方段":
                return ("回填连砂石体积 = 挖方 %s - 砂石基础 %s - 管道 %s = %s（填方段）"
                        % (_q(row.get("vol_cut")), _q(row.get("vol_base")),
                           _q(row.get("vol_pipe")), _q(row.get("vol_back"))))
            return ("回填连砂石体积 = (基准行面积 %s + 本行面积 %s)/2 × 净长 %s - 砂石基础 %s = %s"
                    % (_q(b_aj), _q(row.get("area_back")), _q(ln),
                       _q(row.get("vol_base")), _q(row.get("vol_back"))))
        return ("管顶50cm以上回填土 = 挖方 %s - 砂石基础 %s - 管道 %s - 回填连砂石 %s = %s"
                % (_q(row.get("vol_cut")), _q(row.get("vol_base")),
                   _q(row.get("vol_pipe")), _q(row.get("vol_back")),
                   _q(row.get("vol_soil"))))
    if key == "well_room":
        return "井室高度 = 内径 %s + 1.8 + 0.1 = %s" % (_q(w), _q(row.get("well_room")))
    if key == "shaft_h":
        return ("井筒高度 = 检查井高度 %s - 井室高度 %s = %s"
                % (_q(well_h), _q(row.get("well_room")), _q(row.get("shaft_h"))))
    if key == "shaft_min":
        return ("满足最小井筒高度 = MAX(井筒高度 %s, 0.4) = %s"
                % (_q(row.get("shaft_h")), _q(row.get("shaft_min"))))
    if key == "ring_h":
        return ("井室计算高度 = IF(满足最小=0.4, 检查井高度 %s + 井底厚度 %s - 0.4, 井室高度 %s + 井底厚度 %s) = %s"
                % (_q(well_h), _q(_n(row.get("bot_thick")) or WELL_BOTTOM),
                   _q(row.get("well_room")), _q(_n(row.get("bot_thick")) or WELL_BOTTOM),
                   _q(row.get("ring_h"))))
    if key == "well_ring":
        if road != "是":
            return "砼井周加强：是否车行道 = %s，不做井周加强" % (road or "空")
        return ("砼井周加强 = 2π×1.2×井室计算高度 %s + 2π×0.95×(满足最小井筒 %s - 0.16 - 0.125) = %s"
                % (_q(row.get("ring_h")), _q(row.get("shaft_min")),
                   _q(row.get("well_ring"))))
    return ""


_COL_TITLE = {c["key"]: c["title"] for c in COLUMNS}


def summary(rows):
    """返回 (主管挖方, 支管挖方, 合计挖方, 主管井数, 支管井数)。"""
    main_cut = branch_cut = 0.0
    main_well = branch_well = 0.0
    for kind, seg in segments(rows):
        if kind != "seg":
            continue
        br = _seg_is_branch(seg)
        for row in seg:
            cut = row.get("vol_cut")
            cnt = _n(row.get("well_cnt")) or 0.0
            if br:
                if cut:
                    branch_cut += cut
                branch_well += cnt
            else:
                if cut:
                    main_cut += cut
                main_well += cnt
    return main_cut, branch_cut, main_cut + branch_cut, main_well, branch_well


# -------------------------------------------------------------- 导出 ----

_EXPORT_LAST = "AT"


def export_table(path, project_name, rows, road_thick=ROAD_THICK):
    """照原表布局导出：单行表头 + 数据行（含“断开”行）+ 主管/支管小计 + 合计。"""
    grid = []
    head = [(c["title"], excel_io.S_HEAD) for c in COLUMNS]
    grid.append(head)

    def line_for(row, seq_text=""):
        line = []
        for c in COLUMNS:
            key = c["key"]
            role = c.get("role")
            if c["kind"] == "seq":
                v = seq_text
            elif key == "trench_bottom_in":
                v = row.get("trench_bottom")
            elif role == "result":
                v = row.get(key)
            elif c["kind"] in ("text",):
                v = str(row.get(key, "") or "")
            else:
                v = _n(row.get(key))
            line.append((v, excel_io._col_style(c)))
        return line

    def subtotal_line(tag, sub_rows, title2):
        line = []
        for c in COLUMNS:
            key = c["key"]
            if c["kind"] == "seq":
                line.append((tag, excel_io.S_TEXT_C))
            elif c["kind"] == "text":
                if key == "well":
                    line.append((title2, excel_io.S_TEXT_C))
                else:
                    line.append(("", excel_io.S_TEXT_C))
            elif key in ("len_c", "len_cut", "well_cnt", "vol_cut", "vol_base",
                         "vol_pipe", "vol_back", "vol_soil", "well_ring"):
                total = 0.0
                for r in sub_rows:
                    v = _n(r.get(key))
                    if v is not None:
                        total += v
                line.append((total, excel_io._col_style(c)))
            else:
                line.append(("", excel_io._col_style(c)))
        return line

    # 主管/支管按“段”归属：段内井号带 “-” 即整段算支管
    main_rows, branch_rows, branch_ids = [], [], set()
    for kind, seg in segments(rows):
        if kind != "seg":
            continue
        if _seg_is_branch(seg):
            branch_rows.extend(seg)
            branch_ids.update(id(r) for r in seg)
        else:
            main_rows.extend(seg)

    # 主管区（保留断开行）
    in_branch = False
    seq = 0
    for row in rows:
        if is_break(row):
            if not in_branch:
                line = [("", excel_io.S_TEXT_C)] * len(COLUMNS)
                line[0] = ("断开", excel_io.S_TEXT_C)
                grid.append(line)
            continue
        if is_empty_row(row):
            continue
        if id(row) in branch_ids:
            in_branch = True
            continue
        seq += 1
        grid.append(line_for(row, str(seq)))
    if main_rows:
        grid.append(subtotal_line("主管", main_rows, "小计"))

    for row in branch_rows:
        seq += 1
        grid.append(line_for(row, str(seq)))
    if branch_rows:
        grid.append(subtotal_line("支管", branch_rows, "小计"))

    all_rows = main_rows + branch_rows
    if all_rows:
        grid.append(subtotal_line("", all_rows, "合计"))

    widths = [float(c["width"]) / 7.6 for c in COLUMNS]
    heights = {1: 46}
    for i in range(2, len(grid) + 1):
        heights[i] = 20
    excel_io.write_xlsx(path, (project_name or "雨污水工程")[:31], grid,
                        col_widths=widths, row_heights=heights, freeze_rows=1,
                        print_area=("A", _EXPORT_LAST))


# -------------------------------------------------------------- 导入 ----

_IMPORT_ALIASES = {
    "mark": ("序号/标记", "序号", "标记"),
    "station": ("桩号",),
    "well": ("井号",),
    "well_type": ("检查井型号", "井型号"),
    "road": ("是否车行道", "是否车行道上"),
    "ground": ("原地面高程",),
    "road_base": ("竣工路基高程",),
    "trench_bottom_in": ("沟槽抄测底高程", "沟槽抄测底"),
    "pipe_bottom": ("设计管内底高程",),
    "road_top": ("设计路面高程",),
    "len_c": ("管线中心长度",),
    "len_cut": ("扣减长度",),
    "dia": ("管径",),
    "grade": ("管道级别",),
    "inner": ("管道内径",),
    "wall": ("壁厚t", "壁厚"),
    "c1": ("c1",),
    "c2": ("c2",),
    "work": ("工作面a", "工作面"),
    "well_cnt": ("井个数",),
    "bot_thick": ("井底厚度h1", "井底厚度"),
    "fill_thick": ("换填厚度", "换填"),
}

_TEXT_KEYS = ("station", "well", "well_type", "road", "dia", "grade")


def import_rows(grid, road_thick=ROAD_THICK):
    """从原表（或本程序导出的表）读回数据行；识别不到返回 None。"""
    titles, start = excel_io.read_header(grid)
    if not titles:
        return None
    norm = [excel_io.norm_title(t) for t in titles]
    idx = {}
    for key, aliases in _IMPORT_ALIASES.items():
        for j, t in enumerate(norm):
            if t in aliases:
                idx[key] = j
                break
    if "well" not in idx or "len_c" not in idx:
        return None

    def cell(line, key):
        j = idx.get(key)
        if j is None or j >= len(line):
            return None
        return line[j]

    rows = []
    for line in grid[start:]:
        if not line:
            continue
        mark = str(cell(line, "mark") or "").strip()
        well = str(cell(line, "well") or "").strip()
        if well in ("小计", "合计"):
            continue
        if mark == "断开":
            rows.append(new_break())
            continue
        if not well and not any(str(c or "").strip() for c in line[:6]):
            continue
        row = new_row()
        for key in _TEXT_KEYS:
            v = cell(line, key)
            row[key] = "" if v is None else str(v).strip()
        for key in INPUT_KEYS:
            if key in _TEXT_KEYS:
                continue
            v = _n(cell(line, key))
            row[key] = "" if v is None else fmt.fmt_qty(v)
        # 原表的竣工路基高程是公式算出来的，这里反推出换填厚度
        m = _n(cell(line, "road_top"))
        h = _n(cell(line, "road_base"))
        if m is not None and h is not None:
            thick = 0.0 if str(row.get("road") or "").strip() == "否" else road_thick
            fill = m - thick - h
            if fill > 0.0005:
                row["fill_thick"] = fmt.fmt_qty(fill)
        rows.append(row)
    return rows or None
