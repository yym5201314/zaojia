# -*- coding: utf-8 -*-
"""Excel 读写（纯标准库实现，不依赖第三方包）

- read_table:   读取 .xlsx/.xlsm/.csv，返回二维表
- extract_items: 从二维表中识别 序号/名称/单位/数量 列
- write_xlsx:   写最简 xlsx
- write_sheet:  按表头样式（标题/项目名称/分组表头/列名）写业务表
"""

import csv
import io
import os
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import fmt

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


# ================================================================ 读取 ====

def _col_index(ref):
    """A1 -> 0, AB12 -> 27"""
    m = _REF_RE.match(ref or "")
    if not m:
        return 0
    letters = m.group(1).upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _first_sheet_path(zf):
    names = zf.namelist()
    try:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet = wb.find(NS + "sheets/" + NS + "sheet")
        rid = sheet.get(NS_R + "id") if sheet is not None else None
        if rid:
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rels:
                if rel.get("Id") == rid:
                    target = (rel.get("Target") or "").lstrip("/")
                    if not target.startswith("xl/"):
                        target = "xl/" + target
                    if target in names:
                        return target
    except Exception:
        pass
    for name in names:
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            return name
    raise ValueError("工作簿中没有找到工作表")


def read_xlsx(path):
    """读取 xlsx 第一张工作表，返回 [[值, ...], ...]。"""
    with zipfile.ZipFile(path) as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sroot.findall(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        sheet_path = _first_sheet_path(zf)
        root = ET.fromstring(zf.read(sheet_path))
        data = root.find(NS + "sheetData")
        out = []
        if data is None:
            return out
        for row in data.findall(NS + "row"):
            try:
                ridx = int(row.get("r"))
            except (TypeError, ValueError):
                ridx = len(out) + 1
            while len(out) < ridx - 1:
                out.append([])
            cells = {}
            maxc = -1
            for c in row.findall(NS + "c"):
                ci = _col_index(c.get("r"))
                ctype = c.get("t")
                v = c.find(NS + "v")
                is_el = c.find(NS + "is")
                if ctype == "s" and v is not None:
                    try:
                        val = shared[int(v.text)]
                    except (ValueError, IndexError, TypeError):
                        val = ""
                elif ctype == "inlineStr" and is_el is not None:
                    val = "".join(t.text or "" for t in is_el.iter(NS + "t"))
                elif ctype == "str" and v is not None:
                    val = v.text or ""
                elif v is not None and v.text is not None:
                    txt = v.text.strip()
                    try:
                        fv = float(txt)
                        val = int(fv) if fv == int(fv) and abs(fv) < 1e15 else fv
                    except ValueError:
                        val = txt
                else:
                    val = None
                if val is not None and not (isinstance(val, str) and val == ""):
                    cells[ci] = val
                    if ci > maxc:
                        maxc = ci
            out.append([cells.get(i) for i in range(maxc + 1)])
        return out


def read_csv(path):
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("latin-1")
    out = []
    for row in csv.reader(io.StringIO(text)):
        vals = []
        for cell in row:
            s = cell.strip()
            if not s:
                vals.append(None)
                continue
            try:
                fv = float(s.replace(",", ""))
                vals.append(int(fv) if fv == int(fv) else fv)
            except ValueError:
                vals.append(s)
        out.append(vals)
    return out


def read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext in (".csv", ".txt"):
        return read_csv(path)
    if ext == ".xls":
        raise ValueError("不支持旧版 .xls 文件，请先在 Excel 中另存为 .xlsx 再导入。")
    raise ValueError("不支持的文件类型：%s" % ext)


# =========================================================== 清单识别 ====

_SKIP_NAMES = ("合计", "小计", "总计", "本页小计", "合 计", "合价", "总额")

_TITLE_RE = re.compile(r"^项目名称\s*[:：]")


def _s(cell):
    if cell is None:
        return ""
    if isinstance(cell, float):
        return fmt.fmt_price(cell)
    return str(cell).strip()


def _is_title_cell(text):
    t = (text or "").strip()
    if _TITLE_RE.match(t):
        return True
    return bool(re.match(r"^(项目名称|工程名称|标段名称)\s*[：:]", t))


def _is_name_header(text):
    t = (text or "").strip()
    if not t or _is_title_cell(t):
        return False
    return "名称" in t and len(t) <= 12


def _is_unit_header(text):
    t = (text or "").strip()
    if not t or _is_title_cell(t):
        return False
    return "单位" in t and len(t) <= 12


def _is_qty_header(text):
    t = (text or "").strip()
    if not t or _is_title_cell(t):
        return False
    if "单位" in t:
        return False
    return ("数量" in t or "工程量" in t) and len(t) <= 12


def norm_title(text):
    """表头归一化：去空白、全角括号转半角、转小写。"""
    t = str(text or "")
    t = t.replace("\u3000", "").replace(" ", "")
    t = t.replace("（", "(").replace("）", ")")
    return t.strip().lower()


def read_header(grid, max_scan=15):
    """识别表头，返回 (标题列表, 数据起始行号)，识别不到返回 (None, 0)。

    兼容本程序导出的两行表头（分组行 + 列名行）：同一列有列名时以列名为准。
    """
    for i, row in enumerate(grid[:max_scan]):
        texts = [str(c).strip() if c is not None else "" for c in row]
        joined = "".join(texts)
        if "序号" not in joined or not any("名称" in t for t in texts):
            continue
        nxt = grid[i + 1] if i + 1 < len(grid) else []
        nxt_texts = [str(c).strip() if c is not None else "" for c in nxt]
        filled = numeric = 0
        for cell in nxt:
            if cell is None or cell == "":
                continue
            filled += 1
            if isinstance(cell, (int, float)):
                numeric += 1
        two_rows = filled == 0 or numeric * 2 <= filled
        titles = []
        for j in range(max(len(texts), len(nxt_texts))):
            a = texts[j] if j < len(texts) else ""
            b = nxt_texts[j] if j < len(nxt_texts) else ""
            titles.append((b or a) if two_rows else a)
        return titles, i + (2 if two_rows else 1)
    return None, 0


def extract_items(grid):
    """从二维表中提取 (项目名称, [{"name","unit","qty"}, ...])。"""
    project = ""
    for row in grid[:15]:
        for cell in row:
            s = _s(cell)
            if not s:
                continue
            m = re.search(r"项目名称\s*[:：]\s*(.+)$", s)
            if m:
                project = m.group(1).strip()
                break
            m = re.match(r"^(?:工程名称|标段名称)\s*[:：]\s*(.+)$", s)
            if m:
                project = m.group(1).strip()
                break
        if project:
            break

    head_idx = None
    head_row = None
    for i, row in enumerate(grid[:20]):
        texts = [_s(c) for c in row]
        joined = "".join(texts)
        has_name = any(_is_name_header(t) for t in texts)
        has_more = (any(_is_unit_header(t) for t in texts)
                    or any(_is_qty_header(t) for t in texts)
                    or any("序号" in t for t in texts)
                    or any(k in joined for k in ("结算量", "招标", "单价", "基准价")))
        if has_name and has_more:
            head_idx = i
            head_row = texts
            break

    name_col, unit_col, seq_col, qty_col = 1, 2, 0, None
    start = 0
    if head_row is not None:
        got_name = got_unit = got_seq = got_qty = False
        for j, t in enumerate(head_row):
            if not got_name and _is_name_header(t):
                name_col = j
                got_name = True
            elif not got_unit and _is_unit_header(t):
                unit_col = j
                got_unit = True
            elif not got_seq and "序号" in t:
                seq_col = j
                got_seq = True
            elif not got_qty and _is_qty_header(t):
                qty_col = j
                got_qty = True
        start = head_idx + 1

    items = []
    for row in grid[start:]:
        if not row:
            continue
        name = _s(row[name_col]) if len(row) > name_col else ""
        if not name or name in _SKIP_NAMES or _is_title_cell(name):
            continue
        if any(k in name for k in ("工程量清单", "项目名称及", "金额单位")):
            continue
        if name in ("名称", "项目名称", "清单名称", "分项工程名称", "工程名称",
                    "子目名称", "项目", "计量单位", "单位", "材料名称"):
            continue
        seq = _s(row[seq_col]) if len(row) > seq_col else ""
        if seq and not re.match(r"^\d+(\.\d+)?$", seq) and start == 0:
            continue
        unit = _s(row[unit_col]) if len(row) > unit_col else ""
        qty = None
        if qty_col is not None and len(row) > qty_col:
            v = fmt.parse_num(row[qty_col])
            if v is not None:
                qty = v
        items.append({"name": name, "unit": unit, "qty": qty})
    return project, items


# ================================================================ 写出 ====

# cellXfs 索引
S_DEF = 0
S_TITLE = 1
S_SUB = 2
S_HEAD = 3
S_TEXT = 4
S_TEXT_C = 5
S_NUM_P = 6
S_PCT_P = 7
S_NUM_R = 8
S_MONEY_R = 9
S_MONEY = 10
S_PRICE_R = 11
S_TEXT_WRAP = 12

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4">
<font><sz val="10"/><name val="\u5b8b\u4f53"/><charset val="134"/></font>
<font><b/><sz val="16"/><color rgb="FFC00000"/><name val="\u5fae\u8f6f\u96c5\u9ed1"/><charset val="134"/></font>
<font><b/><sz val="10"/><name val="\u5b8b\u4f53"/><charset val="134"/></font>
<font><b/><sz val="11"/><name val="\u5b8b\u4f53"/><charset val="134"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF8EA9DB"/></left><right style="thin"><color rgb="FF8EA9DB"/></right><top style="thin"><color rgb="FF8EA9DB"/></top><bottom style="thin"><color rgb="FF8EA9DB"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="13">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="2" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="10" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="2" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="2" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="2" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def _wb_xml(sheet_name, print_area=None, last_col="Q", last_row=1):
    q = "'%s'!" % sheet_name.replace("'", "''")
    defined = ""
    if print_area:
        a, b = print_area
        defined = ('<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
                   "%s$%s$1:$%s$%d</definedName></definedNames>" % (q, a, b, last_row))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets><sheet name=\"%s\" sheetId=\"1\" r:id=\"rId1\"/></sheets>%s</workbook>"
        % (escape(sheet_name), defined)
    )


def _col_letter(idx):
    s = ""
    n = idx + 1
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _num_text(v):
    if isinstance(v, int):
        return str(v)
    s = ("%.6f" % v).rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def _cell_xml(row_no, col_no, value, style):
    ref = "%s%d" % (_col_letter(col_no), row_no)
    if value is None or (isinstance(value, str) and value == ""):
        return '<c r="%s" s="%d"/>' % (ref, style)
    if isinstance(value, bool):
        value = int(value)
    if isinstance(value, (int, float)):
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, _num_text(value))
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, escape(str(value))))


def write_xlsx(path, sheet_name, grid, col_widths=None, merges=None,
               row_heights=None, freeze_rows=0, print_area=None,
               outline_cols=None, summary_right=False):
    """grid: [[(值, 样式) 或 值, ...], ...]"""
    n_rows = len(grid)
    n_cols = max([len(r) for r in grid] + [1])
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    if outline_cols:
        parts.append('<sheetPr><outlinePr summaryBelow="1" summaryRight="%d"/></sheetPr>'
                     % (1 if summary_right else 0))
    parts.append('<dimension ref="A1:%s%d"/>' % (_col_letter(n_cols - 1), max(n_rows, 1)))
    if freeze_rows:
        parts.append('<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
                     '<pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>'
                     "</sheetView></sheetViews>" % (freeze_rows, freeze_rows + 1))
    else:
        parts.append('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>')
    parts.append('<sheetFormatPr defaultRowHeight="18"/>')
    if col_widths:
        cols = ["<cols>"]
        for i, w in enumerate(col_widths):
            extra = ""
            if outline_cols and outline_cols[0] <= i <= outline_cols[1]:
                extra = ' outlineLevel="1"'
            cols.append('<col min="%d" max="%d" width="%.2f" customWidth="1"%s/>'
                        % (i + 1, i + 1, w, extra))
        cols.append("</cols>")
        parts.append("".join(cols))
    heights = dict(row_heights or {})
    parts.append("<sheetData>")
    for r, row in enumerate(grid):
        rn = r + 1
        ht = ''
        if rn in heights:
            ht = ' ht="%s" customHeight="1"' % heights[rn]
        cells = []
        for cidx, cell in enumerate(row):
            if isinstance(cell, tuple):
                val, style = cell
            else:
                val, style = cell, S_DEF
            if (val is None or val == "") and style == S_DEF:
                continue
            cells.append(_cell_xml(rn, cidx, val, style))
        parts.append('<row r="%d"%s>%s</row>' % (rn, ht, "".join(cells)))
    parts.append("</sheetData>")
    if merges:
        parts.append('<mergeCells count="%d">' % len(merges))
        for m in merges:
            parts.append('<mergeCell ref="%s"/>' % m)
        parts.append("</mergeCells>")
    parts.append('<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" '
                 'header="0.3" footer="0.3"/>')
    parts.append("</worksheet>")
    sheet_xml = "".join(parts)

    pa = (print_area[0], print_area[1], n_rows) if print_area else None
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        if pa:
            zf.writestr("xl/workbook.xml", _wb_xml(sheet_name, (pa[0], pa[1]), last_row=pa[2]))
        else:
            zf.writestr("xl/workbook.xml", _wb_xml(sheet_name))
        zf.writestr("xl/_rels/workbook.xml.rels", _WB_RELS)
        zf.writestr("xl/styles.xml", _STYLES)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def _col_style(col):
    kind = col.get("kind")
    role = col.get("role")
    if kind == "seq":
        return S_TEXT_C
    if kind in ("text", "text_hint"):
        if col.get("pct"):
            return S_PCT_P
        return S_TEXT_C if col.get("center") else S_TEXT_WRAP
    if role == "result":
        if kind == "money":
            return S_MONEY_R
        if kind == "price":
            return S_PRICE_R
        return S_NUM_R
    if role == "param":
        if kind == "percent":
            return S_PCT_P
        return S_NUM_P
    return S_TEXT


def write_sheet(path, sheet_name, title, project_text, unit_text, columns, data_rows,
                col_widths=None, print_area=None, outline_cols=None):
    """写业务表：标题 / 项目名称+单位 / 分组表头 / 列名 / 数据。

    columns: [{"title","group","kind","role","center"}, ...]
    data_rows: [[值, ...], ...]（与 columns 等长）
    """
    n_cols = len(columns)
    last = _col_letter(n_cols - 1)
    grid = []
    grid.append([(title, S_TITLE)] + [("", S_TITLE)] * (n_cols - 1))
    mid = max(1, int(n_cols * 0.55))
    row2 = [(project_text, S_SUB)] + [("", S_SUB)] * (n_cols - 1)
    row2[-1] = (unit_text, S_SUB)
    grid.append(row2)

    merges = ["A1:%s1" % last, "A2:%s2" % _col_letter(mid - 1),
              "%s2:%s2" % (_col_letter(mid), last)]

    group_row = []
    col_row = []
    i = 0
    while i < n_cols:
        col = columns[i]
        g = col.get("group") or ""
        j = i
        while g and j + 1 < n_cols and (columns[j + 1].get("group") or "") == g:
            j += 1
        if g:
            group_row.append((g, S_HEAD))
            group_row.extend([("", S_HEAD)] * (j - i))
            merges.append("%s3:%s3" % (_col_letter(i), _col_letter(j)))
            for k in range(i, j + 1):
                col_row.append((columns[k]["title"], S_HEAD))
        else:
            # 无分组列：3、4 行合并，标题必须写在合并区左上角（第 3 行）才显示
            group_row.append((col["title"], S_HEAD))
            col_row.append(("", S_HEAD))
            merges.append("%s3:%s4" % (_col_letter(i), _col_letter(i)))
        i = j + 1
    grid.append(group_row)
    grid.append(col_row)

    for values in data_rows:
        line = []
        for c, col in enumerate(columns):
            v = values[c] if c < len(values) else None
            line.append((v, _col_style(col)))
        grid.append(line)

    # 表头行留足高度：列名换行后也不会被裁切、挤压
    heights = {1: 30, 3: 24, 4: 32}
    write_xlsx(path, sheet_name, grid, col_widths=col_widths, merges=merges,
               row_heights=heights, freeze_rows=4, print_area=print_area,
               outline_cols=outline_cols)
