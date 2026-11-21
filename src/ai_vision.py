# -*- coding: utf-8 -*-
"""图纸识读：PDF / 图片 → 页面图像 → 多模态模型 → 雨污水表格行。

支持的输入：PDF（向量图或扫描件）、PNG/JPG/BMP/TIF。
输出：按「排水管道纵断面表」字段结构化的行，供 SewerPage 直接追加。
"""

import json
import os
import re
import shutil
import tempfile

import ai_model

IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
PDF_EXTS = (".pdf",)
SUPPORTED_EXTS = IMAGE_EXTS + PDF_EXTS

MAX_PAGES = 8          # 一次最多识别页数，避免一次请求过大
RENDER_DPI = 200       # PDF 渲染分辨率
IMAGE_MAX_SIDE = 2200  # 送给模型的图片长边上限（图纸文字要够清晰）
VISION_TIMEOUT = 240   # 视觉模型推理慢（120B 级尤甚），单独放宽超时（秒）

# 表格字段（顺序即预览列顺序）
FIELDS = [
    ("station", "桩号"),
    ("well", "井号"),
    ("well_type", "检查井型号"),
    ("ground", "原地面高程"),
    ("pipe_bottom", "设计管内底高程"),
    ("road_top", "设计路面高程"),
    ("len_c", "管道中心长度"),
    ("dia", "管径"),
    ("grade", "管道级别"),
]

PROMPT = """你是市政排水图纸识读助手。图上通常是「排水管道纵断面图」或「检查井一览表」，\
表格里常见列：桩号、井号、检查井型号、管径、管道级别、原地面高程、设计管内底高程、\
设计路面高程、管道中心长度。

请逐行读取表格，只输出 JSON 数组，不要任何解释文字。每个元素一行：
{"station": "", "well": "", "well_type": "", "ground": "", "pipe_bottom": "", \
"road_top": "", "len_c": "", "dia": "", "grade": ""}

要求：
1. 识别不到的字段留空字符串，绝对不要编造数字。
2. 高程保留图纸原样（一般 3 位小数），长度保留 2 位；去掉 m、mm 等单位。
3. 管径按原样写，如 DN400、d600、De315；管道级别如 II级、SN8。
4. 表格有多段（多根管线）时，从上到下、从左到右依次输出全部行。
5. 只输出 [ ... ] 这一段 JSON。"""


def _strip_unit(text):
    value = (text or "").strip()
    value = re.sub(r"[（(].*?[)）]", "", value)
    value = value.replace(",", "").replace("，", "")
    value = re.sub(r"\s*(m|mm|米|毫米|‰)\s*$", "", value, flags=re.I)
    return value.strip()


def normalize_row(item):
    """把模型返回的一行清洗成表格值。"""
    row = {}
    for key, _title in FIELDS:
        raw = item.get(key)
        if raw is None:
            raw = ""
        if isinstance(raw, (int, float)):
            raw = str(raw)
        row[key] = _strip_unit(str(raw))
    return row


def parse_rows(text):
    """从模型回复里抠出 JSON 数组；兼容 ```json 代码块。"""
    if not text:
        return []
    body = text.strip()
    fenced = re.search(r"```(?:json)?\s*(.+?)```", body, re.S)
    if fenced:
        body = fenced.group(1).strip()
    data = None
    try:
        data = json.loads(body)
    except Exception:
        start = body.find("[")
        end = body.rfind("]")
        if start >= 0 and end > start:
            try:
                data = json.loads(body[start:end + 1])
            except Exception:
                data = None
    if isinstance(data, dict):
        for key in ("rows", "data", "items", "list"):
            if isinstance(data.get(key), list):
                data = data[key]
                break
    if not isinstance(data, list):
        return []
    rows = []
    for item in data:
        if isinstance(item, dict):
            row = normalize_row(item)
            if any(row.values()):
                rows.append(row)
    return rows


def page_images(path, temp_dir, dpi=RENDER_DPI):
    """把 PDF 每页渲染成临时 PNG；图片直接返回原路径。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in IMAGE_EXTS:
        return [path]
    if ext not in PDF_EXTS:
        raise RuntimeError("不支持的文件类型：%s" % ext)

    try:
        import fitz  # PyMuPDF
    except Exception:
        raise RuntimeError("缺少 PyMuPDF，无法读取 PDF；请先另存为图片，或安装 pymupdf")

    doc = fitz.open(path)
    try:
        total = len(doc)
        if total == 0:
            raise RuntimeError("PDF 没有页面")
        count = min(total, MAX_PAGES)
        images = []
        for index in range(count):
            pix = doc[index].get_pixmap(dpi=dpi)
            target = os.path.join(temp_dir, "page_%02d.png" % (index + 1))
            pix.save(target)
            images.append(target)
        if total > count:
            raise RuntimeError("PDF 共 %d 页，一次最多识别 %d 页；请先拆分" % (total, MAX_PAGES))
        return images
    finally:
        doc.close()


def ocr_drawing(profile, path, hint="", on_progress=None, temp_dir=None):
    """识别一份图纸，返回结构化行列表。"""
    if not os.path.exists(path):
        raise RuntimeError("找不到文件：%s" % path)
    if not ai_model.is_vision_model(profile.get("model")):
        raise RuntimeError("当前模型不支持看图，请到「AI模型」里换成带 VL / vision 的模型")

    own_temp = temp_dir is None
    workdir = temp_dir or tempfile.mkdtemp(prefix="zaojia_ai_")
    try:
        images = page_images(path, workdir)
        rows = []
        for index, image in enumerate(images):
            if on_progress:
                on_progress(index + 1, len(images))
            prompt = PROMPT
            if hint.strip():
                prompt += "\n\n补充说明：" + hint.strip()
            reply = ai_model.chat_with_image(profile, prompt, [image],
                                             max_side=IMAGE_MAX_SIDE,
                                             timeout=VISION_TIMEOUT)
            rows.extend(parse_rows(reply))
        return rows
    finally:
        if own_temp:
            shutil.rmtree(workdir, ignore_errors=True)
