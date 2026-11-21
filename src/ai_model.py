# -*- coding: utf-8 -*-
"""AI 模型接入层：所有引擎统一走 OpenAI 兼容协议（/chat/completions、/models）。

配置保存在 %APPDATA%\\造价计算工具\\config.json 的 "ai" 段：
    {"engine": "DeepSeek", "temperature": 0.3, "timeout": 30,
     "profiles": {"DeepSeek": {"endpoint": "...", "model": "...", "key_b64": "..."}}}
密钥只做 base64 混淆（不是加密），仅存本机。
"""

import base64
import io
import json
import os
import urllib.error
import urllib.request

# (名称, 默认地址, 文本模型, 可识图模型)
PROVIDERS = [
    ("DeepSeek", "https://api.deepseek.com/v1",
     ["deepseek-chat", "deepseek-reasoner"], []),
    ("通义千问", "https://dashscope.aliyuncs.com/compatible-mode/v1",
     ["qwen-plus", "qwen-max", "qwen-turbo"],
     ["qwen3-vl-plus", "qwen-vl-max", "qwen-vl-plus", "qwen-vl-ocr"]),
    ("OpenAI", "https://api.openai.com/v1",
     ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"], ["gpt-4o", "gpt-4o-mini"]),
    ("硅基流动", "https://api.siliconflow.cn/v1",
     ["Qwen/Qwen2.5-72B-Instruct", "deepseek-ai/DeepSeek-V3"],
     ["Qwen/Qwen3-VL-30B-A3B-Instruct", "Qwen/Qwen3-VL-32B-Instruct",
      "Qwen/Qwen3-VL-8B-Instruct", "PaddlePaddle/PaddleOCR-VL-1.5"]),
    ("智谱", "https://open.bigmodel.cn/api/paas/v4",
     ["glm-4-plus", "glm-4-flash"], ["glm-4v-plus", "glm-4v"]),
    ("豆包", "https://ark.cn-beijing.volces.com/api/v3",
     ["doubao-pro-32k"], ["doubao-1.5-vision-pro", "doubao-vision-pro-32k"]),
    ("Gemini", "https://generativelanguage.googleapis.com/v1beta/openai",
     ["gemini-2.0-flash", "gemini-1.5-pro"], ["gemini-2.0-flash", "gemini-1.5-pro"]),
    ("英伟达", "https://integrate.api.nvidia.com/v1",
     ["nvidia/nemotron-3.5-lightning-30b-a3b", "nvidia/nemotron-3-ultra-550b-a55b",
      "z-ai/glm-5.3-flash"],
     ["nvidia/nemotron-3.5-super-vl-120b-a12b"]),
    ("自定义", "", [], []),
]

VISION_KEYWORDS = ("-vl", "vl-", "vision", "gpt-4o", "gpt-4.1", "gemini",
                   "glm-4v", "internvl", "llava")

ENGINES = [item[0] for item in PROVIDERS]
DEFAULT_TEMPERATURE = 0.3
DEFAULT_TIMEOUT = 30


def config_path():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "造价计算工具", "config.json")


def _read_all():
    try:
        with open(config_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_all(data):
    try:
        path = config_path()
        folder = os.path.dirname(path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=1)
        return True
    except Exception:
        return False


def provider_preset(engine):
    for item in PROVIDERS:
        if item[0] == engine:
            return item
    return PROVIDERS[-1]


def model_choices(engine):
    """下拉候选：可识图模型排在前面，便于选来识别图纸。"""
    _name, _endpoint, models, vision = provider_preset(engine)
    ordered = []
    for item in list(vision) + list(models):
        if item not in ordered:
            ordered.append(item)
    return ordered


def is_vision_model(model):
    name = (model or "").lower()
    return any(key in name for key in VISION_KEYWORDS)


def default_profile(engine):
    _name, endpoint, models, vision = provider_preset(engine)
    first = (vision or models or [""])[0]
    return {"endpoint": endpoint, "model": first, "key": ""}


def _encode(secret):
    if not secret:
        return ""
    return base64.b64encode(secret.encode("utf-8")).decode("ascii")


def _decode(value):
    if not isinstance(value, str) or not value:
        return ""
    try:
        return base64.b64decode(value.encode("ascii")).decode("utf-8")
    except Exception:
        return ""


def load_config():
    raw = _read_all().get("ai")
    ai = raw if isinstance(raw, dict) else {}
    engine = ai.get("engine")
    if engine not in ENGINES:
        engine = ENGINES[0]
    stored = ai.get("profiles") if isinstance(ai.get("profiles"), dict) else {}
    profiles = {}
    for name in ENGINES:
        item = stored.get(name)
        item = item if isinstance(item, dict) else {}
        preset = default_profile(name)
        endpoint = item.get("endpoint")
        model = item.get("model")
        profiles[name] = {
            "endpoint": endpoint if isinstance(endpoint, str) and endpoint else preset["endpoint"],
            "model": model if isinstance(model, str) else preset["model"],
            "key": _decode(item.get("key_b64")),
        }
    temp = ai.get("temperature")
    timeout = ai.get("timeout")
    return {
        "engine": engine,
        "temperature": float(temp) if isinstance(temp, (int, float)) and 0 <= temp <= 2
        else DEFAULT_TEMPERATURE,
        "timeout": int(timeout) if isinstance(timeout, (int, float)) and 5 <= timeout <= 300
        else DEFAULT_TIMEOUT,
        "profiles": profiles,
    }


def save_config(cfg):
    data = _read_all()
    profiles = {}
    for name in ENGINES:
        item = (cfg.get("profiles") or {}).get(name) or {}
        profiles[name] = {
            "endpoint": normalize_endpoint(item.get("endpoint") or ""),
            "model": (item.get("model") or "").strip(),
            "key_b64": _encode((item.get("key") or "").strip()),
        }
    data["ai"] = {
        "engine": cfg.get("engine") if cfg.get("engine") in ENGINES else ENGINES[0],
        "temperature": cfg.get("temperature", DEFAULT_TEMPERATURE),
        "timeout": cfg.get("timeout", DEFAULT_TIMEOUT),
        "profiles": profiles,
    }
    return _write_all(data)


def active_profile(cfg):
    engine = cfg.get("engine") or ENGINES[0]
    item = dict((cfg.get("profiles") or {}).get(engine) or {})
    item["engine"] = engine
    item["temperature"] = cfg.get("temperature", DEFAULT_TEMPERATURE)
    item["timeout"] = cfg.get("timeout", DEFAULT_TIMEOUT)
    return item


def normalize_endpoint(text):
    value = (text or "").strip()
    while value.endswith("/"):
        value = value[:-1]
    for suffix in ("/chat/completions", "/models"):
        if value.endswith(suffix):
            value = value[: -len(suffix)]
            while value.endswith("/"):
                value = value[:-1]
    return value


def _error_text(exc):
    try:
        body = exc.read().decode("utf-8", "replace")
    except Exception:
        body = ""
    detail = ""
    try:
        payload = json.loads(body)
        err = payload.get("error")
        if isinstance(err, dict):
            detail = err.get("message") or err.get("type") or ""
        elif isinstance(err, str):
            detail = err
        if not detail:
            detail = payload.get("message") or ""
    except Exception:
        detail = ""
    if not detail:
        detail = body.strip()[:300] or str(getattr(exc, "reason", "") or exc)
    return "HTTP %s：%s" % (exc.code, detail)


def _request(profile, path, payload=None, timeout=None):
    endpoint = normalize_endpoint(profile.get("endpoint") or "")
    if not endpoint:
        raise RuntimeError("请先填写接口地址")
    url = endpoint + path
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    key = (profile.get("key") or "").strip()
    if key:
        headers["Authorization"] = "Bearer " + key
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers,
                                 method="POST" if data else "GET")
    timeout = int(timeout or profile.get("timeout") or DEFAULT_TIMEOUT)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(_error_text(exc))
    except urllib.error.URLError as exc:
        raise RuntimeError("网络错误：%s" % (getattr(exc, "reason", None) or exc,))
    except Exception as exc:
        raise RuntimeError("请求失败：%s" % (exc,))
    try:
        return json.loads(raw)
    except Exception:
        raise RuntimeError("返回内容不是 JSON：%s" % raw.strip()[:300])


def list_models(profile):
    payload = _request(profile, "/models")
    items = payload.get("data")
    if not isinstance(items, list):
        items = payload.get("models")
    names = []
    for item in items if isinstance(items, list) else []:
        if isinstance(item, dict):
            name = item.get("id") or item.get("name")
        else:
            name = item if isinstance(item, str) else None
        if isinstance(name, str) and name.strip() and name not in names:
            names.append(name)
    names.sort()
    return names


def chat(profile, messages, temperature=None, max_tokens=None, timeout=None):
    model = (profile.get("model") or "").strip()
    if not model:
        raise RuntimeError("请先填写模型名称")
    payload = {
        "model": model,
        "messages": messages,
        "stream": False,
        "temperature": profile.get("temperature", DEFAULT_TEMPERATURE)
        if temperature is None else temperature,
    }
    if max_tokens:
        payload["max_tokens"] = int(max_tokens)
    elif profile.get("engine") == "英伟达":
        # NIM 默认输出上限偏小，表格行多会被截断
        payload["max_tokens"] = 8192
    result = _request(profile, "/chat/completions", payload, timeout=timeout)
    try:
        message = result["choices"][0]["message"]
    except Exception:
        raise RuntimeError("返回格式无法解析：%s"
                           % json.dumps(result, ensure_ascii=False)[:300])
    if not isinstance(message, dict):
        raise RuntimeError("返回格式无法解析：%s"
                           % json.dumps(result, ensure_ascii=False)[:300])
    content = message.get("content")
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict) and isinstance(item.get("text"), str):
                parts.append(item["text"])
        content = "".join(parts)
    text = (content or "").strip() if isinstance(content, str) else ""
    if not text:
        # 带思考的模型（如 NVIDIA nemotron VL）可能只回 reasoning_content
        reasoning = message.get("reasoning_content")
        if isinstance(reasoning, str) and reasoning.strip():
            text = reasoning.strip()
    return text


def test_connection(profile):
    reply = chat(profile, [{"role": "user", "content": "ping"}], max_tokens=8)
    return reply[:60] if reply else "连接正常（无内容返回）"


def _encode_image(path, max_side, quality):
    try:
        from PIL import Image
    except Exception:
        Image = None

    if Image is not None:
        try:
            img = Image.open(path)
            img = img.convert("RGB")
            width, height = img.size
            scale = min(1.0, float(max_side) / max(1, max(width, height)))
            if scale < 1.0:
                img = img.resize((max(1, int(width * scale)), max(1, int(height * scale))),
                                 Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=quality, optimize=True)
            return buf.getvalue(), "image/jpeg"
        except Exception as exc:
            raise RuntimeError("图片读取失败：%s" % exc)

    if not os.path.exists(path):
        raise RuntimeError("找不到图片：%s" % path)
    with open(path, "rb") as fh:
        raw = fh.read()
    mime = "image/png" if path.lower().endswith(".png") else "image/jpeg"
    return raw, mime


def image_data_url(path, max_side=1600, quality=85):
    """把图纸图片压缩成 JPEG 的 data URL；超限时自动降分辨率重压。"""
    size = max(1, int(max_side))
    qual = max(40, int(quality))
    limit = int(3.6 * 1024 * 1024)
    raw, mime = _encode_image(path, size, qual)
    for _ in range(4):
        if len(raw) <= limit:
            break
        size = int(size * 0.75)
        qual = max(60, qual - 8)
        raw, mime = _encode_image(path, size, qual)
    if len(raw) > limit:
        raise RuntimeError("图片压缩后仍超过 3.6MB，请先裁切或降低分辨率")
    return "data:%s;base64,%s" % (mime, base64.b64encode(raw).decode("ascii"))


def chat_with_image(profile, prompt, paths, max_tokens=None, max_side=1600,
                    timeout=None):
    """发送一张或多张图片，让模型读取图纸内容。"""
    if not is_vision_model(profile.get("model")):
        raise RuntimeError("当前模型不支持看图，请改用带 VL / vision 的模型（如 qwen-vl-max）")
    content = [{"type": "text", "text": prompt}]
    for path in paths:
        content.append({"type": "image_url",
                        "image_url": {"url": image_data_url(path, max_side=max_side)}})
    return chat(profile, [{"role": "user", "content": content}],
                max_tokens=max_tokens, timeout=timeout)
