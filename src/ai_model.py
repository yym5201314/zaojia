# -*- coding: utf-8 -*-
"""AI 模型接入层：所有引擎统一走 OpenAI 兼容协议（/chat/completions、/models）。

配置保存在 %APPDATA%\\造价计算工具\\config.json 的 "ai" 段：
    {"engine": "DeepSeek", "temperature": 0.3, "timeout": 30,
     "profiles": {"DeepSeek": {"endpoint": "...", "model": "...", "key_b64": "..."}}}
密钥只做 base64 混淆（不是加密），仅存本机。
"""

import base64
import json
import os
import urllib.error
import urllib.request

# (名称, 默认地址, 常用模型)
PROVIDERS = [
    ("DeepSeek", "https://api.deepseek.com/v1",
     ["deepseek-chat", "deepseek-reasoner"]),
    ("通义千问", "https://dashscope.aliyuncs.com/compatible-mode/v1",
     ["qwen-plus", "qwen-max", "qwen-turbo"]),
    ("OpenAI", "https://api.openai.com/v1",
     ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"]),
    ("硅基流动", "https://api.siliconflow.cn/v1",
     ["Qwen/Qwen2.5-72B-Instruct", "deepseek-ai/DeepSeek-V3"]),
    ("智谱", "https://open.bigmodel.cn/api/paas/v4",
     ["glm-4-plus", "glm-4-flash"]),
    ("豆包", "https://ark.cn-beijing.volces.com/api/v3",
     ["doubao-pro-32k"]),
    ("Gemini", "https://generativelanguage.googleapis.com/v1beta/openai",
     ["gemini-2.0-flash", "gemini-1.5-pro"]),
    ("自定义", "", []),
]

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


def default_profile(engine):
    _name, endpoint, models = provider_preset(engine)
    return {"endpoint": endpoint, "model": (models[0] if models else ""), "key": ""}


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


def _request(profile, path, payload=None):
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
    timeout = int(profile.get("timeout") or DEFAULT_TIMEOUT)
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


def chat(profile, messages, temperature=None, max_tokens=None):
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
    result = _request(profile, "/chat/completions", payload)
    try:
        content = result["choices"][0]["message"]["content"]
    except Exception:
        raise RuntimeError("返回格式无法解析：%s"
                           % json.dumps(result, ensure_ascii=False)[:300])
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict) and isinstance(item.get("text"), str):
                parts.append(item["text"])
        content = "".join(parts)
    return (content or "").strip()


def test_connection(profile):
    reply = chat(profile, [{"role": "user", "content": "ping"}], max_tokens=8)
    return reply[:60] if reply else "连接正常（无内容返回）"
