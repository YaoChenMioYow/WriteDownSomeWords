# -*- coding: utf-8 -*-
"""AI 服务：用法搭配生成与翻译义项精简。

调用 OpenAI 兼容接口；支持多个 API Key 与多个模型同时工作（端点池），
结果缓存到本地，避免重复消耗。配置保存在 %APPDATA%\\WriteDownSomeWords\\config.json。
"""

from __future__ import annotations

import hashlib
import json
import os
import queue
import re
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = Path(os.environ.get("APPDATA") or APP_DIR) / "WriteDownSomeWords"
CONFIG_FILE = CONFIG_DIR / "config.json"
LEGACY_CONFIG_FILE = APP_DIR / "ai_config.json"
USAGE_CACHE_FILE = APP_DIR / ".usage_cache.json"
SENSE_CACHE_FILE = APP_DIR / ".senses_cache.json"
FAMILY_CACHE_FILE = APP_DIR / ".families_cache.json"
CACHE_VERSION = 1

# 每个平台卡片可以指定用于哪些任务
TASKS = ("usage", "senses", "family")
TASK_LABELS = {
    "usage": "用法生成",
    "senses": "精简翻译",
    "family": "精简同根词",
}

DEFAULT_CONFIG = {
    "endpoints": [],
    "base_url": "https://apihub.agnes-ai.com/v1",
    "api_key": "",
    "api_keys": [],
    "model": "agnes-3.0-flash",
    "models": [],
    "parallel": True,
    "max_workers": 6,
    "trim_senses": False,
    "trim_families": False,
    "batch_size": 6,
    "prefetch": 6,
    "timeout": 120,
    "theme": "dark",
    "vocab": "",
    "sections": {},
    "sections_open": False,
    "speak_voice": "",
    "speak_rate": 170,
    "speak_auto": False,
    "zoom": 1.0,
    "hide_mode": False,
}

SYSTEM_PROMPT = (
    "你是资深初中英语老师，熟悉中考考点。只输出合法 JSON，不要任何解释，不要 markdown 代码块。"
)


# ---------- 配置 ----------


def load_config() -> dict:
    config = dict(DEFAULT_CONFIG)
    for path in (CONFIG_FILE, LEGACY_CONFIG_FILE):
        try:
            stored = json.loads(path.read_text("utf-8"))
            if isinstance(stored, dict):
                config.update({key: value for key, value in stored.items() if key in config})
                break
        except (OSError, ValueError):
            continue

    config["base_url"] = os.environ.get("AGNES_API_BASE", config["base_url"])
    config["api_key"] = os.environ.get("AGNES_API_KEY", config["api_key"])
    config["model"] = os.environ.get("AGNES_MODEL", config["model"])
    config["base_url"] = str(config["base_url"]).rstrip("/")

    try:
        config["batch_size"] = max(1, min(12, int(config["batch_size"])))
        config["prefetch"] = max(0, min(20, int(config["prefetch"])))
        config["timeout"] = max(10, int(config["timeout"]))
        config["max_workers"] = max(1, min(8, int(config["max_workers"])))
        config["zoom"] = max(0.7, min(1.8, float(config["zoom"])))
    except (TypeError, ValueError):
        config.update(
            {
                key: DEFAULT_CONFIG[key]
                for key in ("batch_size", "prefetch", "timeout", "max_workers", "zoom")
            }
        )

    keys = config.get("api_keys")
    if not isinstance(keys, list):
        keys = []
    keys = [str(key).strip() for key in keys if str(key).strip()]
    primary = str(config.get("api_key") or "").strip()
    if primary and primary not in keys:
        keys.insert(0, primary)
    config["api_keys"] = keys
    config["api_key"] = keys[0] if keys else ""

    models = config.get("models")
    if not isinstance(models, list):
        models = []
    models = [str(name).strip() for name in models if str(name).strip()]
    primary_model = str(config.get("model") or DEFAULT_CONFIG["model"]).strip()
    if not models:
        models = [primary_model]
    elif primary_model and primary_model not in models:
        models.insert(0, primary_model)
    config["models"] = models
    config["model"] = models[0]

    config["parallel"] = bool(config.get("parallel", True))
    config["trim_senses"] = bool(config.get("trim_senses", False))
    config["trim_families"] = bool(config.get("trim_families", False))
    config["hide_mode"] = bool(config.get("hide_mode", False))
    if config.get("theme") not in ("dark", "light", "system"):
        config["theme"] = "dark"

    config["endpoints"] = normalize_endpoints(config)
    if config["endpoints"]:
        primary = config["endpoints"][0]
        config["base_url"] = primary["base_url"]
        config["api_key"] = primary["api_key"]
        config["model"] = primary["model"]
        config["api_keys"] = [item["api_key"] for item in config["endpoints"]]
        config["models"] = [item["model"] for item in config["endpoints"]]
    return config


def normalize_endpoints(config: dict) -> list[dict]:
    """每张卡片是一个 (base_url, key, model + 用途) 组合；兼容旧的扁平配置。"""
    normalized: list[dict] = []
    raw = config.get("endpoints")
    if isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                continue
            base = str(item.get("base_url") or "").strip().rstrip("/")
            key = str(item.get("api_key") or "").strip()
            model = str(item.get("model") or "").strip()
            if base and key:
                normalized.append(
                    {
                        "base_url": base,
                        "api_key": key,
                        "model": model,
                        "tasks": normalize_tasks(item.get("tasks")),
                    }
                )
    if normalized:
        return normalized

    base = str(config.get("base_url") or DEFAULT_CONFIG["base_url"]).strip().rstrip("/")
    keys = [key for key in (config.get("api_keys") or []) if key]
    if not keys and config.get("api_key"):
        keys = [str(config["api_key"]).strip()]
    models = [name for name in (config.get("models") or []) if name]
    if not models:
        models = [str(config.get("model") or DEFAULT_CONFIG["model"]).strip()]
    if not keys:
        return []
    return [
        {
            "base_url": base,
            "api_key": key,
            "model": models[index % len(models)],
            "tasks": list(TASKS),
        }
        for index, key in enumerate(keys)
    ]


def normalize_tasks(value) -> list[str]:
    if not isinstance(value, list):
        return list(TASKS)
    tasks = [str(item).strip() for item in value]
    kept = [task for task in TASKS if task in tasks]
    return kept or list(TASKS)


def save_config(config: dict) -> bool:
    payload = {key: config.get(key, DEFAULT_CONFIG[key]) for key in DEFAULT_CONFIG}
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        temporary = CONFIG_FILE.with_suffix(".tmp")
        temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")
        os.replace(temporary, CONFIG_FILE)
        return True
    except OSError:
        return False


def fetch_models(base_url: str, api_key: str, timeout: int = 30) -> list[str]:
    base = base_url.rstrip("/")
    request = urllib.request.Request(
        f"{base}/models",
        headers={"Authorization": "Bearer " + api_key, "Accept": "application/json"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = json.loads(response.read().decode("utf-8"))

    items = data.get("data") if isinstance(data, dict) else None
    models: list[str] = []
    if isinstance(items, list):
        for item in items:
            if isinstance(item, dict) and item.get("id"):
                models.append(str(item["id"]))
    return sorted(set(models))


# ---------- 提示词与解析 ----------


def build_prompt(items: list[tuple[str, str]]) -> str:
    lines = "\n".join(f"- {word}: {meaning}" if meaning else f"- {word}" for word, meaning in items)
    return (
        "为下列单词各给出 2-3 条中考常考用法搭配。要求：\n"
        "1. pattern 用 conversation with sb.、be good at doing sth. 这种带 sb./sth./doing 的格式，"
        "只给最常见、确实会考的搭配，不要生造；\n"
        "2. meaning 是简短中文（不超过 12 字）；\n"
        "3. example 是初中水平例句（不超过 15 词）；\n"
        "4. 必须为列表中的每一个单词输出，一个都不能遗漏，不要照抄示例。\n"
        '输出格式：{"items":[{"word":"conversation","usage":[{"pattern":"...","meaning":"...",'
        '"example":"..."}]}]}\n'
        "单词列表：\n" + lines
    )


def build_sense_prompt(items: list[tuple[str, str]]) -> str:
    lines = "\n".join(f"- {word}: {meaning}" for word, meaning in items)
    return (
        "下面每个单词给出的是它的全部释义。请为每个单词只保留中考、高考真正常考的义项（最多 3 个）：\n"
        "1. 删除生僻、古旧、过于专业、口语中几乎不用的义项；\n"
        "2. 可适当缩短义项文字，但不要改变原意，也不要把不同词性合并；\n"
        "3. 如果都常考，就原样保留；pos 用 n. / v. / adj. 这类简短词性；\n"
        "4. 必须为列表中的每一个单词输出，一个都不能遗漏。\n"
        '输出格式：{"items":[{"word":"society","senses":[{"pos":"n.","cn":"社会；社团"}]}]}\n'
        "单词与释义：\n" + lines
    )


def build_family_prompt(items: list[tuple[str, str]]) -> str:
    lines = "\n".join(f"- {word}:\n{family}" for word, family in items if family)
    return (
        "下面是每个单词的同根词（词族）列表。请为每个单词挑出真正值得记的同根词：\n"
        "1. 只保留中考、高考常考常用的派生词（例如 social / society / socialist），"
        "丢弃生僻、专业、罕见后缀（如 -ness、-istic、副词硬凑）和重复项；\n"
        "2. 每个词性最多 5 个词，总数不超过 10 个；\n"
        "3. 词性分组用 n. / v. / adj. / adv. 这类简短写法，中文释义不超过 12 字；\n"
        "4. 如果某个单词确实没有值得记的同根词，groups 返回空数组；\n"
        "5. 必须为列表中的每一个单词输出，一个都不能遗漏。\n"
        '输出格式：{"items":[{"word":"society","root":"social","groups":'
        '[{"pos":"adj.","words":[{"c":"social","cn":"社会的"}]}]}]}\n'
        "单词与同根词：\n" + lines
    )


def _extract_json(text: str) -> dict | None:
    text = (text or "").strip()
    if "```" in text:
        text = re.sub(r"```[a-zA-Z]*", "", text).replace("```", "")
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        data = json.loads(text[start : end + 1])
    except ValueError:
        return None
    return data if isinstance(data, dict) else None


def parse_response(text: str) -> list[dict]:
    data = _extract_json(text)
    if data is None:
        return []
    raw_items = data.get("items")
    if not isinstance(raw_items, list):
        return []

    cleaned: list[dict] = []
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        word = str(raw.get("word", "")).strip()
        usage = raw.get("usage")
        if not word or not isinstance(usage, list):
            continue
        entries: list[dict] = []
        for item in usage:
            if not isinstance(item, dict):
                continue
            pattern = str(item.get("pattern", "")).strip()
            if not pattern:
                continue
            entries.append(
                {
                    "pattern": pattern[:80],
                    "meaning": str(item.get("meaning", "")).strip()[:30],
                    "example": str(item.get("example", "")).strip()[:160],
                }
            )
            if len(entries) >= 4:
                break
        if entries:
            cleaned.append({"word": word, "usage": entries})
    return cleaned


def render_family(item: dict) -> str:
    """把结构化的同根词还原成展示文本。"""
    blocks: list[str] = []
    root = str(item.get("root") or "").strip()
    if root:
        blocks.append(f"词根:{root}")
    for group in item.get("groups") or []:
        if not isinstance(group, dict):
            continue
        pos = str(group.get("pos") or "").strip()
        entries: list[str] = []
        for word in group.get("words") or []:
            if not isinstance(word, dict):
                continue
            name = str(word.get("c") or "").strip()
            cn = str(word.get("cn") or "").strip()
            if name:
                entries.append(f"{name}:{cn}" if cn else name)
        if entries:
            blocks.append((pos + "\n" if pos else "") + "\n".join(entries))
    return "\n\n".join(blocks)


def parse_family(text: str) -> list[dict]:
    data = _extract_json(text)
    if data is None:
        return []
    raw_items = data.get("items")
    if not isinstance(raw_items, list):
        return []

    cleaned: list[dict] = []
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        word = str(raw.get("word", "")).strip()
        if not word:
            continue
        groups: list[dict] = []
        for group in raw.get("groups") or []:
            if not isinstance(group, dict):
                continue
            words: list[dict] = []
            for item in group.get("words") or []:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("c", "")).strip()
                if not name:
                    continue
                words.append({"c": name[:40], "cn": str(item.get("cn", "")).strip()[:40]})
                if len(words) >= 6:
                    break
            if words:
                groups.append({"pos": str(group.get("pos", "")).strip()[:8], "words": words})
            if len(groups) >= 5:
                break
        cleaned.append(
            {
                "word": word,
                "root": str(raw.get("root", "")).strip()[:30],
                "groups": groups,
            }
        )
    return cleaned


def parse_senses(text: str) -> list[dict]:
    data = _extract_json(text)
    if data is None:
        return []
    raw_items = data.get("items")
    if not isinstance(raw_items, list):
        return []

    cleaned: list[dict] = []
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        word = str(raw.get("word", "")).strip()
        senses = raw.get("senses")
        if not word or not isinstance(senses, list):
            continue
        kept: list[dict] = []
        for item in senses:
            if not isinstance(item, dict):
                continue
            cn = str(item.get("cn", "")).strip()
            if not cn:
                continue
            kept.append({"pos": str(item.get("pos", "")).strip()[:8], "cn": cn[:160]})
            if len(kept) >= 4:
                break
        if kept:
            cleaned.append({"word": word, "senses": kept})
    return cleaned


# ---------- 网络请求 ----------


def chat_completion(
    config: dict,
    endpoint: dict,
    user_prompt: str,
    max_tokens: int = 2500,
) -> str:
    payload = {
        "model": endpoint["model"],
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
        "max_tokens": max_tokens,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {
        "Authorization": "Bearer " + endpoint["api_key"],
        "Content-Type": "application/json",
    }
    url = endpoint["base_url"].rstrip("/") + "/chat/completions"

    last_error: Exception | None = None
    for attempt in range(3):
        request = urllib.request.Request(url, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=config["timeout"]) as response:
                body = json.loads(response.read().decode("utf-8"))
            return body["choices"][0]["message"].get("content") or ""
        except urllib.error.HTTPError as exc:
            last_error = exc
            if exc.code in (429, 500, 502, 503, 504) and attempt < 2:
                time.sleep(2 + attempt * 3)
                continue
            raise
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = exc
            if attempt < 2:
                time.sleep(2 + attempt * 3)
                continue
            raise

    raise last_error or RuntimeError("请求失败")


def build_endpoints(config: dict, task: str = "usage") -> list[dict]:
    """按任务类型筛选可用的平台端点，供多个 Worker 并行使用。"""
    endpoints = config.get("endpoints")
    if not isinstance(endpoints, list) or not endpoints:
        endpoints = normalize_endpoints(config)
    endpoints = [
        item
        for item in endpoints
        if item.get("base_url")
        and item.get("api_key")
        and task in normalize_tasks(item.get("tasks"))
    ]
    if not endpoints:
        return []
    if not config.get("parallel", True):
        return endpoints[:1]
    try:
        limit = max(1, min(8, int(config.get("max_workers", 6))))
    except (TypeError, ValueError):
        limit = 6
    return endpoints[:limit]


# ---------- 后台任务池 ----------


class WorkerService:
    """通用后台任务池：多个端点并行，UI 线程只负责入队与取结果。"""

    CACHE_FILE: Path = USAGE_CACHE_FILE
    TASK = "usage"

    def __init__(self, config: dict) -> None:
        self.config = config
        self.cache: dict[str, object] = {}
        self.pending: set[str] = set()
        self.failed: dict[str, float] = {}
        self.urgent: queue.Queue = queue.Queue()
        self.jobs: queue.Queue = queue.Queue()
        self.results: queue.Queue = queue.Queue()
        self._stop = threading.Event()
        self.endpoints = build_endpoints(config, self.TASK)
        self._threads: list[threading.Thread] = []
        self._load_cache()

    # ---------- 缓存 ----------

    def _valid_value(self, value) -> bool:
        return isinstance(value, list)

    def _load_cache(self) -> None:
        try:
            data = json.loads(self.CACHE_FILE.read_text("utf-8"))
        except (OSError, ValueError):
            return
        if not isinstance(data, dict) or data.get("version") != CACHE_VERSION:
            return
        items = data.get("items")
        if not isinstance(items, dict):
            return
        self.cache = {
            str(key).lower(): value for key, value in items.items() if self._valid_value(value)
        }

    def save(self) -> None:
        payload = {"version": CACHE_VERSION, "items": self.cache}
        try:
            temporary = self.CACHE_FILE.with_suffix(".tmp")
            temporary.write_text(json.dumps(payload, ensure_ascii=False), "utf-8")
            os.replace(temporary, self.CACHE_FILE)
        except OSError:
            pass

    # ---------- 查询 ----------

    def is_pending(self, word: str) -> bool:
        return word.lower() in self.pending

    def recently_failed(self, word: str) -> bool:
        key = word.lower()
        return key in self.failed and time.time() - self.failed[key] < 300

    def busy(self) -> bool:
        return not self.urgent.empty() or not self.jobs.empty() or bool(self.pending)

    def parallelism(self) -> int:
        return max(1, len(self.endpoints))

    def _ready(self, word: str, text: str) -> bool:
        return word.lower() in self.cache

    # ---------- 任务 ----------

    def enqueue(
        self,
        items: list[tuple[str, str]],
        urgent: bool = False,
        batch_limit: int | None = None,
    ) -> None:
        target = self.urgent if urgent else self.jobs
        batch: list[tuple[str, str]] = []
        page = max(1, int(self.config.get("batch_size", 6)))
        if batch_limit is not None:
            page = max(1, min(page, int(batch_limit)))
        for word, text in items:
            key = word.lower()
            if not key or key in self.pending or self.recently_failed(word):
                continue
            if self._ready(word, text):
                continue
            self.pending.add(key)
            batch.append((word, text))
            if len(batch) >= page:
                target.put(batch)
                batch = []
        if batch:
            target.put(batch)

    def forget(self, word: str) -> None:
        key = word.lower()
        self.cache.pop(key, None)
        self.failed.pop(key, None)

    def start(self) -> None:
        for endpoint in self.endpoints:
            thread = threading.Thread(target=self._worker, args=(endpoint,), daemon=True)
            thread.start()
            self._threads.append(thread)

    def stop(self) -> None:
        self._stop.set()

    # ---------- 线程 ----------

    def _worker(self, endpoint: dict) -> None:
        while not self._stop.is_set():
            try:
                batch = self.urgent.get_nowait()
            except queue.Empty:
                try:
                    batch = self.jobs.get(timeout=0.5)
                except queue.Empty:
                    continue
            try:
                items = self._request(batch, endpoint)
            except Exception as exc:  # 网络、鉴权、解析失败
                for word, _ in batch:
                    key = word.lower()
                    self.pending.discard(key)
                    self.failed[key] = time.time()
                self.results.put(("error", batch, str(exc)[:200]))
                continue

            for word, _ in batch:
                self.pending.discard(word.lower())
            self.results.put(("ok", items, batch))

    def _request(self, batch: list[tuple[str, str]], endpoint: dict) -> list[dict]:
        raise NotImplementedError


class UsageService(WorkerService):
    """中考常考用法搭配。"""

    CACHE_FILE = USAGE_CACHE_FILE
    TASK = "usage"

    def get(self, word: str) -> list[dict] | None:
        value = self.cache.get(word.lower())
        return value if isinstance(value, list) else None

    def store(self, items: list[dict], batch: list[tuple[str, str]] | None = None) -> None:
        for item in items:
            word = str(item.get("word", "")).strip().lower()
            if word and item.get("usage"):
                self.cache[word] = item["usage"]

    def _request(self, batch, endpoint) -> list[dict]:
        content = chat_completion(self.config, endpoint, build_prompt(batch))
        items = parse_response(content)
        if not items:
            raise ValueError("接口没有返回可解析的用法")
        return items


class SenseService(WorkerService):
    """用 AI 精简翻译义项，只保留常考义项。"""

    CACHE_FILE = SENSE_CACHE_FILE
    TASK = "senses"

    @staticmethod
    def fingerprint(text: str) -> str:
        return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12] if text else ""

    def _valid_value(self, value) -> bool:
        return isinstance(value, dict) and isinstance(value.get("senses"), list)

    def _ready(self, word: str, text: str) -> bool:
        entry = self.cache.get(word.lower())
        return isinstance(entry, dict) and entry.get("src") == self.fingerprint(text)

    def get(self, word: str, text: str) -> list[dict] | None:
        entry = self.cache.get(word.lower())
        if not isinstance(entry, dict) or entry.get("src") != self.fingerprint(text):
            return None
        senses = entry.get("senses")
        return senses if isinstance(senses, list) else None

    def store(self, items: list[dict], batch: list[tuple[str, str]] | None = None) -> None:
        sources = {word.lower(): text for word, text in (batch or [])}
        for item in items:
            word = str(item.get("word", "")).strip().lower()
            senses = item.get("senses")
            if not word or not isinstance(senses, list) or not senses:
                continue
            self.cache[word] = {
                "src": self.fingerprint(sources.get(word, "")),
                "senses": senses,
            }

    def _request(self, batch, endpoint) -> list[dict]:
        content = chat_completion(
            self.config, endpoint, build_sense_prompt(batch), max_tokens=1500
        )
        items = parse_senses(content)
        if not items:
            raise ValueError("接口没有返回可解析的义项")
        return items


class FamilyService(WorkerService):
    """用 AI 精简同根词，只保留常考常用的派生词。"""

    CACHE_FILE = FAMILY_CACHE_FILE
    TASK = "family"

    @staticmethod
    def fingerprint(text: str) -> str:
        return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12] if text else ""

    def _valid_value(self, value) -> bool:
        return isinstance(value, dict) and isinstance(value.get("text"), str)

    def _ready(self, word: str, text: str) -> bool:
        entry = self.cache.get(word.lower())
        return isinstance(entry, dict) and entry.get("src") == self.fingerprint(text)

    def get(self, word: str, text: str) -> str | None:
        entry = self.cache.get(word.lower())
        if not isinstance(entry, dict) or entry.get("src") != self.fingerprint(text):
            return None
        value = entry.get("text")
        return value if isinstance(value, str) and value.strip() else None

    def store(self, items: list[dict], batch: list[tuple[str, str]] | None = None) -> None:
        sources = {word.lower(): text for word, text in (batch or [])}
        for item in items:
            word = str(item.get("word", "")).strip().lower()
            if not word:
                continue
            rendered = render_family(item)
            self.cache[word] = {
                "src": self.fingerprint(sources.get(word, "")),
                "text": rendered,
            }

    def _request(self, batch, endpoint) -> list[dict]:
        content = chat_completion(
            self.config, endpoint, build_family_prompt(batch), max_tokens=2000
        )
        items = parse_family(content)
        if not items:
            raise ValueError("接口没有返回可解析的同根词")
        return items
