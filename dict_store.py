# -*- coding: utf-8 -*-
"""在线词库：从 qwerty-learner 仓库读取词库清单，并把选中的词库下载到本地。

清单来自仓库里的 src/resources/dictionary.ts（一段对象字面量），下载的词典文件是
public/dicts/*.json，条目形如 {"name", "trans": [...], "usphone", "ukphone"}。
下载后重新包装成 {"__meta": {...}, "words": [...]}，vocab.load_json 可直接读取。
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DICT_DIR = APP_DIR / "dicts"
INDEX_CACHE = DICT_DIR / ".index.json"

SOURCE_NAME = "qwerty-learner"
SOURCE_PAGE = "https://github.com/RealKai42/qwerty-learner"

# 依次尝试，第一个能连上就够。国内访问 GitHub 直连常失败，后两个是加速镜像。
MIRRORS = (
    ("GitHub", "https://raw.githubusercontent.com/RealKai42/qwerty-learner/master"),
    ("jsDelivr", "https://cdn.jsdelivr.net/gh/RealKai42/qwerty-learner@master"),
    ("ghfast", "https://ghfast.top/https://raw.githubusercontent.com/RealKai42/qwerty-learner/master"),
    ("gh-proxy", "https://gh-proxy.com/https://raw.githubusercontent.com/RealKai42/qwerty-learner/master"),
)
INDEX_PATH = "/src/resources/dictionary.ts"
DICT_PATH = "/public/dicts"

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WriteDownSomeWords/1.0"

_ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]+')

_FIELD_PATTERNS = {
    "id": re.compile(r"\bid:\s*'([^']*)'"),
    "name": re.compile(r"\bname:\s*'([^']*)'"),
    "description": re.compile(r"\bdescription:\s*'([^']*)'"),
    "category": re.compile(r"\bcategory:\s*'([^']*)'"),
    "tags": re.compile(r"\btags:\s*\[([^\]]*)\]"),
    "url": re.compile(r"\burl:\s*'([^']*)'"),
    "length": re.compile(r"\blength:\s*(\d+)"),
    "language": re.compile(r"\blanguageCategory:\s*'([^']*)'"),
}
_TAG = re.compile(r"'([^']*)'")
_OBJECT = re.compile(r"\{[^{}]*\}")


class StoreError(RuntimeError):
    """词库源不可用或文件无法解析。"""


def safe_filename(name: str, fallback: str = "vocab") -> str:
    text = _ILLEGAL.sub("_", (name or "").strip()).strip(" .")
    return text or fallback


def _read(url: str, timeout: float, progress=None) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        total = int(response.headers.get("Content-Length") or 0)
        chunks: list[bytes] = []
        done = 0
        while True:
            chunk = response.read(65536)
            if not chunk:
                break
            chunks.append(chunk)
            done += len(chunk)
            if progress is not None:
                progress(done, total)
        return b"".join(chunks)


def parse_index(text: str) -> list[dict]:
    metas: list[dict] = []
    seen: set[str] = set()
    for match in _OBJECT.finditer(text):
        block = match.group(0)
        entry: dict = {}
        for key, pattern in _FIELD_PATTERNS.items():
            found = pattern.search(block)
            if not found:
                continue
            entry[key] = found.group(1).strip()
        url = entry.get("url", "")
        if not entry.get("id") or "/" not in url:
            continue
        if entry["id"] in seen:
            continue
        entry["file"] = url.rsplit("/", 1)[-1]
        entry["tags"] = [tag for tag in _TAG.findall(entry.get("tags", "")) if tag]
        entry["length"] = int(entry.get("length") or 0)
        entry.setdefault("name", entry["id"])
        entry.setdefault("description", "")
        entry.setdefault("category", "其他")
        seen.add(entry["id"])
        metas.append(entry)
    return metas


def fetch_index(timeout: float = 15, progress=None, on_try=None) -> tuple[list[dict], str]:
    """联网读取词库清单，返回 (词库列表, 可用镜像名)。"""
    errors: list[str] = []
    for label, root in MIRRORS:
        if on_try is not None:
            on_try(label)
        try:
            raw = _read(root + INDEX_PATH, timeout)
        except (urllib.error.URLError, OSError, ValueError) as exc:
            errors.append(f"{label}: {exc}")
            continue
        metas = parse_index(raw.decode("utf-8", "replace"))
        if not metas:
            errors.append(f"{label}: 清单为空")
            continue
        _write_index_cache(metas, label)
        if progress is not None:
            progress(len(metas), len(metas))
        return metas, label
    detail = "；".join(errors[:2]) or "未知错误"
    raise StoreError(f"无法获取词库清单（{detail}）")


def _write_index_cache(metas: list[dict], mirror: str) -> None:
    try:
        DICT_DIR.mkdir(parents=True, exist_ok=True)
        INDEX_CACHE.write_text(
            json.dumps(
                {
                    "source": SOURCE_PAGE,
                    "mirror": mirror,
                    "saved": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "dicts": metas,
                },
                ensure_ascii=False,
            ),
            "utf-8",
        )
    except OSError:
        pass


def cached_index() -> tuple[list[dict], str, str]:
    """读取上次缓存的清单，返回 (词库列表, 镜像名, 保存时间)；没有缓存时返回空。"""
    try:
        data = json.loads(INDEX_CACHE.read_text("utf-8"))
    except (OSError, ValueError):
        return [], "", ""
    metas = [item for item in data.get("dicts") or [] if isinstance(item, dict) and item.get("id")]
    return metas, str(data.get("mirror") or ""), str(data.get("saved") or "")


def target_path(meta: dict, dest_dir: Path | None = None) -> Path:
    folder = Path(dest_dir or DICT_DIR)
    name = safe_filename(meta.get("name", ""), meta.get("id", "vocab"))
    return folder / f"{name}.json"


def _file_id(path: Path) -> str:
    try:
        data = json.loads(path.read_text("utf-8-sig"))
    except (OSError, ValueError):
        return ""
    if isinstance(data, dict):
        meta = data.get("__meta")
        if isinstance(meta, dict):
            return str(meta.get("id") or "")
    return ""


def _resolve_target(meta: dict, dest_dir: Path) -> Path:
    target = target_path(meta, dest_dir)
    if target.exists() and _file_id(target) not in ("", meta.get("id")):
        target = dest_dir / f"{target.stem} ({meta['id']}).json"
    return target


def download(
    meta: dict,
    mirror: str = "",
    dest_dir: Path | None = None,
    timeout: float = 90,
    progress=None,
) -> Path:
    """下载一个词库，转换成程序可读的格式后返回本地路径。"""
    folder = Path(dest_dir or DICT_DIR)
    folder.mkdir(parents=True, exist_ok=True)
    filename = meta.get("file") or ""
    if not filename:
        raise StoreError("词库信息缺少文件名")

    candidates = [item for item in MIRRORS if not mirror or item[0] == mirror] or list(MIRRORS)
    payload = None
    errors: list[str] = []
    for label, root in candidates:
        try:
            raw = _read(f"{root}{DICT_PATH}/{filename}", timeout, progress)
            payload = json.loads(raw.decode("utf-8-sig"))
            break
        except (urllib.error.URLError, OSError, ValueError) as exc:
            errors.append(f"{label}: {exc}")

    if payload is None:
        detail = "；".join(errors[:2]) or "未知错误"
        raise StoreError(f"下载失败（{detail}）")

    if isinstance(payload, dict):
        entries = (
            payload.get("dictionary")
            or payload.get("words")
            or payload.get("items")
            or payload.get("data")
            or []
        )
    else:
        entries = payload
    entries = [item for item in entries if isinstance(item, dict)]
    if not entries:
        raise StoreError(f"{filename} 里没有单词")

    target = _resolve_target(meta, folder)
    meta = dict(meta)
    meta["downloaded"] = time.strftime("%Y-%m-%d %H:%M:%S")
    meta["source"] = SOURCE_PAGE
    target.write_text(
        json.dumps({"__meta": meta, "words": entries}, ensure_ascii=False),
        "utf-8",
    )
    return target


def _meta_from_head(path: Path) -> tuple[dict, int]:
    """只读文件开头，取出写在最前面的 __meta（下载时就是这么排的）。"""
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            head = handle.read(4096)
    except OSError:
        return {}, 0
    match = re.search(r'"__meta"\s*:\s*(\{)', head)
    if not match:
        return {}, 0
    try:
        meta, _ = json.JSONDecoder().raw_decode(head[match.start(1) :])
    except ValueError:
        return {}, 0
    if not isinstance(meta, dict) or not meta.get("id"):
        return {}, 0
    try:
        count = int(meta.get("length") or 0)
    except (TypeError, ValueError):
        count = 0
    return meta, count


def _meta_from_full(path: Path) -> tuple[dict, int]:
    try:
        data = json.loads(path.read_text("utf-8-sig"))
    except (OSError, ValueError):
        return {}, 0
    if isinstance(data, dict):
        meta = data.get("__meta")
        meta = dict(meta) if isinstance(meta, dict) else {}
        for key in ("words", "dictionary", "items", "data", "list"):
            if isinstance(data.get(key), list):
                return meta, len(data[key])
        return meta, 0
    if isinstance(data, list):
        return {}, len(data)
    return {}, 0


def local_index(dest_dir: Path | None = None) -> dict[str, dict]:
    """扫描下载目录，返回 {词库 id: {名称、路径、词数…}}；同时收录目录里其他 JSON 词表。"""
    folder = Path(dest_dir or DICT_DIR)
    result: dict[str, dict] = {}
    if not folder.is_dir():
        return result
    for path in sorted(folder.glob("*.json")):
        if path.name.startswith("."):
            continue
        meta, count = _meta_from_head(path)
        if not meta:
            meta, count = _meta_from_full(path)
            if not meta and not count:
                continue
        key = str(meta.get("id") or f"file:{path.name}")
        result[key] = {
            "id": str(meta.get("id") or ""),
            "name": str(meta.get("name") or path.stem),
            "category": str(meta.get("category") or ""),
            "description": str(meta.get("description") or ""),
            "path": path,
            "count": count,
            "from_source": bool(meta.get("id")),
        }
    return result


def remove(path: Path) -> None:
    Path(path).unlink(missing_ok=True)
