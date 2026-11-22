# -*- coding: utf-8 -*-
"""词汇表读取：支持 .xlsx 与 .json 两种格式，并带按文件独立的解析缓存。"""

from __future__ import annotations

import hashlib
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CACHE_VERSION = 1
SUPPORTED_SUFFIXES = (".xlsx", ".json")

FIELDS = (
    "word",
    "ipa",
    "ipa_us",
    "translation",
    "phrase",
    "example",
    "synonym",
    "family",
    "origin",
)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HEADER_FIELDS = {
    "单词": "word",
    "音标①": "ipa",
    "音标②": "ipa_us",
    "翻译": "translation",
    "例句": "example",
    "短语": "phrase",
    "近义词": "synonym",
    "同根词": "family",
    "词源": "origin",
}
DEFAULT_COLUMNS = {
    "word": 0,
    "ipa": 1,
    "ipa_us": 2,
    "translation": 4,
    "example": 5,
    "phrase": 6,
    "synonym": 7,
    "family": 8,
    "origin": 9,
}


# ---------- xlsx ----------


def _column_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref)
    index = 0
    for char in letters.group(0) if letters else "A":
        index = index * 26 + ord(char) - 64
    return index - 1


def _all_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.iter(NS + "t"))


def _read_sheet_rows(path: Path) -> list[dict[int, str]]:
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = [_all_text(si) for si in root.findall(NS + "si")]

        sheets = [n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
        sheets.sort(key=lambda n: int(re.search(r"sheet(\d+)\.xml", n).group(1)))
        if not sheets:
            raise ValueError("这个 xlsx 里没有工作表")
        sheet = ET.fromstring(archive.read(sheets[0]))

    rows: list[dict[int, str]] = []
    for row in sheet.iter(NS + "row"):
        cells: dict[int, str] = {}
        for cell in row.findall(NS + "c"):
            kind = cell.get("t")
            value_node = cell.find(NS + "v")
            if kind == "s" and value_node is not None:
                value = shared[int(value_node.text)]
            elif kind == "inlineStr":
                inline = cell.find(NS + "is")
                value = _all_text(inline) if inline is not None else ""
            elif value_node is not None:
                value = value_node.text or ""
            else:
                value = ""
            if value.strip():
                cells[_column_index(cell.get("r", "A1"))] = value
        if cells:
            rows.append(cells)
    return rows


def _clean(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _compact(text: str) -> str:
    """去掉条目之间的空行，靠换行区分即可。"""
    return re.sub(r"\n[ \t]*\n+", "\n", text)


def load_xlsx(path: Path) -> list[dict[str, str]]:
    rows = _read_sheet_rows(path)
    if not rows:
        return []

    columns = dict(DEFAULT_COLUMNS)
    for index, header in rows[0].items():
        for title, field in HEADER_FIELDS.items():
            if header.strip().startswith(title):
                columns[field] = index
                break

    words: list[dict[str, str]] = []
    for cells in rows[1:]:
        word = _clean(cells.get(columns["word"], ""))
        if not word:
            continue
        words.append(
            {
                "word": word,
                "ipa": _clean(cells.get(columns["ipa"], "")),
                "ipa_us": _clean(cells.get(columns["ipa_us"], "")),
                "translation": _clean(cells.get(columns["translation"], "")),
                "phrase": _compact(_clean(cells.get(columns["phrase"], ""))),
                "example": _compact(_clean(cells.get(columns["example"], ""))),
                "synonym": _compact(_clean(cells.get(columns["synonym"], ""))),
                "family": _compact(_clean(cells.get(columns["family"], ""))),
                "origin": _compact(_clean(cells.get(columns["origin"], ""))),
            }
        )
    return words


# ---------- json ----------


def _join_pairs(items, key_english: str, key_chinese: str) -> str:
    lines: list[str] = []
    if not isinstance(items, list):
        return ""
    for item in items:
        if not isinstance(item, dict):
            continue
        english = str(item.get(key_english) or "").strip()
        chinese = str(item.get(key_chinese) or "").strip()
        if english and chinese:
            lines.append(f"{english}\n{chinese}")
        elif english or chinese:
            lines.append(english or chinese)
    return _compact(_clean("\n".join(lines)))


def _json_translation(item: dict) -> str:
    lines: list[str] = []
    for trans in item.get("trans") or []:
        if isinstance(trans, str):
            text = trans.strip()
            if text:
                lines.append(text)
            continue
        if not isinstance(trans, dict):
            continue
        pos = str(trans.get("pos") or "").strip()
        cn = str(trans.get("cn") or "").strip()
        if not (pos or cn):
            continue
        lines.append(f"{pos} {cn}".strip())
    return _clean("\n".join(lines))


def _json_phonetic(item: dict, *keys: str) -> str:
    for key in keys:
        value = item.get(key)
        if isinstance(value, str) and value.strip():
            return _clean(value)
    return ""


def _json_synonym(item: dict) -> str:
    lines: list[str] = []
    for syno in item.get("synos") or []:
        if not isinstance(syno, dict):
            continue
        pos = str(syno.get("pos") or "").strip()
        cn = str(syno.get("cn") or "").strip()
        ws = [str(word).strip() for word in (syno.get("ws") or []) if str(word).strip()]
        head = f"{pos} {cn}".strip()
        if head:
            lines.append(head)
        if ws:
            lines.append(" / ".join(ws))
    return _compact(_clean("\n".join(lines)))


def _json_family(item: dict) -> str:
    rel = item.get("relWords") or {}
    if not isinstance(rel, dict):
        return ""
    blocks: list[str] = []
    root = str(rel.get("root") or "").strip()
    if root:
        blocks.append(f"词根:{root}")
    for group in rel.get("rels") or []:
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
    return _compact(_clean("\n\n".join(blocks)))


def _json_origin(item: dict) -> str:
    lines: list[str] = []
    for entry in item.get("etymology") or []:
        if not isinstance(entry, dict):
            continue
        title = str(entry.get("t") or "").strip()
        detail = str(entry.get("d") or "").strip()
        if detail:
            lines.append(f"{title} {detail}".strip())
        elif title:
            lines.append(title)
    return _compact(_clean("\n".join(lines)))


def load_json(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig") as handle:
        data = json.load(handle)

    if isinstance(data, dict):
        for key in ("words", "items", "data", "list"):
            if isinstance(data.get(key), list):
                data = data[key]
                break
    if not isinstance(data, list):
        raise ValueError("JSON 词汇表应为单词对象数组")

    words: list[dict[str, str]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        word = str(item.get("word") or item.get("name") or "").strip()
        if not word:
            continue
        words.append(
            {
                "word": word,
                "ipa": _json_phonetic(item, "phonetic0", "phonetic", "ukphone"),
                "ipa_us": _json_phonetic(item, "phonetic1", "phonetic_us", "usphone"),
                "translation": _json_translation(item),
                "phrase": _join_pairs(item.get("phrases"), "c", "cn"),
                "example": _join_pairs(item.get("sentences"), "c", "cn"),
                "synonym": _json_synonym(item),
                "family": _json_family(item),
                "origin": _json_origin(item),
            }
        )
    return words


# ---------- 统一入口 ----------


def load_vocabulary(path: Path) -> list[dict[str, str]]:
    suffix = path.suffix.lower()
    if suffix == ".json":
        return load_json(path)
    if suffix == ".xlsx":
        return load_xlsx(path)
    raise ValueError(f"不支持的格式：{suffix}")


def cache_file_for(path: Path) -> Path:
    digest = hashlib.sha1(str(path.resolve()).lower().encode("utf-8")).hexdigest()[:10]
    return APP_DIR / f".vocab_cache_{digest}.json"


def _hide(path: Path) -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x2)
    except Exception:
        pass


def load_cached(path: Path) -> list[dict[str, str]]:
    stat = path.stat()
    cache_file = cache_file_for(path)
    try:
        cached = json.loads(cache_file.read_text("utf-8"))
        if (
            cached.get("version") == CACHE_VERSION
            and cached.get("mtime") == stat.st_mtime
            and cached.get("size") == stat.st_size
            and isinstance(cached.get("words"), list)
        ):
            return cached["words"]
    except (OSError, ValueError, KeyError):
        pass

    words = load_vocabulary(path)
    try:
        cache_file.write_text(
            json.dumps(
                {"version": CACHE_VERSION, "mtime": stat.st_mtime, "size": stat.st_size, "words": words},
                ensure_ascii=False,
            ),
            "utf-8",
        )
        _hide(cache_file)
    except OSError:
        pass
    return words


def clean_stale_caches(active: list[Path] | None = None) -> None:
    """删除已不存在词表的解析缓存，避免 .vocab_cache_* 越积越多。

    保留规则：程序目录与 dicts 目录里现存的词表、以及本次启动用到的词表（active），
    其余缓存一律删除（例如之前从别处打开过、后来移走的文件）。
    """
    keep = {cache_file_for(Path(item)).name for item in (active or [])}
    for folder in (APP_DIR, APP_DIR / "dicts"):
        if not folder.is_dir():
            continue
        for path in folder.iterdir():
            if (
                path.is_file()
                and path.suffix.lower() in SUPPORTED_SUFFIXES
                and not path.name.startswith(("~$", "."))
            ):
                keep.add(cache_file_for(path).name)
    for cache in APP_DIR.glob(".vocab_cache_*.json"):
        if cache.name not in keep:
            try:
                cache.unlink()
            except OSError:
                pass


def find_vocabularies(folder: Path) -> list[Path]:
    found = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in SUPPORTED_SUFFIXES
        and not path.name.startswith(("~$", "."))
    ]
    return sorted(found, key=lambda path: path.name.lower())


def find_default(folder: Path, preferred: str = "") -> Path | None:
    if preferred:
        candidate = Path(preferred)
        if not candidate.is_absolute():
            candidate = folder / candidate
        if candidate.is_file():
            return candidate
    files = find_vocabularies(folder)
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)
