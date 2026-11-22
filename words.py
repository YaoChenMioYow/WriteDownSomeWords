# -*- coding: utf-8 -*-
"""单词抄写展示器。

读取词汇表（.xlsx 或 .json），逐个显示：单词、音标、词性/义项提示、翻译、短语、
例句、近义词，按 Tab 展开同根词与词源。字号会根据词条长短自动适配，右侧边栏用卡片
列出全部单词，点击即可跳转。可在程序内切换词汇表，也可把文件直接拖进窗口。
还能从「在线词库市场」（qwerty-learner，Ctrl+M）挑选词库、下载到程序目录后直接使用。
单词与例句旁的 ▶ 按钮可以朗读（pyttsx3），也能在设置里打开「换词后自动朗读」。
背词模式（H）隐藏释义、回车看答案；数字键 1/2 给当前词标记「已掌握/生词」，
侧边栏可只看生词，词汇表菜单里能一键导出生词本；R 随机跳词。
抄写与背词的进度分开记忆：切模式自动跳回各自的上次进度，互不打扰。
窗口大小位置与字号缩放会被记住，下次启动原样恢复。

用法：
    python words.py [词汇表.xlsx|.json]
"""

from __future__ import annotations

import ctypes
import json
import queue
import random
import re
import sys
import threading
from pathlib import Path

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk

import dict_market
import dict_store
import speech
import usage_ai
import vocab

APP_DIR = Path(__file__).resolve().parent
STATE_FILE = APP_DIR / ".words_state.json"

try:  # 拖放支持：装了 tkinterdnd2 就能直接拖文件进来
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:  # pragma: no cover - 环境缺少该可选依赖
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

THEME_LABELS = {"dark": "深色", "light": "浅色", "system": "跟随系统"}

# 展示区块（顺序即渲染顺序）；单词与音标始终显示
SECTION_ORDER = (
    ("translation", "翻译"),
    ("phrase", "短语"),
    ("usage", "用法"),
    ("example", "例句"),
    ("synonym", "近义词"),
    ("family", "同根词"),
    ("origin", "词源"),
)
DEFAULT_SECTIONS = {
    "translation": True,
    "phrase": True,
    "usage": True,
    "example": True,
    "synonym": True,
    "family": False,
    "origin": False,
}
THEMES = {
    "dark": {
        "bg": "#161719",
        "panel": "#1D1F22",
        "card": "#26282D",
        "card_border": "#34373D",
        "card_active": "#3A2E1F",
        "card_active_border": "#8A6437",
        "ink": "#F0EEE9",
        "body": "#CFCBC3",
        "muted": "#8B8880",
        "accent": "#D9A05C",
        "select": "#3A3F4A",
    },
    "light": {
        "bg": "#FAF8F3",
        "panel": "#F2EDE3",
        "card": "#FFFFFF",
        "card_border": "#E6DFD2",
        "card_active": "#FBEFE1",
        "card_active_border": "#E3C9A6",
        "ink": "#1F1D1A",
        "body": "#33302B",
        "muted": "#8C8579",
        "accent": "#A8743C",
        "select": "#DCE4F0",
    },
}


def system_theme() -> str:
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "light" if value else "dark"
    except OSError:
        return "dark"


def set_dark_titlebar(root: tk.Misc, dark: bool) -> None:
    try:
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(value), ctypes.sizeof(value)
        )
    except Exception:
        pass

def find_vocabulary() -> Path:
    if len(sys.argv) > 1:
        candidate = Path(sys.argv[1]).expanduser()
        if not candidate.is_file():
            raise SystemExit(f"找不到词汇表文件：{candidate}")
        return candidate
    found = vocab.find_default(APP_DIR, str(usage_ai.load_config().get("vocab", "")))
    if found is None:
        raise SystemExit(
            "程序目录里没有找到 .xlsx / .json 词汇表，请把词表文件放进程序目录后再启动"
        )
    return found


POS_RE = re.compile(r"^\s*(vt|vi|aux|adj|adv|abbr|prep|pron|conj|num|art|int|n|v)\s*\.")
ROOT_RE = re.compile(r"词根[:：]\s*([A-Za-z][A-Za-z\-]{1,23})")

MARK_KNOWN = 1  # 已掌握：侧边栏卡片画 ✓
MARK_HARD = 2  # 生词：侧边栏「只看生词」过滤与「导出生词本」使用

MODE_COPY = "copy"  # 抄写模式（正常浏览）
MODE_REVIEW = "review"  # 背词模式（H 遮挡）


def _position_entry(value: object) -> dict[str, str]:
    """词位记录兼容两种格式：旧版是单词字符串，新版按 抄写/背词 分开记。"""
    if isinstance(value, dict):
        return {"copy": str(value.get("copy") or ""), "review": str(value.get("review") or "")}
    if isinstance(value, str) and value:
        return {"copy": value, "review": ""}
    return {"copy": "", "review": ""}


def _pos_tags(translation: str) -> list[str]:
    tags: list[str] = []
    for line in translation.split("\n"):
        match = POS_RE.match(line)
        if match:
            tag = match.group(1) + "."
            if tag not in tags:
                tags.append(tag)
    return tags


def _sense_count(translation: str) -> int:
    if not translation:
        return 0
    count = 0
    for line in translation.split("\n"):
        if line.strip():
            count += line.count("；") + line.count(";") + 1
    return count


def _root_hint(entry: dict[str, str]) -> str:
    for key in ("family", "origin"):
        match = ROOT_RE.search(entry[key])
        if match:
            return match.group(1)
    return ""


def _meaning_hint(entry: dict[str, str]) -> str:
    return entry["translation"].split("\n")[0].strip()[:40]


def _pick_family(candidates: tuple[str, ...], fallback: str) -> str:
    available = {name.lower(): name for name in tkfont.families()}
    for name in candidates:
        if name.lower() in available:
            return available[name.lower()]
    return fallback


class WordApp:
    WORD_SIZE = 44
    IPA_SIZE = 17
    META_SIZE = 10
    LABEL_SIZE = 10
    BODY_SIZE = 15
    UI_SIZE = 10

    MIN_SCALE = 0.35
    MAX_SCALE = 1.8
    # 自动适配最多缩到基准字号的这个比例，避免长词被压得看不清
    SHRINK_FLOOR = 0.6

    CARD_H = 58
    CARD_PAD = 10

    def __init__(self, root: tk.Tk, words: list[dict[str, str]], source: Path) -> None:
        self.root = root
        self.words = words
        self.source = source
        self.index = 0
        self.sections = dict(DEFAULT_SECTIONS)
        self.zoom = 1.0
        self.scale = 1.0
        self._revealed = False  # 背词模式下当前词是否已显示释义
        self.hard_only = False  # 侧边栏「只看生词」
        self.filtered = list(range(len(words)))
        self._search_index = [
            (entry["word"].lower(), entry["translation"].lower()) for entry in words
        ]
        self._running = False
        self._fit_job: str | None = None
        self._scale_cache: dict[tuple, dict[int, float]] = {}
        self._usage_save_job: str | None = None
        self._search_job: str | None = None
        self._poll_job: str | None = None
        self._watch_job: str | None = None
        self._flash_job: str | None = None
        self._settings_window: tk.Toplevel | None = None
        self._market: dict_market.DictMarket | None = None

        self.config = usage_ai.load_config()
        stored_sections = self.config.get("sections")
        if isinstance(stored_sections, dict):
            for key, value in stored_sections.items():
                if key in self.sections:
                    self.sections[key] = bool(value)
        self.sections_open = bool(self.config.get("sections_open", False))
        self.theme_mode = self.config.get("theme", "dark")
        self.theme_name = self._resolve_theme(self.theme_mode)
        self.colors = THEMES[self.theme_name]
        self._state = self._load_state()
        self.marks = self._state["marks"]
        self.mode_positions: dict[str, str] = {"copy": "", "review": ""}
        self.hide_mode = bool(self.config.get("hide_mode"))
        try:
            self.zoom = max(0.7, min(1.8, float(self.config.get("zoom") or 1.0)))
        except (TypeError, ValueError):
            self.zoom = 1.0
        self.speaker = speech.Speaker(
            voice_id=str(self.config.get("speak_voice") or ""),
            rate=int(self.config.get("speak_rate") or 170),
        )
        self._speak_buttons: list[tk.Button] = []
        self.usage_service = None
        self.sense_service = None
        self.family_service = None
        if self.config.get("endpoints") or str(self.config.get("api_key", "")).strip():
            self.usage_service = usage_ai.UsageService(self.config)
            self.usage_service.start()
            if self.config.get("trim_senses"):
                self.sense_service = usage_ai.SenseService(self.config)
                self.sense_service.start()
            if self.config.get("trim_families"):
                self.family_service = usage_ai.FamilyService(self.config)
                self.family_service.start()
            for attr in ("usage_service", "sense_service", "family_service"):
                service = getattr(self, attr)
                if service is not None and not service.endpoints:
                    service.stop()
                    setattr(self, attr, None)

        word_family = _pick_family(("Georgia", "Cambria", "Palatino Linotype"), "Times New Roman")
        text_family = _pick_family(
            ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "SimHei"), "TkDefaultFont"
        )
        self.f_word = tkfont.Font(family=word_family, size=self.WORD_SIZE, weight="bold")
        self.f_ipa = tkfont.Font(family="Segoe UI", size=self.IPA_SIZE)
        self.f_meta = tkfont.Font(family=text_family, size=self.META_SIZE)
        self.f_label = tkfont.Font(family=text_family, size=self.LABEL_SIZE, weight="bold")
        self.f_body = tkfont.Font(family=text_family, size=self.BODY_SIZE)
        self.f_ui = tkfont.Font(family=text_family, size=self.UI_SIZE)
        self.f_card_word = tkfont.Font(family=word_family, size=12, weight="bold")
        self.f_card_sub = tkfont.Font(family="Segoe UI", size=9)
        self.fonts = {
            "word": self.f_word,
            "ipa": self.f_ipa,
            "meta": self.f_meta,
            "label": self.f_label,
            "body": self.f_body,
        }
        # 离屏测量用的独立字体，避免试算字号时带动正文重排
        self.m_fonts = {
            "word": tkfont.Font(family=word_family, size=self.WORD_SIZE, weight="bold"),
            "ipa": tkfont.Font(family="Segoe UI", size=self.IPA_SIZE),
            "meta": tkfont.Font(family=text_family, size=self.META_SIZE),
            "label": tkfont.Font(family=text_family, size=self.LABEL_SIZE, weight="bold"),
            "body": tkfont.Font(family=text_family, size=self.BODY_SIZE),
        }
        self.measure: tk.Text | None = None
        self._chunks: list[tuple[str, str]] = []

        self.root.title(f"单词抄写 · {source.stem}")
        if (APP_DIR / "icon.ico").is_file():
            icon_path = str(APP_DIR / "icon.ico")
            try:
                self.root.iconbitmap(icon_path)
                self.root.iconbitmap(default=icon_path)
            except tk.TclError:
                pass
        self.root.configure(bg=self.colors["bg"])
        self.root.minsize(900, 560)
        geometry = self._state.get("geometry") or ""
        if re.fullmatch(r"\d+x\d+[+-]\d+[+-]\d+", geometry):
            self.root.geometry(geometry)
        else:
            self.root.geometry("1320x880")

        self._build_sidebar()
        self._build_main()
        self.apply_theme(self.theme_name)
        self._restore_position()
        self._bind_keys()
        self._enable_drop()
        self._setup_ime()
        self.render()
        self.root.after(80, self._draw_cards)
        self.root.after(250, self._poll_usage)
        self.root.after(2000, self._watch_theme)
        self._running = True

    # ---------- 界面 ----------

    def _build_main(self) -> None:
        colors = self.colors
        main = tk.Frame(self.root, bg=colors["bg"])
        main.pack(side="left", fill="both", expand=True)
        self.main_frame = main

        header = tk.Frame(main, bg=colors["bg"])
        header.pack(fill="x", padx=44, pady=(22, 2))
        self.header = header
        self.title_label = tk.Label(
            header, text=self.source.stem, bg=colors["bg"], fg=colors["muted"], font=self.f_ui
        )
        self.title_label.pack(side="left")
        self.vocab_button = tk.Menubutton(
            header,
            text="切换词汇表 ▾",
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=8,
            takefocus=0,
            bg=colors["bg"],
            fg=colors["muted"],
            activebackground=colors["panel"],
            activeforeground=colors["ink"],
        )
        self.vocab_button.pack(side="left", padx=(14, 0))
        self.vocab_menu = tk.Menu(
            self.vocab_button,
            tearoff=0,
            bg=colors["panel"],
            fg=colors["ink"],
            activebackground=colors["accent"],
            activeforeground=colors["bg"],
            font=self.f_ui,
        )
        self.vocab_button.configure(menu=self.vocab_menu)
        self.vocab_button.bind("<Button-1>", self._refresh_vocab_menu, add="+")
        self._refresh_vocab_menu()

        self.sections_button = tk.Button(
            header,
            text="显示内容 ▸",
            command=self.toggle_sections_bar,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=8,
            takefocus=0,
            bg=colors["bg"],
            fg=colors["muted"],
            activebackground=colors["panel"],
            activeforeground=colors["ink"],
        )
        self.sections_button.pack(side="left", padx=(6, 0))
        self.progress_label = tk.Label(
            header, text="", bg=colors["bg"], fg=colors["muted"], font=self.f_ui
        )
        self.progress_label.pack(side="right")
        self.scale_label = tk.Label(
            header, text="", bg=colors["bg"], fg=colors["muted"], font=self.f_ui
        )
        self.scale_label.pack(side="right", padx=(0, 18))
        self.settings_button = tk.Button(
            header,
            text="设置",
            command=self.open_settings,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=10,
            takefocus=0,
            bg=colors["bg"],
            fg=colors["muted"],
            activebackground=colors["panel"],
            activeforeground=colors["ink"],
        )
        self.settings_button.pack(side="right", padx=(0, 16))

        area = tk.Frame(main, bg=colors["bg"])
        area.pack(fill="both", expand=True, padx=(40, 34))
        self.area = area
        self.text = tk.Text(
            area,
            bg=colors["bg"],
            fg=colors["body"],
            bd=0,
            highlightthickness=0,
            wrap="word",
            padx=8,
            pady=10,
            cursor="arrow",
            insertwidth=0,
            takefocus=0,
            insertbackground=colors["ink"],
            selectbackground=colors["select"],
            selectforeground=colors["ink"],
        )
        self.scrollbar = ttk.Scrollbar(
            area, orient="vertical", command=self.text.yview, style="Main.Vertical.TScrollbar"
        )
        self.text.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)

        self.text.tag_configure("word", font=self.f_word, foreground=colors["ink"])
        self.text.tag_configure("ipa", font=self.f_ipa, foreground=colors["muted"])
        self.text.tag_configure("meta", font=self.f_meta, foreground=colors["muted"])
        self.text.tag_configure("meta_warn", font=self.f_meta, foreground=colors["accent"])
        self.text.tag_configure("label", font=self.f_label, foreground=colors["accent"])
        self.text.tag_configure("body", font=self.f_body, foreground=colors["body"])
        self.text.tag_configure("usage", font=self.f_body, foreground=colors["ink"])
        self.text.tag_configure("usage_meaning", font=self.f_body, foreground=colors["muted"])
        self.text.tag_configure("usage_example", font=self.f_meta, foreground=colors["muted"])
        self.text.configure(state="disabled")
        self.text.bind("<Configure>", self._on_text_configure)

        hint = tk.Frame(main, bg=colors["bg"])
        hint.pack(fill="x", padx=44, pady=(2, 14))
        self.hint = hint
        self.hint_label = tk.Label(
            hint,
            text=(
                "空格 下一个 · ← 上一个 · H 背词 · 回车 看答案 · 1/2 标记 · R 随机 · Ctrl+F 搜索 · "
                "Ctrl+M 词库 · S 朗读 · Tab 同根词/词源 · U 用法 · Ctrl+U 精简翻译 · Ctrl+Shift+U 精简同根词 · "
                "+ / − 字号 · F11 全屏 · Esc 退出"
            ),
            bg=colors["bg"],
            fg=colors["muted"],
            font=self.f_ui,
        )
        self.hint_label.pack(side="left")

        self._build_sections_bar(main)
        self._layout_sections_bar()

    def _build_sidebar(self) -> None:
        colors = self.colors
        sidebar = tk.Frame(self.root, bg=colors["panel"], width=300)
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)
        self.sidebar = sidebar

        head = tk.Frame(sidebar, bg=colors["panel"])
        head.pack(fill="x", padx=20, pady=(22, 8))
        self.sidebar_head = head
        self.sidebar_title = tk.Label(
            head, text="全部单词", bg=colors["panel"], fg=colors["ink"], font=self.f_ui
        )
        self.sidebar_title.pack(side="left")
        self.sidebar_count = tk.Label(
            head, text=f"{len(self.words)} 词", bg=colors["panel"], fg=colors["muted"], font=self.f_ui
        )
        self.sidebar_count.pack(side="right")

        self.hard_only_var = tk.BooleanVar(value=False)
        self.hard_only_check = tk.Checkbutton(
            head,
            text="只看生词",
            variable=self.hard_only_var,
            command=self._on_hard_filter_toggle,
            bg=colors["panel"],
            fg=colors["ink"],
            activebackground=colors["panel"],
            activeforeground=colors["accent"],
            selectcolor=colors["card"],
            font=self.f_ui,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            takefocus=0,
        )
        self.hard_only_check.pack(side="left", padx=(12, 0))

        self.search_frame = tk.Frame(sidebar, bg=colors["panel"])
        self.search_frame.pack(fill="x", padx=(16, 10), pady=(0, 10))
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            self.search_frame,
            textvariable=self.search_var,
            bg=colors["card"],
            fg=colors["ink"],
            insertbackground=colors["ink"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=colors["card_border"],
            highlightcolor=colors["accent"],
            font=self.f_ui,
            width=10,
        )
        self.search_entry.pack(fill="x", ipady=4)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        self.search_entry.bind("<Return>", self._on_search_enter)
        self.search_entry.bind("<Escape>", self._on_search_escape)
        self.search_entry.bind("<FocusIn>", self._on_search_focus_in)
        self.search_entry.bind("<FocusOut>", self._on_search_focus_out)
        self._show_search_placeholder()

        body = tk.Frame(sidebar, bg=colors["panel"])
        body.pack(fill="both", expand=True, padx=(14, 4), pady=(0, 14))
        self.sidebar_body = body
        self.cards = tk.Canvas(
            body,
            bg=colors["panel"],
            bd=0,
            highlightthickness=0,
            yscrollincrement=self.CARD_H,
        )
        self.cards_sb = ttk.Scrollbar(
            body, orient="vertical", command=self._scroll_cards, style="Side.Vertical.TScrollbar"
        )
        self.cards.configure(yscrollcommand=self.cards_sb.set)
        self.cards_sb.pack(side="right", fill="y", padx=(2, 0))
        self.cards.pack(side="left", fill="both", expand=True)
        self.cards.bind("<Configure>", lambda event: self._draw_cards())
        self.cards.bind("<Button-1>", self._on_card_click)

    # ---------- 搜索 ----------

    SEARCH_PLACEHOLDER = "搜索单词或中文…"

    def _show_search_placeholder(self) -> None:
        self.search_var.set(self.SEARCH_PLACEHOLDER)
        self.search_entry.configure(fg=self.colors["muted"])

    def _on_search_focus_in(self, event=None) -> None:
        if self.search_var.get() == self.SEARCH_PLACEHOLDER:
            self.search_var.set("")
            self.search_entry.configure(fg=self.colors["ink"])

    def _on_search_focus_out(self, event=None) -> None:
        if not self.search_var.get().strip():
            self._show_search_placeholder()

    def focus_search(self, event=None):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        return "break"

    def _on_search_key(self, event=None) -> None:
        if event is not None and event.keysym in ("Escape", "Return", "Up", "Down", "Tab"):
            return
        if self._search_job is not None:
            self.root.after_cancel(self._search_job)
        self._search_job = self.root.after(120, self._run_search)

    def _run_search(self) -> None:
        self._search_job = None
        self.apply_filter(select_first=True)

    def _on_search_enter(self, event=None):
        if self._search_job is not None:
            self.root.after_cancel(self._search_job)
            self._search_job = None
        self.apply_filter(select_first=True, force=True)
        return "break"

    def _on_search_escape(self, event=None):
        if self._search_job is not None:
            self.root.after_cancel(self._search_job)
            self._search_job = None
        self.search_var.set("")
        self.apply_filter()
        self.cards.focus_set()
        self._show_search_placeholder()
        return "break"

    def query(self) -> str:
        text = self.search_var.get().strip()
        return "" if text == self.SEARCH_PLACEHOLDER else text.lower()

    def apply_filter(self, select_first: bool = False, force: bool = False) -> None:
        query = self.query()
        if not query:
            self.filtered = list(range(len(self.words)))
        else:
            exact: list[int] = []
            prefix: list[int] = []
            partial: list[int] = []
            meaning: list[int] = []
            for index, (word, translation) in enumerate(self._search_index):
                if word == query:
                    exact.append(index)
                elif word.startswith(query):
                    prefix.append(index)
                elif query in word:
                    partial.append(index)
                elif query in translation:
                    meaning.append(index)
            self.filtered = exact + prefix + partial + meaning
        if self.hard_only:
            self.filtered = [
                index
                for index in self.filtered
                if self._word_mark(self.words[index]["word"]) == MARK_HARD
            ]

        self.sidebar_title.configure(
            text="搜索结果" if query else ("生词本" if self.hard_only else "全部单词")
        )
        if query:
            self.sidebar_count.configure(text=f"{len(self.filtered)} / {len(self.words)}")
        elif self.hard_only:
            self.sidebar_count.configure(text=f"{len(self.filtered)} 词")
        else:
            self.sidebar_count.configure(text=f"{len(self.words)} 词")
        self.cards.yview_moveto(0)
        self._draw_cards()

        if query and (select_first or force) and self.filtered:
            self.jump_to(self.filtered[0])

    def _filter_active(self) -> bool:
        return len(self.filtered) != len(self.words)

    def _filtered_position(self) -> int | None:
        try:
            return self.filtered.index(self.index)
        except ValueError:
            return None

    # ---------- 词汇表切换 ----------

    def _refresh_vocab_menu(self, event=None):
        menu = self.vocab_menu
        colors = self.colors
        menu.delete(0, "end")
        current = self.source.resolve()
        local = dict_store.local_index()
        local_paths = {str(item["path"].resolve()) for item in local.values()}
        files = [
            path
            for path in vocab.find_vocabularies(self.source.parent)
            if str(path.resolve()) not in local_paths
        ]
        for path in files:
            mark = " ●" if path.resolve() == current else ""
            menu.add_command(
                label=path.name + mark,
                command=lambda target=path: self.switch_vocabulary(target),
            )
        if files and local:
            menu.add_separator()
        if local:
            item_menu = tk.Menu(
                menu,
                tearoff=0,
                bg=colors["panel"],
                fg=colors["ink"],
                activebackground=colors["accent"],
                activeforeground=colors["bg"],
                font=self.f_ui,
            )
            for item in sorted(local.values(), key=lambda it: str(it["name"]).lower()):
                mark = " ●" if item["path"].resolve() == current else ""
                item_menu.add_command(
                    label=f"{item['name']}{mark}",
                    command=lambda target=item["path"]: self.switch_vocabulary(target),
                )
            menu.add_cascade(label="已下载词库 ▸", menu=item_menu)
        if menu.index("end") is not None:
            menu.add_separator()
        menu.add_command(label="在线词库市场…（Ctrl+M）", command=self.open_market)
        menu.add_command(label="选择其他文件…", command=self.choose_vocabulary)
        menu.add_command(label="导出生词本…", command=self.export_hard_words)
        menu.add_command(
            label="在资源管理器中打开文件夹", command=self._open_vocab_folder
        )
        return None

    def open_market(self) -> None:
        existing = self._market
        if existing is not None and not existing.closed:
            existing.win.deiconify()
            existing.win.lift()
            existing.win.focus_set()
            return
        self._market = dict_market.DictMarket(self, chrome=set_dark_titlebar)

    def _open_vocab_folder(self) -> None:
        try:
            import os

            os.startfile(self.source.parent)
        except OSError:
            pass

    def choose_vocabulary(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择词汇表",
            initialdir=str(self.source.parent),
            filetypes=[
                ("词汇表", "*.json *.xlsx"),
                ("JSON 词汇表", "*.json"),
                ("Excel 词汇表", "*.xlsx"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.switch_vocabulary(Path(path))

    def _enable_drop(self) -> None:
        if not DND_AVAILABLE or not hasattr(self.root, "drop_target_register"):
            return
        targets = [self.root]
        targets.extend(self._descendants(self.root))
        for widget in targets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                continue

    @staticmethod
    def _descendants(widget):
        for child in widget.winfo_children():
            yield child
            yield from WordApp._descendants(child)

    def _setup_ime(self) -> None:
        """屏蔽主窗口的中文输入法，让 H/1/2/R/S/U 等单键直达程序。

        不处理的话，输入法开着时按这些键只会弹出拼音候选窗（实测如此）。
        搜索框聚焦时临时恢复输入法（要用中文查词），失焦后再屏蔽；
        设置窗口和词库市场是独立窗口，输入法不受影响。
        """
        self._ime_set = None
        try:
            user32 = ctypes.windll.user32
            imm = ctypes.windll.imm32
            user32.GetParent.argtypes = (ctypes.c_void_p,)
            user32.GetParent.restype = ctypes.c_void_p
            imm.ImmAssociateContext.argtypes = (ctypes.c_void_p, ctypes.c_void_p)
            imm.ImmAssociateContext.restype = ctypes.c_void_p
            hwnd = user32.GetParent(self.root.winfo_id())
            if not hwnd:
                return
            default_ctx = imm.ImmAssociateContext(hwnd, None)

            def _set(enabled: bool) -> None:
                try:
                    imm.ImmAssociateContext(hwnd, default_ctx if enabled else None)
                except Exception:
                    pass

            self._ime_set = _set
            self.search_entry.bind("<FocusIn>", lambda event: _set(True), add="+")
            self.search_entry.bind("<FocusOut>", lambda event: _set(False), add="+")
        except Exception:
            self._ime_set = None

    def _on_drop(self, event):
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        for raw in paths:
            path = Path(str(raw).strip("{}"))
            if path.suffix.lower() in vocab.SUPPORTED_SUFFIXES and path.is_file():
                self.switch_vocabulary(path)
                break
        return event.action if hasattr(event, "action") else None

    def switch_vocabulary(self, path: Path) -> bool:
        path = Path(path)
        if path.resolve() == self.source.resolve():
            return False
        try:
            words = vocab.load_cached(path)
        except Exception as exc:
            self._flash_status(f"读取失败：{exc}")
            return False
        if not words:
            self._flash_status(f"{path.name} 里没有读到单词")
            return False

        self._save_progress()
        self.source = path
        self.words = words
        self._search_index = [
            (entry["word"].lower(), entry["translation"].lower()) for entry in words
        ]
        self.filtered = list(range(len(words)))
        self.index = 0
        self._revealed = False
        self.hard_only = False
        self.hard_only_var.set(False)
        self._scale_cache.clear()
        self._chunks = []
        self.search_var.set("")
        if self.search_entry.focus_get() is not self.search_entry:
            self._show_search_placeholder()
        self.root.title(f"单词抄写 · {path.stem}")
        self.title_label.configure(text=path.stem)
        self.sidebar_title.configure(text="全部单词")
        self.sidebar_count.configure(text=f"{len(self.words)} 词")
        self.progress_label.configure(text="")
        self.cards.yview_moveto(0)
        self._restore_position()
        self.render()

        self.config["vocab"] = str(path)
        usage_ai.save_config(self.config)
        self._flash_status(f"已切换到 {path.stem}（{len(words)} 词）")
        return True

    def _flash_status(self, message: str) -> None:
        if self._flash_job is not None:
            self.root.after_cancel(self._flash_job)
        self.progress_label.configure(text=message)

        def clear() -> None:
            self._flash_job = None
            self.progress_label.configure(text=self._progress_text())

    def _progress_text(self) -> str:
        return ("背 " if self.hide_mode else "") + f"{self.index + 1} / {len(self.words)}"

        self._flash_job = self.root.after(2500, clear)

    # ---------- 主题 ----------

    def _resolve_theme(self, mode: str) -> str:
        if mode == "system":
            return system_theme()
        return mode if mode in THEMES else "dark"

    def apply_theme(self, name: str) -> None:
        if name not in THEMES:
            name = "dark"
        colors = THEMES[name]
        self.theme_name = name
        self.colors = colors
        bg, panel = colors["bg"], colors["panel"]

        self.root.configure(bg=bg)
        for frame in (self.main_frame, self.header, self.area, self.hint):
            frame.configure(bg=bg)
        for frame in (self.sidebar, self.sidebar_head, self.sidebar_body):
            frame.configure(bg=panel)

        for label in (self.title_label, self.progress_label, self.scale_label, self.hint_label):
            label.configure(bg=bg, fg=colors["muted"])
        self.sidebar_title.configure(bg=panel, fg=colors["ink"])
        self.sidebar_count.configure(bg=panel, fg=colors["muted"])
        self.hard_only_check.configure(
            bg=panel,
            fg=colors["ink"],
            activebackground=panel,
            activeforeground=colors["accent"],
            selectcolor=colors["card"],
        )
        self.vocab_button.configure(
            bg=bg, fg=colors["muted"], activebackground=panel, activeforeground=colors["ink"]
        )
        if hasattr(self, "sections_button"):
            self.sections_button.configure(
                bg=bg, fg=colors["muted"], activebackground=panel, activeforeground=colors["ink"]
            )
        if hasattr(self, "sections_bar"):
            self.sections_bar.configure(bg=bg)
            self.sections_tip.configure(bg=bg, fg=colors["muted"])
            for check in self.section_checks.values():
                check.configure(
                    bg=bg,
                    fg=colors["ink"],
                    activebackground=bg,
                    activeforeground=colors["accent"],
                    selectcolor=panel,
                )
        self.vocab_menu.configure(
            bg=panel,
            fg=colors["ink"],
            activebackground=colors["accent"],
            activeforeground=bg,
        )
        self.settings_button.configure(
            bg=bg, fg=colors["muted"], activebackground=panel, activeforeground=colors["ink"]
        )

        self.text.configure(
            bg=bg,
            fg=colors["body"],
            insertbackground=colors["ink"],
            selectbackground=colors["select"],
            selectforeground=colors["ink"],
        )
        for tag, color in (
            ("word", "ink"),
            ("ipa", "muted"),
            ("meta", "muted"),
            ("meta_warn", "accent"),
            ("label", "accent"),
            ("body", "body"),
            ("usage", "ink"),
            ("usage_meaning", "muted"),
            ("usage_example", "muted"),
        ):
            self.text.tag_configure(tag, foreground=colors[color])
        for button in self._speak_buttons:
            button.configure(
                bg=panel,
                fg=colors["muted"],
                activebackground=colors["accent"],
                activeforeground=bg,
            )

        self.cards.configure(bg=panel)
        self.search_frame.configure(bg=panel)
        placeholder = self.search_var.get() == self.SEARCH_PLACEHOLDER
        self.search_entry.configure(
            bg=colors["card"],
            fg=colors["muted"] if placeholder else colors["ink"],
            insertbackground=colors["ink"],
            highlightbackground=colors["card_border"],
            highlightcolor=colors["accent"],
        )
        self._style_scrollbars()
        if self._market is not None and not self._market.closed:
            self._market.apply_colors()
        set_dark_titlebar(self.root, name == "dark")
        self.root.after(60, lambda: set_dark_titlebar(self.root, self.theme_name == "dark"))
        self._draw_cards()

    def _style_scrollbars(self) -> None:
        colors = self.colors
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        for name, trough in (
            ("Main.Vertical.TScrollbar", colors["bg"]),
            ("Side.Vertical.TScrollbar", colors["panel"]),
        ):
            style.configure(
                name,
                background=colors["card"],
                troughcolor=trough,
                bordercolor=trough,
                arrowcolor=colors["muted"],
                lightcolor=colors["card"],
                darkcolor=colors["card"],
                relief="flat",
            )
            style.map(name, background=[("active", colors["accent"])])

    def _watch_theme(self) -> None:
        self._watch_job = None
        if not self._running:
            return
        if self.theme_mode == "system":
            detected = system_theme()
            if detected != self.theme_name:
                self.apply_theme(detected)
        self._watch_job = self.root.after(2000, self._watch_theme)

    # ---------- 设置 ----------

    def open_settings(self) -> None:
        existing = self._settings_window
        if existing is not None and existing.winfo_exists():
            existing.deiconify()
            existing.lift()
            existing.focus_set()
            return

        colors = self.colors
        window = tk.Toplevel(self.root)
        self._settings_window = window
        window.title("设置")
        window.configure(bg=colors["bg"])
        window.transient(self.root)
        window.resizable(False, False)
        try:
            window.iconbitmap(default=str(APP_DIR / "icon.ico"))
        except tk.TclError:
            pass
        window.geometry(f"+{self.root.winfo_rootx() + 180}+{self.root.winfo_rooty() + 80}")
        set_dark_titlebar(window, self.theme_name == "dark")

        def add_label(text: str, row: int) -> None:
            tk.Label(
                window,
                text=text,
                bg=colors["bg"],
                fg=colors["muted"],
                font=self.f_ui,
                anchor="w",
            ).grid(row=row, column=0, sticky="w", padx=(24, 12), pady=(0, 10))

        add_label("主题", 0)
        mode_var = tk.StringVar(value=THEME_LABELS.get(self.theme_mode, "深色"))
        ttk.Combobox(
            window,
            textvariable=mode_var,
            values=list(THEME_LABELS.values()),
            state="readonly",
            width=16,
            font=self.f_ui,
        ).grid(row=0, column=1, sticky="w", pady=(0, 10))

        tk.Label(
            window,
            text="AI 平台（每个平台一张卡片：一把 Key 对应一个模型）",
            bg=colors["bg"],
            fg=colors["ink"],
            font=self.f_ui,
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=24, pady=(0, 6))

        list_wrap = tk.Frame(window, bg=colors["bg"])
        list_wrap.grid(row=2, column=0, columnspan=2, sticky="we", padx=(24, 18))
        card_canvas = tk.Canvas(
            list_wrap,
            bg=colors["bg"],
            bd=0,
            highlightthickness=0,
            height=248,
            width=520,
        )
        card_scroll = ttk.Scrollbar(
            list_wrap, orient="vertical", command=card_canvas.yview, style="Main.Vertical.TScrollbar"
        )
        card_canvas.configure(yscrollcommand=card_scroll.set)
        card_scroll.pack(side="right", fill="y")
        card_canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(card_canvas, bg=colors["bg"])
        card_canvas.create_window((0, 0), window=inner, anchor="nw", tags="inner")
        inner.bind(
            "<Configure>",
            lambda event: card_canvas.configure(scrollregion=card_canvas.bbox("all")),
        )
        card_canvas.bind(
            "<Configure>",
            lambda event: card_canvas.itemconfigure("inner", width=event.width),
        )
        for widget in (card_canvas, inner):
            widget.bind("<MouseWheel>", lambda event: card_canvas.yview_scroll(
                -1 if event.delta > 0 else 1, "units"
            ))

        def make_entry(parent, textvariable) -> tk.Entry:
            return tk.Entry(
                parent,
                textvariable=textvariable,
                bg=colors["bg"],
                fg=colors["ink"],
                insertbackground=colors["ink"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=colors["card_border"],
                highlightcolor=colors["accent"],
                font=self.f_ui,
                width=40,
            )

        cards_state: list[dict] = []

        def collect() -> list[dict]:
            result = []
            for card in cards_state:
                tasks = [task for task, var in card["tasks"].items() if var.get()]
                item = {
                    "base_url": card["base"].get().strip().rstrip("/"),
                    "api_key": card["key"].get().strip(),
                    "model": card["model"].get().strip(),
                    "tasks": tasks,
                }
                if item["base_url"] or item["api_key"] or item["model"]:
                    result.append(item)
            return result

        def rebuild(data: list[dict] | None = None) -> None:
            if data is None:
                data = collect()
            if not data:
                data = [{"base_url": "", "api_key": "", "model": "", "tasks": []}]
            for child in inner.winfo_children():
                child.destroy()
            cards_state.clear()
            for index, item in enumerate(data):
                cards_state.append(make_card(item, index))
            card_canvas.update_idletasks()
            card_canvas.configure(scrollregion=card_canvas.bbox("all"))

        def make_card(item: dict, index: int) -> dict:
            card = tk.Frame(
                inner,
                bg=colors["card"],
                highlightthickness=1,
                highlightbackground=colors["card_border"],
            )
            card.pack(fill="x", pady=(0, 10), padx=(0, 6))

            base_var = tk.StringVar(value=str(item.get("base_url") or ""))
            key_var = tk.StringVar(value=str(item.get("api_key") or ""))
            model_var = tk.StringVar(value=str(item.get("model") or ""))

            tk.Label(
                card, text="Base URL", bg=colors["card"], fg=colors["muted"], font=self.f_ui
            ).grid(row=0, column=0, sticky="w", padx=(12, 8), pady=(10, 4))
            base_entry = make_entry(card, base_var)
            base_entry.grid(row=0, column=1, sticky="we", padx=(0, 8), pady=(10, 4))

            tk.Label(
                card, text="Key", bg=colors["card"], fg=colors["muted"], font=self.f_ui
            ).grid(row=1, column=0, sticky="w", padx=(12, 8), pady=(0, 4))
            key_entry = make_entry(card, key_var)
            key_entry.grid(row=1, column=1, sticky="we", padx=(0, 8), pady=(0, 4))

            tk.Label(
                card, text="模型", bg=colors["card"], fg=colors["muted"], font=self.f_ui
            ).grid(row=2, column=0, sticky="w", padx=(12, 8), pady=(0, 10))
            model_box = ttk.Combobox(
                card,
                textvariable=model_var,
                values=[model_var.get()] if model_var.get() else [],
                width=30,
                font=self.f_ui,
            )
            model_box.grid(row=2, column=1, sticky="we", padx=(0, 8), pady=(0, 6))

            task_vars: dict[str, tk.BooleanVar] = {}
            saved_tasks = item.get("tasks")
            saved_tasks = (
                set(saved_tasks) if isinstance(saved_tasks, list) and saved_tasks else set(usage_ai.TASKS)
            )
            task_row = tk.Frame(card, bg=colors["card"])
            task_row.grid(row=3, column=0, columnspan=3, sticky="w", padx=(12, 8), pady=(0, 6))
            tk.Label(
                task_row, text="用于", bg=colors["card"], fg=colors["muted"], font=self.f_ui
            ).pack(side="left", padx=(0, 6))
            for task in usage_ai.TASKS:
                variable = tk.BooleanVar(value=task in saved_tasks)
                tk.Checkbutton(
                    task_row,
                    text=usage_ai.TASK_LABELS[task],
                    variable=variable,
                    bg=colors["card"],
                    fg=colors["ink"],
                    activebackground=colors["card"],
                    activeforeground=colors["accent"],
                    selectcolor=colors["panel"],
                    font=self.f_ui,
                    bd=0,
                    highlightthickness=0,
                    cursor="hand2",
                ).pack(side="left", padx=(0, 10))
                task_vars[task] = variable

            actions = tk.Frame(card, bg=colors["card"])
            actions.grid(row=0, column=2, rowspan=3, sticky="ne", padx=(0, 10), pady=(8, 8))
            fetch_button = tk.Button(
                actions,
                text="获取模型",
                bd=0,
                relief="flat",
                cursor="hand2",
                font=self.f_ui,
                padx=10,
                pady=2,
                bg=colors["panel"],
                fg=colors["ink"],
                activebackground=colors["accent"],
                activeforeground=colors["bg"],
            )
            fetch_button.pack(fill="x")
            remove_button = tk.Button(
                actions,
                text="删除",
                bd=0,
                relief="flat",
                cursor="hand2",
                font=self.f_ui,
                padx=10,
                pady=2,
                bg=colors["panel"],
                fg=colors["muted"],
                activebackground=colors["accent"],
                activeforeground=colors["bg"],
            )
            remove_button.pack(fill="x", pady=(6, 0))
            status = tk.Label(
                card, text="", bg=colors["card"], fg=colors["muted"], font=self.f_ui, anchor="w"
            )
            status.grid(row=4, column=0, columnspan=3, sticky="w", padx=(12, 8), pady=(0, 8))
            card.columnconfigure(1, weight=1)

            card_state = {
                "base": base_entry,
                "key": key_entry,
                "model": model_box,
                "status": status,
                "fetch": fetch_button,
                "tasks": task_vars,
                # 必须持有引用，否则 StringVar 被回收后输入框会变空
                "vars": (base_var, key_var, model_var, *task_vars.values()),
            }

            remove_button.configure(
                command=lambda idx=index: rebuild(
                    [entry for pos, entry in enumerate(collect()) if pos != idx]
                )
            )

            def fetch(target=card_state) -> None:
                if str(target["fetch"].cget("state")) == "disabled":
                    return
                base = target["base"].get().strip()
                key = target["key"].get().strip()
                if not base or not key:
                    target["status"].configure(text="请先填写 Base URL 和 Key", fg=colors["accent"])
                    return
                target["fetch"].configure(state="disabled")
                target["status"].configure(text="正在获取模型…", fg=colors["muted"])
                results: queue.Queue = queue.Queue()

                def work() -> None:
                    try:
                        results.put(("ok", usage_ai.fetch_models(base, key)))
                    except Exception as exc:  # 网络或鉴权错误
                        results.put(("error", f"{type(exc).__name__}: {exc}"))

                threading.Thread(target=work, daemon=True).start()

                def poll() -> None:
                    try:
                        kind, data = results.get_nowait()
                    except queue.Empty:
                        window.after(150, poll)
                        return
                    target["fetch"].configure(state="normal")
                    if kind == "ok":
                        values = data or []
                        target["model"].configure(values=values)
                        current = target["model"].get().strip()
                        if values and current not in values:
                            target["model"].set(values[0])
                        target["status"].configure(
                            text=f"获取到 {len(values)} 个模型", fg=colors["muted"]
                        )
                    else:
                        target["status"].configure(text=str(data)[:110], fg=colors["accent"])

                window.after(150, poll)

            fetch_button.configure(command=fetch)
            return card_state

        rebuild([dict(item) for item in self.config.get("endpoints") or []])

        add_button = tk.Button(
            window,
            text="+ 添加平台",
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=12,
            pady=3,
            bg=colors["panel"],
            fg=colors["ink"],
            activebackground=colors["accent"],
            activeforeground=colors["bg"],
        )
        add_button.configure(command=lambda: rebuild(collect() + [{"base_url": "", "api_key": "", "model": ""}]))
        add_button.grid(row=3, column=1, sticky="w", pady=(8, 6))

        parallel_var = tk.BooleanVar(value=bool(self.config.get("parallel", True)))
        tk.Checkbutton(
            window,
            text="同时使用所有平台并行生成（更快）",
            variable=parallel_var,
            bg=colors["bg"],
            fg=colors["ink"],
            activebackground=colors["bg"],
            activeforeground=colors["ink"],
            selectcolor=colors["panel"],
            font=self.f_ui,
            anchor="w",
            bd=0,
            highlightthickness=0,
        ).grid(row=4, column=1, sticky="w", pady=(0, 4))

        trim_var = tk.BooleanVar(value=bool(self.config.get("trim_senses", False)))
        tk.Checkbutton(
            window,
            text="用 AI 精简翻译（去掉生僻、不常考的义项）",
            variable=trim_var,
            bg=colors["bg"],
            fg=colors["ink"],
            activebackground=colors["bg"],
            activeforeground=colors["ink"],
            selectcolor=colors["panel"],
            font=self.f_ui,
            anchor="w",
            bd=0,
            highlightthickness=0,
        ).grid(row=5, column=1, sticky="w", pady=(0, 4))

        family_var = tk.BooleanVar(value=bool(self.config.get("trim_families", False)))
        tk.Checkbutton(
            window,
            text="用 AI 精简同根词（只留常考常用，丢弃生僻派生）",
            variable=family_var,
            bg=colors["bg"],
            fg=colors["ink"],
            activebackground=colors["bg"],
            activeforeground=colors["ink"],
            selectcolor=colors["panel"],
            font=self.f_ui,
            anchor="w",
            bd=0,
            highlightthickness=0,
        ).grid(row=6, column=1, sticky="w", pady=(0, 4))

        tk.Label(
            window,
            text="卡片上可勾选该平台用于哪些任务；同一任务勾了几个平台就并行几个。",
            bg=colors["bg"],
            fg=colors["muted"],
            font=self.f_ui,
            anchor="w",
            justify="left",
        ).grid(row=7, column=1, sticky="w", pady=(0, 8))

        voice_var = tk.StringVar()
        rate_var = tk.IntVar(value=int(self.config.get("speak_rate") or 170))
        auto_var = tk.BooleanVar(value=bool(self.config.get("speak_auto")))
        voice_map: dict[str, str] = {}

        add_label("朗读", 8)
        if not self.speaker.available:
            tk.Label(
                window,
                text="未启用语音：未安装 pyttsx3（执行 pip install pyttsx3 后重启程序）",
                bg=colors["bg"],
                fg=colors["muted"],
                font=self.f_ui,
                anchor="w",
            ).grid(row=8, column=1, sticky="w", pady=(0, 8))
        else:
            speak_frame = tk.Frame(window, bg=colors["bg"])
            speak_frame.grid(row=8, column=1, sticky="w", pady=(0, 6))
            voice_box = ttk.Combobox(
                speak_frame,
                textvariable=voice_var,
                state="disabled",
                width=32,
                font=self.f_ui,
            )
            voice_box.pack(side="left")

            def fill_voices() -> None:
                if not window.winfo_exists() or not self.speaker.voices:
                    return
                labels = [self.speaker.voice_label(voice) for voice in self.speaker.voices]
                for label, voice in zip(labels, self.speaker.voices):
                    voice_map[label] = voice["id"]
                voice_box.configure(values=labels, state="readonly")
                current = next(
                    (
                        label
                        for label, voice_id in voice_map.items()
                        if voice_id == self.speaker.voice_id
                    ),
                    labels[0],
                )
                voice_var.set(current)

            fill_voices()
            if not voice_map:  # 引擎还在后台初始化
                window.after(500, fill_voices)

            tk.Button(
                speak_frame,
                text="试听",
                command=lambda: (
                    self.speaker.configure(
                        voice_id=voice_map.get(voice_var.get()) or self.speaker.voice_id,
                        rate=int(rate_var.get()),
                    ),
                    self.speaker.speak(self.words[self.index]["word"]),
                ),
                bd=0,
                relief="flat",
                cursor="hand2",
                font=self.f_ui,
                padx=10,
                pady=2,
                bg=colors["panel"],
                fg=colors["ink"],
                activebackground=colors["accent"],
                activeforeground=colors["bg"],
            ).pack(side="left", padx=(8, 0))

            tk.Label(
                window,
                text="语速",
                bg=colors["bg"],
                fg=colors["muted"],
                font=self.f_ui,
                anchor="w",
            ).grid(row=9, column=0, sticky="w", padx=(24, 12), pady=(0, 6))
            tk.Scale(
                window,
                from_=80,
                to=300,
                resolution=5,
                orient="horizontal",
                variable=rate_var,
                length=250,
                showvalue=True,
                bg=colors["bg"],
                fg=colors["muted"],
                troughcolor=colors["panel"],
                highlightthickness=0,
                bd=0,
                font=self.f_ui,
                activebackground=colors["accent"],
            ).grid(row=9, column=1, sticky="w", pady=(0, 2))

            tk.Checkbutton(
                window,
                text="换词后自动朗读单词",
                variable=auto_var,
                bg=colors["bg"],
                fg=colors["ink"],
                activebackground=colors["bg"],
                activeforeground=colors["ink"],
                selectcolor=colors["panel"],
                font=self.f_ui,
                anchor="w",
                bd=0,
                highlightthickness=0,
            ).grid(row=10, column=1, sticky="w", pady=(0, 8))

        status = tk.Label(
            window, text="", bg=colors["bg"], fg=colors["muted"], font=self.f_ui, anchor="w"
        )
        status.grid(row=11, column=0, columnspan=2, sticky="we", padx=24, pady=(0, 8))

        buttons = tk.Frame(window, bg=colors["bg"])
        buttons.grid(row=12, column=0, columnspan=2, sticky="e", padx=24, pady=(0, 18))

        def close() -> None:
            self._settings_window = None
            window.destroy()

        def save() -> None:
            cards = collect()
            for card in cards:
                if not (card["base_url"] and card["api_key"] and card["model"]):
                    status.configure(
                        text="每个平台都要填 Base URL、Key 和模型", fg=colors["accent"]
                    )
                    return
                if not card["tasks"]:
                    status.configure(
                        text=f"{card['model']} 至少勾选一个用途", fg=colors["accent"]
                    )
                    return
            mode = next(
                (key for key, label in THEME_LABELS.items() if label == mode_var.get()), "dark"
            )
            config = dict(self.config)
            config["endpoints"] = cards
            if cards:
                config["base_url"] = cards[0]["base_url"]
                config["api_key"] = cards[0]["api_key"]
                config["model"] = cards[0]["model"]
                config["api_keys"] = [card["api_key"] for card in cards]
                config["models"] = [card["model"] for card in cards]
            else:
                config["base_url"] = ""
                config["api_key"] = ""
                config["model"] = ""
                config["api_keys"] = []
                config["models"] = []
            config["parallel"] = bool(parallel_var.get())
            config["trim_senses"] = bool(trim_var.get())
            config["trim_families"] = bool(family_var.get())
            config["theme"] = mode
            config["speak_voice"] = voice_map.get(voice_var.get(), self.speaker.voice_id)
            config["speak_rate"] = int(rate_var.get())
            config["speak_auto"] = bool(auto_var.get())
            if not usage_ai.save_config(config):
                status.configure(text="保存失败：配置文件无法写入", fg=colors["accent"])
                return
            self.config = config
            self.theme_mode = mode
            self.speaker.configure(
                voice_id=config["speak_voice"], rate=config["speak_rate"]
            )
            self.apply_theme(self._resolve_theme(mode))
            self._update_service(config)
            close()
            self.render()

        tk.Button(
            buttons,
            text="取消",
            command=close,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=14,
            pady=4,
            bg=colors["panel"],
            fg=colors["muted"],
            activebackground=colors["card_border"],
            activeforeground=colors["ink"],
        ).pack(side="right")
        tk.Button(
            buttons,
            text="保存",
            command=save,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.f_ui,
            padx=18,
            pady=4,
            bg=colors["accent"],
            fg=colors["bg"],
            activebackground=colors["ink"],
            activeforeground=colors["bg"],
        ).pack(side="right", padx=(0, 10))

        window.bind("<Escape>", lambda event: close())
        window.protocol("WM_DELETE_WINDOW", close)

    def _update_service(self, config: dict) -> None:
        buckets: dict[str, dict[str, object]] = {
            "usage_service": {},
            "sense_service": {},
            "family_service": {},
        }
        for attr, bucket in buckets.items():
            service = getattr(self, attr)
            if service is not None:
                service.stop()
                service.save()
                bucket.update(service.cache)
            setattr(self, attr, None)

        if config.get("endpoints") or str(config.get("api_key", "")).strip():
            classes = [("usage_service", usage_ai.UsageService, None)]
            if config.get("trim_senses"):
                classes.append(("sense_service", usage_ai.SenseService, None))
            if config.get("trim_families"):
                classes.append(("family_service", usage_ai.FamilyService, None))
            for attr, cls, _ in classes:
                service = cls(config)
                if not service.endpoints:
                    continue
                service.cache = dict(buckets[attr])
                service.start()
                setattr(self, attr, service)
        self._scale_cache.clear()

    def _bind_keys(self) -> None:
        forward = ("<space>", "<Return>", "<Right>", "<Down>", "<Next>")
        backward = ("<Left>", "<Up>", "<Prior>", "<BackSpace>")
        for sequence in forward:
            self.root.bind(sequence, self.next_word)
        for sequence in backward:
            self.root.bind(sequence, self.prev_word)
        self.root.bind("<Home>", self.first_word)
        self.root.bind("<End>", self.last_word)
        self.root.bind("<Tab>", self.toggle_extra)
        self.text.bind("<Tab>", self.toggle_extra)
        self.root.bind("<Shift-Tab>", self.toggle_extra)
        self.root.bind("<u>", self.regenerate_usage)
        self.root.bind("<U>", self.regenerate_usage)
        self.root.bind("<Control-u>", self.regenerate_senses)
        self.root.bind("<Control-Shift-U>", self.regenerate_family)
        self.root.bind("<Control-e>", self.toggle_sections_bar)
        for sequence in ("<plus>", "<equal>", "<KP_Add>"):
            self.root.bind(sequence, lambda event: self.change_scale(0.1))
        for sequence in ("<minus>", "<underscore>", "<KP_Subtract>"):
            self.root.bind(sequence, lambda event: self.change_scale(-0.1))
        self.root.bind("<F11>", self.toggle_fullscreen)
        self.root.bind("<Control-comma>", lambda event: (self.open_settings(), "break")[1])
        self.root.bind("<Control-f>", self.focus_search)
        self.root.bind("<Control-F>", self.focus_search)
        self.root.bind("<Control-m>", lambda event: (self.open_market(), "break")[1])
        self.root.bind("<Control-M>", lambda event: (self.open_market(), "break")[1])
        self.root.bind("<s>", self._speak_word)
        self.root.bind("<S>", self._speak_word)
        self.root.bind("<h>", self.toggle_hide_mode)
        self.root.bind("<H>", self.toggle_hide_mode)
        self.root.bind("<r>", self.random_word)
        self.root.bind("<R>", self.random_word)
        self.root.bind("<Key-1>", lambda event: self._toggle_mark(MARK_KNOWN))
        self.root.bind("<Key-2>", lambda event: self._toggle_mark(MARK_HARD))
        self.root.bind("<Escape>", lambda event: self.on_close())
        self.root.bind_all("<MouseWheel>", self._on_wheel)
        self.root.bind_all("<Control-MouseWheel>", self._on_ctrl_wheel)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 自动字号 ----------

    def _on_text_configure(self, event=None) -> None:
        if self._fit_job is not None:
            self.root.after_cancel(self._fit_job)
        self._fit_job = self.root.after(120, self._run_refit)

    def _run_refit(self) -> None:
        self._fit_job = None
        self.refit()

    def refit(self) -> None:
        scale = self._compute_scale()
        if scale is None:
            return
        self._apply_fonts(scale)
        self.scale = scale
        self.text.update_idletasks()
        self._update_scale_label()

    def _compute_scale(self) -> float | None:
        if not self.text.winfo_ismapped() or self.text.winfo_width() < 60:
            if self._fit_job is None:
                self._fit_job = self.root.after(60, self._run_refit)
            return None
        self._fit_job = None
        key = (
            self.text.winfo_width(),
            self.text.winfo_height(),
            tuple(sorted(self.sections.items())),
            self.hide_mode,
            self._revealed,
        )
        bucket = self._scale_cache.setdefault(key, {})
        if len(self._scale_cache) > 6:
            self._scale_cache = {key: bucket}
        fitted = bucket.get(self.index)
        if fitted is None:
            self._prepare_measure(self.text.winfo_width(), self.text.winfo_height())
            fitted = self._fit_scale(1.0)
            bucket[self.index] = fitted
        return min(self.MAX_SCALE, max(self.MIN_SCALE, fitted * self.zoom))

    def _update_scale_label(self) -> None:
        label = f"字号 {self.f_body.actual('size')}pt"
        if abs(self.zoom - 1.0) > 0.005:
            label += f"（×{self.zoom:.2f}）"
        self.scale_label.configure(text=label)

    def _create_measure(self) -> None:
        self.measure = tk.Text(
            self.root,
            wrap="word",
            bd=0,
            highlightthickness=0,
            padx=8,
            pady=10,
            cursor="arrow",
            takefocus=0,
            width=10,
            height=5,
        )
        for tag, key in (
            ("word", "word"),
            ("ipa", "ipa"),
            ("meta", "meta"),
            ("meta_warn", "meta"),
            ("label", "label"),
            ("body", "body"),
            ("usage", "body"),
            ("usage_meaning", "body"),
            ("usage_example", "meta"),
        ):
            self.measure.tag_configure(tag, font=self.m_fonts[key])
        self.measure.place(x=-30000, y=-30000, width=1, height=1)

    def _prepare_measure(self, width: int, height: int) -> None:
        if self.measure is None:
            self._create_measure()
        # 测量框设得足够高，保证内容总能完整显示，便于直接量出像素高度
        self.measure.place_configure(x=-30000, y=-30000, width=width, height=4000)
        self.measure.configure(state="normal")
        self.measure.delete("1.0", "end")
        for text, tag in self._chunks:
            self.measure.insert("end", text, tag)
        self.measure.configure(state="disabled")

    def _content_height(self, scale: float) -> int:
        """在离屏文本框里试算字号，返回内容实际占用高度（像素）。"""
        assert self.measure is not None
        measure = self.measure
        self._apply_fonts(scale, self.m_fonts, measure)
        measure.update_idletasks()
        last_line = str(measure.index("end-1c")).split(".")[0]
        info = measure.dlineinfo(f"{last_line}.0")
        if info is None:
            return 10**6
        return info[1] + info[3] + 10

    def _fit_scale(self, base: float) -> float:
        available = max(1, self.text.winfo_height() - 1)
        high_height = self._content_height(base)
        if high_height <= available:
            return base

        floor = round(max(self.MIN_SCALE, base * self.SHRINK_FLOOR), 2)
        low_height = self._content_height(floor)
        if low_height > available:
            return floor

        low_scale, high_scale = floor, base
        best = floor
        for _ in range(2):
            if high_height <= low_height:
                break
            target = low_scale + (available * 0.995 - low_height) * (
                high_scale - low_scale
            ) / (high_height - low_height)
            target = round(min(high_scale, max(low_scale, target)), 2)
            if target <= low_scale or target >= high_scale:
                break
            target_height = self._content_height(target)
            if target_height <= available:
                return target
            high_scale, high_height = target, target_height

        # 字号是整数台阶，插值落点可能正好踩到高一档字号，再二分细化
        low, high = low_scale, high_scale
        for _ in range(3):
            mid = round((low + high) / 2, 2)
            if mid <= low or mid >= high:
                break
            if self._content_height(mid) <= available:
                low = mid
            else:
                high = mid
        return max(best, low)

    def _apply_fonts(
        self,
        scale: float,
        fonts: dict[str, tkfont.Font] | None = None,
        widget: tk.Misc | None = None,
    ) -> None:
        fonts = self.fonts if fonts is None else fonts
        target = self.text if widget is None else widget
        fonts["word"].configure(size=max(16, round(self.WORD_SIZE * scale)))
        fonts["ipa"].configure(size=max(8, round(self.IPA_SIZE * scale)))
        fonts["meta"].configure(size=max(7, round(self.META_SIZE * scale)))
        fonts["label"].configure(size=max(7, round(self.LABEL_SIZE * scale)))
        fonts["body"].configure(size=max(7, round(self.BODY_SIZE * scale)))
        target.tag_configure("word", spacing3=round(4 * scale))
        target.tag_configure("ipa", spacing3=round(2 * scale))
        target.tag_configure("meta", spacing3=round(10 * scale))
        target.tag_configure("meta_warn", spacing3=round(10 * scale))
        target.tag_configure("label", spacing1=round(12 * scale), spacing3=round(2 * scale))
        target.tag_configure("body", spacing3=round(4 * scale))
        target.tag_configure("usage", spacing1=round(2 * scale), spacing3=round(1 * scale))
        target.tag_configure("usage_meaning", spacing1=round(2 * scale), spacing3=round(1 * scale))
        target.tag_configure("usage_example", spacing3=round(4 * scale))

    # ---------- 渲染 ----------

    def _phonetic(self, entry: dict[str, str]) -> str:
        uk, us = entry["ipa"], entry["ipa_us"]
        if uk and us and uk != us:
            return f"英 /{uk}/    美 /{us}/"
        if uk or us:
            return f"/{uk or us}/"
        return ""

    def _meta_parts(self, entry: dict[str, str]) -> list[tuple[str, str]]:
        parts: list[tuple[str, str]] = []
        trimmed = self._trimmed_translation(entry)
        source = trimmed or entry["translation"]
        tags = _pos_tags(source)
        if tags:
            parts.append((" ".join(tags), "meta"))
        senses = _sense_count(source)
        if senses:
            parts.append((f"{senses} 个义项", "meta"))
        root = _root_hint(entry)
        if root:
            parts.append((f"词根 {root}", "meta"))
        if trimmed:
            parts.append(("AI 精简", "meta"))
        elif senses >= 5:
            parts.append(("多义词", "meta_warn"))
        return parts

    def render(self) -> None:
        entry = self.words[self.index]
        self._record_position()
        self._queue_usage()
        self._queue_senses()
        self._queue_families()
        chunks = self._build_chunks(entry)
        self._chunks = chunks

        # 先把字号算好并应用，再写入正文，避免正文用旧字号闪一下
        scale = self._compute_scale()
        if scale is not None:
            self._apply_fonts(scale)
            self.scale = scale

        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        for text, tag in chunks:
            self.text.insert("end", text, tag)
        self._add_speak_buttons(chunks)
        self.text.configure(state="disabled")

        self._update_scale_label()
        self.progress_label.configure(text=self._progress_text())
        self.text.yview_moveto(0)
        self.text.update_idletasks()
        self._draw_cards()
        self._ensure_card_visible()

    def _add_speak_buttons(self, chunks: list[tuple[str, str]]) -> None:
        self._speak_buttons = []
        if not self.speaker.available:
            return
        entry = self.words[self.index]
        spans: list[tuple[int, int, str]] = []
        offset = 0
        for text, tag in chunks:
            spans.append((offset, offset + len(text), tag))
            offset += len(text)

        # (字符偏移, 按钮文字, 朗读内容)；从后往前插，偏移量才不会被挤动
        placements: list[tuple[int, str, str]] = [(len(entry["word"]), "▶ 朗读", entry["word"])]
        for position, (start, end, tag) in enumerate(spans):
            if tag == "label" and "例句" in chunks[position][0]:
                if position + 1 >= len(spans) or spans[position + 1][2] != "body":
                    continue
                cursor, source = spans[position + 1][0], chunks[position + 1][0]
            elif tag == "usage_example" and chunks[position][0].startswith("例："):
                cursor, source = start, chunks[position][0]
            else:
                continue
            for line in source.split("\n"):
                spoken = speech.spoken_english(line.removeprefix("例："))
                if spoken and re.search(r"[A-Za-z]{2}", spoken):
                    placements.append((cursor, "▶", spoken))
                cursor += len(line) + 1

        for char_offset, label, spoken in reversed(placements):
            button = self._make_speak_button(label, spoken)
            self.text.window_create(f"1.0 + {char_offset} chars", window=button, padx=4)
        self._speak_buttons.reverse()

    def _make_speak_button(self, label: str, spoken: str) -> tk.Button:
        colors = self.colors
        compact = label == "▶"

        def play() -> None:
            self.speaker.speak(spoken)
            # 别让按钮抢走键盘焦点，否则空格会再次触发按钮而不是翻词
            self.root.focus_set()

        button = tk.Button(
            self.text,
            text=label,
            command=play,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=("Segoe UI", 8) if compact else self.f_ui,
            padx=3 if compact else 7,
            pady=0,
            takefocus=0,
            bg=colors["panel"],
            fg=colors["muted"],
            activebackground=colors["accent"],
            activeforeground=colors["bg"],
            highlightthickness=0,
        )
        self._speak_buttons.append(button)
        return button

    def _speak_current(self, auto: bool = False) -> None:
        if auto and not self.config.get("speak_auto"):
            return
        if not self.speaker.available:
            return
        self.speaker.speak(self.words[self.index]["word"])

    def _speak_word(self, event=None):
        if self._typing():
            return None
        self._speak_current()
        return "break"

    def _build_chunks(self, entry: dict[str, str]) -> list[tuple[str, str]]:
        chunks: list[tuple[str, str]] = [(entry["word"], "word")]

        phonetic = self._phonetic(entry)
        if phonetic:
            chunks.append(("\n" + phonetic, "ipa"))

        meta = self._meta_parts(entry)
        if meta:
            chunks.append(("\n", "meta"))
            for position, (text, tag) in enumerate(meta):
                if position:
                    chunks.append(("  ·  ", "meta"))
                chunks.append((text, tag))

        if self.hide_mode and not self._revealed:
            # 背词模式：只留单词、音标和提示行，按回车再看全部内容
            chunks.append(("\n…（按 回车 显示全部）", "meta_warn"))
            return chunks

        for key, label in SECTION_ORDER:
            if not self.sections.get(key, True):
                continue
            if key == "usage":
                chunks.extend(self._usage_chunks(entry))
                continue
            content = entry[key]
            label_text = label
            if key == "translation":
                trimmed = self._trimmed_translation(entry)
                if trimmed:
                    content = trimmed
                    label_text = "翻译（AI 精简）"
            elif key == "family":
                trimmed_family = self._trimmed_family(entry)
                if trimmed_family:
                    content = trimmed_family
                    label_text = "同根词（AI 精简）"
            if not content:
                continue
            chunks.append(("\n" + label_text + "\n", "label"))
            chunks.append((content + "\n", "body"))
        return chunks

    def _trimmed_family(self, entry: dict[str, str]) -> str:
        service = self.family_service
        if service is None:
            return ""
        return service.get(entry["word"], entry["family"]) or ""

    def _trimmed_translation(self, entry: dict[str, str]) -> str:
        service = self.sense_service
        if service is None:
            return ""
        senses = service.get(entry["word"], entry["translation"])
        if not senses:
            return ""
        lines = []
        for sense in senses:
            pos = str(sense.get("pos", "")).strip()
            cn = str(sense.get("cn", "")).strip()
            if cn:
                lines.append(f"{pos} {cn}".strip())
        return "\n".join(lines)

    def _usage_chunks(self, entry: dict[str, str]) -> list[tuple[str, str]]:
        service = self.usage_service
        if service is None:
            return []
        word = entry["word"]
        usage = service.get(word)
        if usage:
            chunks: list[tuple[str, str]] = [("\n用法\n", "label")]
            for item in usage:
                chunks.append((item["pattern"], "usage"))
                if item.get("meaning"):
                    chunks.append(("  " + item["meaning"], "usage_meaning"))
                chunks.append(("\n", "usage"))
                if item.get("example"):
                    chunks.append(("例：" + item["example"] + "\n", "usage_example"))
            return chunks
        if service.is_pending(word):
            return [("\n用法\n", "label"), ("正在生成…\n", "usage_example")]
        if service.recently_failed(word):
            return [("\n用法\n", "label"), ("暂时生成失败，按 U 重试\n", "usage_example")]
        return []

    def _queue_usage(self) -> None:
        service = self.usage_service
        if service is None:
            return
        entry = self.words[self.index]
        workers = service.parallelism()
        if not service.get(entry["word"]) and not service.is_pending(entry["word"]):
            # 当前词单独成批，让每个模型都能立刻接到一个短任务，减少等待
            window = self.words[self.index : self.index + max(2, workers)]
            service.enqueue(
                [(item["word"], _meaning_hint(item)) for item in window],
                urgent=True,
                batch_limit=1,
            )
            return
        if service.busy():
            return
        count = max(1, int(service.config.get("prefetch", 6)))
        count = max(count, workers * int(service.config.get("batch_size", 6)))
        upcoming = self.words[self.index + 1 : self.index + 1 + count]
        service.enqueue([(item["word"], _meaning_hint(item)) for item in upcoming])

    def _queue_senses(self) -> None:
        service = self.sense_service
        if service is None:
            return
        entry = self.words[self.index]
        workers = service.parallelism()
        if not service.get(entry["word"], entry["translation"]) and not service.is_pending(
            entry["word"]
        ):
            window = self.words[self.index : self.index + max(2, workers)]
            service.enqueue(
                [(item["word"], item["translation"]) for item in window],
                urgent=True,
                batch_limit=1,
            )
            return
        if service.busy():
            return
        count = max(1, int(service.config.get("prefetch", 6)))
        count = max(count, workers * int(service.config.get("batch_size", 6)))
        upcoming = self.words[self.index + 1 : self.index + 1 + count]
        service.enqueue([(item["word"], item["translation"]) for item in upcoming])

    def _queue_families(self) -> None:
        service = self.family_service
        if service is None or not self.sections.get("family"):
            return
        entry = self.words[self.index]
        if not entry["family"]:
            return
        workers = service.parallelism()
        if not service.get(entry["word"], entry["family"]) and not service.is_pending(
            entry["word"]
        ):
            window = [
                item for item in self.words[self.index : self.index + max(2, workers)]
                if item["family"]
            ]
            service.enqueue(
                [(item["word"], item["family"]) for item in window],
                urgent=True,
                batch_limit=1,
            )
            return
        if service.busy():
            return
        count = max(1, int(service.config.get("prefetch", 6)))
        count = max(count, workers * int(service.config.get("batch_size", 6)))
        upcoming = [item for item in self.words[self.index + 1 : self.index + 1 + count] if item["family"]]
        service.enqueue([(item["word"], item["family"]) for item in upcoming])

    def _poll_service(self, service) -> tuple[set[str], set[str]]:
        changed: set[str] = set()
        failed: set[str] = set()
        if service is None:
            return changed, failed
        while True:
            try:
                kind, payload, batch = service.results.get_nowait()
            except queue.Empty:
                break
            if kind == "ok":
                service.store(payload, batch)
                changed.update(str(item.get("word", "")).lower() for item in payload)
            else:
                # 失败消息是 (kind, batch, 错误文本)，词列表在 payload 里
                failed.update(word.lower() for word, _ in payload)
        return changed, failed

    def _poll_usage(self) -> None:
        self._poll_job = None
        changed, failed = self._poll_service(self.usage_service)
        sense_changed, sense_failed = self._poll_service(self.sense_service)
        family_changed, family_failed = self._poll_service(self.family_service)

        if changed or failed or sense_changed or sense_failed or family_changed or family_failed:
            self._scale_cache.clear()
            current = self.words[self.index]["word"].lower()
            if (
                current in changed
                or current in failed
                or current in sense_changed
                or current in family_changed
            ):
                top = self.text.yview()[0]
                self.render()
                self.text.yview_moveto(top)
            self._schedule_usage_save()

        if self._running:
            self._poll_job = self.root.after(250, self._poll_usage)

    def _services(self):
        return tuple(
            service
            for service in (self.usage_service, self.sense_service, self.family_service)
            if service is not None
        )

    def _schedule_usage_save(self) -> None:
        if self._usage_save_job is not None:
            return
        if not self._services():
            return
        self._usage_save_job = self.root.after(3000, self._save_usage)

    def _save_usage(self) -> None:
        self._usage_save_job = None
        for service in self._services():
            service.save()

    def regenerate_usage(self, event=None):
        if self._typing():
            return None
        service = self.usage_service
        if service is None:
            self._flash_status("未启用 AI 用法生成（设置中开启）")
            return "break"
        word = self.words[self.index]["word"]
        service.forget(word)
        service.enqueue([(word, _meaning_hint(self.words[self.index]))], urgent=True)
        self.render()
        return "break"

    def regenerate_senses(self, event=None):
        if self._typing():
            return None
        service = self.sense_service
        if service is None:
            self._flash_status("未启用 AI 精简翻译（设置中开启）")
            return "break"
        entry = self.words[self.index]
        service.forget(entry["word"])
        service.enqueue([(entry["word"], entry["translation"])], urgent=True, batch_limit=1)
        self.render()
        return "break"

    def regenerate_family(self, event=None):
        if self._typing():
            return None
        service = self.family_service
        if service is None:
            self._flash_status("未启用 AI 精简同根词（设置中开启）")
            return "break"
        entry = self.words[self.index]
        if not entry["family"]:
            self._flash_status("这个词没有同根词")
            return "break"
        service.forget(entry["word"])
        service.enqueue([(entry["word"], entry["family"])], urgent=True, batch_limit=1)
        self.render()
        return "break"

    # ---------- 背词模式与生词标记 ----------

    def _mode(self) -> str:
        return MODE_REVIEW if self.hide_mode else MODE_COPY

    def _record_position(self) -> None:
        """当前词记入当前模式的进度槽（抄写/背词互不影响）。"""
        if self.words:
            self.mode_positions[self._mode()] = self.words[self.index]["word"]

    def _mode_slot_index(self, mode: str) -> int | None:
        word = self.mode_positions.get(mode) or ""
        if not word:
            return None
        for position, entry in enumerate(self.words):
            if entry["word"] == word:
                return position
        return None

    def _slot_text(self, mode: str) -> str:
        index = self._mode_slot_index(mode)
        return f"{index + 1} / {len(self.words)}" if index is not None else "未记录"

    def toggle_hide_mode(self, event=None):
        if self._typing():
            return None
        self._record_position()  # 先记下当前模式的进度再切换
        self.hide_mode = not self.hide_mode
        self._revealed = False
        self.config["hide_mode"] = self.hide_mode
        usage_ai.save_config(self.config)
        target = self._mode_slot_index(self._mode())
        if target is not None:
            self.index = target  # 跳回另一模式的上次进度；没记录就停在当前词
        self._scale_cache.clear()
        self.render()
        if self.hide_mode:
            self._flash_status(
                f"背词模式：释义已隐藏，回车 看答案（抄写进度 {self._slot_text(MODE_COPY)} 已保存）"
            )
        else:
            self._flash_status(
                f"抄写模式：进度 {self._slot_text(MODE_COPY)}"
                f"（背词进度 {self._slot_text(MODE_REVIEW)} 已保存）"
            )
        return "break"

    def random_word(self, event=None):
        if self._typing():
            return None
        pool = self.filtered if self._filter_active() and self.filtered else list(range(len(self.words)))
        candidates = [item for item in pool if item != self.index]
        if candidates:
            self.jump_to(random.choice(candidates))
        return "break"

    def _word_mark(self, word: str) -> int:
        bucket = self.marks.get(self.source.name)
        if not isinstance(bucket, dict):
            return 0
        try:
            return int(bucket.get(word, 0) or 0)
        except (TypeError, ValueError):
            return 0

    def _toggle_mark(self, code: int) -> None:
        if self._typing() or not self.words:
            return
        word = self.words[self.index]["word"]
        bucket = self.marks.setdefault(self.source.name, {})
        if bucket.get(word) == code:
            bucket.pop(word, None)
            label = "已取消标记"
        else:
            bucket[word] = code
            label = "已掌握" if code == MARK_KNOWN else "生词"
        if not bucket:
            self.marks.pop(self.source.name, None)
        self._write_state()
        self._flash_status(f"{word}：{label}")
        if self.hard_only:
            self.apply_filter()
        else:
            self._draw_cards()

    def _on_hard_filter_toggle(self) -> None:
        self.hard_only = bool(self.hard_only_var.get())
        self.apply_filter()

    def export_hard_words(self) -> None:
        entries = [
            entry for entry in self.words if self._word_mark(entry["word"]) == MARK_HARD
        ]
        if not entries:
            self._flash_status("当前词表还没有生词（选中单词按 2 标记）")
            return
        target = APP_DIR / f"{self.source.stem}-生词.json"
        payload = {"words": [dict(entry) for entry in entries]}
        try:
            target.write_text(json.dumps(payload, ensure_ascii=False, indent=1), "utf-8")
        except OSError as exc:
            self._flash_status(f"导出失败：{exc}")
            return
        self._flash_status(f"已导出 {len(entries)} 个生词到 {target.name}")

    # ---------- 侧边栏卡片 ----------

    def _draw_cards(self) -> None:
        canvas = self.cards
        canvas.delete("card")
        if not canvas.winfo_ismapped():
            return
        width = canvas.winfo_width()
        if width <= 1:
            return

        total_height = self.CARD_PAD * 2 + len(self.filtered) * self.CARD_H
        canvas.configure(scrollregion=(0, 0, width, total_height))
        top = canvas.canvasy(0)
        height = canvas.winfo_height()
        first = max(0, int((top - self.CARD_PAD) // self.CARD_H))
        last = min(len(self.filtered), int((top + height - self.CARD_PAD) // self.CARD_H) + 2)
        for position in range(first, last):
            self._draw_card(position, width)

    def _draw_card(self, position: int, width: int) -> None:
        entry_index = self.filtered[position]
        entry = self.words[entry_index]
        left, right = 8, width - 18
        top = self.CARD_PAD + position * self.CARD_H + 2
        bottom = top + self.CARD_H - 10
        colors = self.colors
        selected = entry_index == self.index

        self._rounded(
            left,
            top,
            right,
            bottom,
            10,
            fill=colors["card_active"] if selected else colors["card"],
            outline=colors["card_active_border"] if selected else colors["card_border"],
            width=1,
            tags="card",
        )
        if selected:
            self.cards.create_rectangle(
                left + 2,
                top + 10,
                left + 4,
                bottom - 10,
                fill=colors["accent"],
                outline="",
                tags="card",
            )

        text_width = right - left - 28
        word_font = self.f_card_word
        self.cards.create_text(
            left + 14,
            top + 17,
            text=self._ellipsize(entry["word"], word_font, text_width),
            anchor="w",
            font=word_font,
            fill=colors["accent"] if selected else colors["ink"],
            tags="card",
        )
        subtitle = entry["ipa"] or entry["translation"].split("\n")[0]
        self.cards.create_text(
            left + 14,
            top + 36,
            text=self._ellipsize(subtitle.replace("\n", " "), self.f_card_sub, text_width),
            anchor="w",
            font=self.f_card_sub,
            fill=colors["muted"],
            tags="card",
        )
        mark = self._word_mark(entry["word"])
        if mark:
            self.cards.create_text(
                right - 10,
                top + 17,
                text="✓" if mark == MARK_KNOWN else "●",
                anchor="e",
                font=self.f_card_sub,
                fill=colors["muted"] if mark == MARK_KNOWN else colors["accent"],
                tags="card",
            )

    def _rounded(self, x1: float, y1: float, x2: float, y2: float, radius: float, **options):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.cards.create_polygon(points, smooth=True, **options)

    @staticmethod
    def _ellipsize(text: str, font: tkfont.Font, max_width: int) -> str:
        if font.measure(text) <= max_width:
            return text
        trimmed = text
        while trimmed and font.measure(trimmed + "…") > max_width:
            trimmed = trimmed[:-1]
        return trimmed + "…"

    def _scroll_cards(self, *args) -> None:
        self.cards.yview(*args)
        self._draw_cards()

    def _on_card_click(self, event) -> None:
        position = int((self.cards.canvasy(event.y) - self.CARD_PAD) // self.CARD_H)
        if 0 <= position < len(self.filtered):
            self.index = self.filtered[position]
            self._revealed = False
            self.cards.focus_set()
            self.render()
            self._draw_cards()
            self._ensure_card_visible()

    def _ensure_card_visible(self) -> None:
        canvas = self.cards
        if not canvas.winfo_ismapped():
            return
        position = self._filtered_position()
        if position is None:
            return
        total = self.CARD_PAD * 2 + len(self.filtered) * self.CARD_H
        top = canvas.canvasy(0)
        height = canvas.winfo_height()
        y = self.CARD_PAD + position * self.CARD_H
        if y < top:
            canvas.yview_moveto(max(0.0, (y - self.CARD_PAD) / total))
        elif y + self.CARD_H > top + height:
            canvas.yview_moveto(min(1.0, (y + self.CARD_H + self.CARD_PAD - height) / total))
        self._draw_cards()

    # ---------- 操作 ----------

    def _typing(self) -> bool:
        try:
            widget = self.root.focus_get()
        except (KeyError, tk.TclError):
            return False
        return isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox))

    def _step(self, delta: int) -> None:
        if self._filter_active() and self.filtered:
            position = self._filtered_position()
            if position is None:
                target = 0 if delta > 0 else len(self.filtered) - 1
            else:
                target = max(0, min(len(self.filtered) - 1, position + delta))
            self.jump_to(self.filtered[target])
            return
        self.jump_to(self.index + delta)

    def jump_to(self, index: int) -> None:
        index = max(0, min(len(self.words) - 1, index))
        if index == self.index:
            self._draw_cards()
            self._ensure_card_visible()
            return
        self.index = index
        self._revealed = False
        self.render()
        self._speak_current(auto=True)

    def next_word(self, event=None):
        if self._typing():
            return None
        # 背词模式下回车先显示释义，再按一次才翻页
        if (
            self.hide_mode
            and not self._revealed
            and event is not None
            and event.keysym in ("Return", "KP_Enter")
        ):
            self._revealed = True
            self.render()
            return "break"
        if self._filter_active():
            self._step(1)  # _step 内部经 jump_to 移动时会重置遮挡
            return "break"
        if self.index < len(self.words) - 1:
            self.index += 1
            self._revealed = False
            self.render()
            self._speak_current(auto=True)
        return "break"

    def prev_word(self, event=None):
        if self._typing():
            return None
        if self._filter_active():
            self._step(-1)
            return "break"
        if self.index > 0:
            self.index -= 1
            self._revealed = False
            self.render()
            self._speak_current(auto=True)
        return "break"

    def first_word(self, event=None):
        if self._typing():
            return None
        if self._filter_active() and self.filtered:
            self.jump_to(self.filtered[0])
            return "break"
        self.jump_to(0)
        return "break"

    def last_word(self, event=None):
        if self._typing():
            return None
        if self._filter_active() and self.filtered:
            self.jump_to(self.filtered[-1])
            return "break"
        self.jump_to(len(self.words) - 1)
        return "break"

    def toggle_extra(self, event=None):
        if self._typing():
            return None
        value = not (self.sections.get("family") or self.sections.get("origin"))
        for key in ("family", "origin"):
            self.sections[key] = value
            variable = self.section_vars.get(key)
            if variable is not None:
                variable.set(value)
        self._sections_changed()
        return "break"

    # ---------- 显示内容扩展栏 ----------

    def _build_sections_bar(self, parent) -> None:
        colors = self.colors
        bar = tk.Frame(parent, bg=colors["bg"])
        self.sections_bar = bar
        self.section_vars: dict[str, tk.BooleanVar] = {}
        self.section_checks: dict[str, tk.Checkbutton] = {}
        for key, label in SECTION_ORDER:
            variable = tk.BooleanVar(value=bool(self.sections.get(key)))
            check = tk.Checkbutton(
                bar,
                text=label,
                variable=variable,
                command=lambda k=key: self._on_section_toggle(k),
                bg=colors["bg"],
                fg=colors["ink"],
                activebackground=colors["bg"],
                activeforeground=colors["accent"],
                selectcolor=colors["panel"],
                font=self.f_ui,
                bd=0,
                highlightthickness=0,
                padx=2,
                cursor="hand2",
                takefocus=0,
            )
            check.pack(side="left", padx=(0, 12))
            self.section_vars[key] = variable
            self.section_checks[key] = check

        tip = tk.Label(
            bar,
            text="只保留需要抄写的内容",
            bg=colors["bg"],
            fg=colors["muted"],
            font=self.f_ui,
        )
        tip.pack(side="left", padx=(6, 0))
        self.sections_tip = tip

    def toggle_sections_bar(self, event=None):
        self.sections_open = not self.sections_open
        self._layout_sections_bar()
        self.config["sections_open"] = self.sections_open
        usage_ai.save_config(self.config)
        return "break"

    def _layout_sections_bar(self) -> None:
        if self.sections_open:
            self.sections_bar.pack(fill="x", padx=40, pady=(2, 0), after=self.header)
            self.sections_button.configure(text="显示内容 ▾")
        else:
            self.sections_bar.pack_forget()
            self.sections_button.configure(text="显示内容 ▸")

    def _on_section_toggle(self, key: str) -> None:
        self.sections[key] = bool(self.section_vars[key].get())
        self._sections_changed()

    def _sections_changed(self) -> None:
        self._scale_cache.clear()
        self._save_sections()
        self.render()

    def _save_sections(self) -> None:
        self.config["sections"] = dict(self.sections)
        self.config["sections_open"] = self.sections_open
        usage_ai.save_config(self.config)

    def change_scale(self, delta: float) -> None:
        if self._typing():
            return
        zoom = round(min(1.8, max(0.7, self.zoom + delta)), 2)
        if zoom == self.zoom:
            return
        self.zoom = zoom
        self.config["zoom"] = zoom
        usage_ai.save_config(self.config)
        self.refit()

    def _on_wheel(self, event):
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is self.cards or widget is self.cards_sb:
            step = -1 if event.delta > 0 else 1
            self.cards.yview_scroll(step * 2, "units")
            self._draw_cards()
            return "break"
        return None

    def _on_ctrl_wheel(self, event):
        self.change_scale(0.1 if event.delta > 0 else -0.1)
        return "break"

    def toggle_fullscreen(self, event=None):
        self.root.attributes("-fullscreen", not self.root.attributes("-fullscreen"))
        return "break"

    # ---------- 进度记忆 ----------

    def _load_state(self) -> dict:
        state: dict = {"positions": {}, "marks": {}, "geometry": ""}
        try:
            stored = json.loads(STATE_FILE.read_text("utf-8"))
        except (OSError, ValueError):
            return state
        if not isinstance(stored, dict):
            return state
        positions = stored.get("positions")
        if isinstance(positions, dict):
            state["positions"] = positions
        elif stored.get("file") and stored.get("word"):
            state["positions"][str(stored["file"])] = str(stored["word"])
        marks = stored.get("marks")
        if isinstance(marks, dict):
            state["marks"] = {
                str(name): dict(bucket) if isinstance(bucket, dict) else {}
                for name, bucket in marks.items()
            }
        geometry = stored.get("geometry")
        if isinstance(geometry, str):
            state["geometry"] = geometry
        return state

    def _write_state(self) -> None:
        try:
            STATE_FILE.write_text(json.dumps(self._state, ensure_ascii=False, indent=1), "utf-8")
        except OSError:
            pass

    def _restore_position(self) -> None:
        self.index = 0
        entry = _position_entry(self._state["positions"].get(self.source.name))
        self.mode_positions = {"copy": entry["copy"], "review": entry["review"]}
        word = self.mode_positions[MODE_REVIEW if self.hide_mode else MODE_COPY]
        if not word:
            return
        for position, item in enumerate(self.words):
            if item["word"] == word:
                self.index = position
                break

    def _save_progress(self) -> None:
        if not self.words:
            return
        self._record_position()
        if self.root.winfo_width() > 200:
            self._state["geometry"] = self.root.geometry()
        self._state["positions"][self.source.name] = {
            slot: word for slot, word in self.mode_positions.items() if word
        }
        self._write_state()

    def on_close(self) -> None:
        self._running = False
        if self._market is not None:
            self._market.close()
            self._market = None
        for job in (self._poll_job, self._watch_job, self._flash_job, self._search_job):
            if job is not None:
                try:
                    self.root.after_cancel(job)
                except tk.TclError:
                    pass
        self._poll_job = self._watch_job = self._flash_job = self._search_job = None
        self._save_progress()
        self.speaker.stop()
        if self._usage_save_job is not None:
            self.root.after_cancel(self._usage_save_job)
            self._usage_save_job = None
        for service in self._services():
            service.stop()
            service.save()
        self.root.destroy()


def main() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    source = find_vocabulary()
    vocab.clean_stale_caches([source])
    words = vocab.load_cached(source)
    if not words:
        raise SystemExit(f"词汇表里没有读到单词：{source.name}")

    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    WordApp(root, words, source)
    root.mainloop()


if __name__ == "__main__":
    main()
