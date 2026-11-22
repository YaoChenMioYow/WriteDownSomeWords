# -*- coding: utf-8 -*-
"""在线词库窗口：浏览 qwerty-learner 的词库清单，下载到本地并直接打开。"""

from __future__ import annotations

import os
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import dict_store

APP_DIR = Path(__file__).resolve().parent
ALL_CATEGORIES = "全部"

COLUMNS = (
    ("name", "词库", 260),
    ("count", "词数", 70),
    ("category", "分类", 110),
    ("state", "本地状态", 130),
)


class DictMarket:
    """一个独立的 Toplevel；所有网络操作都在后台线程里做。"""

    def __init__(self, app, chrome=None) -> None:
        self.app = app
        self.chrome = chrome
        self.root = app.root
        self.colors = app.colors
        self.metas: list[dict] = []
        self.by_id: dict[str, dict] = {}
        self.local: dict[str, dict] = {}
        self.mirror = ""
        self.busy = False
        self.closed = False
        self.messages: queue.Queue = queue.Queue()
        self._poll_job: str | None = None
        self._last_progress = 0.0

        win = tk.Toplevel(self.root)
        self.win = win
        win.title("在线词库 · qwerty-learner")
        win.configure(bg=self.colors["bg"])
        win.transient(self.root)
        win.minsize(680, 480)
        win.geometry("760x560")
        win.geometry(f"+{self.root.winfo_rootx() + 120}+{self.root.winfo_rooty() + 60}")
        try:
            win.iconbitmap(default=str(APP_DIR / "icon.ico"))
        except tk.TclError:
            pass
        self._apply_titlebar()
        win.protocol("WM_DELETE_WINDOW", self.close)

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self._muted_labels: list[tk.Label] = []
        self._ink_labels: list[tk.Label] = []
        self._frames: list[tk.Frame] = []

        head = tk.Frame(win, bg=self.colors["bg"])
        head.pack(fill="x", padx=22, pady=(18, 6))
        self._frames.append(head)
        title = tk.Label(
            head,
            text="在线词库",
            bg=self.colors["bg"],
            fg=self.colors["ink"],
            font=app.f_ui,
            anchor="w",
        )
        title.pack(side="left")
        self._ink_labels.append(title)
        self.source_label = tk.Label(
            head,
            text=f"来源：{dict_store.SOURCE_NAME} · 下载到 {dict_store.DICT_DIR}",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=app.f_ui,
            anchor="w",
        )
        self.source_label.pack(side="left", padx=(12, 0))
        self._muted_labels.append(self.source_label)

        tools = tk.Frame(win, bg=self.colors["bg"])
        tools.pack(fill="x", padx=22, pady=(0, 8))
        self._frames.append(tools)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            tools,
            textvariable=self.search_var,
            bg=self.colors["panel"],
            fg=self.colors["ink"],
            insertbackground=self.colors["ink"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.colors["card_border"],
            highlightcolor=self.colors["accent"],
            font=app.f_ui,
            width=22,
        )
        self.search_entry.pack(side="left", ipady=4, padx=(0, 8))
        tk.Label(
            tools,
            text="搜索",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=app.f_ui,
        ).pack(side="left", padx=(0, 14))
        self._muted_labels.append(tools.winfo_children()[-1])

        self.category_var = tk.StringVar(value=ALL_CATEGORIES)
        self.category_box = ttk.Combobox(
            tools,
            textvariable=self.category_var,
            values=[ALL_CATEGORIES],
            state="readonly",
            width=14,
            font=app.f_ui,
        )
        self.category_box.pack(side="left")
        self.refresh_button = self._button(tools, "刷新清单", self.refresh_online)
        self.refresh_button.pack(side="right")
        self.folder_button = self._button(tools, "打开下载目录", self.open_folder)
        self.folder_button.pack(side="right", padx=(0, 8))

        body = tk.Frame(win, bg=self.colors["bg"])
        body.pack(fill="both", expand=True, padx=22, pady=(0, 6))
        self._frames.append(body)
        self.tree = ttk.Treeview(
            body,
            columns=[column[0] for column in COLUMNS],
            show="headings",
            style="Market.Treeview",
            selectmode="browse",
        )
        for key, label, width in COLUMNS:
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor="w" if key != "count" else "e", stretch=True)
        scroll = ttk.Scrollbar(
            body,
            orient="vertical",
            command=self.tree.yview,
            style="Main.Vertical.TScrollbar",
        )
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda event: self.download_selected(open_after=True))
        self.tree.bind("<Return>", lambda event: self.download_selected(open_after=True))
        self.tree.bind("<<TreeviewSelect>>", lambda event: self._update_buttons())

        bottom = tk.Frame(win, bg=self.colors["bg"])
        bottom.pack(fill="x", padx=22, pady=(0, 18))
        self._frames.append(bottom)
        self.progress = ttk.Progressbar(
            bottom, mode="determinate", maximum=100, style="Market.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(0, 8))

        self.status_label = tk.Label(
            bottom,
            text="准备就绪",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=app.f_ui,
            anchor="w",
        )
        self.status_label.pack(side="left")
        self._muted_labels.append(self.status_label)
        self.count_label = tk.Label(
            bottom, text="", bg=self.colors["bg"], fg=self.colors["muted"], font=app.f_ui
        )
        self.count_label.pack(side="left", padx=(12, 0))
        self._muted_labels.append(self.count_label)

        self.delete_button = self._button(bottom, "删除本地副本", self.delete_local)
        self.delete_button.pack(side="right")
        self.open_button = self._button(bottom, "下载并打开", self.download_and_open)
        self.open_button.pack(side="right", padx=(0, 8))
        self.download_button = self._button(bottom, "下载", self.download_only)
        self.download_button.pack(side="right", padx=(0, 8))

        self.search_var.trace_add("write", lambda *args: self.refresh())
        self.category_var.trace_add("write", lambda *args: self.refresh())

        self.apply_colors()
        self._poll_job = self.win.after(120, self._poll)
        self._bootstrap()
        self.search_entry.focus_set()

    # ---------- 界面辅助 ----------

    def _button(self, parent, text: str, command) -> tk.Button:
        colors = self.colors
        return tk.Button(
            parent,
            text=text,
            command=command,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=self.app.f_ui,
            padx=10,
            pady=4,
            bg=colors["card"],
            fg=colors["ink"],
            activebackground=colors["accent"],
            activeforeground=colors["bg"],
            highlightthickness=0,
        )

    def _apply_titlebar(self) -> None:
        if self.chrome is None:
            return
        try:
            self.chrome(self.win, self.app.theme_name == "dark")
        except Exception:
            pass

    def apply_colors(self) -> None:
        colors = self.colors = self.app.colors
        self.win.configure(bg=colors["bg"])
        for frame in self._frames:
            frame.configure(bg=colors["bg"])
        for label in self._muted_labels:
            label.configure(bg=colors["bg"], fg=colors["muted"])
        for label in self._ink_labels:
            label.configure(bg=colors["bg"], fg=colors["ink"])
        self.search_entry.configure(
            bg=colors["panel"],
            fg=colors["ink"],
            insertbackground=colors["ink"],
            highlightbackground=colors["card_border"],
            highlightcolor=colors["accent"],
        )
        self.source_label.configure(
            text=f"来源：{dict_store.SOURCE_NAME} · 下载到 {dict_store.DICT_DIR}"
        )
        style = ttk.Style(self.win)
        style.configure(
            "Market.Treeview",
            background=colors["panel"],
            fieldbackground=colors["panel"],
            foreground=colors["ink"],
            rowheight=26,
            borderwidth=0,
            font=self.app.f_ui,
        )
        style.configure(
            "Market.Treeview.Heading",
            background=colors["card"],
            foreground=colors["muted"],
            relief="flat",
            font=self.app.f_ui,
        )
        style.map(
            "Market.Treeview",
            background=[("selected", colors["select"])],
            foreground=[("selected", colors["ink"])],
        )
        style.map("Market.Treeview.Heading", background=[("active", colors["card"])])
        style.configure(
            "Market.Horizontal.TProgressbar",
            background=colors["accent"],
            troughcolor=colors["panel"],
            bordercolor=colors["panel"],
            lightcolor=colors["accent"],
            darkcolor=colors["accent"],
        )
        for button in (
            self.refresh_button,
            self.folder_button,
            self.download_button,
            self.open_button,
            self.delete_button,
        ):
            button.configure(
                bg=colors["card"],
                fg=colors["ink"],
                activebackground=colors["accent"],
                activeforeground=colors["bg"],
            )
        self._apply_titlebar()

    def _poll(self) -> None:
        self._poll_job = None
        if self.closed:
            return
        try:
            while True:
                try:
                    message = self.messages.get_nowait()
                except queue.Empty:
                    break
                try:
                    self._dispatch(message)
                except tk.TclError:
                    continue
        finally:
            if not self.closed:
                self._poll_job = self.win.after(120, self._poll)

    def _dispatch(self, message: tuple) -> None:
        kind = message[0]
        if kind == "index":
            self._fetch_done(message[1], message[2])
        elif kind == "index_failed":
            self._fetch_failed(message[1])
        elif kind == "trying":
            self.set_status(f"正在连接 {message[1]}…")
        elif kind == "progress":
            self._set_progress(message[1], message[2])
        elif kind == "done":
            self._download_done(message[1], message[2], message[3])
        elif kind == "failed":
            self._download_failed(message[1])

    def set_status(self, text: str) -> None:
        if self.closed:
            return
        self.status_label.configure(text=text)

    # ---------- 数据 ----------

    def _bootstrap(self) -> None:
        self.set_status("正在读取本地已下载词库…")
        self.win.update_idletasks()
        self.local = dict_store.local_index()
        metas, mirror, saved = dict_store.cached_index()
        if metas:
            self.metas = metas
            self.mirror = mirror
            self._fill_categories()
            self.refresh()
            self.set_status(f"缓存清单：{len(metas)} 个词库（更新于 {saved}），正在联网刷新…")
        else:
            self.set_status("正在获取在线词库清单…")
        self.refresh_online()

    def refresh_online(self) -> None:
        if self.busy:
            return
        self.busy = True
        self._update_buttons()
        self.set_status("正在获取在线词库清单…")
        threading.Thread(target=self._fetch_worker, daemon=True).start()

    def _fetch_worker(self) -> None:
        try:
            metas, mirror = dict_store.fetch_index(
                timeout=15, on_try=lambda label: self.messages.put(("trying", label))
            )
        except Exception as exc:
            self.messages.put(("index_failed", str(exc)))
            return
        self.messages.put(("index", metas, mirror))

    def _fetch_failed(self, message: str) -> None:
        self.busy = False
        self._update_buttons()
        suffix = "" if self.metas else "，可以先检查网络或稍后重试"
        self.set_status(f"获取失败：{message}{suffix}")

    def _fetch_done(self, metas: list[dict], mirror: str) -> None:
        self.busy = False
        self.metas = metas
        self.mirror = mirror
        self.local = dict_store.local_index()
        self._fill_categories()
        self.refresh()
        self._update_buttons()
        self.set_status(f"已通过 {mirror} 获取 {len(metas)} 个词库")

    def _fill_categories(self) -> None:
        categories: list[str] = []
        for meta in self.metas:
            name = str(meta.get("category") or "")
            if name and name not in categories:
                categories.append(name)
        current = self.category_var.get()
        self.category_box.configure(values=[ALL_CATEGORIES] + categories)
        if current not in categories and current != ALL_CATEGORIES:
            self.category_var.set(ALL_CATEGORIES)

    def refresh(self) -> None:
        if self.closed:
            return
        keyword = self.search_var.get().strip().lower()
        category = self.category_var.get()
        shown = 0
        children = self.tree.get_children()
        if children:
            self.tree.delete(*children)
        self.by_id.clear()
        for meta in self.metas:
            if category and category != ALL_CATEGORIES and meta.get("category") != category:
                continue
            if keyword:
                haystack = " ".join(
                    [
                        str(meta.get("name") or ""),
                        str(meta.get("description") or ""),
                        str(meta.get("category") or ""),
                        " ".join(str(tag) for tag in meta.get("tags") or []),
                        str(meta.get("id") or ""),
                        str(meta.get("file") or ""),
                    ]
                ).lower()
                if keyword not in haystack:
                    continue
            key = str(meta.get("id"))
            self.by_id[key] = meta
            local = self.local.get(key)
            state = ""
            if local:
                count = local.get("count") or 0
                state = f"✓ 已下载 {count} 词" if count else "✓ 已下载"
            length = meta.get("length") or 0
            self.tree.insert(
                "",
                "end",
                iid=key,
                values=(meta.get("name", ""), length or "", meta.get("category", ""), state),
            )
            shown += 1
        self.count_label.configure(text=f"{shown} / {len(self.metas)} 个词库")
        self._update_buttons()

    def _selected(self) -> dict | None:
        selection = self.tree.selection()
        if not selection:
            self.set_status("请先在列表里选中一个词库")
            return None
        return self.by_id.get(selection[0])

    def _update_buttons(self) -> None:
        if self.closed:
            return
        states = "disabled" if self.busy else "normal"
        for button in (self.download_button, self.open_button, self.refresh_button):
            button.configure(state=states)
        meta = None
        selection = self.tree.selection()
        if selection:
            meta = self.by_id.get(selection[0])
        has_local = bool(meta and self.local.get(str(meta.get("id"))))
        self.delete_button.configure(state="normal" if has_local and not self.busy else "disabled")

    def _set_progress(self, done: int, total: int) -> None:
        if self.closed:
            return
        if total:
            self.progress.configure(value=min(100, done * 100 / total))
            self.set_status(
                f"正在下载… {done / 1024 / 1024:.1f} / {total / 1024 / 1024:.1f} MB"
            )
        else:
            self.progress.configure(value=0)
            self.set_status(f"正在下载… {done / 1024:.0f} KB")

    def _on_progress(self, done: int, total: int) -> None:
        now = time.time()
        if total and done < total and now - self._last_progress < 0.12:
            return
        self._last_progress = now
        self.messages.put(("progress", done, total))

    # ---------- 下载 / 打开 ----------

    def download_only(self) -> None:
        meta = self._selected()
        if meta:
            self._start_download(meta, open_after=False)

    def download_and_open(self) -> None:
        meta = self._selected()
        if not meta:
            return
        local = self.local.get(str(meta.get("id")))
        if local:
            path = Path(local["path"])
            if path.is_file() and self.app.switch_vocabulary(path):
                self.set_status(f"已切换到 {local['name']}")
                return
        self._start_download(meta, open_after=True)

    def _start_download(self, meta: dict, open_after: bool) -> None:
        if self.busy:
            self.set_status("清单正在刷新，请稍候再下载")
            return
        self.busy = True
        self._update_buttons()
        self.progress.configure(value=0)
        self.set_status(f"正在下载 {meta.get('name', '')}…")
        self._last_progress = 0.0

        def worker() -> None:
            try:
                path = dict_store.download(
                    meta, mirror=self.mirror, progress=self._on_progress
                )
            except Exception as exc:
                self.messages.put(("failed", str(exc)))
                return
            self.messages.put(("done", meta, path, open_after))

        threading.Thread(target=worker, daemon=True).start()

    def _download_failed(self, message: str) -> None:
        self.busy = False
        self.progress.configure(value=0)
        self._update_buttons()
        self.set_status(f"下载失败：{message}")

    def _download_done(self, meta: dict, path: Path, open_after: bool) -> None:
        self.busy = False
        self.progress.configure(value=0)
        self.local = dict_store.local_index()
        self.refresh()
        self._update_buttons()
        key = str(meta.get("id"))
        if self.tree.exists(key):
            self.tree.selection_set(key)
            self.tree.see(key)
        self.set_status(f"已下载到 {path.name}（{path.parent}）")
        if open_after and self.app.switch_vocabulary(path):
            self.set_status(f"已切换到 {meta.get('name', '')}（{path.name}）")

    def delete_local(self) -> None:
        meta = self._selected()
        if not meta:
            return
        local = self.local.get(str(meta.get("id")))
        if not local:
            self.set_status("这个词库还没有下载")
            return
        path = Path(local["path"])
        if not messagebox.askyesno(
            "删除本地副本", f"从本地删除「{local['name']}」？\n{path}", parent=self.win
        ):
            return
        try:
            dict_store.remove(path)
        except OSError as exc:
            self.set_status(f"删除失败：{exc}")
            return
        self.local = dict_store.local_index()
        self.refresh()
        self.set_status(f"已删除 {path.name}")

    def open_folder(self) -> None:
        try:
            dict_store.DICT_DIR.mkdir(parents=True, exist_ok=True)
            os.startfile(dict_store.DICT_DIR)  # noqa: S606 - 打开本地目录
        except OSError as exc:
            self.set_status(f"无法打开目录：{exc}")

    def close(self) -> None:
        if self.closed:
            return
        self.closed = True
        if self._poll_job is not None:
            try:
                self.win.after_cancel(self._poll_job)
            except tk.TclError:
                pass
            self._poll_job = None
        if getattr(self.app, "_market", None) is self:
            self.app._market = None
        try:
            self.win.destroy()
        except tk.TclError:
            pass
