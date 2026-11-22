# -*- coding: utf-8 -*-
"""语音朗读：把 pyttsx3 放在后台线程里，界面线程只负责投递文本。

pyttsx3 的引擎不能跨线程调用，而且 runAndWait 会阻塞，所以这里用一条常驻线程 +
队列：新的朗读请求会丢弃还没播的旧请求，保证用户连续点按钮时只读最新的内容。
"""

from __future__ import annotations

import gc
import queue
import re
import threading
import time

try:  # 语音是可选的：没装 pyttsx3 时其余功能照常
    import pyttsx3

    PYTTSX3_AVAILABLE = True
    PYTTSX3_ERROR = ""
except Exception as exc:  # pragma: no cover - 取决于本机环境
    pyttsx3 = None
    PYTTSX3_AVAILABLE = False
    PYTTSX3_ERROR = str(exc)

MAX_CHARS = 600
_CONTROL = re.compile(r"[\x00-\x08\x0b-\x1f]")
_CJK = re.compile(r"[\u3000-\u303f\u3400-\u9fff\uff00-\uffef]+")


def spoken_english(text: str) -> str:
    """去掉中文与中文标点，只留英文朗读内容（英文音色读汉字会变噪音）。"""
    cleaned = _CJK.sub(" ", text or "")
    cleaned = _CONTROL.sub(" ", cleaned)
    return cleaned.strip(" \t·-—–:;,.!?")


def _voice_languages(voice) -> str:
    parts: list[str] = []
    for item in getattr(voice, "languages", None) or []:
        if isinstance(item, bytes):
            text = item.decode("utf-8", "ignore")
        else:
            text = str(item)
        text = text.replace("\x05", "").strip()
        if text:
            parts.append(text)
    return ",".join(parts)


def _english_score(voice: dict) -> int:
    lang = voice.get("lang", "").lower()
    name = voice.get("name", "").lower()
    if "en-us" in lang:
        return 0
    if lang.startswith("en") or "en-gb" in lang:
        return 1
    if "english" in name:
        return 2
    return 9


def pick_english(voices: list[dict]) -> str:
    if not voices:
        return ""
    best = min(voices, key=_english_score)
    if _english_score(best) < 9:
        return best["id"]
    return voices[0]["id"]


class Speaker:
    """后台朗读器；界面只需要 speak()/stop() 和 voices 列表。"""

    def __init__(self, voice_id: str = "", rate: int = 170, volume: float = 1.0) -> None:
        self.available = PYTTSX3_AVAILABLE
        self.error = PYTTSX3_ERROR
        self.voices: list[dict] = []
        self.voice_id = str(voice_id or "")
        self.rate = int(rate)
        self.volume = float(volume)
        self.ready = threading.Event()
        self._queue: queue.Queue = queue.Queue()
        self._thread = None
        if self.available:
            self._thread = threading.Thread(target=self._loop, daemon=True, name="speech")
            self._thread.start()
        else:
            self.ready.set()

    # ---------- 对外接口 ----------

    def speak(self, text: str) -> None:
        text = (text or "").strip()[:MAX_CHARS]
        if not self.available or not text:
            return
        self._queue.put(("say", text))

    def configure(self, voice_id: str | None = None, rate: int | None = None,
                  volume: float | None = None) -> None:
        if voice_id is not None:
            self.voice_id = str(voice_id)
        if rate is not None:
            self.rate = int(rate)
        if volume is not None:
            self.volume = float(volume)
        self._queue.put(("config",))

    def stop(self) -> None:
        self._queue.put(("quit",))

    def voice_label(self, voice: dict) -> str:
        lang = voice.get("lang", "")
        return f"{voice.get('name', '')}（{lang}）" if lang else str(voice.get("name", ""))

    # ---------- 后台线程 ----------

    def _apply(self, engine) -> None:
        engine.setProperty("rate", int(self.rate))
        engine.setProperty("volume", max(0.0, min(1.0, float(self.volume))))
        if self.voice_id:
            try:
                engine.setProperty("voice", self.voice_id)
            except Exception:  # 旧配置里的音色可能已经不在了
                pass

    def _new_engine(self):
        engine = pyttsx3.init()
        self._apply(engine)
        return engine

    def _loop(self) -> None:
        try:
            engine = self._new_engine()
        except Exception as exc:
            self.available = False
            self.error = str(exc)
            self.ready.set()
            return

        try:
            raw_voices = engine.getProperty("voices") or []
            self.voices = [
                {
                    "id": str(getattr(voice, "id", "")),
                    "name": str(getattr(voice, "name", "")),
                    "lang": _voice_languages(voice),
                }
                for voice in raw_voices
            ]
            known = {voice["id"] for voice in self.voices}
            if self.voice_id not in known:
                self.voice_id = pick_english(self.voices)
            self._apply(engine)
        except Exception as exc:
            self.error = str(exc)
        self.ready.set()

        while True:
            item = self._queue.get()
            kind = item[0]
            if kind == "quit":
                break
            if kind == "config":
                try:
                    self._apply(engine)
                except Exception as exc:
                    self.error = str(exc)
                continue
            if kind != "say":
                continue

            text = item[1]
            # 排队中还有请求就只读最后一条
            while True:
                try:
                    nxt = self._queue.get_nowait()
                except queue.Empty:
                    break
                if nxt[0] == "quit":
                    return
                if nxt[0] == "config":
                    try:
                        self._apply(engine)
                    except Exception:
                        pass
                elif nxt[0] == "say":
                    text = nxt[1]
            if self._play(engine, text):
                # 被打断的引擎之后会哑火；必须先彻底释放它，再建新的，否则新引擎也不出声
                engine = None
                gc.collect()
                try:
                    engine = self._new_engine()
                except Exception as exc:
                    self.available = False
                    self.error = str(exc)
                    return

    def _play(self, engine, text: str) -> bool:
        """念一句。用 startLoop(False)+iterate 手动推进，runAndWait 重复调用会哑火。
        返回 True 表示这句被打断，调用方需要换一个新引擎。"""
        interrupted = False
        try:
            engine.say(text)
            engine.startLoop(False)
            while engine.isBusy():
                engine.iterate()
                time.sleep(0.02)
                if not self._queue.empty():
                    engine.stop()  # 又有新点击：掐断这句，马上去读新的
                    interrupted = True
                    break
        except Exception as exc:
            self.error = str(exc)
            interrupted = True
        finally:
            try:
                engine.endLoop()
            except Exception:
                pass
        return interrupted
