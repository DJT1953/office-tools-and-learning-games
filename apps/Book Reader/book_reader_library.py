"""The Book Reader's library: which books there are, and what you did in each.

Settings, the library list (a details cache so later starts are quick), cover
thumbnails and one small file per book (place, bookmarks, highlights, notes,
favourite, finished, collections, reading speed) live in a state folder,
normally %APPDATA%\\Book Reader. Every file is written to a temporary file and
then swapped in, so a crash or power cut never leaves half a file. A file that
can't be read is kept as name.bad and defaults are used instead.

A book is known by a fingerprint of its contents (size + first 1 MiB), so its
notes follow it when the file is moved or renamed.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import tempfile
import threading
import time
from dataclasses import asdict, dataclass, field
from typing import Any, Callable

import book_reader_formats as formats
from book_reader_formats import READABLE, BookInfo, LockInfo

DEFAULT_KINDLE_FOLDER = os.path.join(os.path.expanduser("~"), "Documents", "My Kindle Content")
FONTS = ["Georgia", "Times New Roman", "Palatino Linotype", "Cambria", "Verdana", "Arial",
         "Segoe UI", "Calibri", "The book's own"]
LINE_SPACINGS = [1.2, 1.5, 1.8, 2.0]
MARGINS = ["small", "medium", "large"]
COLOURS = ["white", "sepia", "night"]
LAYOUTS = ["one", "two", "scroll"]
SORTS = ["recent", "title", "author", "genre", "added"]
HIGHLIGHT_COLOURS = ["yellow", "green", "blue", "pink"]


def default_state_folder() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(base, "Book Reader")


# ---------------------------------------------------------------- files

def read_json(path: str) -> Any | None:
    """The file's data, or None if it is missing. A damaged file becomes .bad."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (OSError, ValueError, UnicodeDecodeError):
        try:
            os.replace(path, path + ".bad")
        except OSError:
            pass
        return None


def write_json(path: str, data: Any) -> None:
    """Write data so the old file stays whole until the new one is complete."""
    folder = os.path.dirname(path)
    os.makedirs(folder, exist_ok=True)
    text = json.dumps(data, ensure_ascii=False, indent=1, allow_nan=False)
    fd, temp = tempfile.mkstemp(prefix=".", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp, path)
    except BaseException:
        try:
            os.remove(temp)
        except OSError:
            pass
        raise


def _int(value: Any, default: int, low: int, high: int) -> int:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        return default
    return int(min(max(value, low), high))


def _float(value: Any, default: float, low: float, high: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        return default
    return float(value) if low <= value <= high else default


def _str(value: Any, default: str = "", choices: list[str] | None = None) -> str:
    if not isinstance(value, str):
        return default
    if choices is not None and value not in choices:
        return default
    return value


def _bool(value: Any, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def _strings(value: Any) -> list[str]:
    return [v for v in value if isinstance(v, str)] if isinstance(value, list) else []


def book_id(path: str) -> str:
    """A fingerprint of the file's contents: same book, same id, wherever it is."""
    h = hashlib.sha1()
    size = os.path.getsize(path)
    h.update(str(size).encode())
    with open(path, "rb") as f:
        h.update(f.read(1 << 20))
    return h.hexdigest()


# ---------------------------------------------------------------- settings

@dataclass
class Settings:
    folders: list[str] = field(default_factory=list)
    font_px: int = 28
    font: str = "Georgia"
    line_spacing: float = 1.5
    margins: str = "medium"
    colours: str = "white"
    layout: str = "one"
    justify: bool = True
    sort: str = "recent"
    voice: str = ""
    speech_rate: float = 0.0
    last_book: str = ""
    window: str = ""

    @classmethod
    def load(cls, folder: str) -> "Settings":
        data = read_json(os.path.join(folder, "settings.json"))
        data = data if isinstance(data, dict) else {}
        s = cls()
        folders = _strings(data.get("folders"))
        s.folders = folders if "folders" in data and isinstance(data.get("folders"), list) else (
            [DEFAULT_KINDLE_FOLDER] if os.path.isdir(DEFAULT_KINDLE_FOLDER)
            else [os.path.join(os.path.expanduser("~"), "Documents")])
        s.font_px = _int(data.get("font_px"), 28, 12, 60)
        s.font = _str(data.get("font"), "Georgia", FONTS)
        s.line_spacing = _float(data.get("line_spacing"), 1.5, 1.0, 3.0)
        s.margins = _str(data.get("margins"), "medium", MARGINS)
        s.colours = _str(data.get("colours"), "white", COLOURS)
        s.layout = _str(data.get("layout"), "one", LAYOUTS)
        s.justify = _bool(data.get("justify"), True)
        s.sort = _str(data.get("sort"), "recent", SORTS)
        s.voice = _str(data.get("voice"))
        s.speech_rate = _float(data.get("speech_rate"), 0.0, -1.0, 1.0)
        s.last_book = _str(data.get("last_book"))
        s.window = _str(data.get("window"))
        return s

    def save(self, folder: str) -> None:
        write_json(os.path.join(folder, "settings.json"), asdict(self))


# ---------------------------------------------------------------- book state

@dataclass
class Bookmark:
    section: int
    offset: int
    text: str = ""
    made: float = 0.0


@dataclass
class Highlight:
    id: str
    section: int
    start: int
    end: int
    text: str = ""
    colour: str = "yellow"
    note: str = ""
    made: float = 0.0


@dataclass
class BookState:
    place: tuple[int, int] | None = None
    bookmarks: list[Bookmark] = field(default_factory=list)
    highlights: list[Highlight] = field(default_factory=list)
    favorite: bool = False
    finished: bool = False
    collections: list[str] = field(default_factory=list)
    opened: float = 0.0
    added: float = 0.0
    progress: float = 0.0
    chars_per_minute: float = 1000.0
    section_lengths: list[int] = field(default_factory=list)
    zoom: float = 0.0               # PDFs: 0 = fit width

    @classmethod
    def from_dict(cls, data: Any) -> "BookState":
        data = data if isinstance(data, dict) else {}
        s = cls()
        place = data.get("place")
        if (isinstance(place, list) and len(place) == 2
                and all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in place)):
            s.place = (place[0], place[1])
        for b in data.get("bookmarks") if isinstance(data.get("bookmarks"), list) else []:
            if isinstance(b, dict) and _int(b.get("section"), -1, -1, 10**6) >= 0:
                s.bookmarks.append(Bookmark(_int(b["section"], 0, 0, 10**6), _int(b.get("offset"), 0, 0, 10**9),
                                            _str(b.get("text")), _float(b.get("made"), 0.0, 0, 1e12)))
        for h in data.get("highlights") if isinstance(data.get("highlights"), list) else []:
            if isinstance(h, dict) and isinstance(h.get("id"), str) and _int(h.get("section"), -1, -1, 10**6) >= 0:
                start = _int(h.get("start"), 0, 0, 10**9)
                end = _int(h.get("end"), start, start, 10**9)
                s.highlights.append(Highlight(h["id"], _int(h["section"], 0, 0, 10**6), start, end,
                                              _str(h.get("text")), _str(h.get("colour"), "yellow", HIGHLIGHT_COLOURS),
                                              _str(h.get("note")), _float(h.get("made"), 0.0, 0, 1e12)))
        s.favorite = _bool(data.get("favorite"), False)
        s.finished = _bool(data.get("finished"), False)
        s.collections = _strings(data.get("collections"))
        s.opened = _float(data.get("opened"), 0.0, 0, 1e12)
        s.added = _float(data.get("added"), 0.0, 0, 1e12)
        s.progress = _float(data.get("progress"), 0.0, 0.0, 1.0)
        s.chars_per_minute = _float(data.get("chars_per_minute"), 1000.0, 100.0, 20000.0)
        lengths = data.get("section_lengths")
        if isinstance(lengths, list) and all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in lengths):
            s.section_lengths = lengths
        s.zoom = _float(data.get("zoom"), 0.0, 0.0, 8.0)
        return s

    def to_dict(self) -> dict:
        data = asdict(self)
        data["place"] = list(self.place) if self.place else None
        return data


def update_speed(state: BookState, chars: int, seconds: float) -> None:
    """Learn how fast this person reads from one page: chars read in seconds.
    Very short looks (flicking past) and long gaps (left the room) are ignored."""
    if chars <= 0 or not 2.0 <= seconds <= 300.0:
        return
    speed = min(max(chars / seconds * 60.0, 100.0), 20000.0)
    state.chars_per_minute = 0.8 * state.chars_per_minute + 0.2 * speed


# ---------------------------------------------------------------- library

@dataclass
class LibraryEntry:
    id: str
    path: str
    info: BookInfo
    size: int
    mtime: float
    added: float


def info_to_dict(info: BookInfo) -> dict:
    return {"format": info.format, "title": info.title, "authors": info.authors,
            "lock": asdict(info.lock) if info.lock else None, "error": info.error, "subjects": info.subjects}


def info_from_dict(path: str, data: dict) -> BookInfo:
    lock = data.get("lock")
    lock_info = None
    if isinstance(lock, dict) and isinstance(lock.get("message"), str):
        lock_info = LockInfo(_str(lock.get("owner")), lock["message"],
                             lock.get("free_copy_url") if isinstance(lock.get("free_copy_url"), str) else None)
    return BookInfo(path, _str(data.get("format"), "Book"), _str(data.get("title")),
                    _strings(data.get("authors")), None, lock_info,
                    data.get("error") if isinstance(data.get("error"), str) else None,
                    _strings(data.get("subjects")))


class Library:
    """Books in the chosen folders plus single books added by hand."""

    def __init__(self, state_folder: str | None = None,
                 peek: Callable[[str], BookInfo] = formats.peek,
                 thumbnail: Callable[[bytes], bytes | None] | None = None) -> None:
        self.folder = state_folder or default_state_folder()
        self.peek = peek
        self.thumbnail = thumbnail
        self.settings = Settings.load(self.folder)
        self._lock = threading.Lock()
        self._states: dict[str, BookState] = {}
        data = read_json(os.path.join(self.folder, "library.json"))
        data = data if isinstance(data, dict) else {}
        self._extra = _strings(data.get("extra_files"))
        self._removed = set(_strings(data.get("removed")))
        self._cache: dict[str, dict] = {}
        for row in data.get("books") if isinstance(data.get("books"), list) else []:
            if isinstance(row, dict) and isinstance(row.get("path"), str) and isinstance(row.get("id"), str):
                self._cache[os.path.normcase(row["path"])] = row
        self._entries: list[LibraryEntry] = self._entries_from_cache()

    # ---- the list
    def _entries_from_cache(self) -> list[LibraryEntry]:
        out, seen = [], set()
        for row in self._cache.values():
            if row["id"] in seen or row["id"] in self._removed:
                continue
            seen.add(row["id"])
            out.append(LibraryEntry(row["id"], row["path"], info_from_dict(row["path"], row),
                                    _int(row.get("size"), 0, 0, 1 << 62), _float(row.get("mtime"), 0.0, 0, 1e12),
                                    _float(row.get("added"), 0.0, 0, 1e12)))
        return out

    def entries(self) -> list[LibraryEntry]:
        with self._lock:
            return list(self._entries)

    def book_files(self) -> list[str]:
        """Every book file in the folders (hidden folders skipped) plus added files."""
        found = []
        for root in self.settings.folders:
            if not os.path.isdir(root):
                continue
            for folder, dirs, files in os.walk(root):
                dirs[:] = sorted(d for d in dirs if not d.startswith((".", "$")) and d.lower() != "system volume information")
                for name in sorted(files):
                    if os.path.splitext(name)[1].lower() in READABLE and not name.startswith(("~$", ".")):
                        found.append(os.path.join(folder, name))
        for path in self._extra:
            if os.path.isfile(path):
                found.append(path)
        return found

    def scan(self, stop: threading.Event, progress: Callable[[int], None]) -> list[LibraryEntry]:
        """Look through the folders; unchanged files come from the cache."""
        new_cache: dict[str, dict] = {}
        done = 0
        stopped = False
        for path in self.book_files():
            if stop.is_set():
                stopped = True
                break
            key = os.path.normcase(path)
            try:
                st = os.stat(path)
            except OSError:
                continue
            row = self._cache.get(key)
            # A book listed by an older Book Reader has no subject tags saved yet: read it once more.
            if not (row and row.get("size") == st.st_size and row.get("mtime") == st.st_mtime
                    and "subjects" in row):
                old = row
                try:
                    row = self._read_details(path, st, old)
                except Exception:          # one odd file must never stop the whole scan
                    row = self._damaged_row(path, st, old)
                if row is None:
                    continue
                if old and old.get("id") and old["id"] != row["id"]:
                    self._carry_state(old["id"], row["id"])
            new_cache[key] = row
            done += 1
            progress(done)
        if stopped:                       # keep what wasn't looked at yet
            for key, row in self._cache.items():
                new_cache.setdefault(key, row)
        with self._lock:
            self._cache = new_cache
            self._entries = self._entries_from_cache()
            self._save_library()
            return list(self._entries)

    def _read_details(self, path: str, st: os.stat_result, old: dict | None) -> dict | None:
        try:
            ident = book_id(path)
        except OSError:
            return None
        info = self.peek(path)
        row = {"id": ident, "path": path, "size": st.st_size, "mtime": st.st_mtime,
               "added": (old or {}).get("added") or time.time(), **info_to_dict(info)}
        if info.cover and ident not in self._removed:
            cover = self.thumbnail(info.cover) if self.thumbnail else info.cover
            if cover:
                target = self._cover_path(ident)
                temp = target + ".part"
                try:                      # written whole, then swapped in: never half a picture
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    with open(temp, "wb") as f:
                        f.write(cover)
                    os.replace(temp, target)
                except OSError:
                    try:
                        os.remove(temp)
                    except OSError:
                        pass              # no new cover is fine; the old or a drawn one is used
        return row

    def _damaged_row(self, path: str, st: os.stat_result, old: dict | None) -> dict | None:
        try:
            ident = book_id(path)
        except OSError:
            return None
        title, authors = formats.title_from_file_name(path)
        info = BookInfo(path, "Book", title, authors, error="This file couldn't be read; it may be damaged.")
        return {"id": ident, "path": path, "size": st.st_size, "mtime": st.st_mtime,
                "added": (old or {}).get("added") or time.time(), **info_to_dict(info)}

    def _carry_state(self, old_id: str, new_id: str) -> None:
        """The same file with new contents (its details edited, say) keeps its place and notes."""
        old_file = os.path.join(self.folder, "books", old_id + ".json")
        new_file = os.path.join(self.folder, "books", new_id + ".json")
        if os.path.exists(new_file) or not os.path.exists(old_file):
            return
        data = read_json(old_file)
        if isinstance(data, dict):
            try:
                write_json(new_file, data)
            except OSError:
                pass

    def _save_library(self) -> None:
        write_json(os.path.join(self.folder, "library.json"), {
            "version": 1, "books": list(self._cache.values()),
            "extra_files": self._extra, "removed": sorted(self._removed)})

    def _cover_path(self, ident: str) -> str:
        return os.path.join(self.folder, "covers", ident + ".img")

    def cover(self, ident: str) -> bytes | None:
        try:
            with open(self._cover_path(ident), "rb") as f:
                return f.read()
        except OSError:
            return None

    def add_files(self, paths: list[str]) -> None:
        """Remember single book files that aren't inside the library folders."""
        with self._lock:
            for path in paths:
                path = os.path.abspath(path)
                inside = any(os.path.normcase(path).startswith(os.path.normcase(os.path.join(f, "")))
                             for f in self.settings.folders)
                if not inside and path not in self._extra:
                    self._extra.append(path)
                try:
                    self._removed.discard(book_id(path))
                except OSError:
                    pass
            self._save_library()

    def remove(self, ident: str) -> None:
        """Take a book off the list (the file itself is left alone)."""
        with self._lock:
            self._removed.add(ident)
            self._extra = [p for p in self._extra
                           if not any(r["id"] == ident and os.path.normcase(r["path"]) == os.path.normcase(p)
                                      for r in self._cache.values())]
            self._entries = [e for e in self._entries if e.id != ident]
            self._save_library()

    def entry_for_path(self, path: str) -> LibraryEntry | None:
        key = os.path.normcase(os.path.abspath(path))
        with self._lock:
            for e in self._entries:
                if os.path.normcase(e.path) == key:
                    return e
        return None

    # ---- per-book state
    def state(self, ident: str) -> BookState:
        with self._lock:
            if ident not in self._states:
                self._states[ident] = BookState.from_dict(
                    read_json(os.path.join(self.folder, "books", ident + ".json")))
            return self._states[ident]

    def save_state(self, ident: str, state: BookState) -> None:
        with self._lock:
            self._states[ident] = state
            write_json(os.path.join(self.folder, "books", ident + ".json"), state.to_dict())

    def all_states(self) -> dict[str, BookState]:
        return {e.id: self.state(e.id) for e in self.entries()}

    def save_settings(self) -> None:
        self.settings.save(self.folder)

    def notes_text(self, ident: str, title: str, section_names: list[str] | None = None) -> str:
        """Bookmarks and highlights (with notes) as plain text, in book order."""
        state = self.state(ident)
        names = section_names or []

        def where(section: int) -> str:
            name = names[section] if section < len(names) and names[section] else f"Part {section + 1}"
            return name

        rows = [((b.section, b.offset), f"Bookmark ({where(b.section)}): {b.text}".rstrip(": "))
                for b in state.bookmarks]
        for h in state.highlights:
            line = f"Highlight, {h.colour} ({where(h.section)}): “{h.text}”"
            if h.note:
                line += f"\n    Note: {h.note}"
            rows.append(((h.section, h.start), line))
        rows.sort(key=lambda r: r[0])
        lines = [f"Notes from “{title}”", time.strftime("Saved %d %B %Y"), ""]
        lines += [text for _key, text in rows] or ["(no bookmarks or highlights yet)"]
        return "\n".join(lines) + "\n"
