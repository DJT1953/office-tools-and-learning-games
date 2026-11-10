"""The page view for reflowable books (EPUB, Kindle, text, web pages).

ReaderView shows one chapter at a time in Qt WebEngine. The page script
(book_reader_page.js) lays it out as pages and reports the place; this class
turns pages across chapters, follows links (with Back), and passes on what the
reader selects or clicks.
"""

from __future__ import annotations

import json
import os

from PySide6.QtCore import QFile, QIODevice, QObject, Qt, QTimer, QUrl, Signal, Slot
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineScript
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QVBoxLayout, QWidget

import book_reader_web as web

HERE = os.path.dirname(os.path.abspath(__file__))
MARGIN_PX = {"small": 20, "medium": 44, "large": 80}
_SCRIPT: str | None = None


def page_script_source() -> str:
    """qwebchannel.js (from Qt) followed by the reader's page script."""
    global _SCRIPT
    if _SCRIPT is None:
        f = QFile(":/qtwebchannel/qwebchannel.js")
        if not f.open(QIODevice.OpenModeFlag.ReadOnly):
            raise RuntimeError("Qt's qwebchannel.js is missing")
        channel_js = bytes(f.readAll()).decode("utf-8")
        f.close()
        with open(os.path.join(HERE, "book_reader_page.js"), encoding="utf-8") as fh:
            _SCRIPT = channel_js + "\n" + fh.read()
    return _SCRIPT


def style_for(settings) -> dict:
    return {"font": settings.font, "size": settings.font_px, "line": settings.line_spacing,
            "margin": MARGIN_PX.get(settings.margins, 44), "colours": settings.colours,
            "layout": settings.layout, "justify": settings.justify}


def book_url(href: str) -> QUrl:
    url = QUrl()
    url.setScheme("book")
    url.setHost("b")
    url.setPath("/" + href)
    return url


class Bridge(QObject):
    """The object the page script talks to."""

    message = Signal(dict)

    @Slot(str)
    def event(self, text: str) -> None:
        try:
            data = json.loads(text)
        except ValueError:
            return
        if isinstance(data, dict):
            self.message.emit(data)


class BookPage(QWebEnginePage):
    """Keeps link clicks for the reader to handle; never opens new windows."""

    link_url = Signal(QUrl)

    def acceptNavigationRequest(self, url: QUrl, nav_type, is_main_frame: bool) -> bool:  # noqa: N802
        if nav_type == QWebEnginePage.NavigationType.NavigationTypeLinkClicked:
            self.link_url.emit(QUrl(url))
            return False
        return url.scheme() in ("book", "about")

    def createWindow(self, _kind):  # noqa: N802
        return None

    def javaScriptConsoleMessage(self, level, message, line, source):  # noqa: N802
        if os.environ.get("BOOK_READER_DEBUG"):
            print(f"[page {line}] {message}")


class ReaderView(QWidget):
    place_changed = Signal(int, int, int, int)      # section, offset, page, pages in section
    selected = Signal(dict)                          # start, end, text, x, y (page coordinates)
    tapped = Signal()
    highlight_clicked = Signal(str, int, int)
    web_link_clicked = Signal(str)
    image_clicked = Signal(str)                      # path inside the book
    section_changed = Signal(int)
    load_failed = Signal(str)
    edge_reached = Signal(str)                       # "end" when Next is pressed on the last page
    ready = Signal()                                 # a chapter is laid out and placed

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.book = None
        self.section = 0
        self.offset = 0
        self.page = 0
        self.pages = 1
        self.total = 0
        self.at_book_end = False
        self.history: list[tuple[int, int]] = []
        self.highlights: list[dict] = []
        self._ready = False
        self._target: dict = {}
        self._generation = 0
        self._style: dict = {}
        self._sent_style: dict = {}
        self._view = QWebEngineView(self)
        self._profile = web.make_profile(self, self._resolve)
        self._page = BookPage(self._profile, self._view)
        self._page.link_url.connect(self._on_link_url)
        self._view.setPage(self._page)
        self._view.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
        self._bridge = Bridge(self)
        self._bridge.message.connect(self._on_message)
        self._channel = QWebChannel(self._page)
        self._channel.registerObject("bridge", self._bridge)
        self._page.setWebChannel(self._channel, web.APP_WORLD)
        script = QWebEngineScript()
        script.setName("book-reader")
        script.setWorldId(web.APP_WORLD)
        script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentReady)
        script.setRunsOnSubFrames(False)
        script.setSourceCode(page_script_source())
        self._page.scripts().insert(script)
        self._page.loadFinished.connect(self._on_loaded)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._view)

    # ---------------------------------------------------------------- for others
    def web_page(self) -> QWebEnginePage:
        return self._page

    def web_view(self) -> QWebEngineView:
        return self._view

    def is_ready(self) -> bool:
        return self._ready

    def place(self) -> tuple[int, int]:
        return self.section, self.offset

    def run_js(self, code: str, callback=None) -> None:
        """Run code in the page's reader world; nothing happens before the page script is there."""
        code = "if (window.reader) { %s }" % code
        if callback is None:
            self._page.runJavaScript(code, web.APP_WORLD)
        else:
            self._page.runJavaScript(code, web.APP_WORLD, callback)

    # ---------------------------------------------------------------- opening and moving
    def open(self, book, state, settings) -> None:
        self.book = book
        self._style = style_for(settings)
        self.history.clear()
        self.highlights = [h.__dict__.copy() for h in getattr(state, "highlights", [])]
        section, offset = state.place or (0, 0)
        if not 0 <= section < len(book.sections):
            section, offset = 0, 0
        self._load(section, {"offset": offset})

    def _load(self, section: int, target: dict) -> None:
        self._generation += 1
        self._ready = False
        self.at_book_end = False
        self.section = section
        self.offset = target.get("offset", 0) if isinstance(target.get("offset"), int) else 0
        self.page = 0
        self._target = target
        self.section_changed.emit(section)
        self._page.load(book_url(self.book.sections[section].href))

    def _on_loaded(self, ok: bool) -> None:
        if not ok:
            # A load replaced by a newer one also ends "not ok": complain only if nothing came of it.
            QTimer.singleShot(2000, lambda g=self._generation: self._still_not_ready(g))
            return
        self._send_start()

    def _send_start(self) -> None:
        """Start the page script for this load (harmless if it was started already)."""
        self._sent_style = dict(self._style)
        self.run_js("reader.setHighlights(%s); reader.start(%s, %s, %d)" % (
            json.dumps(self._marks_here()), json.dumps(self._style), json.dumps(self._target), self._generation))

    def _still_not_ready(self, generation: int) -> None:
        if generation == self._generation and not self._ready:
            self.load_failed.emit("This part of the book couldn't be shown.")

    def _marks_here(self) -> list[dict]:
        return [{"id": h["id"], "start": h["start"], "end": h["end"], "colour": h["colour"]}
                for h in self.highlights if h.get("section") == self.section]

    def set_highlights(self, highlights: list) -> None:
        """All the book's highlights (Highlight objects or dicts); this chapter's are painted."""
        self.highlights = [h if isinstance(h, dict) else h.__dict__.copy() for h in highlights]
        self.run_js("reader.setHighlights(%s)" % json.dumps(self._marks_here()))

    def go_to_find(self, section: int, query: str, occurrence: int) -> None:
        """Go to the occurrence-th place query is found in section, and mark it."""
        if section == self.section and self._ready:
            self.run_js("reader.find(%s, %d)" % (json.dumps(query), occurrence))
        else:
            self._load(section, {"find": [query, occurrence]})

    def page_info(self, callback) -> None:
        """callback({start, end, text}) for the page on screen."""
        def got(text):
            try:
                data = json.loads(text) if isinstance(text, str) else None
            except ValueError:
                data = None
            if isinstance(data, dict):
                callback(data)
        self.run_js("JSON.stringify(reader.pageInfo())", got)

    def clear_selection(self) -> None:
        self.run_js("reader.clearSelection()")

    def section_count(self) -> int:
        return len(self.book.sections) if self.book else 0

    def sentences_from(self, offset: int, callback) -> None:
        """callback([{start, end, text}]) for the sentences from offset to the chapter's end."""
        def got(text):
            try:
                data = json.loads(text) if isinstance(text, str) and text else []
            except ValueError:
                data = []
            callback(data if isinstance(data, list) else [])
        self.run_js("JSON.stringify(reader.sentencesFrom(%d))" % max(0, offset), got)

    def mark_sentence(self, start: int, end: int) -> None:
        self.run_js("reader.markSentence(%d, %d)" % (start, end))

    def clear_speech(self) -> None:
        self.run_js("reader.clearSpeech()")

    def next_page(self) -> None:
        if self._ready:
            self.run_js("reader.next()")

    def prev_page(self) -> None:
        if self._ready:
            self.run_js("reader.prev()")

    def go_to(self, section: int, offset: int = 0, remember: bool = False) -> None:
        if self.book is None or not 0 <= section < len(self.book.sections):
            return
        if remember:
            self.history.append(self.place())
        if section == self.section and self._ready:
            self.run_js("reader.goToOffset(%d)" % max(0, offset))
        else:
            self._load(section, {"offset": max(0, offset)})

    def go_to_href(self, href: str, remember: bool = True) -> bool:
        target = self.book.find_href(href) if self.book else None
        if target is None:
            return False
        section, anchor = target
        if remember:
            self.history.append(self.place())
        if section == self.section and self._ready:
            self.run_js("reader.goToAnchor(%s)" % json.dumps(anchor) if anchor else "reader.goToStart()")
        else:
            self._load(section, {"anchor": anchor} if anchor else {"offset": 0})
        return True

    def go_back(self) -> bool:
        if not self.history:
            return False
        section, offset = self.history.pop()
        self.go_to(section, offset)
        return True

    def apply_settings(self, settings) -> None:
        """New looks; if the page is still busy, they are sent as soon as it is free."""
        self._style = style_for(settings)
        if self._ready:
            self._send_style()

    def _send_style(self) -> None:
        self._ready = False
        self._sent_style = dict(self._style)
        self.run_js("reader.setStyle(%s)" % json.dumps(self._style))

    # ---------------------------------------------------------------- from the page
    def _resolve(self, path: str):
        if self.book is None or not path.startswith("b/"):
            return None
        return self.book.resource(path[2:])

    def _href_of(self, url: QUrl) -> str | None:
        if url.scheme() != "book" or url.host() != "b":
            return None
        href = web.request_path(url)[2:]
        if url.hasFragment():
            href += "#" + url.fragment(QUrl.ComponentFormattingOption.FullyDecoded)
        return href

    def _on_link_url(self, url: QUrl) -> None:
        href = self._href_of(url)
        if href is not None:
            self.go_to_href(href)
        elif url.scheme() in ("http", "https", "mailto", "ftp"):
            self.web_link_clicked.emit(url.toString())

    def _on_message(self, data: dict) -> None:
        kind = data.get("type")
        if kind == "hello":                          # a page's script has just arrived
            if self.book is not None and web.request_path(QUrl(str(data.get("href", "")))) == \
                    "b/" + self.book.sections[self.section].href:
                self._send_start()
            return
        if data.get("gen") != self._generation:      # from a page that has been left
            return
        if kind == "busy":
            self._ready = False
        elif kind == "ready":
            self.total = int(data.get("total") or 0)
            self._ready = True
            if self._style != self._sent_style:      # changed while the page was busy
                self._send_style()
            else:
                self.ready.emit()
        elif kind == "place":
            self.offset = int(data.get("offset") or 0)
            self.page = int(data.get("page") or 0)
            self.pages = max(1, int(data.get("pages") or 1))
            self.total = int(data.get("total") or self.total)
            self.place_changed.emit(self.section, self.offset, self.page, self.pages)
        elif kind == "edge":
            if data.get("dir") == "next":
                if self.book and self.section + 1 < len(self.book.sections):
                    self._load(self.section + 1, {"offset": 0})
                else:
                    self.at_book_end = True
                    self.edge_reached.emit("end")
            elif data.get("dir") == "prev" and self.section > 0:
                self._load(self.section - 1, {"end": True})
            elif data.get("dir") == "prev":
                self.edge_reached.emit("start")
        elif kind == "link":
            self._on_link_url(QUrl(str(data.get("href", ""))))
        elif kind == "image":
            href = self._href_of(QUrl(str(data.get("src", ""))))
            if href:
                self.image_clicked.emit(href)
        elif kind == "selected":
            self.selected.emit(data)
        elif kind == "tap":
            self.tapped.emit()
        elif kind == "highlight":
            self.highlight_clicked.emit(str(data.get("id", "")), int(data.get("x", 0)), int(data.get("y", 0)))
