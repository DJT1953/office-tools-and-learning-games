"""The page view for PDFs, using Qt's PDF engine.

PdfView behaves like ReaderView where the reading page needs it to: a "section"
is one PDF page, places are (page, 0), and it can give its page text for
search, bookmarks and reading aloud. PDF pages are fixed, so there are no
highlights or colour changes; there is zoom instead.
"""

from __future__ import annotations

import re

from PySide6.QtCore import QEvent, QObject, QPointF, Qt, QTimer, Signal
from PySide6.QtPdf import QPdfBookmarkModel, QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import QVBoxLayout, QWidget

import book_reader_formats as formats

SENTENCE_END = re.compile(r"[.!?…]+[\"'”’)\]]*(?=\s|$)")


def split_sentences(text: str) -> list[dict]:
    """[{start, end, text}] for the sentences of a page's text."""
    out: list[dict] = []
    pos = 0
    for m in list(SENTENCE_END.finditer(text)) + [None]:
        end = m.end() if m else len(text)
        piece = text[pos:end]
        words = " ".join(piece.split())
        if words and re.search(r"\w", words):
            lead = len(piece) - len(piece.lstrip())
            out.append({"start": pos + lead, "end": end, "text": words})
        pos = end
    return out


class PdfBook:
    """The library-and-panels side of an open PDF: pages as sections, outline as contents."""

    def __init__(self, path: str, document: QPdfDocument, toc: list[formats.TocEntry]) -> None:
        self.path = path
        self.format = "PDF"
        self.document = document
        title = document.metaData(QPdfDocument.MetaDataField.Title)
        self.title = str(title).strip() if title else ""
        self.title = self.title or formats.title_from_file_name(path)[0]
        author = document.metaData(QPdfDocument.MetaDataField.Author)
        self.authors = [str(author).strip()] if author and str(author).strip() else []
        self.cover = None
        self.sections = [formats.Section(f"page{n + 1}", f"Page {n + 1}") for n in range(document.pageCount())]
        self.toc = toc
        self._texts: dict[int, str] = {}

    def section_text(self, i: int) -> str:
        if i not in self._texts:
            self._texts[i] = self.document.getAllText(i).text() if 0 <= i < len(self.sections) else ""
        return self._texts[i]

    def close(self) -> None:
        pass


class PdfView(QWidget):
    place_changed = Signal(int, int, int, int)
    ready = Signal()
    edge_reached = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.document = QPdfDocument(self)
        self.view = QPdfView(self)
        self.view.setDocument(self.document)
        self.view.setPageMode(QPdfView.PageMode.SinglePage)
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.navigator = self.view.pageNavigator()
        self.navigator.currentPageChanged.connect(self._page_changed)
        self.outline = QPdfBookmarkModel(self)
        self.outline.setDocument(self.document)
        self.book: PdfBook | None = None
        self.path = ""
        self.page = self.section = 0
        self.offset = 0
        self.pages = 0
        self.total = 0
        self.at_book_end = False
        self.history: list[tuple[int, int]] = []
        self._ready = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.view)
        self.view.viewport().installEventFilter(self)

    # ---------------------------------------------------------------- opening
    def open(self, path: str, state, ask_password) -> str | None:
        """Open the PDF; returns a plain message if it can't be opened."""
        self.close()
        error = self.document.load(path)
        tries = 0
        while error == QPdfDocument.Error.IncorrectPassword:
            if tries >= 3:
                return "That password isn't right, so the PDF can't be opened."
            password = ask_password(tries)
            if password is None:
                return "The PDF wasn't opened (it needs its password)."
            tries += 1
            self.document.setPassword(password)
            error = self.document.load(path)
        if error != QPdfDocument.Error.None_ or self.document.pageCount() < 1:
            reason = {QPdfDocument.Error.FileNotFound: "the file isn't there any more",
                      QPdfDocument.Error.InvalidFileFormat: "it is damaged or isn't really a PDF",
                      QPdfDocument.Error.UnsupportedSecurityScheme: "it is locked in a way this reader can't open"}
            return f"This PDF can't be opened ({reason.get(error, 'it is damaged')})."
        self.path = path
        self.pages = self.document.pageCount()
        self.book = PdfBook(path, self.document, self._contents())
        page = state.place[0] if state.place and 0 <= state.place[0] < self.pages else 0
        if state.zoom > 0:
            self.view.setZoomMode(QPdfView.ZoomMode.Custom)
            self.view.setZoomFactor(state.zoom)
        else:
            self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.history.clear()
        self._ready = True
        self.navigator.jump(page, QPointF(0, 0), 0)
        self.page = self.section = page
        QTimer.singleShot(0, self._announce)
        return None

    def _announce(self) -> None:
        self.at_book_end = self.page == self.pages - 1
        self.place_changed.emit(self.page, 0, self.page, self.pages)
        self.ready.emit()

    def _contents(self) -> list[formats.TocEntry]:
        found: list[formats.TocEntry] = []
        model = self.outline

        def walk(parent, depth: int) -> None:
            for row in range(model.rowCount(parent)):
                index = model.index(row, 0, parent)
                title = str(model.data(index, QPdfBookmarkModel.Role.Title) or "").strip()
                page = model.data(index, QPdfBookmarkModel.Role.Page)
                if title and isinstance(page, int) and 0 <= page < self.document.pageCount():
                    found.append(formats.TocEntry(title, page, "", depth))
                walk(index, depth + 1)
        from PySide6.QtCore import QModelIndex
        walk(QModelIndex(), 0)
        return found

    def close(self) -> None:
        self._ready = False
        self.book = None
        self.document.close()

    def is_ready(self) -> bool:
        return self._ready

    def place(self) -> tuple[int, int]:
        return self.page, 0

    def section_count(self) -> int:
        return self.pages

    # ---------------------------------------------------------------- moving
    def _page_changed(self, page: int) -> None:
        if not self._ready:
            return
        self.page = self.section = page
        self.at_book_end = page == self.pages - 1
        self.place_changed.emit(page, 0, page, self.pages)

    def go_to(self, section: int, offset: int = 0, remember: bool = False) -> None:
        if not self._ready or not 0 <= section < self.pages:
            return
        if remember:
            self.history.append(self.place())
        self.navigator.jump(section, QPointF(0, 0), 0)
        QTimer.singleShot(0, self.ready.emit)

    def next_page(self) -> None:
        if self.page + 1 < self.pages:
            self.go_to(self.page + 1)
        else:
            self.at_book_end = True
            self.edge_reached.emit("end")

    def prev_page(self) -> None:
        if self.page > 0:
            self.go_to(self.page - 1)

    def go_back(self) -> bool:
        if not self.history:
            return False
        page, _ = self.history.pop()
        self.go_to(page)
        return True

    def go_to_href(self, href: str, remember: bool = True) -> bool:
        m = re.fullmatch(r"page(\d+)(#.*)?", href)
        if not m:
            return False
        self.go_to(int(m.group(1)) - 1, 0, remember)
        return True

    def go_to_find(self, section: int, query: str, occurrence: int) -> None:
        self.go_to(section)

    # ---------------------------------------------------------------- zoom
    def zoom(self) -> float:
        return self.view.zoomFactor()

    def zoom_mode(self) -> str:
        mode = self.view.zoomMode()
        return {QPdfView.ZoomMode.FitToWidth: "width", QPdfView.ZoomMode.FitInView: "page"}.get(mode, "custom")

    def zoom_by(self, factor: float) -> None:
        current = self.view.zoomFactor()
        if self.view.zoomMode() != QPdfView.ZoomMode.Custom:
            width = self.view.viewport().width()
            page_width = self.document.pagePointSize(self.page).width() * self.logicalDpiX() / 72.0
            current = width / page_width if page_width > 0 else 1.0
        self.view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.view.setZoomFactor(max(0.25, min(8.0, current * factor)))

    def fit_width(self) -> None:
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)

    def fit_page(self) -> None:
        self.view.setZoomMode(QPdfView.ZoomMode.FitInView)

    def set_continuous(self, on: bool) -> None:
        self.view.setPageMode(QPdfView.PageMode.MultiPage if on else QPdfView.PageMode.SinglePage)
        if self._ready:
            self.navigator.jump(self.page, QPointF(0, 0), 0)

    def zoom_for_state(self) -> float:
        return self.view.zoomFactor() if self.view.zoomMode() == QPdfView.ZoomMode.Custom else 0.0

    # ---------------------------------------------------------------- for the reading page
    def page_info(self, callback) -> None:
        text = self.book.section_text(self.page) if self.book else ""
        QTimer.singleShot(0, lambda: callback({"start": 0, "end": 1, "text": " ".join(text.split())[:160]}))

    def sentences_from(self, offset: int, callback) -> None:
        text = self.book.section_text(self.page) if self.book else ""
        sentences = [s for s in split_sentences(text) if s["end"] > offset]
        QTimer.singleShot(0, lambda: callback(sentences))

    def mark_sentence(self, start: int, end: int) -> None:
        pass

    def clear_speech(self) -> None:
        pass

    def set_highlights(self, _highlights) -> None:
        pass

    def clear_selection(self) -> None:
        pass

    def apply_settings(self, _settings) -> None:
        pass

    # ---------------------------------------------------------------- clicks, keys, wheel
    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        kind = event.type()
        if kind == QEvent.Type.MouseButtonRelease and event.button() == Qt.MouseButton.LeftButton:
            x, width = event.position().x(), max(1, self.view.viewport().width())
            if x < width * 0.2:
                self.prev_page()
                return True
            if x > width * 0.8:
                self.next_page()
                return True
        elif kind == QEvent.Type.KeyPress:
            key = event.key()
            if key in (Qt.Key.Key_Right, Qt.Key.Key_PageDown, Qt.Key.Key_Space):
                self.next_page()
                return True
            if key in (Qt.Key.Key_Left, Qt.Key.Key_PageUp):
                self.prev_page()
                return True
        elif kind == QEvent.Type.Wheel and self.view.pageMode() == QPdfView.PageMode.SinglePage:
            bar = self.view.verticalScrollBar()
            down = event.angleDelta().y() < 0
            if down and bar.value() >= bar.maximum():
                self.next_page()
                return True
            if not down and bar.value() <= bar.minimum() and self.page > 0:
                self.prev_page()
                return True
        return False
