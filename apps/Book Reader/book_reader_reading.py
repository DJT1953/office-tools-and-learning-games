"""The reading page: the book, big buttons above and below, and side lists.

Top row: Library, Contents, Search, Bookmark, Notes, Text, Read aloud, Full
screen, Save a copy. Below the page: Previous page / Next page, a slider for
the whole book and a line saying where you are. Picked words get a bar of big
buttons (highlight colours, note, copy, look up, read aloud from here).
"""

from __future__ import annotations

import os
import re
import threading
import time
import uuid
from typing import NamedTuple

from PySide6.QtCore import QEventLoop, QObject, Qt, Signal
from PySide6.QtGui import QFont, QKeySequence, QPixmap, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSlider,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

import book_reader_formats as formats
import book_reader_library as lib
import book_reader_speech as speech
import book_reader_ui as ui
from book_reader_pdf import PdfView
from book_reader_view import ReaderView

PDF_TEXT_NOTE = ("A PDF's pages are fixed, so its words can't be made bigger or reco"
                 'loured here. Use Zoom in and Zoom out, under the page, to make them bigger.')
COLOUR_BUTTONS = {"Yellow": "#fff176", "Green": "#b9f6ca", "Blue": "#b3e5fc", "Pink": "#f8bbd0"}


class SearchHit(NamedTuple):
    section: int
    occurrence: int
    snippet: str


def search_book(book, query: str, stop: threading.Event | None = None, limit: int = 500) -> list[SearchHit]:
    """Every place query is found (any capitals), with a few words around it."""
    q = query.lower().strip()
    hits: list[SearchHit] = []
    if not q:
        return hits
    for i in range(len(book.sections)):
        if stop is not None and stop.is_set():
            break
        text = book.section_text(i)
        low = text.lower()
        n, at = 0, low.find(q)
        while at >= 0:
            before = " ".join(text[max(0, at - 70):at].split()[1:] if at > 70 else text[:at].split())
            after = " ".join(text[at + len(q):at + len(q) + 70].split()[:-1])
            snippet = ("… " if at > 70 else "") + before + (" " if before else "") + \
                " ".join(text[at:at + len(q)].split()) + " " + after + " …"
            hits.append(SearchHit(i, n, snippet.strip()))
            if len(hits) >= limit:
                return hits
            n += 1
            at = low.find(q, at + 1)
    return hits


def plural(n: int, word: str) -> str:
    return f"{n:,} {word}" if n == 1 else f"{n:,} {word}s"


def big_list() -> QListWidget:
    lst = QListWidget()
    font = QFont(lst.font())
    font.setPointSizeF(ui.words(1.0))
    lst.setFont(font)
    lst.setWordWrap(True)
    lst.setStyleSheet(f"QListWidget {{ font-size: {ui.words(1.0):g}pt; }}"
                      " QListWidget::item { padding: 8px 6px; border-bottom: 1px solid #cccccc; }"
                      " QListWidget::item:selected { background: #cfe8ff; color: black; }")
    lst.setVerticalScrollMode(QListWidget.ScrollMode.ScrollPerPixel)
    return lst


# ---------------------------------------------------------------- side panel

class SidePanel(QFrame):
    """Contents, Search, Notes and Text settings, one at a time, beside the page."""

    go = Signal(int, str)                 # section, anchor
    go_offset = Signal(int, int)          # section, offset
    go_find = Signal(int, str, int)       # section, query, occurrence
    search_requested = Signal(str)
    save_notes_requested = Signal()
    text_changed = Signal()
    closed = Signal()

    def __init__(self, settings: lib.Settings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setMinimumWidth(380)
        self.book = None
        self.kind = ""
        self._toc_sections: list[tuple[int, str]] = []
        self._current = -1
        self._hits: list[SearchHit] = []
        self._notes: list[tuple[int, int]] = []
        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        self.title = QLabel("")
        self.title.setStyleSheet(f"font-weight: bold; font-size: {ui.words(1.1):g}pt;")
        top.addWidget(self.title, 1)
        self.close_button = ui.wide_button("Close", self.hide_panel, 52, "Hide this list")
        top.addWidget(self.close_button)
        layout.addLayout(top)
        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

        # contents
        self.contents_list = big_list()
        self.contents_list.itemClicked.connect(lambda item: self.activate_contents(self.contents_list.row(item)))
        self.contents_message = QLabel("")
        self.contents_message.setWordWrap(True)
        page = QWidget()
        box = QVBoxLayout(page)
        box.setContentsMargins(0, 0, 0, 0)
        box.addWidget(self.contents_message)
        box.addWidget(self.contents_list, 1)
        self.stack.addWidget(page)

        # search
        page = QWidget()
        box = QVBoxLayout(page)
        box.setContentsMargins(0, 0, 0, 0)
        row = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Words to find")
        self.search_box.setMinimumHeight(52)
        self.search_box.returnPressed.connect(self.start_search)
        row.addWidget(self.search_box, 1)
        row.addWidget(ui.wide_button("Find", self.start_search, 52, "Find these words in the book"))
        box.addLayout(row)
        self.search_status = QLabel("Type some words and press Find.")
        self.search_status.setWordWrap(True)
        box.addWidget(self.search_status)
        self.search_list = big_list()
        self.search_list.itemClicked.connect(lambda item: self.activate_search(self.search_list.row(item)))
        box.addWidget(self.search_list, 1)
        self.stack.addWidget(page)

        # notes
        page = QWidget()
        box = QVBoxLayout(page)
        box.setContentsMargins(0, 0, 0, 0)
        self.notes_message = QLabel("")
        self.notes_message.setWordWrap(True)
        box.addWidget(self.notes_message)
        self.notes_list = big_list()
        self.notes_list.itemClicked.connect(lambda item: self.activate_note(self.notes_list.row(item)))
        box.addWidget(self.notes_list, 1)
        box.addWidget(ui.wide_button("Save notes to a file", self.save_notes_requested.emit, 56,
                                     "Write your bookmarks, highlights and notes into a text file"))
        self.stack.addWidget(page)

        # text settings
        self.text_buttons: dict[str, QPushButton] = {}
        page = QWidget()
        grid = QGridLayout(page)
        grid.setContentsMargins(0, 0, 0, 0)
        row = 0

        def label(text: str) -> None:
            nonlocal row
            caption = QLabel(text)
            caption.setStyleSheet("font-weight: bold;")
            grid.addWidget(caption, row, 0, 1, 3)
            row += 1

        def buttons(names: list[str], slot) -> None:
            nonlocal row
            for col, name in enumerate(names):
                b = ui.wide_button(name, lambda _checked=False, n=name: slot(n), 52)
                b.setCheckable(True)
                self.text_buttons[name] = b
                grid.addWidget(b, row, col)
            row += 1

        label("Text size")
        size_row = QHBoxLayout()
        for name in ("Smaller", "Bigger"):
            b = ui.wide_button(name, lambda _c=False, n=name: self.change_size(n), 56)
            self.text_buttons[name] = b
            size_row.addWidget(b)
        self.size_label = QLabel("")
        size_row.insertWidget(1, self.size_label)
        holder = QWidget()
        holder.setLayout(size_row)
        grid.addWidget(holder, row, 0, 1, 3)
        row += 1
        label("Font")
        self.font_box = QComboBox()
        self.font_box.addItems(lib.FONTS)
        self.font_box.setMinimumHeight(52)
        self.font_box.currentTextChanged.connect(self.change_font)
        grid.addWidget(self.font_box, row, 0, 1, 3)
        row += 1
        label("Colours")
        buttons(["White", "Sepia", "Night"], self.change_colours)
        label("Pages")
        buttons(["One page", "Two pages", "Scroll"], self.change_layout)
        label("Space between lines")
        buttons(["Close", "Normal", "Wide"], self.change_spacing)
        label("Margins")
        buttons(["Small margins", "Medium margins", "Large margins"], self.change_margins)
        self.justify_box = QCheckBox("Straight right-hand edge")
        self.justify_box.toggled.connect(self.change_justify)
        grid.addWidget(self.justify_box, row, 0, 1, 3)
        row += 1
        grid.setRowStretch(row, 1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(page)
        self.stack.addWidget(scroll)
        self.refresh_text()
        self.hide()

    # ---- general
    def hide_panel(self) -> None:
        self.hide()
        self.closed.emit()

    def set_book(self, book) -> None:
        self.book = book
        self._hits = []
        self.search_list.clear()
        self.search_status.setText("Type some words and press Find.")
        self.contents_list.clear()
        self._toc_sections = []
        toc = book.toc if book is not None else []
        for entry in toc:
            item = QListWidgetItem(("      " * min(entry.depth, 4)) + entry.title)
            self.contents_list.addItem(item)
            self._toc_sections.append((entry.section, entry.anchor))
        self.contents_message.setText("" if toc else
                                      "This book has no list of chapters. Use the slider under the page "
                                      "to move through it, or Search to find words.")
        self._current = -1

    def show_kind(self, kind: str) -> None:
        self.kind = kind
        index = {"contents": 0, "search": 1, "notes": 2, "text": 3}[kind]
        self.title.setText({"contents": "Contents", "search": "Search the book", "notes": "Bookmarks and notes",
                            "text": "Text and colours"}[kind])
        self.stack.setCurrentIndex(index)
        self.show()
        if kind == "search":
            self.search_box.setFocus()
            self.search_box.selectAll()
        if kind == "text":
            self.refresh_text()

    # ---- contents
    def mark_current(self, section: int) -> None:
        current = -1
        for i, (s, _a) in enumerate(self._toc_sections):
            if s <= section and (current < 0 or s > self._toc_sections[current][0]):
                current = i
        if current == self._current:
            return
        for i in range(self.contents_list.count()):
            item = self.contents_list.item(i)
            text = item.text().removeprefix("▶ ")
            font = item.font()
            font.setBold(i == current)
            item.setFont(font)
            item.setText(("▶ " + text) if i == current else text)
        self._current = current
        if current >= 0:
            self.contents_list.scrollToItem(self.contents_list.item(current))

    def contents_rows(self) -> list[tuple[str, bool]]:
        return [(self.contents_list.item(i).text().removeprefix("▶ ").strip(), i == self._current)
                for i in range(self.contents_list.count())]

    def activate_contents(self, row: int) -> None:
        if 0 <= row < len(self._toc_sections):
            self.go.emit(*self._toc_sections[row])

    # ---- search
    def start_search(self) -> None:
        text = self.search_box.text().strip()
        if not text:
            self.search_status.setText("Type some words first, then press Find.")
            return
        self.search_status.setText(f"Looking for “{text}”…")
        self.search_list.clear()
        self._hits = []
        self.search_requested.emit(text)

    def set_results(self, query: str, hits: list[SearchHit], names: list[str]) -> None:
        self._hits = hits
        self._query = query
        self.search_list.clear()
        for hit in hits:
            where = names[hit.section] if hit.section < len(names) else f"Part {hit.section + 1}"
            self.search_list.addItem(f"{where}:  {hit.snippet}")
        if not hits:
            self.search_status.setText(f"No places in this book have “{query}”.")
        else:
            more = " (the first 500)" if len(hits) >= 500 else ""
            self.search_status.setText(f"Found in {plural(len(hits), 'place')}{more}. Click one to go there.")

    def search_hits(self) -> list[SearchHit]:
        return list(self._hits)

    def activate_search(self, row: int) -> None:
        if 0 <= row < len(self._hits):
            hit = self._hits[row]
            self.go_find.emit(hit.section, self._query, hit.occurrence)

    # ---- notes
    def set_notes(self, state: lib.BookState, names: list[str]) -> None:
        self.notes_list.clear()
        rows = []
        for b in state.bookmarks:
            rows.append(((b.section, b.offset), (b.section, b.offset), f"Bookmark — {self._where(b.section, names)}: "
                         f"{' '.join(b.text.split())[:120]}"))
        for h in state.highlights:
            text = f"Highlight ({h.colour}) — “{' '.join(h.text.split())[:160]}”"
            if h.note:
                text += f"\n      Note: {h.note}"
            rows.append(((h.section, h.start), (h.section, h.start), text))
        rows.sort(key=lambda r: r[0])
        self._notes = [where for _key, where, _t in rows]
        for _key, _where, text in rows:
            self.notes_list.addItem(text)
        if rows:
            self.notes_message.setText("Click one to go there.")
        else:
            self.notes_message.setText(
                "You have no bookmarks or highlights yet in this book.\n\n"
                "To bookmark a page, press Bookmark at the top.\n"
                "To highlight words, pick them with the mouse (hold the left button down and drag "
                "across them), then choose a colour or Add note.")

    @staticmethod
    def _where(section: int, names: list[str]) -> str:
        return names[section] if section < len(names) and names[section] else f"Part {section + 1}"

    def note_rows(self) -> list[str]:
        return [self.notes_list.item(i).text() for i in range(self.notes_list.count())]

    def activate_note(self, row: int) -> None:
        if 0 <= row < len(self._notes):
            self.go_offset.emit(*self._notes[row])

    # ---- text settings
    def refresh_text(self) -> None:
        s = self.settings
        self.size_label.setText(f"  {s.font_px}  ")
        self.size_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.font_box.blockSignals(True)
        self.font_box.setCurrentText(s.font)
        self.font_box.blockSignals(False)
        picks = {"White": s.colours == "white", "Sepia": s.colours == "sepia", "Night": s.colours == "night",
                 "One page": s.layout == "one", "Two pages": s.layout == "two", "Scroll": s.layout == "scroll",
                 "Close": s.line_spacing <= 1.3, "Normal": 1.3 < s.line_spacing <= 1.65,
                 "Wide": s.line_spacing > 1.65,
                 "Small margins": s.margins == "small", "Medium margins": s.margins == "medium",
                 "Large margins": s.margins == "large"}
        for name, on in picks.items():
            self.text_buttons[name].setChecked(on)
        self.justify_box.blockSignals(True)
        self.justify_box.setChecked(s.justify)
        self.justify_box.blockSignals(False)

    def _changed(self) -> None:
        self.refresh_text()
        self.text_changed.emit()

    def change_size(self, name: str) -> None:
        step = 2 if self.settings.font_px < 30 else 4
        self.settings.font_px = max(12, min(60, self.settings.font_px + (step if name == "Bigger" else -step)))
        self._changed()

    def change_font(self, name: str) -> None:
        if name in lib.FONTS:
            self.settings.font = name
            self._changed()

    def change_colours(self, name: str) -> None:
        self.settings.colours = name.lower()
        self._changed()

    def change_layout(self, name: str) -> None:
        self.settings.layout = {"One page": "one", "Two pages": "two", "Scroll": "scroll"}[name]
        self._changed()

    def change_spacing(self, name: str) -> None:
        self.settings.line_spacing = {"Close": 1.2, "Normal": 1.5, "Wide": 1.8}[name]
        self._changed()

    def change_margins(self, name: str) -> None:
        self.settings.margins = name.split()[0].lower()
        self._changed()

    def change_justify(self, on: bool) -> None:
        self.settings.justify = bool(on)
        self._changed()


# ---------------------------------------------------------------- picked words

class SelectionBar(ui.Bar):
    """Big buttons for words just picked, or for a highlight clicked."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("SelectionBar { background: #fff8e1; }")
        self.mode = ""
        self.buttons: dict[str, QPushButton] = {}
        layout = QVBoxLayout(self)
        self.caption = QLabel("")
        self.caption.setWordWrap(True)
        layout.addWidget(self.caption)
        self.row = QHBoxLayout()
        layout.addLayout(self.row)
        self.hide()

    def _fill(self, names: list[str], handlers: dict) -> None:
        while self.row.count():
            item = self.row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.buttons = {}
        for name in names:
            b = ui.wide_button(name, handlers[name], 56)
            if name in COLOUR_BUTTONS:
                b.setStyleSheet(f"QPushButton {{ background: {COLOUR_BUTTONS[name]}; color: black; }}")
            self.buttons[name] = b
            self.row.addWidget(b)

    def show_new(self, text: str, handlers: dict) -> None:
        self.mode = "new"
        short = " ".join(text.split())
        self.caption.setText("The words you picked: “" + (short[:90] + "…" if len(short) > 90 else short)
                             + "”.  Highlight them in a colour, or:")
        self._fill(["Yellow", "Green", "Blue", "Pink", "Add note", "Copy", "Look up", "Read aloud from here",
                    "Cancel"], handlers)
        self.show()

    def show_existing(self, highlight: lib.Highlight, handlers: dict) -> None:
        self.mode = "existing"
        short = " ".join(highlight.text.split())
        note = f"  Your note: {highlight.note}" if highlight.note else ""
        self.caption.setText("Your highlight: “" + (short[:90] + "…" if len(short) > 90 else short)
                             + "”." + note + "  Change its colour, or:")
        self._fill(["Yellow", "Green", "Blue", "Pink", "Edit note", "Remove", "Cancel"], handlers)
        self.show()

    def close_bar(self) -> None:
        self.mode = ""
        self.hide()


class SpeechBar(ui.Bar):
    """Shown while reading aloud: Pause / Go on, Stop, Slower, Faster and the voice."""

    def __init__(self, page: "ReadingPage") -> None:
        super().__init__(page)
        self.page = page
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("SpeechBar { background: #e0f7fa; }")
        row = QHBoxLayout(self)
        self.caption = QLabel("Reading aloud")
        self.caption.setStyleSheet("font-weight: bold;")
        row.addWidget(self.caption)
        self.buttons: dict[str, QPushButton] = {}
        for name, slot in (("Pause", page.pause_or_go_on), ("Stop", page.stop_reading),
                           ("Slower", lambda: page.change_speech_rate(-0.1)),
                           ("Faster", lambda: page.change_speech_rate(0.1))):
            b = ui.wide_button(name, slot, 56)
            self.buttons[name] = b
            row.addWidget(b)
        self.voice_box = QComboBox()
        self.voice_box.setMinimumHeight(52)
        self.voice_box.setToolTip("Choose the voice that reads")
        self.voice_box.currentTextChanged.connect(page.change_voice)
        row.addWidget(QLabel("Voice:"))
        row.addWidget(self.voice_box, 1)
        self.hide()

    def set_voices(self, names: list[str], current: str) -> None:
        self.voice_box.blockSignals(True)
        self.voice_box.clear()
        self.voice_box.addItems(names)
        if current in names:
            self.voice_box.setCurrentText(current)
        self.voice_box.blockSignals(False)


class PdfBar(ui.Bar):
    """Zoom and page buttons, shown only for PDFs."""

    def __init__(self, page: "ReadingPage") -> None:
        super().__init__(page)
        row = QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 0)
        self.page = page
        self.buttons: dict[str, QPushButton] = {}
        for name, slot in (("Zoom out", lambda: page.pdf.zoom_by(1 / 1.25)),
                           ("Zoom in", lambda: page.pdf.zoom_by(1.25)),
                           ("Fit width", page.pdf.fit_width), ("Whole page", page.pdf.fit_page),
                           ("Continuous pages", self.toggle_continuous)):
            b = ui.wide_button(name, slot, 52)
            self.buttons[name] = b
            row.addWidget(b)
        self.buttons["Continuous pages"].setCheckable(True)
        row.addStretch(1)
        self.hide()

    def toggle_continuous(self) -> None:
        self.page.pdf.set_continuous(self.buttons["Continuous pages"].isChecked())


class PictureWindow(QDialog):
    """A picture from the book, as big as the screen allows."""

    def __init__(self, data: bytes, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Picture")
        self.original = QPixmap()
        self.original.loadFromData(data)
        layout = QVBoxLayout(self)
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label, 1)
        layout.addWidget(ui.wide_button("Close", self.close, 56))
        screen = (parent.screen() if parent else QApplication.primaryScreen()).availableGeometry()
        box = screen.size() * 0.85
        shown = self.original
        if not shown.isNull():
            scale = min(box.width() / max(1, shown.width()), (box.height() - 100) / max(1, shown.height()), 6.0)
            shown = shown.scaled(int(shown.width() * scale), int(shown.height() * scale),
                                 Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.label.setPixmap(shown)

    def picture_size(self) -> tuple[int, int]:
        return self.original.width(), self.original.height()


# ---------------------------------------------------------------- the page

class Worker(QObject):
    """Runs a job in a thread and hands its result back on the window's thread."""

    done = Signal(object, object)

    def run(self, key, job) -> None:
        def go():
            try:
                result = job()
            except Exception as e:  # shown to the person as a plain message
                result = e
            self.done.emit(key, result)
        threading.Thread(target=go, daemon=True).start()


class ReadingPage(QWidget):
    library_requested = Signal()
    full_screen_requested = Signal(bool)

    def __init__(self, library: lib.Library, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.library = library
        self.settings = library.settings
        self.book = None
        self.book_id = ""
        self.path = ""
        self.state: lib.BookState | None = None
        self.picture_window: PictureWindow | None = None
        self.lookup_window: speech.LookupWindow | None = None
        self.dictionary_url = speech.DICTIONARY_URL
        self.read_aloud_engine = speech.ReadAloud(self)
        self.read_aloud_engine.state_changed.connect(self._speech_state)
        self._restart_reading = False
        self._lengths: list[int] = []
        self._names: list[str] = []
        self._selection: dict = {}
        self._highlight_id = ""
        self._last_turn = (0, 0, 0.0)
        self._page_info: dict = {}
        self._search_stop = threading.Event()
        self.worker = Worker(self)
        self.worker.done.connect(self._job_done)
        self.full_screen = False

        self.buttons: dict = {}
        top = QHBoxLayout()
        top.setSpacing(6)
        for name, glyph, tip, slot in (
                ("Library", "☰", "Back to all your books", self.go_library),
                ("Contents", "≡", "The list of chapters", lambda: self.toggle_panel("contents")),
                ("Search", "\U0001F50D", "Find words in this book (Ctrl+F)", lambda: self.toggle_panel("search")),
                ("Bookmark", "\U0001F516", "Mark this page so you can find it again (Ctrl+B)", self.toggle_bookmark),
                ("Notes", "✎", "Your bookmarks, highlights and notes", lambda: self.toggle_panel("notes")),
                ("Text", "A", "Text size, font, colours and pages", lambda: self.toggle_panel("text")),
                ("Read aloud", "\U0001F50A", "Read the book out loud", self.read_aloud),
                ("Full screen", "⛶", "Use the whole screen for the page (F11)", self.toggle_full_screen),
                ("Save a copy", "\U0001F4BE", "Save this book as EPUB, PDF, Word or plain text", self.save_copy)):
            b = ui.tool_button(name, glyph, tip, slot)
            self.buttons[name] = b
            top.addWidget(b)
        top.addStretch(1)
        self.top_bar = ui.SideScroll(ui.holder(top))

        self.reader = ReaderView()
        self.reader.place_changed.connect(self._on_place)
        self.reader.selected.connect(self._on_selected)
        self.reader.tapped.connect(self._close_selection)
        self.reader.highlight_clicked.connect(self._on_highlight_clicked)
        self.reader.web_link_clicked.connect(self._on_web_link)
        self.reader.image_clicked.connect(self._on_image)
        self.reader.section_changed.connect(lambda _s: self._close_selection())
        self.reader.edge_reached.connect(self._on_edge)
        self.reader.load_failed.connect(self._page_failed)
        self.pdf = PdfView()
        self.pdf.place_changed.connect(self._on_place)
        self.pdf.edge_reached.connect(self._on_edge)
        self.views = QStackedWidget()
        self.views.addWidget(self.reader)
        self.views.addWidget(self.pdf)
        self.current = self.reader
        self.is_pdf = False

        self.panel = SidePanel(self.settings)
        self.panel.go.connect(lambda section, anchor: self._go_anchor(section, anchor))
        self.panel.go_offset.connect(lambda section, offset: self.current.go_to(section, offset, remember=True))
        self.panel.go_find.connect(self._go_find)
        self.panel.search_requested.connect(self._search)
        self.panel.save_notes_requested.connect(self.save_notes)
        self.panel.text_changed.connect(self._text_changed)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.addWidget(self.panel)
        self.splitter.addWidget(self.views)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.setChildrenCollapsible(False)

        self.selection_bar = SelectionBar()
        self.speech_bar = SpeechBar(self)
        self.pdf_bar = PdfBar(self)

        bottom = QHBoxLayout()
        self.back_button = ui.wide_button("◀ Back", self.go_back, ui.SECOND_ROW,
                                          "Go back to where you were before you followed a link")
        self.back_button.hide()
        bottom.addWidget(self.back_button)
        self.prev_button = ui.wide_button("◀  Previous page", self.prev_page, ui.SECOND_ROW)
        self.next_button = ui.wide_button("Next page  ▶", self.next_page, ui.SECOND_ROW)
        bottom.addWidget(self.prev_button)
        middle = QVBoxLayout()
        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setRange(0, 1000)
        self.slider.setMinimumHeight(30)
        self.slider.setToolTip("Drag to move through the whole book")
        self.slider.sliderReleased.connect(self.slider_released)
        self.slider.valueChanged.connect(self._slider_moving)
        middle.addWidget(self.slider)
        self.status = QLabel("")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status.setWordWrap(True)                     # two lines rather than cut off
        self.status.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Minimum)
        self.status.setMinimumHeight(self.status.fontMetrics().lineSpacing() * 2 + 6)
        middle.addWidget(self.status)
        bottom.addLayout(middle, 1)
        bottom.addWidget(self.next_button)
        self.bottom_bar = QWidget()
        self.bottom_bar.setLayout(bottom)

        self.exit_full = ui.wide_button("Exit full screen", self.toggle_full_screen, 52)
        self.exit_full.hide()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.addWidget(self.top_bar)
        layout.addWidget(self.exit_full, 0, Qt.AlignmentFlag.AlignRight)
        layout.addWidget(self.splitter, 1)
        layout.addWidget(ui.SideScroll(self.selection_bar))
        layout.addWidget(ui.SideScroll(self.speech_bar))
        layout.addWidget(ui.SideScroll(self.pdf_bar))
        layout.addWidget(self.bottom_bar)

        for keys, slot in (("Ctrl+F", lambda: self.toggle_panel("search", force=True)),
                           ("Ctrl+B", self.toggle_bookmark), ("F11", self.toggle_full_screen),
                           ("Escape", self._escape)):
            shortcut = QShortcut(QKeySequence(keys), self)
            shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            shortcut.activated.connect(slot)

    # ---------------------------------------------------------------- opening and closing
    def open_path(self, path: str, entry_id: str | None = None) -> bool:
        """Open a book file; on a problem, say so plainly and return False."""
        if path.lower().endswith(".pdf"):
            return self._open_pdf(path, entry_id)
        self.status.setText("Opening\u2026")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        QApplication.processEvents(QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents)
        problem = None
        try:
            book = formats.open_book(path)
        except (formats.BookLocked, formats.BookError) as e:
            book, problem = None, e
        finally:
            QApplication.restoreOverrideCursor()
        if isinstance(problem, formats.BookLocked):
            explain_lock(self, problem.info)
            return False
        if problem is not None:
            ui.inform(self, "This book can't be opened", str(problem))
            return False
        self.close_book()
        self.book = book
        self.path = path
        self.current, self.is_pdf = self.reader, False
        self.pdf_bar.hide()
        try:
            self.book_id = entry_id or lib.book_id(path)
        except OSError:
            self.book_id = os.path.basename(path)
        self.state = self.library.state(self.book_id)
        self.state.opened = time.time()
        if not self.state.added:
            self.state.added = self.state.opened
        self._names = section_names(book)
        self.panel.set_book(book)
        self.panel.hide()
        self._lengths = []
        if len(self.state.section_lengths) == len(book.sections):
            self._lengths = list(self.state.section_lengths)
        elif sum(len(book.section_html(i)) for i in range(len(book.sections))) < 3_000_000:
            self._lengths = [len(book.section_text(i)) for i in range(len(book.sections))]
        else:
            self.worker.run("lengths", lambda b=book: [len(b.section_text(i)) for i in range(len(b.sections))])
        self.views.setCurrentWidget(self.reader)
        self.reader.open(book, self.state, self.settings)
        self.settings.last_book = path
        self._save_settings()
        self._update_back()
        self.status.setText("Opening…")
        return True

    def close_book(self) -> None:
        if self.book is None:
            return
        self._search_stop.set()
        self.read_aloud_engine.stop()
        self._save_state()
        self._close_selection()
        self.book.close()
        if self.is_pdf:
            self.pdf.close()
        self.book = None
        if self.picture_window is not None:
            self.picture_window.close()

    def lengths_known(self) -> bool:
        return bool(self._lengths) and self.book is not None and len(self._lengths) == len(self.book.sections)

    def _job_done(self, key, result) -> None:
        if key == "lengths" and isinstance(result, list) and self.book is not None \
                and len(result) == len(self.book.sections):
            self._lengths = result
            self._update_status()
        elif isinstance(key, tuple) and key[0] == "search" and self.book is not None:
            if isinstance(result, Exception):
                self.panel.search_status.setText("The search didn't work: " + str(result))
            elif key[1] is self.book:
                self.panel.set_results(key[2], result, self._names)

    def _save_state(self) -> None:
        if self.book is None or self.state is None:
            return
        self.state.place = self.current.place()
        self.state.progress = self.fraction()
        if self.is_pdf:
            self.state.zoom = self.pdf.zoom_for_state()
        if self.lengths_known():
            self.state.section_lengths = list(self._lengths)
        try:
            self.library.save_state(self.book_id, self.state)
        except OSError:
            pass

    def _save_settings(self) -> None:
        try:
            self.library.save_settings()
        except OSError:
            pass

    def _open_pdf(self, path: str, entry_id: str | None) -> bool:
        try:
            ident = entry_id or lib.book_id(path)
        except OSError as e:
            ui.inform(self, "This book can't be opened", f"The file can't be read ({e.strerror or e}).")
            return False
        state = self.library.state(ident)

        def ask_password(tries: int) -> str | None:
            label = ("This PDF needs its password to open. Type the password:" if not tries else
                     "That password wasn't right. Type the PDF's password again:")
            return ui.ask_text(self, "PDF password", label, "", multiline=False)

        self.close_book()
        problem = self.pdf.open(path, state, ask_password)
        if problem:
            ui.inform(self, "This PDF can't be opened", problem)
            return False
        self.book, self.path, self.book_id, self.state = self.pdf.book, path, ident, state
        self.current, self.is_pdf = self.pdf, True
        state.opened = time.time()
        state.added = state.added or state.opened
        self._names = [f"Page {n + 1}" for n in range(self.pdf.pages)]
        self._lengths = []
        self.panel.set_book(self.book)
        self.panel.hide()
        self.views.setCurrentWidget(self.pdf)
        self.pdf_bar.show()
        self.settings.last_book = path
        self._save_settings()
        self._update_back()
        return True

    # ---------------------------------------------------------------- moving
    def next_page(self) -> None:
        self._take_over_from_voice()
        self.current.next_page()

    def prev_page(self) -> None:
        self._take_over_from_voice()
        self.current.prev_page()

    def _take_over_from_voice(self) -> None:
        """Turning the page while it reads: carry on reading from the new page."""
        if self.read_aloud_engine.state == "reading":
            self.read_aloud_engine.stop(quiet=True)
            self._restart_reading = True

    def go_back(self) -> None:
        self.current.go_back()
        self._update_back()

    def _go_anchor(self, section: int, anchor: str) -> None:
        href = self.book.sections[section].href + ("#" + anchor if anchor else "")
        if not self.current.go_to_href(href):
            self.current.go_to(section, 0, remember=True)
        self._update_back()

    def _go_find(self, section: int, query: str, occurrence: int) -> None:
        self.current.history.append(self.current.place())
        self.current.go_to_find(section, query, occurrence)
        self._update_back()

    def _page_failed(self, text: str) -> None:
        """A part of the book couldn't be shown: say so and go back to the library."""
        ui.inform(self, "Book Reader", text + " The book is closed; you are back in the library.")
        self.go_library()

    def _on_edge(self, where: str) -> None:
        if where == "end":
            self.status.setText("This is the end of the book.")

    def fraction(self) -> float:
        if self.is_pdf:
            return self.pdf.page / max(1, self.pdf.pages - 1)
        if not self.lengths_known():
            return 0.0
        total = sum(self._lengths) or 1
        section, offset = self.current.place()
        return min(1.0, (sum(self._lengths[:section]) + min(offset, self._lengths[section])) / total)

    def _on_place(self, section: int, offset: int, page: int, pages: int) -> None:
        now = time.monotonic()
        last_section, last_page, when = self._last_turn
        if self.state is not None and when and (section, page) in ((last_section, last_page + 1), (last_section + 1, 0)):
            lib.update_speed(self.state, max(1, self.current.total // max(1, pages)), now - when)
        self._last_turn = (section, page, now)
        if self._restart_reading:
            self._restart_reading = False
            self.read_aloud()
        if self.state is not None:
            self.state.place = (section, offset)
            self.state.progress = self.fraction()
            self.state.finished = self.state.finished or (self.current.at_book_end and self.state.progress > 0.95)
            self._save_state()
        self.panel.mark_current(section)
        self._update_status()
        self._update_back()
        self._update_bookmark_button()
        if not self.slider.isSliderDown():
            self.slider.blockSignals(True)
            self.slider.setValue(int(self.fraction() * 1000))
            self.slider.blockSignals(False)

    def _update_status(self) -> None:
        if self.book is None:
            return
        section, offset = self.current.place()
        if self.is_pdf:
            self.status.setText(f"page {self.pdf.page + 1} of {self.pdf.pages} - "
                                f"{int(self.fraction() * 100)}% of the book")
            return
        name = self._names[section] if section < len(self._names) else f"Part {section + 1}"
        text = f"{name} - page {self.current.page + 1} of {self.current.pages}"
        if self.lengths_known():
            text += f" - {int(self.fraction() * 100)}% of the book"
            left = max(0, self._lengths[section] - offset)
            speed = self.state.chars_per_minute if self.state else 1000.0
            minutes = max(1, round(left / speed))
            text += f" - about {plural(minutes, 'minute')} left in this chapter"
        self.status.setText(text)

    def _update_back(self) -> None:
        self.back_button.setVisible(bool(self.current.history))

    def _slider_moving(self, value: int) -> None:
        if self.slider.isSliderDown():
            self.status.setText(f"Let go to jump to {value // 10}% of the book")

    def slider_released(self) -> None:
        if self.is_pdf:
            self.pdf.go_to(round(self.slider.value() / 1000 * max(0, self.pdf.pages - 1)), 0, remember=True)
            self._update_back()
            return
        if not self.lengths_known():
            self.status.setText("Still working out how long the book is; try again in a moment.")
            return
        target = self.slider.value() / 1000 * sum(self._lengths)
        for section, length in enumerate(self._lengths):
            if target < length or section == len(self._lengths) - 1:
                self.current.go_to(section, int(max(0, min(target, length))), remember=True)
                break
            target -= length
        self._update_back()

    # ---------------------------------------------------------------- panels
    def toggle_panel(self, kind: str, force: bool = False) -> None:
        if self.book is None:
            return
        if kind == "text" and self.is_pdf:
            ui.inform(self, "Text and colours", PDF_TEXT_NOTE)
            return
        if self.panel.isVisible() and self.panel.kind == kind and not force:
            self.panel.hide()
            return
        if kind == "notes":
            self.panel.set_notes(self.state, self._names)
        self.panel.show_kind(kind)
        if kind == "contents":
            self.panel.mark_current(self.current.section)

    def _search(self, query: str) -> None:
        if self.is_pdf:                       # Qt's PDF engine is used on this thread only
            self.panel.set_results(query, search_book(self.book, query), self._names)
            return
        self._search_stop.set()
        self._search_stop = threading.Event()
        stop, book = self._search_stop, self.book
        self.worker.run(("search", book, query), lambda: search_book(book, query, stop))

    def _text_changed(self) -> None:
        self._save_settings()
        self.current.apply_settings(self.settings)

    def save_notes(self) -> None:
        if self.book is None:
            return
        suggested = os.path.join(os.path.expanduser("~"), "Documents",
                                 re.sub(r'[<>:"/\\|?*]', "", self.book.title or "Book")[:80] + " - notes.txt")
        name = ui.save_file_name(self, "Save your notes", suggested, "Text files (*.txt)")
        if not name:
            return
        try:
            with open(name, "w", encoding="utf-8") as f:
                f.write(self.library.notes_text(self.book_id, self.book.title, self._names))
        except OSError as e:
            ui.inform(self, "Notes not saved", f"The notes couldn't be saved there ({e.strerror or e}).")
            return
        ui.inform(self, "Notes saved", f"Your notes are saved in:\n{name}")

    # ---------------------------------------------------------------- bookmarks
    def toggle_bookmark(self) -> None:
        if self.book is None:
            return

        def got(info: dict) -> None:
            start, end = int(info.get("start", 0)), int(info.get("end", 0))
            section = self.current.section
            here = [b for b in self.state.bookmarks if b.section == section and start <= b.offset < max(end, start + 1)]
            if here:
                for b in here:
                    self.state.bookmarks.remove(b)
            else:
                self.state.bookmarks.append(lib.Bookmark(section, start, " ".join(str(info.get("text", "")).split()),
                                                         time.time()))
            self._save_state()
            self._update_bookmark_button()
            if self.panel.isVisible() and self.panel.kind == "notes":
                self.panel.set_notes(self.state, self._names)
        self.current.page_info(got)

    def _update_bookmark_button(self) -> None:
        if self.book is None or self.state is None:
            return
        section = self.current.section
        mine = [b for b in self.state.bookmarks if b.section == section]
        if not mine:
            self.buttons["Bookmark"].setText("Bookmark")
            return

        def got(info: dict) -> None:
            start, end = int(info.get("start", 0)), int(info.get("end", 0))
            on = any(start <= b.offset < max(end, start + 1) for b in mine)
            self.buttons["Bookmark"].setText("Remove bookmark" if on else "Bookmark")
        self.current.page_info(got)

    # ---------------------------------------------------------------- highlights
    def _on_selected(self, data: dict) -> None:
        self._selection = data
        handlers = {name: (lambda _c=False, n=name: self._highlight_new(n.lower())) for name in COLOUR_BUTTONS}
        handlers.update({"Add note": self._note_new, "Copy": self._copy, "Look up": self._look_up_selection,
                         "Read aloud from here": self._read_from_selection, "Cancel": self._close_selection})
        self.selection_bar.show_new(str(data.get("text", "")), handlers)

    def _close_selection(self) -> None:
        if self.selection_bar.isVisible():
            self.selection_bar.close_bar()
            self.current.clear_selection()
        self._selection = {}

    def _add_highlight(self, colour: str, note: str = "") -> None:
        data = self._selection
        if not data or self.state is None:
            return
        start, end = int(data.get("start", 0)), int(data.get("end", 0))
        if end <= start:
            return
        self.state.highlights.append(lib.Highlight(uuid.uuid4().hex, self.current.section, start, end,
                                                   str(data.get("text", "")), colour, note, time.time()))
        self._save_state()
        self.current.set_highlights(self.state.highlights)
        self._close_selection()

    def _highlight_new(self, colour: str) -> None:
        self._add_highlight(colour)

    def _note_new(self) -> None:
        note = ui.ask_text(self, "Add a note", "Type your note about these words:")
        if note is not None:
            self._add_highlight("yellow", note.strip())

    def _copy(self) -> None:
        QApplication.clipboard().setText(str(self._selection.get("text", "")))
        self._close_selection()

    def _look_up_selection(self) -> None:
        self.look_up(str(self._selection.get("text", "")))
        self._close_selection()

    def _read_from_selection(self) -> None:
        start = int(self._selection.get("start", 0))
        self._close_selection()
        self.read_aloud(from_offset=start)

    def _find_highlight(self, ident: str) -> lib.Highlight | None:
        return next((h for h in (self.state.highlights if self.state else []) if h.id == ident), None)

    def _on_highlight_clicked(self, ident: str, _x: int, _y: int) -> None:
        mark = self._find_highlight(ident)
        if mark is None:
            return
        self._highlight_id = ident
        handlers = {name: (lambda _c=False, n=name: self._recolour(n.lower())) for name in COLOUR_BUTTONS}
        handlers.update({"Edit note": self._edit_note, "Remove": self._remove_highlight,
                         "Cancel": self.selection_bar.close_bar})
        self.selection_bar.show_existing(mark, handlers)

    def _recolour(self, colour: str) -> None:
        mark = self._find_highlight(self._highlight_id)
        if mark:
            mark.colour = colour
            self._save_state()
            self.current.set_highlights(self.state.highlights)
        self.selection_bar.close_bar()

    def _edit_note(self) -> None:
        mark = self._find_highlight(self._highlight_id)
        if mark is None:
            return
        note = ui.ask_text(self, "Your note", "Change your note (leave it empty to have no note):", mark.note)
        if note is not None:
            mark.note = note.strip()
            self._save_state()
        self.selection_bar.close_bar()

    def _remove_highlight(self) -> None:
        mark = self._find_highlight(self._highlight_id)
        if mark is not None:
            self.state.highlights.remove(mark)
            self._save_state()
            self.current.set_highlights(self.state.highlights)
        self.selection_bar.close_bar()

    # ---------------------------------------------------------------- links and pictures
    def _on_web_link(self, url: str) -> None:
        if ui.ask(self, "Open a web page?", f"This link goes to a web page:\n\n{url}\n\n"
                                            "Open it in your web browser?", "Open it", "No"):
            ui.open_url(url)

    def _on_image(self, href: str) -> None:
        found = self.book.resource(href) if self.book else None
        if not found:
            return
        if self.picture_window is not None:
            self.picture_window.close()
        self.picture_window = PictureWindow(found[0], self)
        self.picture_window.show()

    # ---------------------------------------------------------------- other buttons
    def go_library(self) -> None:
        self.close_book()
        if self.full_screen:
            self.toggle_full_screen()
        self.library_requested.emit()

    def toggle_full_screen(self) -> None:
        self.set_full_screen(not self.full_screen)
        self.full_screen_requested.emit(self.full_screen)

    def set_full_screen(self, on: bool) -> None:
        self.full_screen = on
        self.top_bar.setVisible(not on)
        self.bottom_bar.setVisible(not on)
        self.exit_full.setVisible(on)
        if on:
            self.panel.hide()

    def _escape(self) -> None:
        if self.selection_bar.isVisible():
            self._close_selection()
        elif self.full_screen:
            self.toggle_full_screen()
        elif self.panel.isVisible():
            self.panel.hide()

    def look_up(self, text: str) -> None:
        word = speech.clean_word(text)
        if not word:
            ui.inform(self, "Look up", "Pick a word first: hold the left mouse button down and drag across it.")
            return
        if self.lookup_window is not None:
            self.lookup_window.close()
        self.lookup_window = speech.LookupWindow(word, self)
        self.lookup_window.show()
        speech.lookup_word(word, self.lookup_window.show_result, self.dictionary_url)

    # ---------------------------------------------------------------- reading aloud
    def read_aloud(self, from_offset: int | None = None) -> None:
        """The Read aloud button: start reading here, or stop if reading."""
        if self.book is None:
            return
        engine = self.read_aloud_engine
        if engine.state in ("reading", "paused") and from_offset is None and not self._restart_reading:
            self.stop_reading()
            return
        problem = engine.start(self.current, from_offset, self.settings.voice, self.settings.speech_rate)
        if problem:
            ui.inform(self, "Read aloud", problem)
            return
        self.speech_bar.set_voices(engine.voice_names(), self.settings.voice)

    def stop_reading(self) -> None:
        self._restart_reading = False
        self.read_aloud_engine.stop()

    def pause_or_go_on(self) -> None:
        engine = self.read_aloud_engine
        if engine.state == "reading":
            engine.pause()
        elif engine.state == "paused":
            engine.resume()

    def change_speech_rate(self, step: float) -> None:
        self.settings.speech_rate = round(max(-1.0, min(1.0, self.settings.speech_rate + step)), 2)
        self.read_aloud_engine.set_rate(self.settings.speech_rate)
        self._save_settings()
        speed = int(round((self.settings.speech_rate + 1) * 50))
        self.speech_bar.caption.setText(f"Reading aloud (speed {speed} of 100)")

    def change_voice(self, name: str) -> None:
        if name:
            self.settings.voice = name
            self.read_aloud_engine.set_voice(name)
            self._save_settings()

    def _speech_state(self, state: str) -> None:
        reading = state in ("reading", "paused")
        self.buttons["Read aloud"].setText("Stop reading" if reading else "Read aloud")
        self.speech_bar.setVisible(reading)
        self.speech_bar.buttons["Pause"].setText("Go on" if state == "paused" else "Pause")
        if state == "finished":
            self.status.setText("Finished reading: this is the end of the book.")

    def save_copy(self) -> None:
        """Save a copy. It is made from its own opening of the file, so going to the library or
        opening another book while it is made doesn't disturb it."""
        if self.book is None:
            return
        import book_reader_copy
        if self.is_pdf:                      # PDFs are copied as text straight away
            self._copy_job = book_reader_copy.ask_and_save(self, self.book)
            return
        try:
            own = formats.open_book(self.path)
        except (formats.BookLocked, formats.BookError) as e:
            ui.inform(self, "Save a copy", str(e))
            return
        job = book_reader_copy.ask_and_save(self, own, on_done=lambda _ok, _m, b=own: b.close())
        if job is None:
            own.close()
        self._copy_job = job


def section_names(book) -> list[str]:
    """A name for each section: the contents entry it falls under, else 'Part N'."""
    names = [""] * len(book.sections)
    entries = sorted(((e.section, e.title) for e in book.toc if 0 <= e.section < len(names)), key=lambda e: e[0])
    current = ""
    i = 0
    for section in range(len(names)):
        while i < len(entries) and entries[i][0] <= section:
            current = entries[i][1] if entries[i][0] == section or not names[section] else current
            if entries[i][0] == section and not names[section]:
                names[section] = entries[i][1]
            i += 1
        names[section] = names[section] or current or f"Part {section + 1}"
    return [n[:60] for n in names]


def explain_lock(parent: QWidget | None, info: formats.LockInfo) -> None:
    """Say why a book is locked; for free Gutenberg books, offer the free copy."""
    if info.free_copy_url:
        if ui.ask(parent, "This book is locked", info.message + "\n\nOpen the page where you can download "
                  "the free, unlocked copy?", "Open the page", "No"):
            ui.open_url(info.free_copy_url)
        return
    ui.inform(parent, "This book is locked", info.message)
