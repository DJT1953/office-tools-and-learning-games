"""Book Reader: reads EPUB, Kindle (MOBI, AZW, AZW3, PRC), PDF, text and saved
web pages, with a library of covers, bookmarks, highlights, notes, search,
reading aloud and more. Locked books are listed with the reason; nothing here
removes copy protection.

Start it with Run_BookReader.bat. Books dragged onto that file (or onto the
window) open in it. Only one Book Reader window runs at a time: a second start
hands its books to the open window.
"""

from __future__ import annotations

import getpass
import importlib.util
import json
import os
import re
import site
import subprocess
import sys
import threading
import traceback

sys.dont_write_bytecode = True

TITLE = "Book Reader"
NEEDS = {
    # module to look for: (what pip installs, what it is for, download size in MB)
    "PySide6.QtWebEngineWidgets": ("pyside6>=6.10", "PySide6, which shows the books and the window", 200),
    "lxml": ("lxml", "lxml, which reads EPUB books", 4),
    "comtypes": ("comtypes", "comtypes, which lets Microsoft Word save copies", 1),
}


# ------------------------------------------------------- start-up and crashes

def tk_message(kind: str, text: str):
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    answer = getattr(messagebox, kind)(TITLE, text)
    root.destroy()
    return answer


def show_crash(kind, value, tb) -> None:
    """A window started without a console closes the moment it fails, taking the
    error with it, so show anything unexpected in a box."""
    text = "".join(traceback.format_exception(kind, value, tb))
    print(text, file=sys.stderr)
    message = "Something went wrong:\n\n" + text[-1500:]
    widgets = sys.modules.get("PySide6.QtWidgets")
    if widgets and widgets.QApplication.instance():
        widgets.QMessageBox.critical(None, TITLE, message)
    else:
        tk_message("showerror", message)


def has_module(name: str) -> bool:
    try:
        spec = importlib.util.find_spec(name)
    except ModuleNotFoundError:
        return False
    return spec is not None and spec.origin not in (None, "namespace")


def install_missing(needs: dict = NEEDS) -> None:
    """On a PC where they aren't installed yet, offer to install the free add-ons."""
    missing = [need for module, need in needs.items() if not has_module(module)]
    if not missing:
        return
    names = ", ".join(what for _, what, _ in missing)
    size = sum(mb for *_, mb in missing)
    if not tk_message("askyesno",
                      f"This copy of Python ({sys.version.split()[0]}) still needs {names}."
                      f"\n\nDownload and install now? This only happens once. It is about "
                      f"{size} MB, so it can take a few minutes; a window shows the progress."):
        sys.exit()
    command = [sys.executable, "-m", "pip", "install", *(pip for pip, _, _ in missing)]
    flags = subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0   # show pip's progress
    if subprocess.run(command, creationflags=flags).returncode:
        raise RuntimeError("The add-ons couldn't be installed. Check the internet connection "
                           "and try again.")
    site.addsitedir(site.getusersitepackages())   # a first install may have just made it
    importlib.invalidate_caches()


if __name__ == "__main__":
    sys.excepthook = show_crash
    install_missing()

import book_reader_web as web  # noqa: E402  (its scheme must be registered before QApplication)

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QRect, QSize, Qt, Signal  # noqa: E402
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QPen, QPixmap  # noqa: E402
from PySide6.QtNetwork import QLocalServer, QLocalSocket  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QAbstractItemView,
    QApplication,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QStyle,
    QStyledItemDelegate,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

import book_reader_formats as formats  # noqa: E402
import book_reader_genres as genres  # noqa: E402
import book_reader_library as lib  # noqa: E402
import book_reader_ui as ui  # noqa: E402
from book_reader_reading import ReadingPage, explain_lock, plural  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
HELP_FILE = os.path.join(HERE, "Book Reader - How to use.txt")
TILE = QSize(260, 480)
COVER = QSize(200, 280)
COVER_COLOURS = ["#5d4037", "#1a237e", "#004d40", "#4a148c", "#b71c1c", "#33691e", "#0d47a1", "#e65100",
                 "#263238", "#880e4f"]
FILTERS = ["All", "Reading", "Not started", "Finished", "Favorites", "Locked"]
SORT_NAMES = {"recent": "Most recent", "title": "Title", "author": "Author", "genre": "Genre",
              "added": "Date added"}
BOOK_FILTER = "Books (*.epub *.mobi *.azw *.azw3 *.prc *.pdf *.txt *.htm *.html *.mht *.mhtml);;All files (*)"
FREE_BOOKS = [
    # Websites that let anyone download books free without signing up (checked 2026-11-13).
    # name, address shown, page opened, what it has
    ("Project Gutenberg", "gutenberg.org", "https://www.gutenberg.org/", "tens of thousands of free classic books"),
    ("Standard Ebooks", "standardebooks.org", "https://standardebooks.org/ebooks",
     "free classics, carefully edited"),
]


def server_name() -> str:
    return "BookReader-" + re.sub(r"\W", "", getpass.getuser() or "me")


def send_to_running(name: str, paths: list[str]) -> bool:
    """Hand paths to a Book Reader already open; False if none is open."""
    sock = QLocalSocket()
    sock.connectToServer(name)
    if not sock.waitForConnected(700):
        return False
    sock.write(json.dumps({"open": paths}).encode("utf-8") + b"\n")
    sock.flush()
    sock.waitForBytesWritten(2000)
    sock.waitForReadyRead(3000)              # the open window answers "ok" once it has the books
    sock.disconnectFromServer()
    return True


def make_thumbnail(data: bytes) -> bytes | None:
    """A cover picture made small (420 pixels tall) for the library."""
    image = QImage.fromData(data)
    if image.isNull():
        return None
    if image.height() > 420:
        image = image.scaledToHeight(420, Qt.TransformationMode.SmoothTransformation)
    out = QByteArray()
    buffer = QBuffer(out)
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buffer, "PNG")
    return bytes(out)


def drawn_cover(title: str, author: str, size: QSize = COVER) -> QPixmap:
    """A plain cover for a book without one: its title and author on a colour."""
    pixmap = QPixmap(size)
    colour = QColor(COVER_COLOURS[sum(map(ord, title)) % len(COVER_COLOURS)])
    pixmap.fill(colour)
    painter = QPainter(pixmap)
    painter.setPen(QPen(QColor("white")))
    painter.drawRect(QRect(8, 8, size.width() - 17, size.height() - 17))
    font = QFont("Georgia")
    font.setPixelSize(24)
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(QRect(18, 24, size.width() - 36, size.height() - 110),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter | Qt.TextFlag.TextWordWrap, title[:90])
    font.setPixelSize(17)
    font.setBold(False)
    painter.setFont(font)
    painter.drawText(QRect(18, size.height() - 84, size.width() - 36, 64),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop | Qt.TextFlag.TextWordWrap, author[:60])
    painter.end()
    return pixmap


def sort_title(title: str) -> str:
    return re.sub(r"^(the|a|an)\s+", "", title.strip().casefold())


def progress_text(entry: lib.LibraryEntry, state: lib.BookState) -> str:
    if entry.info.lock:
        owner = entry.info.lock.owner
        return f"Locked by {owner}" if owner in ("Amazon", "Samsung", "Adobe", "Apple") else "Locked"
    if entry.info.error:
        return "Can't be opened"
    if state.finished:
        return "Finished"
    if state.progress > 0:
        return f"{max(1, int(state.progress * 100))}% read"
    return "Started" if state.opened else "Not started"


# ---------------------------------------------------------------- library page

class TileDelegate(QStyledItemDelegate):
    """Draws a book tile: cover, title, author, how far read."""

    def __init__(self, page: "LibraryPage") -> None:
        super().__init__(page)
        self.page = page

    def sizeHint(self, option, index) -> QSize:  # noqa: N802
        return TILE

    def paint(self, painter: QPainter, option, index) -> None:
        data = index.data(Qt.ItemDataRole.UserRole)
        if not isinstance(data, dict):
            return
        rect = option.rect.adjusted(6, 6, -6, -6)
        painter.save()
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        painter.fillRect(rect, QColor("#cfe8ff") if selected else QColor("white"))
        painter.setPen(QPen(QColor("#1565c0") if selected else QColor("#bbbbbb"), 3 if selected else 1))
        painter.drawRect(rect)
        cover = self.page.cover_for(data)
        scaled = cover.scaled(COVER, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        x = rect.x() + (rect.width() - scaled.width()) // 2
        painter.drawPixmap(x, rect.y() + 8 + (COVER.height() - scaled.height()), scaled)
        if data["locked"]:
            sign = QFont("Segoe UI Emoji")
            sign.setPixelSize(30)
            painter.setFont(sign)
            painter.fillRect(QRect(x + 4, rect.y() + 12, 42, 42), QColor(255, 255, 255, 220))
            painter.drawText(QRect(x + 4, rect.y() + 12, 42, 42), Qt.AlignmentFlag.AlignCenter,
                             "\U0001F512" if data["progress"].startswith("Locked") else "\u26A0")
        if data["favorite"]:
            star = QFont("Segoe UI Symbol")
            star.setPixelSize(30)
            painter.setFont(star)
            painter.setPen(QColor("#f9a825"))
            painter.drawText(QRect(rect.right() - 40, rect.y() + 4, 36, 36), Qt.AlignmentFlag.AlignCenter, "★")
        top = rect.y() + COVER.height() + 14
        font = QFont(self.page.font())
        font.setPixelSize(24)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QColor("black"))
        painter.drawText(QRect(rect.x() + 6, top, rect.width() - 12, 62),
                         Qt.AlignmentFlag.AlignHCenter | Qt.TextFlag.TextWordWrap, data["title"][:70])
        font.setBold(False)
        font.setPixelSize(21)
        painter.setFont(font)
        metrics = painter.fontMetrics()
        painter.drawText(QRect(rect.x() + 6, top + 64, rect.width() - 12, 28), Qt.AlignmentFlag.AlignHCenter,
                         metrics.elidedText(data["author"], Qt.TextElideMode.ElideRight, rect.width() - 12))
        painter.setPen(QColor("#8b0000") if data["locked"] else QColor("black"))
        painter.drawText(QRect(rect.x() + 6, top + 94, rect.width() - 12, 28), Qt.AlignmentFlag.AlignHCenter,
                         data["progress"])
        if data.get("genre"):
            painter.setPen(QColor("#0d47a1"))
            painter.drawText(QRect(rect.x() + 6, top + 124, rect.width() - 12, 28), Qt.AlignmentFlag.AlignHCenter,
                             metrics.elidedText(data["genre"], Qt.TextElideMode.ElideRight, rect.width() - 12))
        painter.restore()


class SearchBox(QLineEdit):
    """A search box wide enough for its whole hint, measured in the words it really has
    (the window's big words apply only after it is made)."""

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        size = super().minimumSizeHint()
        return QSize(max(size.width(), self.fontMetrics().horizontalAdvance(self.placeholderText()) + 60),
                     size.height())


class LibraryPage(QWidget):
    open_requested = Signal(str, str)        # entry id, path
    add_books_requested = Signal()
    add_folder_requested = Signal()
    continue_requested = Signal()
    save_copy_requested = Signal(str, str)   # entry id, path
    free_books_requested = Signal()
    help_requested = Signal()

    def __init__(self, library: lib.Library, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.library = library
        self.filter = "All"
        self.collection = ""
        self.genre = ""
        self._covers: dict[str, QPixmap] = {}
        self._entries: dict[str, lib.LibraryEntry] = {}
        self.scan_text = ""

        self.top_buttons: dict[str, QToolButton] = {}
        top = QHBoxLayout()
        for name, glyph, tip, signal in (
                ("Continue reading", "▶", "Open the book you read last, where you left off", self.continue_requested),
                ("Add books", "➕", "Add book files to the library", self.add_books_requested),
                ("Add folder", "\U0001F4C1", "Add a folder of books to the library", self.add_folder_requested),
                ("Free books", "\U0001F4D6", "Websites where you can download books free",
                 self.free_books_requested),
                ("Help", "?", "How to use the Book Reader", self.help_requested)):
            b = ui.tool_button(name, glyph, tip, signal.emit)
            # Never narrower than its whole word: a squeezed button shows "A…s" instead of "Add books".
            # (A set minimum width would win over this, so it goes; the row scrolls sideways instead.)
            b.setMinimumWidth(0)
            b.setSizePolicy(QSizePolicy.Policy.Minimum, b.sizePolicy().verticalPolicy())
            self.top_buttons[name] = b
            top.addWidget(b)
        top.addStretch(1)
        # The buttons, "Last read", and the search each get a row: side by side they don't fit
        # a usual window with big words, and a long title beside the buttons would wrap and cut them.
        self.continue_label = QLabel("")
        self.continue_label.setWordWrap(True)
        row = QHBoxLayout()
        self.search_box = SearchBox()
        self.search_box.setPlaceholderText("Find a title or author")
        self.search_box.setMinimumHeight(52)
        self.search_box.setClearButtonEnabled(True)
        self.search_box.textChanged.connect(lambda _t: self.refresh())
        row.addWidget(self.search_box, 1)
        self.genre_button = ui.wide_button("Genre", self._genre_menu, 52, "Show only one kind of book")
        row.addWidget(self.genre_button)
        # Just "Sort": with the chosen order in it too, this row wouldn't fit a usual window.
        self.sort_button = ui.wide_button("Sort", self._sort_menu, 52, "Change the order of the covers")
        row.addWidget(self.sort_button)

        filters = QHBoxLayout()
        self.filter_buttons: dict[str, QPushButton] = {}
        for name in FILTERS:
            b = ui.wide_button(name, lambda n=name: self.set_filter(n), ui.SECOND_ROW - 6)
            b.setCheckable(True)
            self.filter_buttons[name] = b
            filters.addWidget(b)
        self.collections_button = QPushButton("Collections ▾")
        self.collections_button.setMinimumHeight(ui.SECOND_ROW - 6)
        self.collections_button.setToolTip("Show one of your collections")
        self.collections_button.clicked.connect(self._collections_menu)
        filters.addWidget(self.collections_button)
        filters.addStretch(1)

        self.grid = QListWidget()
        self.grid.setViewMode(QListView.ViewMode.IconMode)
        self.grid.setResizeMode(QListView.ResizeMode.Adjust)
        self.grid.setMovement(QListView.Movement.Static)
        self.grid.setUniformItemSizes(True)
        self.grid.setGridSize(TILE)
        self.grid.setSpacing(4)
        self.grid.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.grid.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.grid.setItemDelegate(TileDelegate(self))
        self.grid.itemSelectionChanged.connect(self._selection_changed)
        self.grid.itemDoubleClicked.connect(lambda _item: self.open_selected())
        self.grid.itemActivated.connect(lambda _item: self.open_selected())
        self.empty_message = QLabel("")
        self.empty_message.setWordWrap(True)
        self.empty_message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_message.hide()

        self.strip = QWidget()
        strip = QHBoxLayout(self.strip)
        strip.setContentsMargins(0, 0, 0, 0)
        self.selected_label = QLabel("")
        self.selected_label.setStyleSheet("font-weight: bold;")
        strip.addWidget(self.selected_label, 1)
        self.strip_buttons: dict[str, QPushButton] = {}
        for name, slot, tip in (
                ("Open", self.open_selected, "Read this book"),
                ("Favorite", self.toggle_favorite, "Mark this book as a favorite (or not)"),
                ("Finished", self.toggle_finished, "Mark this book as finished (or not)"),
                ("Collections", self._book_collections_menu, "Put this book in one of your collections"),
                ("Save a copy", self.save_copy_selected, "Save this book as EPUB, PDF, Word or plain text"),
                ("Show in folder", self.show_selected_in_folder, "Show where this book's file is"),
                ("Remove from library", self.remove_selected, "Take this book off the list (the file stays)")):
            b = ui.wide_button(name, slot, ui.SECOND_ROW, tip)
            self.strip_buttons[name] = b
            strip.addWidget(b)
        self.status = QLabel("")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.addWidget(ui.SideScroll(ui.holder(top)))
        layout.addWidget(self.continue_label)
        layout.addWidget(ui.SideScroll(ui.holder(row)))
        layout.addWidget(ui.SideScroll(ui.holder(filters)))
        layout.addWidget(self.grid, 1)
        layout.addWidget(self.empty_message, 1)
        self.strip_holder = ui.SideScroll(self.strip)
        layout.addWidget(self.strip_holder)
        layout.addWidget(self.status)
        self.filter_buttons["All"].setChecked(True)

    # ---------------------------------------------------------------- what shows
    def refresh(self, keep: str | None = None) -> None:
        keep = keep if keep is not None else (self.selected_entry().id if self.selected_entry() else "")
        entries = self.library.entries()
        self._entries = {e.id: e for e in entries}
        query = self.search_box.text().strip().casefold()
        shown = []
        for entry in entries:
            state = self.library.state(entry.id)
            if query and query not in entry.info.title.casefold() and \
                    not any(query in a.casefold() for a in entry.info.authors):
                continue
            if not self._passes(entry, state):
                continue
            shown.append((entry, state))
        key = self.library.settings.sort
        if key == "title":
            shown.sort(key=lambda es: sort_title(es[0].info.title))
        elif key == "author":
            shown.sort(key=lambda es: (not es[0].info.authors, (es[0].info.authors or [""])[0].casefold(),
                                       sort_title(es[0].info.title)))
        elif key == "genre":
            shown.sort(key=lambda es: (genres.sort_key(self.genres_of(es[0])[0]), sort_title(es[0].info.title)))
        elif key == "added":
            shown.sort(key=lambda es: -(es[1].added or es[0].added))
        else:
            shown.sort(key=lambda es: (-es[1].opened, -(es[1].added or es[0].added), sort_title(es[0].info.title)))
        self.grid.blockSignals(True)
        self.grid.clear()
        reselect = None
        for entry, state in shown:
            item = QListWidgetItem(entry.info.title)
            main_genre = self.genres_of(entry)[0]
            item.setData(Qt.ItemDataRole.UserRole, {
                "id": entry.id, "title": entry.info.title or "(no title)",
                "author": ", ".join(entry.info.authors) or "Unknown author",
                "progress": progress_text(entry, state), "locked": bool(entry.info.lock or entry.info.error),
                "favorite": state.favorite, "has_cover": self.library.cover(entry.id) is not None,
                "genre": "" if main_genre in (genres.NO_GENRE, genres.OTHER) else main_genre})
            item.setSizeHint(TILE)
            self.grid.addItem(item)
            if entry.id == keep:
                reselect = item
        self.grid.blockSignals(False)
        if reselect is not None:
            self.grid.setCurrentItem(reselect)
        self._selection_changed()
        empty = not shown
        self.grid.setVisible(not empty)
        self.empty_message.setVisible(empty)
        if empty:
            self.empty_message.setText(self._empty_text(query))
        locked = sum(1 for e in entries if e.info.lock)
        text = plural(len(entries), "book") + (f" - {locked:,} locked" if locked else "")
        if len(shown) != len(entries):
            text += f" - showing {len(shown):,}"
        self.status.setText(text + (f" - {self.scan_text}" if self.scan_text else ""))
        self._continue_text()

    def _passes(self, entry: lib.LibraryEntry, state: lib.BookState) -> bool:
        locked = bool(entry.info.lock or entry.info.error)
        if self.collection:
            return self.collection in state.collections
        if self.genre:
            return self.genre in self.genres_of(entry)
        f = self.filter
        if f == "Locked":
            return locked
        if f == "Reading":
            return not locked and not state.finished and (state.progress > 0 or state.opened > 0)
        if f == "Not started":
            return not locked and not state.finished and state.progress == 0 and not state.opened
        if f == "Finished":
            return state.finished
        if f == "Favorites":
            return state.favorite
        return True

    def _empty_text(self, query: str) -> str:
        if query:
            return f"No books have “{self.search_box.text().strip()}” in their title or author."
        if self.collection:
            return f"The collection “{self.collection}” is empty."
        if self.genre:
            return f"No books are in the genre “{self.genre}”."
        return {"Favorites": "You have no favorites yet. Click a book, then press Favorite.",
                "Finished": "No books are marked finished yet.",
                "Reading": "You aren't part-way through any book yet.",
                "Not started": "You have started every book.",
                "Locked": "None of your books are locked."}.get(
            self.filter, "There are no books here yet. Press Add books or Add folder, or drag books "
                         "onto this window.")

    def _continue_text(self) -> None:
        last = self.library.settings.last_book
        entry = self.library.entry_for_path(last) if last else None
        if last and os.path.exists(last):
            title = entry.info.title if entry else formats.title_from_file_name(last)[0]
            self.continue_label.setText(f"Last read: {title}")
        else:
            self.continue_label.setText("Click a book's cover to start reading.")

    def tiles(self) -> list[dict]:
        return [self.grid.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.grid.count())] \
            if self.grid.isVisible() or self.grid.count() else []

    def cover_for(self, data: dict) -> QPixmap:
        ident = data["id"]
        if ident not in self._covers:
            pixmap = QPixmap()
            raw = self.library.cover(ident)
            if raw is None or not pixmap.loadFromData(raw):
                pixmap = drawn_cover(data["title"], data["author"] if data["author"] != "Unknown author" else "")
            self._covers[ident] = pixmap
        return self._covers[ident]

    def forget_covers(self) -> None:
        self._covers.clear()

    # ---------------------------------------------------------------- choosing
    def set_filter(self, name: str) -> None:
        self.filter, self.collection, self.genre = name, "", ""
        for n, b in self.filter_buttons.items():
            b.setChecked(n == name)
        self.collections_button.setText("Collections ▾")
        self.genre_button.setText("Genre")
        self.refresh()

    def show_collection(self, name: str) -> None:
        self.collection, self.genre = name, ""
        for b in self.filter_buttons.values():
            b.setChecked(False)
        self.collections_button.setText(f"Collection: {name} ▾")
        self.genre_button.setText("Genre")
        self.refresh()

    def genres_of(self, entry: lib.LibraryEntry) -> list[str]:
        """The book's plain genres, its main one first (No genre when it has no tags)."""
        return genres.genres(entry.info.subjects)

    def show_genre(self, name: str) -> None:
        """Show only the books of one genre; All (or any other button below) shows every book again."""
        self.genre, self.collection = name, ""
        for b in self.filter_buttons.values():
            b.setChecked(False)
        self.collections_button.setText("Collections ▾")
        self.genre_button.setText(f"Genre: {name}")
        self.refresh()

    def _genre_menu(self) -> None:
        counts: dict[str, int] = {}
        for entry in self.library.entries():
            for genre in self.genres_of(entry):
                counts[genre] = counts.get(genre, 0) + 1
        menu = QMenu(self)
        menu.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        for genre in sorted(counts, key=genres.sort_key):
            menu.addAction(f"{genre} ({counts[genre]:,})", lambda g=genre: self.show_genre(g))
        if not counts:
            menu.addAction("There are no books yet.").setEnabled(False)
        self.genre_menu = menu
        menu.popup(self.genre_button.mapToGlobal(self.genre_button.rect().bottomLeft()))

    def _sort_menu(self) -> None:
        menu = QMenu(self)
        menu.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        for key, name in SORT_NAMES.items():
            action = menu.addAction(name, lambda k=key: self.set_sort(k))
            action.setCheckable(True)
            action.setChecked(key == self.library.settings.sort)
        self.sort_menu = menu
        menu.popup(self.sort_button.mapToGlobal(self.sort_button.rect().bottomLeft()))

    def collection_names(self) -> list[str]:
        names = set()
        for entry in self.library.entries():
            names.update(self.library.state(entry.id).collections)
        return sorted(names, key=str.casefold)

    def _collections_menu(self) -> None:
        menu = QMenu(self)
        names = self.collection_names()
        if not names:
            action = menu.addAction("You have no collections yet. Click a book, then press Collections.")
            action.setEnabled(False)
        for name in names:
            menu.addAction(name, lambda n=name: self.show_collection(n))
        menu.popup(self.collections_button.mapToGlobal(self.collections_button.rect().bottomLeft()))

    def set_sort(self, key: str) -> None:
        if key in SORT_NAMES:
            self.library.settings.sort = key
            try:
                self.library.save_settings()
            except OSError:
                pass
            self.refresh()

    def select_title(self, title: str) -> bool:
        for i in range(self.grid.count()):
            if self.grid.item(i).data(Qt.ItemDataRole.UserRole)["title"] == title:
                self.grid.setCurrentRow(i)
                return True
        return False

    def selected_entry(self) -> lib.LibraryEntry | None:
        item = self.grid.currentItem()
        if item is None or not item.isSelected():
            return None
        return self._entries.get(item.data(Qt.ItemDataRole.UserRole)["id"])

    def _selection_changed(self) -> None:
        entry = self.selected_entry()
        self.strip.setVisible(entry is not None)
        self.strip_holder.setVisible(entry is not None)
        if entry is None:
            return
        state = self.library.state(entry.id)
        self.selected_label.setText(entry.info.title)
        self.strip_buttons["Favorite"].setText("Not a favorite" if state.favorite else "Favorite")
        self.strip_buttons["Finished"].setText("Not finished" if state.finished else "Finished")

    # ---------------------------------------------------------------- the strip
    def _need_book(self) -> lib.LibraryEntry | None:
        entry = self.selected_entry()
        if entry is None:
            ui.inform(self, "Pick a book", "Click a book's cover first, then press this button.")
        return entry

    def open_selected(self) -> None:
        entry = self._need_book()
        if entry is None:
            return
        if entry.info.lock:
            explain_lock(self, entry.info.lock)
            return
        if entry.info.error:
            ui.inform(self, "This book can't be opened", entry.info.error)
            return
        self.open_requested.emit(entry.id, entry.path)

    def _change_state(self, change) -> None:
        entry = self._need_book()
        if entry is None:
            return
        state = self.library.state(entry.id)
        change(state)
        try:
            self.library.save_state(entry.id, state)
        except OSError:
            pass
        self.refresh(keep=entry.id)

    def toggle_favorite(self) -> None:
        self._change_state(lambda s: setattr(s, "favorite", not s.favorite))

    def toggle_finished(self) -> None:
        self._change_state(lambda s: setattr(s, "finished", not s.finished))

    def new_collection(self) -> None:
        entry = self._need_book()
        if entry is None:
            return
        name = ui.ask_text(self, "New collection", "Name the new collection (for example “To read”):",
                           "", multiline=False)
        if name and name.strip():
            name = name.strip()[:60]
            self._change_state(lambda s: s.collections.append(name) if name not in s.collections else None)

    def _book_collections_menu(self) -> None:
        entry = self._need_book()
        if entry is None:
            return
        state = self.library.state(entry.id)
        menu = QMenu(self)
        for name in self.collection_names():
            action = menu.addAction(name)
            action.setCheckable(True)
            action.setChecked(name in state.collections)
            action.triggered.connect(lambda on, n=name: self._change_state(
                lambda s: s.collections.append(n) if on and n not in s.collections
                else (s.collections.remove(n) if not on and n in s.collections else None)))
        menu.addSeparator()
        menu.addAction("New collection…", self.new_collection)
        button = self.strip_buttons["Collections"]
        menu.popup(button.mapToGlobal(button.rect().topLeft()))

    def save_copy_selected(self) -> None:
        entry = self._need_book()
        if entry is not None:
            self.save_copy_requested.emit(entry.id, entry.path)

    def show_selected_in_folder(self) -> None:
        entry = self._need_book()
        if entry is not None:
            ui.show_in_folder(entry.path)

    def remove_selected(self) -> None:
        entry = self._need_book()
        if entry is None:
            return
        if ui.ask(self, "Take it off the list?", f"Take “{entry.info.title}” off the library list?\n\n"
                  "The book's file itself stays where it is.", "Take it off", "Keep it"):
            self.library.remove(entry.id)
            self.refresh(keep="")


# ---------------------------------------------------------------- help

class HelpWindow(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("How to use the Book Reader")
        self.resize(1000, 760)
        layout = QVBoxLayout(self)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        try:
            with open(HELP_FILE, encoding="utf-8") as f:
                text.setPlainText(f.read())
        except OSError:
            text.setPlainText("The help file (Book Reader - How to use.txt) is missing.")
        layout.addWidget(text, 1)
        layout.addWidget(ui.wide_button("Close", self.close, 56))


# ---------------------------------------------------------------- the window

class MainWindow(QMainWindow):
    scan_progress = Signal(int)
    scan_finished = Signal()

    def __init__(self, library: lib.Library, server_name: str | None = None,
                 start_paths: list[str] | None = None) -> None:
        super().__init__()
        app = QApplication.instance()
        ui.black_words(app)
        screen = self.screen().availableGeometry() if self.screen() else QRect(0, 0, 1280, 900)
        self.setStyleSheet(ui.window_style(screen.width()))
        self.setWindowTitle("Book Reader")
        self.setAcceptDrops(True)
        self.library = library
        self.stack = QStackedWidget()
        self.library_page = LibraryPage(library)
        self.reading_page = ReadingPage(library)
        self.stack.addWidget(self.library_page)
        self.stack.addWidget(self.reading_page)
        self.setCentralWidget(self.stack)
        self.library_page.open_requested.connect(lambda ident, path: self.open_book_path(path, ident))
        self.library_page.add_books_requested.connect(self.add_books)
        self.library_page.add_folder_requested.connect(self.add_folder)
        self.library_page.continue_requested.connect(self.continue_reading)
        self.library_page.save_copy_requested.connect(self.save_copy)
        self.library_page.free_books_requested.connect(self.free_books)
        self.library_page.help_requested.connect(self.show_help)
        self.reading_page.library_requested.connect(self.show_library)
        self.reading_page.full_screen_requested.connect(self._full_screen)
        self.scan_progress.connect(self._scan_progress)
        self.scan_finished.connect(self._scan_finished)
        self._scan_thread: threading.Thread | None = None
        self._scan_busy = False
        self._scan_stop = threading.Event()
        self._rescan = False
        self._was_maximized = False
        self.help_window: HelpWindow | None = None
        self.server = QLocalServer(self)
        if server_name:
            if not self.server.listen(server_name):
                QLocalServer.removeServer(server_name)
                self.server.listen(server_name)
            self.server.newConnection.connect(self._incoming)
        self._first_size(screen)
        self.library_page.refresh()
        files = [p for p in (start_paths or []) if os.path.isfile(p)]
        folders = [p for p in (start_paths or []) if os.path.isdir(p)]
        self._add_folders(folders, rescan=False)
        if files:
            self.library.add_files(files)
        self.start_scan()
        if files:
            self.open_book_path(files[0])

    def _first_size(self, screen: QRect) -> None:
        saved = self.library.settings.window
        if saved:
            try:
                self.restoreGeometry(QByteArray.fromHex(saved.encode("ascii")))
            except (ValueError, UnicodeError):
                pass
            frame = self.frameGeometry()
            if frame.width() <= screen.width() + 16 and frame.height() <= screen.height() + 16:
                return
        self.resize(min(1400, screen.width() - 40), min(960, screen.height() - 60))
        self.move(screen.x() + max(0, (screen.width() - self.width()) // 2),
                  screen.y() + max(0, (screen.height() - self.height()) // 3))

    # ---------------------------------------------------------------- scanning
    def scanning(self) -> bool:
        """True from the start of a scan until the library list has been refreshed after it."""
        return self._scan_busy

    def start_scan(self) -> None:
        if self.scanning():
            self._rescan = True
            return
        self._scan_busy = True
        self._scan_stop = threading.Event()
        stop = self._scan_stop

        def run() -> None:
            try:
                self.library.scan(stop, lambda n: self.scan_progress.emit(n) if n % 10 == 0 else None)
            finally:
                self.scan_finished.emit()
        self.library_page.scan_text = "looking through your folders…"
        self._scan_thread = threading.Thread(target=run, daemon=True)
        self._scan_thread.start()

    def _scan_progress(self, count: int) -> None:
        self.library_page.scan_text = f"looking through your folders… {count:,} so far"
        self.library_page.status.setText(self.library_page.status.text().split(" - looking")[0]
                                          + " - " + self.library_page.scan_text)

    def _scan_finished(self) -> None:
        if self._scan_thread is not None:
            self._scan_thread.join(1)
        self._scan_thread = None
        self.library_page.scan_text = ""
        self.library_page.forget_covers()
        self.library_page.refresh()
        self._scan_busy = False
        if self._rescan:
            self._rescan = False
            self.start_scan()

    def stop_scan(self) -> None:
        self._scan_stop.set()
        if self._scan_thread is not None:
            self._scan_thread.join(10)
        self._scan_thread = None
        self._scan_busy = False

    # ---------------------------------------------------------------- books
    def open_book_path(self, path: str, ident: str | None = None) -> bool:
        if ident is None:
            entry = self.library.entry_for_path(path)
            ident = entry.id if entry else None
        before = self.stack.currentWidget()
        self.stack.setCurrentWidget(self.reading_page)      # pages are laid out at their real size
        if not self.reading_page.open_path(path, ident):
            self.stack.setCurrentWidget(before)
            return False
        title = self.reading_page.book.title if self.reading_page.book else ""
        self.setWindowTitle(f"{title} - Book Reader" if title else "Book Reader")
        return True

    def show_library(self) -> None:
        self.stack.setCurrentWidget(self.library_page)
        self.setWindowTitle("Book Reader")
        self.library_page.refresh()

    def continue_reading(self) -> None:
        last = self.library.settings.last_book
        if last and os.path.exists(last):
            self.open_book_path(last)
        else:
            ui.inform(self, "Continue reading", "You haven't opened a book yet (or it has been moved). "
                      "Click a book's cover to start.")

    def handle_paths(self, paths: list[str]) -> None:
        """Books or folders given to the window (dragged, or from a second start)."""
        files = [p for p in paths if os.path.isfile(p)]
        folders = [p for p in paths if os.path.isdir(p)]
        if files:
            self.library.add_files(files)
        self._add_folders(folders)
        if files:
            self.start_scan()
            self.open_book_path(files[0])
        if self.isMinimized():
            self.showNormal()
        self.raise_()
        self.activateWindow()

    def _add_folders(self, folders: list[str], rescan: bool = True) -> None:
        added = False
        for folder in folders:
            folder = os.path.abspath(folder)
            if folder not in self.library.settings.folders:
                self.library.settings.folders.append(folder)
                added = True
        if added:
            try:
                self.library.save_settings()
            except OSError:
                pass
            if rescan:
                self.start_scan()

    def add_books(self) -> None:
        names = ui.open_file_names(self, "Add books", os.path.join(os.path.expanduser("~"), "Documents"), BOOK_FILTER)
        if names:
            self.library.add_files(names)
            self.start_scan()

    def add_folder(self) -> None:
        folder = ui.choose_folder(self, "Add a folder of books", os.path.expanduser("~"))
        if folder:
            self._add_folders([folder])

    def save_copy(self, ident: str, path: str) -> None:
        """Save a copy of a book picked in the library (it is opened just for this)."""
        import book_reader_copy
        entry = self.library.entry_for_path(path)
        document = None
        if entry is not None and entry.info.lock:
            book_reader_copy.ask_and_save(self, formats.Book(path, entry.info.format), locked=True)
            return
        if path.lower().endswith(".pdf"):
            from PySide6.QtPdf import QPdfDocument
            from book_reader_pdf import PdfBook
            document = QPdfDocument(self)
            if document.load(path) != QPdfDocument.Error.None_:
                ui.inform(self, "Save a copy", "This PDF can't be opened, so no copy can be made.")
                return
            book = PdfBook(path, document, [])
        else:
            try:
                book = formats.open_book(path)
            except formats.BookLocked as e:
                explain_lock(self, e.info)
                return
            except formats.BookError as e:
                ui.inform(self, "This book can't be opened", str(e))
                return
        def finished(_ok=False, _message="", b=book, d=document) -> None:
            b.close()
            if d is not None:                 # the PDF opened just for this copy
                d.close()
                d.deleteLater()
        job = book_reader_copy.ask_and_save(self, book, on_done=finished)
        if job is None:
            finished()
        else:
            self._copy_job = job

    def free_books(self) -> None:
        """A menu of websites that give books away free without signing up: the one picked opens in
        the web browser; the last item says how to read what was downloaded."""
        menu = QMenu(self)
        menu.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        for name, shown, url, what in FREE_BOOKS:
            menu.addAction(f"{name} ({shown}): {what}", lambda u=url: ui.open_url(u))
        menu.addSeparator()
        menu.addAction("How to read a book you downloaded", self._free_books_help)
        self.free_books_menu = menu
        button = self.library_page.top_buttons["Free books"]
        menu.popup(button.mapToGlobal(button.rect().bottomLeft()))

    def _free_books_help(self) -> None:
        ui.inform(self, "Free books",
                  "When a website asks which kind of file you want, pick EPUB.\n\n"
                  "To read a book you downloaded, drag it from your Downloads folder onto this window.\n\n"
                  "Or press Add folder once and pick your Downloads folder: books you download then join "
                  "the library each time the Book Reader starts.")

    def show_help(self) -> None:
        if self.help_window is None:
            self.help_window = HelpWindow(self)
        self.help_window.show()
        self.help_window.raise_()

    # ---------------------------------------------------------------- window
    def _full_screen(self, on: bool) -> None:
        if on:
            self._was_maximized = self.isMaximized()
            self.showFullScreen()
        elif self._was_maximized:
            self.showMaximized()
        else:
            self.showNormal()

    def _incoming(self) -> None:
        """A second start sent its books: read every line, including what arrives with the hang-up."""
        while self.server.hasPendingConnections():
            sock = self.server.nextPendingConnection()
            buffer = bytearray()

            def take(s=sock, buf=buffer, final=False) -> None:
                buf.extend(bytes(s.readAll()))
                while b"\n" in buf or (final and buf):
                    line, _, rest = bytes(buf).partition(b"\n")
                    buf[:] = rest
                    if not final:
                        s.write(b"ok\n")
                        s.flush()
                    self._message(line)

            def gone(s=sock) -> None:
                take(s, final=True)
                s.deleteLater()
            sock.readyRead.connect(take)
            sock.disconnected.connect(gone)
            take()

    def _message(self, line: bytes) -> None:
        try:
            data = json.loads(line.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return
        paths = [p for p in data.get("open", []) if isinstance(p, str)] if isinstance(data, dict) else []
        self.handle_paths(paths)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.handle_paths(paths)

    def closeEvent(self, event) -> None:  # noqa: N802
        for job in (getattr(self, "_copy_job", None), getattr(self.reading_page, "_copy_job", None)):
            if job is not None and job.running():
                job.stop_and_wait()         # never leave Word or a half-made copy behind
        self.stop_scan()
        self.reading_page.close_book()
        if not self.isFullScreen():
            self.library.settings.window = bytes(self.saveGeometry().toHex()).decode("ascii")
        try:
            self.library.save_settings()
        except OSError:
            pass
        self.server.close()
        super().closeEvent(event)


def main(argv: list[str]) -> int:
    web.register_scheme()
    app = QApplication(argv)
    app.setApplicationName("Book Reader")
    paths = [os.path.abspath(a) for a in argv[1:] if os.path.exists(a)]
    name = os.environ.get("BOOK_READER_SERVER") or server_name()
    if send_to_running(name, paths):
        return 0
    library = lib.Library(thumbnail=make_thumbnail)
    window = MainWindow(library, name, paths)
    if window.width() >= window.screen().availableGeometry().width() - 60:
        window.showMaximized()
    else:
        window.show()
    code = app.exec()
    window.deleteLater()
    app.processEvents()
    return code


if __name__ == "__main__":
    # Leave without Python tearing Qt WebEngine down piece by piece (that can crash);
    # everything was saved when the window closed.
    os._exit(main(sys.argv))
