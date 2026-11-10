"""Save a copy: ask which kind and where, then make it with a Cancel button.

EPUB, plain text and Word copies are made on a worker thread; PDF copies are
printed by Chromium from the one-page book (PdfPrinter). PDF books are read by
Qt's PDF engine, which is used on the window's thread only.
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading

from PySide6.QtCore import QMarginsF, QObject, QUrl, Signal
from PySide6.QtGui import QPageLayout, QPageSize
from PySide6.QtWebEngineCore import (
    QWebEnginePage,
    QWebEngineProfile,
    QWebEngineSettings,
    QWebEngineUrlRequestInterceptor,
)
from PySide6.QtWidgets import QDialog, QLabel, QVBoxLayout, QWidget

import book_reader_convert as convert
import book_reader_ui as ui

FILTERS = {"EPUB": "EPUB books (*.epub)", "PDF": "PDF files (*.pdf)", "Word": "Word documents (*.docx)",
           "Plain text": "Text files (*.txt)"}


class LocalFilesOnly(QWebEngineUrlRequestInterceptor):
    def interceptRequest(self, info) -> None:  # noqa: N802
        if info.requestUrl().scheme().lower() not in ("file", "data", "qrc"):
            info.block(True)


class PdfPrinter(QObject):
    """Prints a book to a PDF file through Chromium (Letter pages, 3/4 inch margins)."""

    finished = Signal(bool, str)        # ok, the file or what went wrong

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.profile = QWebEngineProfile(self)
        self.profile.setUrlRequestInterceptor(LocalFilesOnly(self.profile))
        settings = self.profile.settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, False)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, False)
        self.page = QWebEnginePage(self.profile, self)
        self.page.loadFinished.connect(self._loaded)
        self.page.pdfPrintingFinished.connect(self._printed)
        self._out = ""

    def print_book(self, book, out: str, work_folder: str) -> None:
        self._out = out
        try:
            path = convert.single_html(book, work_folder)
        except OSError as e:
            self.finished.emit(False, f"The book couldn't be got ready ({e.strerror or e}).")
            return
        self.page.load(QUrl.fromLocalFile(path))

    def _loaded(self, ok: bool) -> None:
        if not ok:
            self.finished.emit(False, "The book couldn't be laid out for the PDF.")
            return
        layout = QPageLayout(QPageSize(QPageSize.PageSizeId.Letter), QPageLayout.Orientation.Portrait,
                             QMarginsF(0.75, 0.75, 0.75, 0.75), QPageLayout.Unit.Inch)
        self.page.printToPdf(self._out, layout)

    def _printed(self, path: str, ok: bool) -> None:
        self.finished.emit(ok, path if ok else "The PDF couldn't be written there.")


class CopyJob(QObject):
    """One copy being made, with a small window saying so and a Cancel button."""

    done = Signal(bool, str)
    _thread_done = Signal(bool, str)

    def __init__(self, parent: QWidget, book, kind: str, out: str) -> None:
        super().__init__(parent)
        self.parent_widget, self.book, self.kind, self.out = parent, book, kind, out
        self.stop = threading.Event()
        self.work = tempfile.mkdtemp(prefix="book_copy_", dir=os.path.dirname(os.path.abspath(out)) or None)
        self.window = QDialog(parent)
        self.window.setWindowTitle("Saving a copy")
        box = QVBoxLayout(self.window)
        note = " This can take a minute or two." if kind in ("Word", "PDF") else ""
        label = QLabel(f"Saving a copy as {kind}…{note}")
        label.setWordWrap(True)
        box.addWidget(label)
        box.addWidget(ui.wide_button("Cancel", self.cancel, 56))
        self._thread_done.connect(self._finish)
        self.printer: PdfPrinter | None = None
        self.thread: threading.Thread | None = None

    def start(self) -> None:
        self.window.show()
        if self.kind == "PDF":
            self.printer = PdfPrinter(self)
            self.printer.finished.connect(self._finish)
            self.printer.print_book(self.book, self.out, self.work)
        elif self.book.format == "PDF":
            self._run()                    # Qt's PDF engine: this thread only
        else:
            self.thread = threading.Thread(target=self._run, daemon=True)
            self.thread.start()

    def _run(self) -> None:
        try:
            if self.kind == "EPUB":
                convert.save_epub(self.book, self.out, self.stop)
            elif self.kind == "Plain text":
                convert.save_text(self.book, self.out)
            elif self.kind == "Word":
                page = convert.single_html(self.book, self.work, self.stop)
                convert.save_word(page, self.out, self.stop)
            self._thread_done.emit(True, self.out)
        except convert.Cancelled:
            self._thread_done.emit(False, "")
        except OSError as e:
            self._thread_done.emit(False, f"The copy couldn't be saved there ({e.strerror or e}).")
        except Exception as e:                       # Word not working, a damaged book ...
            text = "Microsoft Word didn't make the copy." if self.kind == "Word" else "The copy couldn't be made."
            self._thread_done.emit(False, f"{text} ({type(e).__name__}: {str(e)[:200]})")

    def running(self) -> bool:
        return self.window is not None

    def stop_and_wait(self, timeout: float = 120) -> None:
        """Cancel and wait until the work (and Word) has stopped; used when the window closes."""
        self.stop.set()
        if self.thread is not None:
            self.thread.join(timeout)
        self._finish(False, "")

    def cancel(self) -> None:
        self.stop.set()
        if self.printer is not None:
            self._finish(False, "")

    def _finish(self, ok: bool, message: str) -> None:
        if self.window is None:
            return
        self.window.close()
        self.window = None
        shutil.rmtree(self.work, ignore_errors=True)
        if ok:
            ui.inform(self.parent_widget, "Copy saved", f"The copy is saved as:\n{self.out}")
        elif message:
            ui.inform(self.parent_widget, "The copy wasn't saved", message)
        self.done.emit(ok, message)


def ask_and_save(parent: QWidget, book, locked: bool = False, on_done=None) -> CopyJob | None:
    """The Save a copy button: which kind, where, then make it. on_done(ok, message) is called
    when the copy is finished or cancelled (it may be at once, for a PDF's text)."""
    kinds = convert.copy_kinds(book.format, locked)
    if not kinds:
        ui.inform(parent, "Save a copy", "This book is locked with copy protection, so no copy can be made of it.")
        return None
    kind = ui.choose(parent, "Save a copy", "Save a copy of this book as:", kinds)
    if not kind:
        return None
    name = re.sub(r'[<>:"/\\|?*]', "", book.title or "Book").strip()[:80] or "Book"
    suggested = os.path.join(os.path.expanduser("~"), "Documents", name + convert.EXTENSIONS[kind])
    out = ui.save_file_name(parent, "Save a copy", suggested, FILTERS[kind])
    if not out:
        return None
    if os.path.normcase(os.path.abspath(out)) == os.path.normcase(os.path.abspath(book.path)):
        ui.inform(parent, "Save a copy", "That is the book itself. Choose another name for the copy.")
        return None
    job = CopyJob(parent, book, kind, out)
    if on_done is not None:
        job.done.connect(on_done)
    job.start()
    return job
