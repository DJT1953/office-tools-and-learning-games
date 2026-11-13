"""
proofreader.py - the Proofreader's window.

Add Word, RTF and text documents (or whole folders), tick what to fix, press
Start. A free checker that runs only on this computer (LanguageTool) goes
first, then Microsoft Word's own checker. Corrected copies are saved in a
"corrected" folder beside the originals, with a report of every change. The
originals are never changed, and swear words are never touched.

Started by Run_Proofreader.bat. See "Proofreader - How to use.txt".
"""

import importlib.util
import json
import os
import subprocess
import sys
import threading
import time
import traceback

TITLE = "Proofreader"
NEEDS = {"PySide6": "PySide6", "comtypes": "comtypes", "docx": "python-docx"}


def tk_message(kind, text):
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    answer = getattr(messagebox, kind)(TITLE, text)
    root.destroy()
    return answer


def show_crash(kind, value, tb):
    """Started without a console, a failure would otherwise vanish unseen."""
    text = "".join(traceback.format_exception(kind, value, tb))
    try:
        tk_message("showerror", "Something went wrong:\n\n" + text[-1500:])
    except Exception:
        print(text)


def install_missing():
    missing = [pip for module, pip in NEEDS.items() if importlib.util.find_spec(module) is None]
    if not missing:
        return
    if not tk_message("askyesno", f"This copy of Python still needs: {', '.join(missing)}.\n\n"
                                  "Download and install now? This only happens once."):
        sys.exit()
    result = subprocess.run([sys.executable, "-m", "pip", "install"] + missing,
                            stderr=subprocess.PIPE, text=True,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if result.returncode:
        raise RuntimeError("Could not install " + ", ".join(missing) + ":\n" + result.stderr[-800:])
    importlib.invalidate_caches()


sys.excepthook = show_crash
if __name__ == "__main__":
    install_missing()

from PySide6.QtCore import Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QDialog, QDialogButtonBox, QFileDialog, QGroupBox, QHBoxLayout,
    QLabel, QListWidget, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QVBoxLayout,
    QWidget)

import proofreader_job as job
import proofreader_tools as tools
from proofreader_engine import Options

FILE_FILTER = "Documents (*.doc *.docx *.rtf *.txt)"
BOXES = (
    ("spelling", "Spelling"),
    ("grammar", "Grammar"),
    ("punctuation", "Punctuation and spacing"),
    ("capitals", "Capital letters"),
    ("runons", "Run-on sentences (fix the clear ones, list the rest)"),
)
CHOICES = (
    ("use_word", "Have Microsoft Word check after the free checker", True),
    ("word_grammar", "Also list sentences Word's grammar check questions (much slower)", False),
    ("show_changes", "Also save a copy that shows each change, to accept or reject in Word", False),
    ("subfolders", "Include the folders inside any folder I add", False),
)
WORD_ONLY = ("use_word", "word_grammar", "show_changes")    # greyed out without Word


def settings_file():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "Proofreader", "settings.json")


def progress_file():
    """Where a batch's progress is kept, so it can be picked up next time."""
    return os.path.join(os.path.dirname(settings_file()), "unfinished batch.jsonl")


def time_left(seconds):
    """A rough "about 20 minutes left", or "" while it is too early to tell."""
    if seconds is None:
        return ""
    minutes = round(seconds / 60)
    if minutes < 1:
        return "less than a minute left"
    if minutes < 90:
        return f"about {minutes} minute{'s' if minutes != 1 else ''} left"
    return f"about {seconds / 3600:.1f} hours left"


def load_settings():
    try:
        with open(settings_file(), encoding="utf-8") as f:
            saved = json.load(f)
        return saved if isinstance(saved, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(values):
    try:
        os.makedirs(os.path.dirname(settings_file()), exist_ok=True)
        with open(settings_file(), "w", encoding="utf-8") as f:
            json.dump(values, f, indent=1)
    except OSError:
        pass


def parse_words(text):
    """Words typed one per line or separated by commas."""
    words = []
    for piece in text.replace(",", "\n").splitlines():
        piece = piece.strip()
        if piece and piece.lower() not in (w.lower() for w in words):
            words.append(piece)
    return words


class Worker(QThread):
    """Does the proofreading away from the window so it stays responsive."""
    step = Signal(int, int, str)        # done, total, what is happening
    note = Signal(str)                  # words for the line under the bar, nothing more
    paused = Signal(bool)               # True while it is waiting for Resume
    finished_ok = Signal(list, list, bool, bool)    # results, reports, Word missing, all done
    failed = Signal(str)

    def __init__(self, files, settings, download, has_word=True, journal=None, earlier=(),
                 word_missing_before=False):
        super().__init__()
        self.files = files
        self.settings = settings
        self.download = download
        self.has_word = has_word
        self.journal = journal
        self.earlier = list(earlier)            # documents done before a pick-up
        self.word_missing_before = word_missing_before
        self.stopping = False
        self.carry_on = threading.Event()
        self.carry_on.set()
        self.started_at = None
        self.needs_word = False
        self.word_missing = False
        self.parts_begun = 0

    # ------------------------------------------------------------ pausing
    def pause(self):
        self.carry_on.clear()

    def resume(self):
        self.carry_on.set()

    def stop(self):
        self.stopping = True
        self.carry_on.set()                     # a paused run must wake up to stop

    def wait_while_paused(self):
        """Called between steps. Word and the checker stay ready meanwhile."""
        if self.carry_on.is_set():
            return
        self.paused.emit(True)
        held = time.monotonic()
        while not self.carry_on.wait(0.25):
            pass
        if self.started_at is not None:         # time spent paused is not working time
            self.started_at += time.monotonic() - held
        self.paused.emit(False)

    def report_step(self, done, total, name):
        if self.started_at is None:
            self.started_at = time.monotonic()  # timing starts with the first document
        before = len(self.earlier)
        left = None
        if done >= 2 and self.started_at is not None:
            per_document = (time.monotonic() - self.started_at) / done
            left = per_document * (total - done)
        text = f"Proofreading {name}" if name else "Finishing..."
        if name and left is not None:
            text += f"   - {time_left(left)}"
        self.step.emit(before + done, before + total, text)

    # ------------------------------------------------------------ Word and the checker
    def open_tools(self):
        """Start the free checker and, when wanted, Word. Done afresh for every
        50 documents, so the memory they gather is handed back."""
        first = self.parts_begun == 0
        self.parts_begun += 1
        if first:
            self.step.emit(0, 0, "Starting the free checker...")
        else:
            self.note.emit("Freeing memory: starting Word and the free checker afresh "
                           "(a few seconds every 50 documents)...")
        checker = tools.LanguageTool().start()
        word = None
        if self.needs_word and self.has_word:
            if first:
                self.step.emit(0, 0, "Starting Microsoft Word...")
            try:
                import proofreader_word
                word = proofreader_word.Word()
            except Exception:
                word = None         # installed but would not start: go on without it
        if self.needs_word and word is None and not self.word_missing:
            self.word_missing = True
            if self.journal:
                self.journal.word_was_missing()
        return checker, word

    @staticmethod
    def close_tools(checker, word):
        try:
            if word is not None:
                word.close()
        finally:
            checker.stop()

    def run(self):
        try:
            if self.download:
                tools.install(lambda label, part: self.step.emit(int(part * 1000), 1000, label))
            self.needs_word = self.settings.use_word or any(
                not f.lower().endswith(".txt") for f in self.files)
            new = job.run_in_parts(self.files, self.settings, self.open_tools, self.close_tools,
                                   self.report_step, lambda: self.stopping,
                                   self.wait_while_paused, self.journal)
            missing = self.word_missing or self.word_missing_before
            all_done = len(new) == len(self.files) and not self.stopping
            if all_done and self.journal:
                self.journal.finish()           # nothing left to pick up
            results = self.earlier + new
            reports = job.save_report(results, word_missing=missing) if results else []
            self.finished_ok.emit(results, reports, missing, all_done)
        except Exception as error:
            self.failed.emit(f"{error}\n\n{traceback.format_exc()[-900:]}")


class KeepWords(QDialog):
    def __init__(self, parent, words):
        super().__init__(parent)
        self.setWindowTitle("Words to leave alone")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Names and special words that should never be changed.\n"
                                "One on each line. Capital or small letters make no difference."))
        self.box = QPlainTextEdit("\n".join(words))
        layout.addWidget(self.box)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(460, 420)

    def words(self):
        return parse_words(self.box.toPlainText())


class Window(QWidget):
    def __init__(self, paths=()):
        super().__init__()
        self.setWindowTitle(TITLE)
        self.setAcceptDrops(True)
        self.saved = load_settings()
        self.keep_words = list(self.saved.get("keep_words", []))
        self.has_word = job.word_installed()
        self.journal = job.Journal(progress_file())
        self.worker = None
        self.reports = []
        self.folders = []
        self.build()
        self.add_paths(paths)
        self.refresh()
        QTimer.singleShot(0, self.offer_pick_up)    # once the window is on screen

    # ------------------------------------------------------------ layout
    def build(self):
        outer = QVBoxLayout(self)
        outer.addWidget(QLabel("Documents and folders to proofread  "
                               "(Word .doc and .docx, .rtf, .txt). You can also drag them here."))
        top = QHBoxLayout()
        self.list = QListWidget()
        self.list.setSelectionMode(QListWidget.ExtendedSelection)
        self.list.itemSelectionChanged.connect(self.refresh)
        top.addWidget(self.list, 1)
        side = QVBoxLayout()
        self.add_files_button = self.button("Add files...", self.pick_files, side)
        self.add_folder_button = self.button("Add a folder...", self.pick_folder, side)
        self.remove_button = self.button("Remove from list", self.remove_selected, side)
        self.clear_button = self.button("Clear the list", self.clear_list, side)
        side.addStretch(1)
        top.addLayout(side)
        outer.addLayout(top, 2)

        row = QHBoxLayout()
        fix = QGroupBox("What to fix")
        fix_layout = QVBoxLayout(fix)
        self.boxes = {}
        for key, label in BOXES:
            box = QCheckBox(label)
            box.setChecked(bool(self.saved.get(key, True)))
            fix_layout.addWidget(box)
            self.boxes[key] = box
        row.addWidget(fix, 1)
        how = QGroupBox("How")
        how_layout = QVBoxLayout(how)
        for key, label, default in CHOICES:
            box = QCheckBox(label)
            box.setChecked(bool(self.saved.get(key, default)))
            how_layout.addWidget(box)
            self.boxes[key] = box
        self.no_word_label = QLabel(
            "Microsoft Word was not found on this computer, so its check is skipped. "
            ".docx and .txt documents are still corrected by the free checker. "
            ".doc and .rtf documents need Word.")
        self.no_word_label.setWordWrap(True)
        self.no_word_label.setVisible(not self.has_word)
        how_layout.addWidget(self.no_word_label)
        if not self.has_word:
            for key in WORD_ONLY:
                self.boxes[key].setChecked(False)
        self.keep_button = self.button("Words to leave alone...", self.edit_keep_words, how_layout)
        how_layout.addStretch(1)
        row.addWidget(how, 1)
        outer.addLayout(row)

        promise = QLabel("Your original documents are never changed. Swear words are never "
                         "changed or flagged.")
        promise.setWordWrap(True)
        outer.addWidget(promise)

        go = QHBoxLayout()
        self.start_button = self.button("Start proofreading", lambda: self.start(), go)
        self.pause_button = self.button("Pause", self.toggle_pause, go)
        self.stop_button = self.button("Stop", self.stop, go)
        bold = QFont(self.start_button.font())
        bold.setBold(True)
        self.start_button.setFont(bold)
        for button in (self.start_button, self.pause_button, self.stop_button):
            button.setMinimumHeight(44)
        go.setStretch(0, 3)
        go.setStretch(1, 1)
        go.setStretch(2, 1)
        outer.addLayout(go)

        self.bar = QProgressBar()
        self.bar.setTextVisible(False)
        outer.addWidget(self.bar)
        self.status = QLabel("Add documents, then press Start proofreading.")
        self.status.setWordWrap(True)
        outer.addWidget(self.status)

        self.report = QPlainTextEdit()
        self.report.setReadOnly(True)
        self.report.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.report.setFont(QFont("Consolas", 11))
        self.report.setPlaceholderText("The report of what was corrected appears here.")
        outer.addWidget(self.report, 3)

        after = QHBoxLayout()
        self.open_folder_button = self.button("Open the corrected folder", self.open_folder, after)
        self.open_report_button = self.button("Open the report", self.open_report, after)
        after.addStretch(1)
        outer.addLayout(after)
        self.resize(980, 860)

    def button(self, text, action, layout):
        button = QPushButton(text)
        button.setMinimumHeight(34)
        button.clicked.connect(action)
        layout.addWidget(button)
        return button

    # ------------------------------------------------------------ the list
    def paths(self):
        return [self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())]

    def add_paths(self, paths):
        have = {os.path.normcase(p) for p in self.paths()}
        skipped = 0
        for path in paths:
            path = os.path.abspath(path)
            if os.path.normcase(path) in have:
                continue
            if os.path.isdir(path):
                label = path + "   (folder)"
            elif os.path.isfile(path) and job.is_document(path):
                label = path
            else:
                skipped += 1
                continue
            self.list.addItem(label)
            self.list.item(self.list.count() - 1).setData(Qt.UserRole, path)
            have.add(os.path.normcase(path))
        if skipped:
            self.status.setText(f"{skipped} item(s) were not added: only .doc, .docx, .rtf and "
                                ".txt documents can be proofread.")
        self.refresh()

    def start_folder(self):
        folder = self.saved.get("last_folder", "")
        return folder if os.path.isdir(folder) else os.path.expanduser("~\\Documents")

    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add documents", self.start_folder(),
                                                FILE_FILTER)
        if files:
            self.saved["last_folder"] = os.path.dirname(files[0])
            self.add_paths(files)

    def pick_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Add a folder of documents",
                                                  self.start_folder())
        if folder:
            self.saved["last_folder"] = folder
            self.add_paths([folder])

    def remove_selected(self):
        for item in self.list.selectedItems():
            self.list.takeItem(self.list.row(item))
        self.refresh()

    def clear_list(self):
        self.list.clear()
        self.refresh()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and not self.busy():
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.add_paths([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])

    def edit_keep_words(self):
        dialog = KeepWords(self, self.keep_words)
        if dialog.exec() == QDialog.Accepted:
            self.keep_words = dialog.words()
            self.remember()

    # ------------------------------------------------------------ settings
    def settings(self):
        ticked = {key: box.isChecked() for key, box in self.boxes.items()}
        options = Options(keep_words=tuple(self.keep_words),
                          **{key: ticked[key] for key, _ in BOXES})
        return job.Settings(options, ticked["use_word"], ticked["show_changes"],
                            ticked["subfolders"], ticked["word_grammar"])

    def remember(self):
        # without Word its boxes are forced off; keep what the user last chose
        self.saved.update({key: box.isChecked() for key, box in self.boxes.items()
                           if self.has_word or key not in WORD_ONLY})
        self.saved["keep_words"] = self.keep_words
        save_settings(self.saved)

    # ------------------------------------------------------------ running
    def busy(self):
        return self.worker is not None

    def refresh(self):
        busy = self.busy()
        for widget in (self.add_files_button, self.add_folder_button, self.keep_button):
            widget.setEnabled(not busy)
        for key, box in self.boxes.items():
            box.setEnabled(not busy and (self.has_word or key not in WORD_ONLY))
        self.remove_button.setEnabled(not busy and bool(self.list.selectedItems()))
        self.clear_button.setEnabled(not busy and self.list.count() > 0)
        self.start_button.setEnabled(not busy and self.list.count() > 0)
        stopping = busy and self.worker.stopping
        self.pause_button.setEnabled(busy and not stopping)
        if not busy:
            self.pause_button.setText("Pause")
        self.stop_button.setEnabled(busy and not stopping)
        self.open_folder_button.setEnabled(bool(self.folders))
        self.open_report_button.setEnabled(bool(self.reports))

    def start(self, unfinished=None):
        """Start a new batch, or carry on with an unfinished one."""
        if unfinished:
            settings, files = unfinished.settings, unfinished.left
        else:
            settings = self.settings()
            if not any(settings.options.wants(key) for key, _ in BOXES):
                QMessageBox.information(self, TITLE, "Tick at least one thing to fix.")
                return
            files = job.collect(self.paths(), settings.subfolders)
            if not files:
                QMessageBox.information(self, TITLE, "There are no .doc, .docx, .rtf or .txt "
                                                     "documents in what you added.")
                return
            waiting = self.journal.load()
            if waiting and not self.replace_unfinished(waiting):
                return
        download = not tools.installed()
        if download and QMessageBox.question(
                self, TITLE,
                "The free checker (LanguageTool) has to be downloaded once, with its own copy "
                f"of Java: about {tools.DOWNLOAD_MB} MB, from languagetool.org and adoptium.net."
                "\n\nIt goes into its own folder and nothing else on the computer is changed. "
                "After that the Proofreader works without the internet.\n\nDownload it now?"
                ) != QMessageBox.Yes:
            return
        self.remember()
        self.report.clear()
        self.reports, self.folders = [], []
        if unfinished:
            job.tidy_half_done(unfinished.half_done)
            self.worker = Worker(files, settings, download, self.has_word, self.journal,
                                 unfinished.done, unfinished.word_missing)
        else:
            self.journal.begin(self.paths(), files, settings)
            self.worker = Worker(files, settings, download, self.has_word, self.journal)
        self.worker.step.connect(self.show_step)
        self.worker.paused.connect(self.show_paused)
        self.worker.note.connect(self.status.setText)
        self.worker.finished_ok.connect(self.done)
        self.worker.failed.connect(self.went_wrong)
        self.worker.finished.connect(self.worker_gone)
        self.worker.start()
        self.refresh()

    def stop(self):
        if self.worker:
            self.worker.stop()
            self.status.setText("Stopping... Your place is saved, so next time you can pick "
                                "up where you left off.")
            self.refresh()

    def toggle_pause(self):
        if not self.worker:
            return
        if self.pause_button.text() == "Pause":
            self.worker.pause()
            self.pause_button.setText("Resume")
            self.status.setText("Pausing after the step in hand...")
        else:
            self.worker.resume()
            self.pause_button.setText("Pause")
            self.status.setText("Carrying on...")

    def show_paused(self, paused):
        if paused:
            self.status.setText(f"Paused after {self.bar.value()} of {self.bar.maximum()} "
                                "documents. Press Resume to carry on. You can also close the "
                                "Proofreader and pick up where you left off next time.")

    # ------------------------------------------------------------ picking up
    def offer_pick_up(self):
        """If a batch was left unfinished last time, offer to carry on with it."""
        if self.busy():
            return
        unfinished = self.journal.load()
        if not unfinished:
            return
        choice = self.ask_pick_up(unfinished)
        if choice == "pick up":
            self.pick_up(unfinished)
        elif choice == "forget":
            self.journal.finish()
            self.status.setText("The unfinished batch was put aside. Its corrected copies "
                                "stay where they are.")
        else:
            self.status.setText(f"Your unfinished batch ({len(unfinished.left)} documents "
                                "still to do) is kept and will be offered again next time.")

    def ask_pick_up(self, unfinished):
        """The question shown when the Proofreader opens. Returns "pick up",
        "later" or "forget"."""
        box = QMessageBox(self)
        box.setWindowTitle(TITLE)
        box.setIcon(QMessageBox.Question)
        done = len(unfinished.files) - len(unfinished.left)
        box.setText("Pick up where you left off last time?")
        box.setInformativeText(
            f"The batch you started on {unfinished.when()} stopped after {done} of "
            f"{len(unfinished.files)} documents.\n\n"
            f"Picking up proofreads the {len(unfinished.left)} documents still to do, with the "
            "same choices as before. The ones already done are not done again.")
        pick = box.addButton("Pick up where I left off", QMessageBox.AcceptRole)
        box.addButton("Not now", QMessageBox.RejectRole)
        forget = box.addButton("Put it aside for good", QMessageBox.DestructiveRole)
        box.setDefaultButton(pick)
        box.exec()
        clicked = box.clickedButton()
        return "pick up" if clicked is pick else "forget" if clicked is forget else "later"

    def replace_unfinished(self, waiting):
        """Starting a new batch while one is unfinished: True to go ahead."""
        return QMessageBox.question(
            self, TITLE,
            f"You have an unfinished batch from {waiting.when()}, with {len(waiting.left)} of "
            f"{len(waiting.files)} documents still to do.\n\n"
            "Starting a new batch means that one can no longer be picked up. "
            "(Its corrected copies stay where they are.)\n\nStart the new batch?"
            ) == QMessageBox.Yes

    def pick_up(self, unfinished):
        self.list.clear()
        self.add_paths(unfinished.sources)
        options = unfinished.settings.options
        for key, _ in BOXES:
            self.boxes[key].setChecked(options.wants(key))
        self.boxes["use_word"].setChecked(unfinished.settings.use_word and self.has_word)
        self.boxes["show_changes"].setChecked(unfinished.settings.show_changes and self.has_word)
        self.boxes["word_grammar"].setChecked(unfinished.settings.word_grammar and self.has_word)
        self.boxes["subfolders"].setChecked(unfinished.settings.subfolders)
        self.start(unfinished)

    def show_step(self, done, total, text):
        self.bar.setRange(0, total)             # 0 to 0 shows a moving bar
        self.bar.setValue(done)
        self.status.setText(text + (f"   ({done + 1} of {total})"
                                    if text.startswith("Proofreading") else ""))

    def done(self, results, reports, word_missing, all_done=True):
        self.reports = reports
        self.folders = sorted({os.path.dirname(r.output) for r in results if r.output})
        self.report.setPlainText(job.report_text(results, word_missing=word_missing))
        fixed = sum(len(r.changes) for r in results)
        copies = sum(1 for r in results if r.output)
        failed = sum(1 for r in results if r.error)
        text = (f"Finished. {fixed} correction(s) in {copies} of {len(results)} document(s). "
                "Corrected copies are in the \"corrected\" folder beside the originals.")
        if failed:
            text += f"  {failed} document(s) could not be corrected; see the report."
        if word_missing:
            text += "  Microsoft Word was not available, so its check was skipped."
        if not all_done:
            left = self.bar.maximum() - len(results)
            text = (f"Stopped with {left} document(s) still to do. Your place is saved: open "
                    "the Proofreader again any time to pick up where you left off.  " + text)
        self.status.setText(text)

    def went_wrong(self, text):
        self.status.setText("The proofreading could not be done.")
        QMessageBox.critical(self, TITLE, text)

    def worker_gone(self):
        self.worker = None
        self.bar.setRange(0, 1)
        self.bar.setValue(0)
        self.refresh()

    def open_folder(self):
        for folder in self.folders[:3]:
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    def open_report(self):
        for report in self.reports[:3]:
            QDesktopServices.openUrl(QUrl.fromLocalFile(report))

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()
            self.status.setText("Saving your place before closing...")
            QApplication.processEvents()
            if not self.worker.wait(120000):
                # closing now would cut Word off in the middle of a document
                self.status.setText("Still finishing a long document. Please close again "
                                    "in a moment.")
                event.ignore()
                return
        self.remember()
        event.accept()


def black_words(app):
    """Every word black, not Windows' soft grey-black, whether the window is in
    front or not. Greyed-out buttons stay grey: that shows they can't be
    pressed just now."""
    palette = app.palette()
    roles = (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text,
             QPalette.ColorRole.ButtonText, QPalette.ColorRole.PlaceholderText,
             QPalette.ColorRole.ToolTipText, QPalette.ColorRole.BrightText)
    for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive):
        for role in roles:
            palette.setColor(group, role, QColor("black"))
    app.setPalette(palette)


def main():
    app = QApplication(sys.argv)
    font = app.font()
    font.setPointSize(max(font.pointSize(), 11))
    app.setFont(font)
    black_words(app)
    window = Window(sys.argv[1:])
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
