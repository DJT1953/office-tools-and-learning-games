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

from PySide6.QtCore import Qt, QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QFont
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
    ("show_changes", "Also save a copy that shows each change, to accept or reject in Word", False),
    ("subfolders", "Include the folders inside any folder I add", False),
)
WORD_ONLY = ("use_word", "show_changes")        # greyed out when Word is not installed


def settings_file():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "Proofreader", "settings.json")


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
    finished_ok = Signal(list, list, bool)      # results, report files, Word was missing
    failed = Signal(str)

    def __init__(self, files, settings, download, has_word=True):
        super().__init__()
        self.files = files
        self.settings = settings
        self.download = download
        self.has_word = has_word
        self.stopping = False

    def run(self):
        checker = word = None
        try:
            if self.download:
                tools.install(lambda label, part: self.step.emit(int(part * 1000), 1000, label))
            self.step.emit(0, 0, "Starting the free checker...")
            checker = tools.LanguageTool().start()
            needs_word = self.settings.use_word or any(
                not f.lower().endswith(".txt") for f in self.files)
            if needs_word and self.has_word:
                self.step.emit(0, 0, "Starting Microsoft Word...")
                try:
                    import proofreader_word
                    word = proofreader_word.Word()
                except Exception:
                    word = None         # installed but would not start: go on without it
            missing = needs_word and word is None
            results = job.run(self.files, self.settings, checker, word,
                              lambda done, total, name: self.step.emit(
                                  done, total, f"Proofreading {name}" if name else "Finishing..."),
                              lambda: self.stopping)
            reports = job.save_report(results, word_missing=missing) if results else []
            self.finished_ok.emit(results, reports, missing)
        except Exception as error:
            self.failed.emit(f"{error}\n\n{traceback.format_exc()[-900:]}")
        finally:
            if word is not None:
                word.close()
            if checker is not None:
                checker.stop()


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
        self.worker = None
        self.reports = []
        self.folders = []
        self.build()
        self.add_paths(paths)
        self.refresh()

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
        self.start_button = self.button("Start proofreading", self.start, go)
        self.stop_button = self.button("Stop", self.stop, go)
        bold = QFont(self.start_button.font())
        bold.setBold(True)
        self.start_button.setFont(bold)
        self.start_button.setMinimumHeight(44)
        self.stop_button.setMinimumHeight(44)
        go.setStretch(0, 3)
        go.setStretch(1, 1)
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
                            ticked["subfolders"])

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
        self.stop_button.setEnabled(busy)
        self.open_folder_button.setEnabled(bool(self.folders))
        self.open_report_button.setEnabled(bool(self.reports))

    def start(self):
        settings = self.settings()
        if not any(settings.options.wants(key) for key, _ in BOXES):
            QMessageBox.information(self, TITLE, "Tick at least one thing to fix.")
            return
        files = job.collect(self.paths(), settings.subfolders)
        if not files:
            QMessageBox.information(self, TITLE, "There are no .doc, .docx, .rtf or .txt "
                                                 "documents in what you added.")
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
        self.worker = Worker(files, settings, download, self.has_word)
        self.worker.step.connect(self.show_step)
        self.worker.finished_ok.connect(self.done)
        self.worker.failed.connect(self.went_wrong)
        self.worker.finished.connect(self.worker_gone)
        self.worker.start()
        self.refresh()

    def stop(self):
        if self.worker:
            self.worker.stopping = True
            self.status.setText("Stopping after the document in hand...")

    def show_step(self, done, total, text):
        self.bar.setRange(0, total)             # 0 to 0 shows a moving bar
        self.bar.setValue(done)
        self.status.setText(text + (f"   ({done + 1} of {total})"
                                    if text.startswith("Proofreading") else ""))

    def done(self, results, reports, word_missing):
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
        if self.worker and self.worker.stopping:
            text = "Stopped early. " + text
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
            self.worker.stopping = True
            self.status.setText("Finishing the document in hand before closing...")
            QApplication.processEvents()
            if not self.worker.wait(120000):
                # closing now would cut Word off in the middle of a document
                self.status.setText("Still finishing a long document. Please close again "
                                    "in a moment.")
                event.ignore()
                return
        self.remember()
        event.accept()


def main():
    app = QApplication(sys.argv)
    font = app.font()
    font.setPointSize(max(font.pointSize(), 11))
    app.setFont(font)
    window = Window(sys.argv[1:])
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
