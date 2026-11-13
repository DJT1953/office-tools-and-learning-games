"""Song Cleaner: tidies a folder of songs recorded from the radio.

1. Scan reads every MP3 in the folder (nothing is changed).
2. The pages then show what it proposes: better names, copies of the same
   song to remove, and talk or noise to trim from the ends. Untick anything
   you do not want. Play buttons let you hear before deciding.
3. Apply does the ticked things. Nothing is deleted: removed copies and the
   untouched originals of trimmed songs go to folders beside the music
   folder, and Undo last Apply puts everything back.

The first time, run Install_SongCleaner.bat. Then start it with Run_SongCleaner.bat.
"""

import os
import sys
import traceback

sys.dont_write_bytecode = True

if __name__ == "__main__":
    # Started without the add-ons (say, by double-clicking this file before running
    # Install_SongCleaner.bat): offer to fetch them before anything needs them.
    import song_cleaner_setup
    song_cleaner_setup.offer_missing()

from PySide6.QtCore import QLockFile, QSettings, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QColor, QPalette
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtWidgets import (QApplication, QDialog, QDialogButtonBox, QFileDialog, QHBoxLayout,
                               QInputDialog, QLabel, QLineEdit, QMainWindow, QMessageBox,
                               QPlainTextEdit, QProgressBar, QPushButton, QTabWidget,
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

import song_cleaner_apply as applying
import song_cleaner_engine as engine
import song_cleaner_job as job
import song_cleaner_names as names
import song_cleaner_online as online

TEXT_PT = 20
NUDGE_S = 0.5
PLAY_S = 8.0
STYLE = f"""
QWidget {{ font-size: {TEXT_PT}pt; color: black; }}
QPushButton {{ padding: 8px 16px; }}
QPushButton:disabled {{ color: #777777; }}
QTreeWidget {{ alternate-background-color: #f0f4f8; }}
QTabBar::tab {{ padding: 10px 18px; }}
QTreeView::indicator {{ width: 28px; height: 28px; }}
QMessageBox QLabel, QInputDialog QLabel {{ min-width: 700px; }}
"""
HELP = __doc__.strip() + """

The pages

Names: the name now, and the name proposed. "Edit name" lets you type your own.
A name marked doubtful could not be worked out; "Look up doubtful names online"
can ask a free music database (it needs a free key, see Settings).

Copies: files that are the same recording. One is marked KEEP; the others are
ticked to be removed. "Keep this one instead" changes which. Where the program
cannot tell which copy is fuller, nothing is ticked: listen and choose.

Trims: cuts the program is sure of, because two signs agree (a gap in the
sound, a change in the sound, another copy of the song).

Listen and decide: cuts it is not sure of. Often what looks extra is really
more of the song. Play, and tick only the ones you want.

Problems: files that could not be read, songs with the same name that are
different recordings, and anything Apply had to skip.

Trimming never lowers the sound quality. A cut cannot be faded, so it lands
on the quietest moment nearby.
"""


class Worker(QThread):
    """Runs one long job off the window's thread."""
    progress = Signal(int, int, str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, work, parent=None):
        super().__init__(parent)
        self.work = work
        self.stop_asked = False

    def run(self):
        try:
            result = self.work(lambda d, t, text="": self.progress.emit(d, t, text),
                               lambda: self.stop_asked)
            self.done.emit(result)
        except engine.SongError as error:
            self.failed.emit(str(error))
        except Exception:  # noqa: BLE001 - anything else is a fault in the program
            self.failed.emit("Something went wrong inside the program:\n\n" + traceback.format_exc(limit=4))


class Page(QWidget):
    """A list with a row of worded buttons under it."""

    def __init__(self, headers, buttons, window):
        super().__init__()
        self.window = window
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(headers)
        self.tree.setUniformRowHeights(True)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(False)
        self.tree.itemChanged.connect(self._changed)
        self.tree.itemSelectionChanged.connect(window.stop_playing)
        self.buttons = {}
        row = QHBoxLayout()
        for label, action in buttons:
            button = QPushButton(label)
            button.clicked.connect(action)
            row.addWidget(button)
            self.buttons[label] = button
        row.addStretch(1)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.addWidget(self.tree, 1)
        layout.addLayout(row)
        self.filling = False
        self.on_tick = None
        self.on_tick_all = None

    def _changed(self, item, column):
        if not self.filling and column == 0 and self.on_tick and item.data(0, Qt.UserRole) is not None:
            self.on_tick(item.data(0, Qt.UserRole), item.checkState(0) == Qt.Checked)

    def fill(self, rows):
        """rows: (key, ticked or None for no box, [texts], tooltip)."""
        self.filling = True
        self.tree.clear()
        for key, ticked, texts, tip in rows:
            item = QTreeWidgetItem(texts)
            item.setData(0, Qt.UserRole, key)
            if ticked is not None:
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(0, Qt.Checked if ticked else Qt.Unchecked)
            for column in range(len(texts)):
                item.setToolTip(column, tip or texts[column])
            self.tree.addTopLevelItem(item)
        for column in range(self.tree.columnCount()):
            self.tree.resizeColumnToContents(column)
            self.tree.setColumnWidth(column, min(self.tree.columnWidth(column) + 30, 1100))
        self.filling = False

    def current(self):
        item = self.tree.currentItem()
        return item.data(0, Qt.UserRole) if item else None

    def select(self, key):
        for index in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(index)
            if item.data(0, Qt.UserRole) == key:
                self.tree.setCurrentItem(item)
                return item
        return None

    def set_all(self, ticked):
        """Tick or untick every row, telling the window once at the end (not once a row)."""
        keys = []
        self.filling = True
        for index in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(index)
            if item.flags() & Qt.ItemIsUserCheckable and item.data(0, Qt.UserRole) is not None:
                item.setCheckState(0, Qt.Checked if ticked else Qt.Unchecked)
                keys.append(item.data(0, Qt.UserRole))
        self.filling = False
        if self.on_tick_all:
            self.on_tick_all(keys, ticked)


class SettingsBox(QDialog):
    def __init__(self, window):
        super().__init__(window)
        self.setWindowTitle("Settings")
        words = QLabel(
            "Looking up doubtful names online uses AcoustID, a free public music database.\n"
            "It needs a free key:\n\n"
            f"   1. Open  {online.KEY_PAGE}  in your web browser and sign in.\n"
            "   2. Register an application (any name, such as Song Cleaner).\n"
            "   3. Copy the API key it shows and paste it here.\n\n"
            "The key is kept only on this computer. Leave the box empty to stay offline.")
        words.setWordWrap(True)
        words.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.key = QLineEdit(window.settings.value("acoustid key", "") or "")
        self.key.setPlaceholderText("AcoustID key")
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout(self)
        layout.addWidget(words)
        layout.addWidget(self.key)
        layout.addWidget(buttons)


class MainWindow(QMainWindow):
    def __init__(self, settings_path=None, cache_folder=None, log_path=None):
        super().__init__()
        self.setWindowTitle("Song Cleaner")
        self.settings = QSettings(settings_path or os.path.join(engine.app_folder(), "settings.ini"),
                                  QSettings.IniFormat)
        self.cache = engine.Cache(cache_folder)
        self.log_path = log_path or applying.default_log()
        self.folder = str(self.settings.value("folder", "") or "")
        self.found = job.Proposals()
        self.worker = None
        self.after_apply_skipped = []
        self.setStyleSheet(STYLE)

        self.player = QMediaPlayer(self)
        self.sound = QAudioOutput(self)
        self.player.setAudioOutput(self.sound)
        self.play_from = self.play_to = 0
        self.player.mediaStatusChanged.connect(self._media_ready)
        self.player.positionChanged.connect(self._played_to)

        self.folder_label = QLabel()
        self.folder_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top = QHBoxLayout()
        self.choose_button = self._button("Choose folder", self.choose_folder, top)
        top.addWidget(self.folder_label, 1)
        self._button("Settings", self.show_settings, top)
        self._button("Help", self.show_help, top)
        steps = QHBoxLayout()
        self.scan_button = self._button("1. Scan", self.start_scan, steps)
        self.lookup_button = self._button("2. Look up doubtful names online", self.start_lookup, steps)
        self.apply_button = self._button("3. Apply ticked changes", self.start_apply, steps)
        self.undo_button = self._button("Undo last Apply", self.start_undo, steps)
        steps.addStretch(1)

        self.names_page = Page(["Name now", "New name", "Why"], [
            ("Play", self.play_selected), ("Stop", self.stop_playing), ("Edit name", self.edit_name),
            ("Tick all", lambda: self.names_page.set_all(True)),
            ("Untick all", lambda: self.names_page.set_all(False))], self)
        self.names_page.on_tick = lambda path, on: self._tick(self.found.name_ticks, path, on)
        self.names_page.on_tick_all = lambda paths, on: self._tick_all(self.found.name_ticks, paths, on)
        self.copies_page = Page(["Remove?", "File", "Length", "Sound", "What it has"], [
            ("Play", self.play_selected), ("Play the end", self.play_end), ("Stop", self.stop_playing),
            ("Keep this one instead", self.keep_instead)], self)
        self.copies_page.on_tick = self._tick_remove
        trim_buttons = lambda page: [
            ("Play new start", lambda: self.play_trim(page(), "new start")),
            ("Play new end", lambda: self.play_trim(page(), "new end")),
            ("Play what is cut", lambda: self.play_trim(page(), "cut")), ("Stop", self.stop_playing),
            ("Tick all", lambda: page().set_all(True)), ("Untick all", lambda: page().set_all(False)),
            ("Start earlier", lambda: self.nudge(page(), "start", -NUDGE_S)),
            ("Start later", lambda: self.nudge(page(), "start", NUDGE_S)),
            ("End earlier", lambda: self.nudge(page(), "end", -NUDGE_S)),
            ("End later", lambda: self.nudge(page(), "end", NUDGE_S))]
        trim_headers = ["Song", "Cut from start", "Cut from end", "Why"]
        self.trims_page = Page(trim_headers, trim_buttons(lambda: self.trims_page), self)
        self.listen_page = Page(trim_headers, trim_buttons(lambda: self.listen_page), self)
        for page in (self.trims_page, self.listen_page):
            page.on_tick = lambda path, on: self._tick(self.found.trim_ticks, path, on)
            page.on_tick_all = lambda paths, on: self._tick_all(self.found.trim_ticks, paths, on)
        self.problems_page = Page(["File", "What is wrong"], [
            ("Play", self.play_selected), ("Stop", self.stop_playing)], self)
        self.tabs = QTabWidget()
        self.pages = [(self.names_page, "Names"), (self.copies_page, "Copies"), (self.trims_page, "Trims"),
                      (self.listen_page, "Listen and decide"), (self.problems_page, "Problems")]
        for page, title in self.pages:
            self.tabs.addTab(page, title)
        self.tabs.currentChanged.connect(lambda _: self.stop_playing())

        self.status = QLabel("Choose the folder your recorded songs are in, then press 1. Scan.")
        self.status.setWordWrap(True)
        self.bar = QProgressBar()
        self.bar.setVisible(False)
        self.stop_button = QPushButton("Stop")
        self.stop_button.setVisible(False)
        self.stop_button.clicked.connect(self.stop_work)
        bottom = QHBoxLayout()
        bottom.addWidget(self.status, 1)
        bottom.addWidget(self.bar)
        bottom.addWidget(self.stop_button)

        middle = QWidget()
        layout = QVBoxLayout(middle)
        layout.addLayout(top)
        layout.addLayout(steps)
        layout.addWidget(self.tabs, 1)
        layout.addLayout(bottom)
        self.setCentralWidget(middle)
        self.resize(1500, 900)
        self.refresh()

    # ---- small helpers --------------------------------------------------------
    def _button(self, label, action, layout):
        button = QPushButton(label)
        button.clicked.connect(lambda: action())
        layout.addWidget(button)
        return button

    def tell(self, text, title="Song Cleaner"):
        QMessageBox.information(self, title, text)

    def ask(self, text, title="Song Cleaner"):
        return QMessageBox.question(self, title, text, QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.No) == QMessageBox.Yes

    def ask_text(self, label, text):
        value, ok = QInputDialog.getText(self, "Song Cleaner", label, text=text)
        return value if ok else None

    def busy(self):
        return self.worker is not None

    def _tick(self, ticks, path, on):
        ticks[path] = on
        self.refresh_buttons()

    def _tick_all(self, ticks, paths, on):
        for path in paths:
            ticks[path] = on
        self.refresh_buttons()

    def _tick_remove(self, path, on):
        self.found.remove_ticks[path] = on
        self.fill_trims()
        self.refresh_buttons()

    # ---- showing --------------------------------------------------------------
    def refresh(self):
        self.folder_label.setText(self.folder or "No folder chosen yet")
        self.fill_names()
        self.fill_copies()
        self.fill_trims()
        self.fill_problems()
        self.refresh_buttons()

    def refresh_buttons(self):
        found, busy = self.found, self.busy()
        counts = found.counts()
        total = sum(counts.values())
        self.choose_button.setEnabled(not busy)
        self.scan_button.setEnabled(not busy and bool(self.folder))
        self.lookup_button.setEnabled(not busy and bool(found.doubtful()))
        self.apply_button.setEnabled(not busy and total > 0)
        self.apply_button.setText(f"3. Apply {total} ticked changes" if total else "3. Apply ticked changes")
        self.undo_button.setEnabled(not busy and applying.has_undo(self.log_path))
        numbers = [len(found.names), len(found.groups), len(found.sure_trims()),
                   len(found.listen_trims()), self.problems_page.tree.topLevelItemCount()]
        for index, ((_, title), number) in enumerate(zip(self.pages, numbers)):
            self.tabs.setTabText(index, f"{title} ({number:,})" if found.songs or number else title)

    def fill_names(self):
        rows = []
        for change in self.found.names:
            new = change.new_name if change.changes else "(no change)"
            why = change.reason or "Tidied by rule"
            box = self.found.name_ticks.get(change.path, False) if change.changes else None
            rows.append((change.path, box, [change.old_name, new, why], None))
        self.names_page.fill(rows)

    def fill_copies(self):
        rows = []
        for group in self.found.groups:
            for n, member in enumerate(group.members):
                song = member.song
                keep = n == group.keep
                sound = f"{song.bitrate} kbps" + ("" if song.channels != 1 else ", mono")
                rows.append((song.path, None if keep else self.found.remove_ticks.get(song.path, False),
                             ["KEEP" if keep else "remove", song.name, job.clock(song.length), sound, member.note],
                             None))
            rows.append((None, None, ["", "", "", "", ""], None))
        self.copies_page.fill(rows)
        for index in range(self.copies_page.tree.topLevelItemCount()):
            item = self.copies_page.tree.topLevelItem(index)
            if item.text(0) == "KEEP":
                font = item.font(1)
                font.setBold(True)
                for column in range(5):
                    item.setFont(column, font)

    def fill_trims(self):
        def rows(trims):
            out = []
            for cut in sorted(trims, key=lambda t: os.path.basename(t.path).lower()):
                out.append((cut.path, self.found.trim_ticks.get(cut.path, False),
                            [os.path.basename(cut.path), self._seconds(cut.cut_start),
                             self._seconds(cut.cut_end), cut.reason], None))
            return out
        self.trims_page.fill(rows(self.found.sure_trims()))
        self.listen_page.fill(rows(self.found.listen_trims()))

    @staticmethod
    def _seconds(value):
        return f"{value:.1f} s" if value >= 0.05 else "nothing"

    def fill_problems(self):
        rows = [(None, None, [name, why], None) for name, why in self.found.problems]
        for group in self.found.different:
            for song in group:
                rows.append((song.path, None, [song.name, "Same name as another file, but a different "
                                               "recording (or played at another speed). Both are kept."], None))
        rows += [(None, None, [name, "Apply skipped it: " + why], None) for name, why in self.after_apply_skipped]
        self.problems_page.fill(rows)

    # ---- long jobs ------------------------------------------------------------
    def run(self, work, then, doing):
        self.stop_playing()
        self.worker = Worker(work, self)
        self.worker.progress.connect(self._progress)
        self.worker.done.connect(lambda result: self._finished(then, result))
        self.worker.failed.connect(lambda text: self._finished(None, None, text))
        self.status.setText(doing)
        self.bar.setRange(0, 0)
        self.bar.setVisible(True)
        self.stop_button.setVisible(True)
        self.stop_button.setEnabled(True)
        self.refresh_buttons()
        self.worker.start()

    def _progress(self, done, total, text):
        self.bar.setRange(0, max(total, 1))
        self.bar.setValue(done)
        self.status.setText(f"{done:,} of {total:,}   {text}")

    def _finished(self, then, result, error=None):
        worker, self.worker = self.worker, None
        if worker:
            worker.wait()
        self.bar.setVisible(False)
        self.stop_button.setVisible(False)
        if error:
            self.status.setText("Stopped: there was a problem.")
            self.refresh_buttons()
            self.tell(error)
            return
        then(result)
        self.refresh_buttons()

    def stop_work(self):
        if self.worker:
            self.worker.stop_asked = True
            self.stop_button.setEnabled(False)
            self.status.setText("Stopping after the song in hand...")

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Folder with the recorded songs", self.folder)
        if folder:
            self.set_folder(folder)

    def set_folder(self, folder):
        self.folder = os.path.normpath(folder)
        self.settings.setValue("folder", self.folder)
        self.found = job.Proposals()
        self.after_apply_skipped = []
        self.status.setText("Press 1. Scan to read the songs. Nothing is changed by scanning.")
        self.refresh()

    def start_scan(self, quiet=False):
        if self.busy() or not self.folder:
            return
        folder, cache = self.folder, self.cache

        def work(progress, should_stop):
            songs = engine.scan(folder, cache, should_stop=should_stop,
                                progress=lambda d, t, name: progress(d, t, "Reading " + name))
            if should_stop():
                return None
            found = job.analyse(songs, progress, should_stop)
            others = engine.other_sounds(folder)
            if others:
                kinds = sorted({os.path.splitext(n)[1].lower() for n in others})
                found.problems.append((
                    f"{len(others):,} files that are not MP3 ({', '.join(kinds)})",
                    "Left alone: this program only works on MP3s. The first is " + others[0]))
            return found

        def then(found):
            if found is None:
                self.status.setText("Scan stopped. Press 1. Scan to carry on where it left off.")
                return
            self.found = found
            self.refresh()
            counts = found.counts()
            self.status.setText(
                f"{len(found.songs):,} songs read. Proposed: {counts['rename']:,} renames, "
                f"{counts['remove']:,} copies to remove, {counts['trim']:,} trims. "
                "Look through the pages, untick what you do not want, then press 3. Apply.")

        if not quiet:
            self.after_apply_skipped = []
        self.run(work, then, "Reading the songs...")

    def start_lookup(self):
        if self.busy():
            return
        doubtful = self.found.doubtful()
        if not doubtful:
            return
        key = str(self.settings.value("acoustid key", "") or "").strip()
        if not key:
            self.tell("Looking names up online needs a free AcoustID key.\n\n"
                      "Press Settings to see how to get one and where to paste it.\n\n"
                      "Without a key, use Edit name on the Names page to type a name yourself.")
            return
        need_program = online.fpcalc_path() is None
        if need_program and not self.ask(
                "To recognise songs, a small free program (Chromaprint fpcalc, about 4 MB) has to be "
                "downloaded once from the Chromaprint project's page on GitHub.\n\nDownload it now?"):
            return
        found = self.found

        def work(progress, should_stop):
            if need_program:
                progress(0, len(doubtful), "Downloading the fingerprint program")
                online.fetch_fpcalc()
            lookup = online.Lookup(key)
            hits = 0
            for done, change in enumerate(doubtful, 1):
                if should_stop():
                    break
                progress(done, len(doubtful), "Asking about " + change.old_name)
                try:
                    match = lookup.find(change.path)
                except online.OnlineError as error:
                    if "fingerprinted" in str(error):
                        continue                         # this one file; carry on with the rest
                    raise
                if match:
                    found.found_online(change, match)
                    hits += 1
            return hits

        def then(hits):
            self.refresh()
            self.status.setText(f"Found {hits:,} of {len(doubtful):,} doubtful names online. "
                                "The rest can be typed with Edit name.")

        self.run(work, then, "Looking up names...")

    def start_apply(self):
        if self.busy():
            return
        changes = self.found.changes()
        if not changes:
            return
        words = job.describe_changes(self.found.counts(), self.folder)
        if not self.ask(words + "\n\nNothing is deleted, and Undo last Apply puts it all back."
                                "\n\nDo this now?"):
            return
        folder, log = self.folder, self.log_path

        def work(progress, should_stop):
            return applying.apply(changes, folder, log, progress, should_stop)

        def then(report):
            self.after_apply_skipped = list(report.skipped)
            text = f"Done: {report.done:,} changes made."
            if report.stopped:
                text += " Stopped before the end; the rest were left alone."
            if report.skipped:
                text += f"\n\n{len(report.skipped):,} were skipped. They are listed on the Problems page."
            self.tell(text)
            QTimer.singleShot(0, lambda: self.start_scan(quiet=True))

        self.run(work, then, "Making the changes...")

    def start_undo(self):
        if self.busy() or not applying.has_undo(self.log_path):
            return
        if not self.ask("Put back everything the last Apply did?\n\nRenamed songs get their old names, "
                        "removed copies come back, and trimmed songs get their original sound back."):
            return
        log = self.log_path

        def work(progress, should_stop):
            return applying.undo(log, progress)

        def then(report):
            self.after_apply_skipped = []
            text = f"Put back: {report.done:,} changes."
            if report.skipped:
                text += "\n\nCould not put back:\n" + "\n".join(f"{n}: {w}" for n, w in report.skipped[:15])
            self.tell(text)
            if self.folder:
                QTimer.singleShot(0, lambda: self.start_scan(quiet=True))

        self.run(work, then, "Putting things back...")

    # ---- changing proposals ---------------------------------------------------
    def edit_name(self):
        path = self.names_page.current()
        change = next((c for c in self.found.names if c.path == path), None)
        if change is None:
            self.tell("Click a song in the list first, then press Edit name.")
            return
        now = f"{change.artist} - {change.title}" if change.artist else change.title
        typed = self.ask_text("Type the name as  Artist - Title", now)
        if typed is None or not typed.strip():
            return
        artist, title = names.split_name(typed.strip())
        self.found.rename(change, artist, title)
        change.reason = "Typed by you"
        self.fill_names()
        self.names_page.select(path)
        self.refresh_buttons()

    def keep_instead(self):
        path = self.copies_page.current()
        for group in self.found.groups:
            for index, member in enumerate(group.members):
                if member.song.path == path:
                    self.found.keep_instead(group, index)
                    self.fill_copies()
                    self.fill_trims()
                    self.copies_page.select(path)
                    self.refresh_buttons()
                    return
        self.tell("Click one of the copies in the list first, then press Keep this one instead.")

    def nudge(self, page, side, seconds):
        path = page.current()
        if path not in self.found.trims:
            self.tell("Click a song in the list first.")
            return
        self.found.nudge(path, side, seconds)
        self.fill_trims()
        page.select(path)
        self.play_trim(page, "new start" if side == "start" else "new end")

    # ---- playing --------------------------------------------------------------
    def play(self, path, start=0.0, end=None):
        """Play part of a file: from start, to end (or for a few seconds)."""
        self.stop_playing()
        if not path or not os.path.exists(path):
            return
        self.play_from = int(max(0.0, start) * 1000)
        self.play_to = int((end if end is not None else start + PLAY_S) * 1000)
        self.player.setSource(QUrl.fromLocalFile(path))
        self.player.play()

    def _media_ready(self, status):
        if status in (QMediaPlayer.LoadedMedia, QMediaPlayer.BufferedMedia) and self.play_from:
            start, self.play_from = self.play_from, 0
            self.player.setPosition(start)

    def _played_to(self, position):
        if self.play_to and not self.play_from and position >= self.play_to:
            self.stop_playing()

    def stop_playing(self):
        self.play_to = 0
        self.play_from = 0
        self.player.stop()
        self.player.setSource(QUrl())                    # lets go of the file, so Apply can move it

    def play_selected(self):
        path = self.tabs.currentWidget().current()
        if path:
            self.play(path, 0.0, 30.0)
        else:
            self.tell("Click a song in the list first, then press Play.")

    def play_end(self):
        path = self.copies_page.current()
        song = self.found.songs.get(path)
        if song:
            self.play(path, max(0.0, song.length - 15.0), song.length + 1)

    def play_trim(self, page, what):
        cut = self.found.trims.get(page.current())
        if cut is None:
            self.tell("Click a song in the list first.")
            return
        if what == "new start":
            self.play(cut.path, cut.start, cut.start + PLAY_S)
        elif what == "new end":
            self.play(cut.path, max(cut.start, cut.end - PLAY_S), cut.end)
        elif cut.cut_start >= 0.05:
            self.play(cut.path, 0.0, cut.start)
        else:
            self.play(cut.path, cut.end, cut.length + 1)

    # ---- other windows --------------------------------------------------------
    def show_settings(self):
        box = SettingsBox(self)
        if box.exec() == QDialog.Accepted:
            self.settings.setValue("acoustid key", box.key.text().strip())

    def show_help(self):
        box = QDialog(self)
        box.setWindowTitle("Song Cleaner: how to use it")
        text = QPlainTextEdit(HELP)
        text.setReadOnly(True)
        close = QDialogButtonBox(QDialogButtonBox.Close)
        close.rejected.connect(box.reject)
        layout = QVBoxLayout(box)
        layout.addWidget(text)
        layout.addWidget(close)
        box.resize(1200, 800)
        box.exec()

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop_asked = True
            self.worker.wait(20000)
        self.stop_playing()
        self.settings.sync()
        event.accept()


def black_words(app):
    palette = app.palette()
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        palette.setColor(QPalette.Active, role, QColor("black"))
        palette.setColor(QPalette.Inactive, role, QColor("black"))
    app.setPalette(palette)


def show_crash(kind, error, trace):
    text = "".join(traceback.format_exception(kind, error, trace, limit=6))
    try:
        QMessageBox.critical(None, "Song Cleaner", "Something went wrong inside the program.\n\n" + text)
    except Exception:  # noqa: BLE001
        print(text, file=sys.stderr)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Song Cleaner")
    black_words(app)
    sys.excepthook = show_crash
    lock = QLockFile(os.path.join(engine.app_folder(), "running.lock"))
    if not lock.tryLock(100):
        QMessageBox.information(None, "Song Cleaner", "Song Cleaner is already open.")
        return 0
    window = MainWindow()
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
