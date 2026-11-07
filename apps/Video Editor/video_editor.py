"""
video_editor.py

A video and sound editor built on video_player.py. Open videos or sound files,
cut and join them, then save the result as a new file:

  Trim ............ select a piece and type where it starts and ends in its
                    file, or play to the spot and click Start here / End here
  Split ........... cut the piece under the playhead in two (S)
  Delete a part ... Mark start (I) and Mark end (O), then Delete marked part
  Keep a part ..... mark it, then Keep only marked part
  Insert .......... Insert at playhead puts other videos or sounds in at the
                    playhead; Copy marked part and Paste here repeat a part
  Splice .......... pieces play one after another, so adding, deleting and
                    moving pieces joins them; Move up / Move down change the order
  Sound ........... per piece: mute, volume (up to 200%), fade in, fade out.
                    Add music or narration lays a sound over the edit from the
                    playhead, mixed in or replacing the pieces' own sound, with
                    its own volume and fades
  Open ............ the big button in the middle when nothing is open, the blue
                    Open / add files button, File > Open, Ctrl+O, or drag files in
  Save as ......... pick a folder, a name and a format: MP4, MOV, M4V, MKV, WebM,
                    AVI, WMV or MPG video, or MP3, M4A, WAV, FLAC, OGG, Opus or
                    WMA sound (an edit with video saved as sound keeps its sound)

The files you open are never changed. Undo and Redo cover every edit. The
preview plays the edit with its cuts, volumes and added sound; volumes over
100% sound like 100% until saved, and fades are approximate.

Keys: Space play/pause, Left/Right 1 second, comma/period one frame,
Home/End, S split, I/O mark start/end, Esc clear marks, Delete delete the
marked part (or the selected piece), Ctrl+C/Ctrl+V copy and paste the marked
part, Ctrl+Z/Ctrl+Y undo and redo, Ctrl+O add files, Ctrl+S save as.

Requirements (one-time): PySide6, plus imageio-ffmpeg, which brings the ffmpeg
used for saving. If the Python running this lacks them, it offers to install them.
"""

import os
import subprocess
import sys
import tempfile

import video_player as vp      # checks for PySide6 and shows crashes in a dialog

TITLE = "Video Editor"
vp.TITLE = TITLE               # its install and crash messages then carry this name
vp.install_missing({"imageio_ffmpeg": ("imageio-ffmpeg",
                                       "imageio-ffmpeg, which saves the edited files", 40)})

import imageio_ffmpeg
import video_editor_engine as engine
from video_editor_engine import Edit, ProbeError, fade_gain, fit_fades
from PySide6.QtCore import (QObject, QProcess, QRectF, QSettings, QSize, QStandardPaths, Qt, QUrl,
                            Signal)
from PySide6.QtGui import QColor, QFont, QIcon, QKeySequence, QPainter, QPen, QShortcut
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QDoubleSpinBox, QFileDialog, QFormLayout, QHBoxLayout, QLabel,
    QListWidget, QMainWindow, QMessageBox, QProgressDialog, QPushButton, QSlider, QSplitter,
    QStackedWidget, QStyle, QTabWidget, QToolButton, QVBoxLayout, QWidget,
)

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
MAX_UNDO = 100
LONG_GRAPH = 8000      # characters; longer filter graphs go to ffmpeg in a file


def precise(t):
    """A time to a tenth of a second: 1:05.3"""
    tenths = int(round(max(0.0, t) * 10))
    whole, tenth = divmod(tenths, 10)
    return f"{vp.format_time(whole * 1000)}.{tenth}"


def open_filter():
    both = " ".join("*" + ext for ext in sorted(vp.VIDEO_EXTENSIONS | vp.AUDIO_EXTENSIONS))
    return f"Videos and sound files ({both});;" + vp.file_filter()


def file_name(path):
    return os.path.basename(path)


# ---------------------------------------------------------------- timeline

class Timeline(QWidget):
    """The edit drawn to scale: pieces on top, added sounds below, the marked
    part shaded and the playhead in red. Click or drag to move the playhead."""

    seek = Signal(float)
    empty_clicked = Signal()
    piece_clicked = Signal(int)
    sound_clicked = Signal(int)
    LANE, SOUND_LANE, GAP, PAD = 46, 24, 6, 4

    def __init__(self):
        super().__init__()
        self.edit = Edit()
        self.position = 0.0
        self.marks = (None, None)
        self.selected_piece = self.selected_sound = -1
        self.setMinimumHeight(self.LANE + self.SOUND_LANE + self.GAP + 2 * self.PAD)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setToolTip("Click or drag to move the playhead.\n"
                        "Top strip: the pieces in order. Bottom strip: added sound.")

    def x_of(self, t):
        return self.PAD + (self.width() - 2 * self.PAD) * t / max(self.edit.duration, 0.001)

    def t_of(self, x):
        span = max(1, self.width() - 2 * self.PAD)
        return min(max(0.0, (x - self.PAD) / span * self.edit.duration), self.edit.duration)

    def block(self, painter, rect, color, text, chosen):
        painter.fillRect(rect, color)
        if chosen:
            painter.setPen(QPen(QColor("#ffcc00"), 3))
            painter.drawRect(rect.adjusted(1.5, 1.5, -1.5, -1.5))
        if rect.width() > 24:
            painter.setPen(QColor("white"))
            label = painter.fontMetrics().elidedText(text, Qt.TextElideMode.ElideRight,
                                                     int(rect.width()) - 8)
            painter.drawText(rect.adjusted(4, 0, -4, 0),
                             Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, label)

    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), self.palette().base())
        total = self.edit.duration
        top, low = self.PAD, self.PAD + self.LANE + self.GAP
        if total <= 0:
            p.setPen(self.palette().placeholderText().color())
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "Click here, or Open / add files, to add videos or sound")
            return
        for i, (start, piece) in enumerate(zip(self.edit.starts(), self.edit.pieces)):
            left, right = self.x_of(start), self.x_of(start + piece.length)
            color = QColor("#2f74b5" if piece.has_video else "#7b4fb8")
            if i % 2:
                color = color.lighter(118)
            label = file_name(piece.source) + ("  (muted)" if piece.silent else "")
            self.block(p, QRectF(left, top, max(1.0, right - left - 1), self.LANE), color,
                       label, i == self.selected_piece)
        for j, sound in enumerate(self.edit.sounds):
            if sound.at >= total:
                continue
            left, right = self.x_of(sound.at), self.x_of(min(sound.finish, total))
            color = QColor("#c0702a" if sound.replace else "#2e8f57")
            self.block(p, QRectF(left, low, max(1.0, right - left), self.SOUND_LANE), color,
                       file_name(sound.source), j == self.selected_sound)
        a, b = self.marks
        bottom = low + self.SOUND_LANE
        if a is not None and b is not None:
            p.fillRect(QRectF(self.x_of(a), top, self.x_of(b) - self.x_of(a), bottom - top),
                       QColor(255, 255, 255, 90))
        p.setPen(QPen(QColor("#ffcc00"), 2))
        for mark in (a, b):
            if mark is not None:
                p.drawLine(int(self.x_of(mark)), top, int(self.x_of(mark)), bottom)
        p.setPen(QPen(QColor("#e02020"), 2))
        x = int(self.x_of(self.position))
        p.drawLine(x, 0, x, self.height())

    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        if self.edit.duration <= 0:
            self.empty_clicked.emit()
            return
        t = self.t_of(event.position().x())
        y = event.position().y()
        if y >= self.PAD + self.LANE + self.GAP / 2:
            hit = [j for j, s in enumerate(self.edit.sounds) if s.at <= t < s.finish]
            if hit:
                self.sound_clicked.emit(hit[-1])
        else:
            self.piece_clicked.emit(self.edit.locate(t)[0])
        self.seek.emit(t)

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton and self.edit.duration > 0:
            self.seek.emit(self.t_of(event.position().x()))


# ------------------------------------------------------------------ saving

class ExportJob(QObject):
    """Runs ffmpeg for one save. Writes to a temporary name beside the target
    and renames it only when ffmpeg succeeds, so a failed or stopped save
    never leaves half a file or damages one that was there before."""

    progress = Signal(float)            # 0 to 1
    done = Signal(bool, str)            # succeeded, the saved path or what went wrong

    def __init__(self, plan, path, parent=None):
        super().__init__(parent)
        self.path = path
        self.duration = max(plan.duration, 0.001)
        stem, ext = os.path.splitext(path)
        self.temp = f"{stem}.saving{ext}"
        plan.output = plan.output[:-1] + [self.temp]
        self.graph_file = None
        if len(plan.graph) > LONG_GRAPH:
            handle, self.graph_file = tempfile.mkstemp(suffix=".txt", prefix="video-editor-")
            with os.fdopen(handle, "w", encoding="utf-8") as f:
                f.write(plan.graph)
        command = plan.command(FFMPEG, self.graph_file)
        self.errors = self.unread = ""
        self.stopped = False
        self.process = QProcess(self)
        self.process.readyReadStandardOutput.connect(self.read_progress)
        self.process.readyReadStandardError.connect(self.read_errors)
        self.process.finished.connect(self.finished)
        self.process.errorOccurred.connect(self.failed_to_start)
        self.process.start(command[0], command[1:])

    def read_progress(self):
        self.unread += bytes(self.process.readAllStandardOutput()).decode("utf-8", "replace")
        *lines, self.unread = self.unread.split("\n")
        for line in lines:
            seconds = engine.progress_seconds(line)
            if seconds is not None:
                self.progress.emit(min(1.0, seconds / self.duration))

    def read_errors(self):
        text = bytes(self.process.readAllStandardError()).decode("utf-8", "replace")
        self.errors = (self.errors + text)[-4000:]

    def stop(self):
        self.stopped = True
        self.process.kill()

    def failed_to_start(self, error):
        if error == QProcess.ProcessError.FailedToStart:
            self.cleanup()
            self.done.emit(False, f"ffmpeg could not start ({FFMPEG}).")

    def finished(self, code, status):
        ok = (not self.stopped and status == QProcess.ExitStatus.NormalExit and code == 0
              and os.path.isfile(self.temp))
        if ok:
            try:
                os.replace(self.temp, self.path)
            except OSError as err:
                ok, self.errors = False, f"Could not finish the file: {err}"
        self.cleanup()
        if ok:
            self.done.emit(True, self.path)
        elif self.stopped:
            self.done.emit(False, "")
        else:
            self.done.emit(False, self.errors.strip()[-1500:] or f"ffmpeg stopped with code {code}.")

    def cleanup(self):
        for leftover in (self.temp, self.graph_file):
            if leftover and os.path.exists(leftover):
                try:
                    os.remove(leftover)
                except OSError:
                    pass


# ------------------------------------------------------------------ window

class EditorWindow(QMainWindow):

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.edit = Edit()
        self.undo_stack, self.redo_stack = [], []
        self.t = 0.0
        self.marks = [None, None]
        self.clipboard = []
        self.selected_piece = self.selected_sound = -1
        self.dirty = False
        self.job = None
        # preview playback
        self.playing = False
        self.current = -1
        self.loaded = self.pending = None
        self.sound_loaded = self.sound_pending = None

        self.player = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.player.setAudioOutput(self.audio)
        self.video = QVideoWidget()
        self.player.setVideoOutput(self.video)
        self.sound_player = QMediaPlayer(self)
        self.sound_audio = QAudioOutput(self)
        self.sound_player.setAudioOutput(self.sound_audio)

        self.build_window()
        self.build_menu()
        self.build_shortcuts()
        self.player.positionChanged.connect(self.position_changed)
        self.player.mediaStatusChanged.connect(self.media_status_changed)
        self.player.errorOccurred.connect(self.player_error)
        self.sound_player.mediaStatusChanged.connect(self.sound_status_changed)
        self.restore_settings()
        self.refresh()

    # ---------------------------------------------------------------- layout
    def icon(self, theme_name, fallback):
        icon = QIcon.fromTheme(getattr(QIcon.ThemeIcon, theme_name)) if hasattr(
            QIcon.ThemeIcon, theme_name) else QIcon()
        return icon if not icon.isNull() else self.style().standardIcon(fallback)

    def button(self, text, icon, tip, slot, theme=None):
        b = QToolButton()
        b.setText(text)
        b.setIcon(self.icon(theme, icon) if theme else self.style().standardIcon(icon))
        b.setToolTip(tip)
        b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        b.clicked.connect(slot)
        return b

    def row(self, *widgets):
        layout = QHBoxLayout()
        layout.setSpacing(4)
        for w in widgets:
            if w is None:
                layout.addStretch(1)
            else:
                layout.addWidget(w)
        return layout

    def build_window(self):
        self.setWindowTitle(TITLE)
        self.setAcceptDrops(True)
        SP = QStyle.StandardPixmap

        self.start_page = self.build_start_page()
        self.sound_only = QLabel("Sound only")
        self.sound_only.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sound_only.setStyleSheet("background: black; color: #bbbbbb; font-size: 15pt;")
        self.screen_stack = QStackedWidget()
        for w in (self.start_page, self.video, self.sound_only):
            self.screen_stack.addWidget(w)
        self.screen_stack.setMinimumSize(320, 180)

        self.timeline = Timeline()
        self.timeline.seek.connect(lambda t: self.seek(t))
        self.timeline.empty_clicked.connect(self.add_to_end)
        self.timeline.piece_clicked.connect(self.select_piece)
        self.timeline.sound_clicked.connect(self.select_sound)

        self.play_button = self.button("Play", SP.SP_MediaPlay, "Play or pause (Space)",
                                       self.toggle_play)
        self.play_button.setMinimumWidth(84)
        self.time_label = QLabel("0:00.0 / 0:00.0")
        self.time_label.setMinimumWidth(130)
        self.marks_label = QLabel("")
        self.preview_volume = QSlider(Qt.Orientation.Horizontal)
        self.preview_volume.setRange(0, 100)
        self.preview_volume.setFixedWidth(90)
        self.preview_volume.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.preview_volume.setToolTip("Preview volume (doesn't change the saved file)")
        self.preview_volume.valueChanged.connect(lambda _: self.apply_sound_levels())
        transport = self.row(
            self.button("Start", SP.SP_MediaSkipBackward, "Go to the start (Home)",
                        lambda: self.seek(0)),
            self.button("1 s", SP.SP_MediaSeekBackward, "Back 1 second (Left arrow)",
                        lambda: self.seek(self.t - 1)),
            self.button("Frame", SP.SP_ArrowLeft, "Back one frame (comma)",
                        lambda: self.step_frame(-1)),
            self.play_button,
            self.button("Frame", SP.SP_ArrowRight, "Forward one frame (period)",
                        lambda: self.step_frame(1)),
            self.button("1 s", SP.SP_MediaSeekForward, "Forward 1 second (Right arrow)",
                        lambda: self.seek(self.t + 1)),
            self.button("End", SP.SP_MediaSkipForward, "Go to the end (End)",
                        lambda: self.seek(self.edit.duration)),
            None, QLabel("Preview volume"), self.preview_volume)
        readout = self.row(self.time_label, None, self.marks_label)

        self.undo_button = self.button("Undo", SP.SP_ArrowBack, "Undo the last edit (Ctrl+Z)",
                                       self.undo, "EditUndo")
        self.redo_button = self.button("Redo", SP.SP_ArrowForward, "Redo (Ctrl+Y)",
                                       self.redo, "EditRedo")
        self.save_button = self.button("Save as...", SP.SP_DialogSaveButton,
                                       "Save the edit as a new file, in the format and folder "
                                       "you choose (Ctrl+S)", self.save_as, "DocumentSaveAs")
        self.save_button.setStyleSheet("font-weight: bold;")
        self.open_button = self.button(
            "Open / add files...", SP.SP_DialogOpenButton,
            "Browse your files and add videos or sound to the end of the edit (Ctrl+O)",
            self.add_to_end)
        self.open_button.setStyleSheet(
            "QToolButton { background: #2f74b5; color: white; font-weight: bold;"
            " padding: 4px 10px; border-radius: 4px; }"
            "QToolButton:hover { background: #3d8ad6; }")
        files_row = self.row(
            self.open_button,
            self.button("Insert at playhead", SP.SP_FileDialogDetailedView,
                        "Put videos or sound files in where the red playhead is",
                        self.insert_files, "InsertLink"),
            self.button("Add music or narration", SP.SP_MediaVolume,
                        "Lay a sound file over the edit, starting at the playhead",
                        self.add_sound, "AudioVolumeHigh"),
            None, self.undo_button, self.redo_button, self.save_button)
        cut_row = self.row(
            self.button("Split here", SP.SP_DialogResetButton,
                        "Cut the piece under the playhead in two (S)", self.split_here, "EditCut"),
            self.button("Mark start", SP.SP_ArrowDown, "Mark the start of a part (I)",
                        self.mark_start),
            self.button("Mark end", SP.SP_ArrowDown, "Mark the end of a part (O)",
                        self.mark_end),
            self.button("Clear marks", SP.SP_DialogCancelButton, "Clear the marks (Esc)",
                        self.clear_marks, "EditClear"),
            self.button("Delete marked part", SP.SP_TrashIcon,
                        "Take out the marked part and join what is left (Delete)",
                        self.delete_marked, "EditDelete"),
            self.button("Keep only marked part", SP.SP_DialogApplyButton,
                        "Throw away everything outside the marked part", self.keep_marked),
            self.button("Copy marked part", SP.SP_FileDialogContentsView,
                        "Copy the marked part (Ctrl+C)", self.copy_marked, "EditCopy"),
            self.button("Paste here", SP.SP_FileDialogListView,
                        "Put the copied part in at the playhead (Ctrl+V)", self.paste, "EditPaste"),
            None)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(6, 6, 6, 6)
        left_layout.addWidget(self.screen_stack, 1)
        left_layout.addLayout(readout)
        left_layout.addWidget(self.timeline)
        for layout in (transport, files_row, cut_row):
            left_layout.addLayout(layout)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.build_piece_tab(), "Pieces")
        self.tabs.addTab(self.build_sound_tab(), "Added sound")
        # Wide enough for its times and buttons at any text size Windows is set to.
        panel = max(self.tabs.minimumSizeHint().width(), self.fontMetrics().averageCharWidth() * 46)
        self.tabs.setMinimumWidth(panel)

        splitter = QSplitter()
        splitter.addWidget(left)
        splitter.addWidget(self.tabs)
        splitter.setStretchFactor(0, 1)
        splitter.setSizes([980, panel])
        self.setCentralWidget(splitter)

    def build_start_page(self):
        """What the picture area shows before anything is opened: one big button."""
        page = QWidget()
        page.setObjectName("startPage")
        page.setStyleSheet("#startPage { background: black; }")
        self.open_big = QPushButton("  Open a video or sound file...")
        self.open_big.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        self.open_big.setIconSize(QSize(32, 32))
        self.open_big.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_big.setToolTip("Browse your files and pick the videos or sound to edit (Ctrl+O)")
        self.open_big.setStyleSheet(
            "QPushButton { background: #2f74b5; color: white; font-size: 17pt; font-weight: bold;"
            " padding: 16px 30px; border: 3px solid #2f74b5; border-radius: 10px; }"
            "QPushButton:hover { background: #3d8ad6; border-color: #3d8ad6; }"
            "QPushButton:pressed { background: #245a8c; }"
            "QPushButton:focus { border-color: #ffcc00; }")
        self.open_big.clicked.connect(self.add_to_end)
        also = QLabel("or drag videos and sound files onto this window")
        also.setStyleSheet("color: #bbbbbb; font-size: 12pt;")
        layout = QVBoxLayout(page)
        layout.addStretch(1)
        layout.addWidget(self.open_big, 0, Qt.AlignmentFlag.AlignHCenter)
        layout.addSpacing(12)
        layout.addWidget(also, 0, Qt.AlignmentFlag.AlignHCenter)
        layout.addStretch(1)
        return page

    def build_menu(self):
        """A File menu, where most programs keep Open and Save."""
        menu = self.menuBar().addMenu("&File")
        for text, key, slot in (("&Open / add files...", "Ctrl+O", self.add_to_end),
                                ("&Insert files at the playhead...", None, self.insert_files),
                                ("Add &music or narration...", None, self.add_sound),
                                (None, None, None),
                                ("&Save as...", "Ctrl+S", self.save_as),
                                (None, None, None),
                                ("&Close", None, self.close)):
            if text is None:
                menu.addSeparator()
                continue
            action = menu.addAction(text)
            if key:
                action.setShortcut(QKeySequence(key))
            action.triggered.connect(lambda checked=False, run=slot: run())

    def spin(self, maximum, tip, slot, step=0.1):
        box = QDoubleSpinBox()
        box.setDecimals(2)
        box.setRange(0, maximum)
        box.setSingleStep(step)
        box.setSuffix(" s")
        box.setKeyboardTracking(False)          # change only on Enter, leaving the box or arrows
        box.setToolTip(tip)
        box.setMinimumWidth(self.fontMetrics().horizontalAdvance("0000.00 s") + 40)
        box.valueChanged.connect(slot)
        return box

    def volume_slider(self, slot):
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(0, 200)
        slider.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        label = QLabel("100%")
        label.setMinimumWidth(40)
        slider.value_label = label
        slider.valueChanged.connect(lambda v: label.setText(f"{v}%"))
        slider.valueChanged.connect(lambda v: None if slider.isSliderDown() else slot())
        slider.sliderReleased.connect(slot)
        return slider, label

    def build_piece_tab(self):
        SP = QStyle.StandardPixmap
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.piece_list = QListWidget()
        self.piece_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.piece_list.setWordWrap(True)
        self.piece_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.piece_list.currentRowChanged.connect(self.piece_row_chosen)
        layout.addWidget(self.piece_list, 1)
        layout.addLayout(self.row(
            self.button("Move up", SP.SP_ArrowUp, "Play this piece earlier",
                        lambda: self.move_piece(-1), "GoUp"),
            self.button("Move down", SP.SP_ArrowDown, "Play this piece later",
                        lambda: self.move_piece(1), "GoDown"),
            None,
            self.button("Remove", SP.SP_TrashIcon, "Take this piece out of the edit",
                        self.remove_piece, "ListRemove")))

        self.piece_form = QWidget()
        form = QFormLayout(self.piece_form)
        self.piece_name = QLabel("")
        self.piece_name.setWordWrap(True)
        form.addRow(self.piece_name)
        self.piece_start = self.spin(0, "Where this piece starts, in its file", self.piece_trimmed)
        self.piece_end = self.spin(0, "Where this piece ends, in its file", self.piece_trimmed)
        form.addRow("Starts at", self.row(self.piece_start, self.button(
            "Start here", SP.SP_ArrowRight, "Trim the piece under the playhead so it "
            "starts at the playhead", lambda: self.trim_here(True))))
        form.addRow("Ends at", self.row(self.piece_end, self.button(
            "End here", SP.SP_ArrowLeft, "Trim the piece under the playhead so it "
            "ends at the playhead", lambda: self.trim_here(False))))
        self.piece_mute = QCheckBox("Mute this piece")
        self.piece_mute.toggled.connect(self.piece_sound_changed)
        form.addRow("Sound", self.piece_mute)
        self.piece_volume, volume_label = self.volume_slider(self.piece_sound_changed)
        form.addRow("Volume", self.row(self.piece_volume, volume_label))
        self.piece_fade_in = self.spin(30, "Fade the sound in over this many seconds",
                                       self.piece_sound_changed)
        self.piece_fade_out = self.spin(30, "Fade the sound out over this many seconds",
                                        self.piece_sound_changed)
        form.addRow("Fade in", self.piece_fade_in)
        form.addRow("Fade out", self.piece_fade_out)
        layout.addWidget(self.piece_form)
        return tab

    def build_sound_tab(self):
        SP = QStyle.StandardPixmap
        tab = QWidget()
        layout = QVBoxLayout(tab)
        note = QLabel("Music or narration laid over the edit. Add it with "
                      "\"Add music or narration\"; it starts at the playhead.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.sound_list = QListWidget()
        self.sound_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.sound_list.setWordWrap(True)
        self.sound_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.sound_list.currentRowChanged.connect(self.sound_row_chosen)
        layout.addWidget(self.sound_list, 1)
        layout.addLayout(self.row(None, self.button(
            "Remove", SP.SP_TrashIcon, "Take this sound out of the edit",
            self.remove_sound, "ListRemove")))

        self.sound_form = QWidget()
        form = QFormLayout(self.sound_form)
        self.sound_name = QLabel("")
        self.sound_name.setWordWrap(True)
        form.addRow(self.sound_name)
        self.sound_at = self.spin(0, "When this sound starts, in the edit", self.sound_changed)
        form.addRow("Starts in the edit at", self.row(self.sound_at, self.button(
            "Playhead", SP.SP_ArrowDown, "Start this sound at the playhead",
            self.sound_to_playhead)))
        self.sound_from = self.spin(0, "Where to start in the sound file", self.sound_changed)
        self.sound_to = self.spin(0, "Where to stop in the sound file", self.sound_changed)
        form.addRow("From (in its file)", self.sound_from)
        form.addRow("To (in its file)", self.sound_to)
        self.sound_volume, volume_label = self.volume_slider(self.sound_changed)
        form.addRow("Volume", self.row(self.sound_volume, volume_label))
        self.sound_fade_in = self.spin(30, "Fade in over this many seconds", self.sound_changed)
        self.sound_fade_out = self.spin(30, "Fade out over this many seconds", self.sound_changed)
        form.addRow("Fade in", self.sound_fade_in)
        form.addRow("Fade out", self.sound_fade_out)
        self.sound_replace = QCheckBox("Silence the video's own sound while this plays")
        self.sound_replace.toggled.connect(self.sound_changed)
        form.addRow(self.sound_replace)
        layout.addWidget(self.sound_form)
        return tab

    def build_shortcuts(self):
        keys = {
            "Space": self.toggle_play, "Left": lambda: self.seek(self.t - 1),
            "Right": lambda: self.seek(self.t + 1), ",": lambda: self.step_frame(-1),
            ".": lambda: self.step_frame(1), "Home": lambda: self.seek(0),
            "End": lambda: self.seek(self.edit.duration), "S": self.split_here,
            "I": self.mark_start, "O": self.mark_end, "Esc": self.clear_marks,
            "Del": self.delete_key, "Ctrl+C": self.copy_marked, "Ctrl+V": self.paste,
            "Ctrl+Z": self.undo, "Ctrl+Y": self.redo, "Ctrl+Shift+Z": self.redo,
        }
        for key, slot in keys.items():
            QShortcut(QKeySequence(key), self, slot)

    # -------------------------------------------------------------- settings
    def restore_settings(self):
        self.preview_volume.setValue(int(self.settings.value("preview_volume", 80)))
        geometry = self.settings.value("geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)
        else:
            self.resize(1340, 820)

    def folder(self, key, fallback):
        folder = self.settings.value(key, "")
        if folder and os.path.isdir(folder):
            return folder
        return QStandardPaths.writableLocation(fallback)

    def closeEvent(self, event):
        if self.job is not None:
            if QMessageBox.question(self, TITLE, "A file is still being saved. Stop saving "
                                    "and close?") != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            job = self.job                # finishing may clear self.job
            job.stop()
            job.process.waitForFinished(5000)
            job.cleanup()
        elif self.dirty and self.edit.pieces:
            if QMessageBox.question(self, TITLE, "Your edit hasn't been saved as a file. "
                                    "Close anyway and lose it?") != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.settings.setValue("preview_volume", self.preview_volume.value())
        self.settings.setValue("geometry", self.saveGeometry())
        self.player.stop()
        self.sound_player.stop()
        super().closeEvent(event)

    # ---------------------------------------------------------- opening files
    def choose_files(self, title, many=True):
        start = self.folder("open_folder", QStandardPaths.StandardLocation.MoviesLocation)
        if many:
            paths, _ = QFileDialog.getOpenFileNames(self, title, start, open_filter())
        else:
            path, _ = QFileDialog.getOpenFileName(self, title, start, open_filter())
            paths = [path] if path else []
        if paths:
            self.settings.setValue("open_folder", os.path.dirname(paths[0]))
        return paths

    def read_files(self, paths):
        """Look inside each file; report the ones that can't be used."""
        infos, problems = [], []
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            for path in paths:
                try:
                    infos.append(engine.probe(path, FFMPEG))
                except (ProbeError, OSError) as err:
                    problems.append(str(err))
        finally:
            QApplication.restoreOverrideCursor()
        if problems:
            QMessageBox.warning(self, TITLE, "Some files can't be used:\n\n" + "\n".join(problems))
        return infos

    def add_to_end(self, paths=None):
        paths = paths or self.choose_files("Add videos or sound files to the end")
        infos = self.read_files(paths) if paths else []
        if infos:
            first = len(self.edit.pieces)
            self.change(f"Added {len(infos)} file(s) to the end",
                        lambda: self.edit.append([engine.piece_from(i) for i in infos]))
            self.selected_piece = first
            self.refresh()
            if first == 0:
                self.seek(0, play=False)

    def insert_files(self):
        paths = self.choose_files("Insert videos or sound files at the playhead")
        infos = self.read_files(paths) if paths else []
        if infos:
            at = self.t
            index = self.change(f"Inserted {len(infos)} file(s) at {precise(at)}",
                                lambda: self.edit.insert(at, [engine.piece_from(i) for i in infos]))
            self.selected_piece = index
            self.refresh()
            self.seek(self.edit.starts()[index])

    def add_sound(self):
        if not self.edit.pieces:
            self.say("Add some videos or sound first; music or narration goes over them.")
            return
        paths = self.choose_files("Add music or narration at the playhead", many=False)
        infos = self.read_files(paths) if paths else []
        if not infos:
            return
        try:
            sound = engine.sound_from(infos[0], self.t)
        except ProbeError as err:
            QMessageBox.warning(self, TITLE, str(err))
            return
        self.change(f"Added {file_name(sound.source)} at {precise(sound.at)}",
                    lambda: self.edit.sounds.append(sound))
        self.selected_sound = len(self.edit.sounds) - 1
        self.tabs.setCurrentIndex(1)
        self.refresh()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        files = []
        for path in paths:
            files += vp.videos_in_folder(path) if os.path.isdir(path) else [path]
        if files:
            self.add_to_end(files)

    # ----------------------------------------------------------------- edits
    def change(self, description, action):
        """Make an edit that Undo can take back."""
        before = self.edit.snapshot()
        result = action()
        self.undo_stack.append(before)
        del self.undo_stack[:-MAX_UNDO]
        self.redo_stack.clear()
        self.dirty = True
        self.after_change(description)
        return result

    def after_change(self, description):
        self.t = min(self.t, self.edit.duration)
        if self.selected_piece >= len(self.edit.pieces):
            self.selected_piece = len(self.edit.pieces) - 1
        if self.selected_sound >= len(self.edit.sounds):
            self.selected_sound = len(self.edit.sounds) - 1
        for i, mark in enumerate(self.marks):
            if mark is not None and mark > self.edit.duration:
                self.marks[i] = None
        self.refresh()
        self.say(description)
        self.seek(self.t)

    def undo(self):
        self.swap_history(self.undo_stack, self.redo_stack, "Undid the last edit",
                          "Nothing to undo.")

    def redo(self):
        self.swap_history(self.redo_stack, self.undo_stack, "Redid the edit", "Nothing to redo.")

    def swap_history(self, take, keep, done, empty):
        if not take:
            self.say(empty)
            return
        keep.append(self.edit.snapshot())
        self.edit = take.pop()
        self.dirty = True
        self.after_change(done)

    def split_here(self):
        trial = self.edit.snapshot()
        trial.split(self.t)
        if len(trial.pieces) == len(self.edit.pieces):
            self.say("There is already a cut here (or the playhead is at an end).")
            return
        at = self.t
        self.change(f"Split at {precise(at)}", lambda: self.edit.split(at))

    def marked(self):
        a, b = self.marks
        if a is None or b is None or b - a < engine.MIN_PIECE:
            self.say("First mark the start (I) and the end (O) of a part.")
            return None
        return a, b

    def mark_start(self):
        self.marks[0] = self.t
        if self.marks[1] is not None and self.marks[1] <= self.t:
            self.marks[1] = None
        self.refresh_marks()

    def mark_end(self):
        self.marks[1] = self.t
        if self.marks[0] is not None and self.marks[0] > self.t:
            self.marks = [self.t, self.marks[0]]
        self.refresh_marks()

    def clear_marks(self):
        self.marks = [None, None]
        self.refresh_marks()

    def delete_marked(self):
        part = self.marked()
        if part:
            self.marks = [None, None]
            self.t = part[0]
            self.change(f"Deleted {precise(part[0])} to {precise(part[1])}",
                        lambda: self.edit.delete_range(*part))

    def keep_marked(self):
        part = self.marked()
        if part:
            self.marks = [None, None]
            self.t = 0.0
            self.change(f"Kept only {precise(part[0])} to {precise(part[1])}",
                        lambda: self.edit.keep_range(*part))

    def copy_marked(self):
        part = self.marked()
        if part:
            self.clipboard = self.edit.copy_range(*part)
            self.say(f"Copied {precise(part[1] - part[0])} s. Move the playhead and click Paste here.")

    def paste(self):
        if not self.clipboard:
            self.say("Copy a marked part first.")
            return
        at = self.t
        index = self.change(f"Pasted at {precise(at)}", lambda: self.edit.insert(at, self.clipboard))
        self.selected_piece = index
        self.refresh()

    def delete_key(self):
        if self.marks != [None, None]:
            self.delete_marked()          # tells the user if a mark is still missing
        else:
            self.remove_piece()

    def remove_piece(self):
        i = self.selected_piece
        if 0 <= i < len(self.edit.pieces):
            self.change(f"Removed {file_name(self.edit.pieces[i].source)}",
                        lambda: self.edit.remove(i))
        else:
            self.say("Select a piece first.")

    def move_piece(self, step):
        i = self.selected_piece
        if not 0 <= i < len(self.edit.pieces):
            self.say("Select a piece first.")
            return
        if not 0 <= i + step < len(self.edit.pieces):
            return
        self.selected_piece = self.change("Moved a piece", lambda: self.edit.move(i, step))
        self.refresh()
        self.seek(self.edit.starts()[self.selected_piece])

    def trim_here(self, keep_after):
        at = self.t
        index = self.edit.snapshot().trim_to_playhead(at, keep_after)     # a trial run
        if index < 0:
            self.say("The playhead is already at the edge of a piece.")
            return
        word = "starts" if keep_after else "ends"
        self.change(f"Piece {index + 1} now {word} at the playhead",
                    lambda: self.edit.trim_to_playhead(at, keep_after))
        self.selected_piece = index
        self.refresh()
        self.seek(self.edit.starts()[index] if keep_after else at)

    def piece_trimmed(self):
        i = self.selected_piece
        if not 0 <= i < len(self.edit.pieces):
            return
        start, end = self.piece_start.value(), self.piece_end.value()
        p = self.edit.pieces[i]
        if abs(start - p.start) < 0.005 and abs(end - p.end) < 0.005:
            return
        self.change(f"Trimmed piece {i + 1}", lambda: self.edit.set_trim(i, start, end))
        self.timeline.setFocus()

    def piece_sound_changed(self, *_):
        i = self.selected_piece
        if not 0 <= i < len(self.edit.pieces):
            return
        p = self.edit.pieces[i]
        new = (self.piece_mute.isChecked(), self.piece_volume.value() / 100,
               self.piece_fade_in.value(), self.piece_fade_out.value())
        if new == (p.muted, p.volume, p.fade_in, p.fade_out):
            return

        def apply():
            p.muted, p.volume, p.fade_in, p.fade_out = new
            fit_fades(p)
        self.change(f"Changed the sound of piece {i + 1}", apply)
        self.timeline.setFocus()

    def sound_changed(self, *_):
        j = self.selected_sound
        if not 0 <= j < len(self.edit.sounds):
            return
        s = self.edit.sounds[j]
        start = self.sound_from.value()
        end = max(self.sound_to.value(), start + engine.MIN_PIECE)
        new = (self.sound_at.value(), start, min(end, s.source_length),
               self.sound_volume.value() / 100, self.sound_fade_in.value(),
               self.sound_fade_out.value(), self.sound_replace.isChecked())
        if new == (s.at, s.start, s.end, s.volume, s.fade_in, s.fade_out, s.replace):
            return

        def apply():
            s.at, s.start, s.end, s.volume, s.fade_in, s.fade_out, s.replace = new
            fit_fades(s)
        self.change(f"Changed {file_name(s.source)}", apply)
        self.timeline.setFocus()

    def sound_to_playhead(self):
        if 0 <= self.selected_sound < len(self.edit.sounds):
            self.sound_at.setValue(self.t)

    def remove_sound(self):
        j = self.selected_sound
        if 0 <= j < len(self.edit.sounds):
            self.change(f"Removed {file_name(self.edit.sounds[j].source)}",
                        lambda: self.edit.sounds.pop(j))
        else:
            self.say("Select an added sound first.")

    # ------------------------------------------------------------- display
    def say(self, text):
        self.statusBar().showMessage(text, 8000)

    def refresh(self):
        """Bring the lists, forms and timeline up to date with the edit."""
        starts = self.edit.starts()
        self.piece_list.blockSignals(True)
        self.piece_list.clear()
        for i, (start, p) in enumerate(zip(starts, self.edit.pieces)):
            notes = []
            if not p.has_video:
                notes.append("sound only")
            elif not p.has_audio:
                notes.append("no sound")
            if p.muted:
                notes.append("muted")
            elif p.has_audio and abs(p.volume - 1) > 0.005:
                notes.append(f"volume {round(p.volume * 100)}%")
            if p.fade_in or p.fade_out:
                notes.append("fades")
            extra = f"  ({', '.join(notes)})" if notes else ""
            self.piece_list.addItem(
                f"{i + 1}. {file_name(p.source)}\n    {precise(p.start)} to {precise(p.end)} "
                f"of the file, {precise(p.length)} long{extra}")
        self.piece_list.setCurrentRow(self.selected_piece)
        self.piece_list.blockSignals(False)

        self.sound_list.blockSignals(True)
        self.sound_list.clear()
        for s in self.edit.sounds:
            how = "replaces the video's sound" if s.replace else "mixed in"
            past = "  (starts after the end)" if s.at >= self.edit.duration else ""
            self.sound_list.addItem(f"{file_name(s.source)}\n    at {precise(s.at)}, "
                                    f"{precise(s.length)} long, {how}{past}")
        self.sound_list.setCurrentRow(self.selected_sound)
        self.sound_list.blockSignals(False)

        self.fill_piece_form()
        self.fill_sound_form()
        self.timeline.edit = self.edit
        self.timeline.selected_piece = self.selected_piece
        self.timeline.selected_sound = self.selected_sound
        self.undo_button.setEnabled(bool(self.undo_stack))
        self.redo_button.setEnabled(bool(self.redo_stack))
        self.save_button.setEnabled(bool(self.edit.pieces))
        if not self.edit.pieces:
            self.screen_stack.setCurrentWidget(self.start_page)
        self.setWindowTitle(TITLE + (" - not saved yet" if self.dirty else ""))
        self.refresh_marks()

    def fill(self, widgets_values):
        for widget, value in widgets_values:
            widget.blockSignals(True)
            if isinstance(widget, QCheckBox):
                widget.setChecked(value)
            elif isinstance(widget, QDoubleSpinBox):
                widget.setMaximum(max(value, widget.maximum()))
                widget.setValue(value)
            else:
                widget.setValue(value)
                if hasattr(widget, "value_label"):
                    widget.value_label.setText(f"{value}%")
            widget.blockSignals(False)

    def fill_piece_form(self):
        i = self.selected_piece
        ok = 0 <= i < len(self.edit.pieces)
        self.piece_form.setEnabled(ok)
        if not ok:
            self.piece_name.setText("Select a piece to trim it or change its sound.")
            return
        p = self.edit.pieces[i]
        kind = "video" if p.has_video else "sound"
        self.piece_name.setText(f"Piece {i + 1}: {file_name(p.source)} "
                                f"({kind}, {precise(p.source_length)} long)")
        for box in (self.piece_start, self.piece_end):
            box.setMaximum(p.source_length)
        self.fill([(self.piece_start, p.start), (self.piece_end, p.end),
                   (self.piece_mute, p.muted), (self.piece_volume, round(p.volume * 100)),
                   (self.piece_fade_in, p.fade_in), (self.piece_fade_out, p.fade_out)])
        for w in (self.piece_mute, self.piece_volume, self.piece_fade_in, self.piece_fade_out):
            w.setEnabled(p.has_audio)

    def fill_sound_form(self):
        j = self.selected_sound
        ok = 0 <= j < len(self.edit.sounds)
        self.sound_form.setEnabled(ok)
        if not ok:
            self.sound_name.setText("Select an added sound to change it.")
            return
        s = self.edit.sounds[j]
        self.sound_name.setText(f"{file_name(s.source)} ({precise(s.source_length)} long)")
        for box in (self.sound_from, self.sound_to):
            box.setMaximum(s.source_length)
        self.sound_at.setMaximum(max(self.edit.duration, s.at))
        self.fill([(self.sound_at, s.at), (self.sound_from, s.start), (self.sound_to, s.end),
                   (self.sound_volume, round(s.volume * 100)), (self.sound_fade_in, s.fade_in),
                   (self.sound_fade_out, s.fade_out), (self.sound_replace, s.replace)])

    def refresh_marks(self):
        a, b = self.marks
        if a is not None and b is not None:
            self.marks_label.setText(f"Marked {precise(a)} to {precise(b)} "
                                     f"({precise(b - a)} long)")
        elif a is not None:
            self.marks_label.setText(f"Marked start {precise(a)}; now mark the end (O)")
        else:
            self.marks_label.setText("")
        self.timeline.marks = tuple(self.marks)
        self.update_position_display()

    def update_position_display(self):
        self.time_label.setText(f"{precise(self.t)} / {precise(self.edit.duration)}")
        self.timeline.position = self.t
        self.timeline.update()

    def piece_row_chosen(self, row):
        if 0 <= row < len(self.edit.pieces):
            self.select_piece(row)
            self.seek(self.edit.starts()[row])

    def sound_row_chosen(self, row):
        if 0 <= row < len(self.edit.sounds):
            self.select_sound(row)
            self.seek(self.edit.sounds[row].at)

    def select_piece(self, i):
        self.selected_piece = i
        self.tabs.setCurrentIndex(0)
        self.refresh()

    def select_sound(self, j):
        self.selected_sound = j
        self.tabs.setCurrentIndex(1)
        self.refresh()

    # ------------------------------------------------------------- preview
    def seek(self, t, play=None):
        play = self.playing if play is None else play
        self.t = min(max(0.0, t), self.edit.duration)
        if not self.edit.pieces:
            self.player.stop()
            self.sound_player.stop()
            self.playing = False
            self.current = -1
            self.loaded = None
            self.update_play_button()
            self.update_position_display()
            return
        i, inside = self.edit.locate(self.t)
        self.load(i, inside, play)
        self.update_position_display()

    def load(self, i, inside, play):
        p = self.edit.pieces[i]
        self.current = i
        self.screen_stack.setCurrentWidget(self.video if p.has_video else self.sound_only)
        target = p.start + inside
        S = QMediaPlayer.MediaStatus
        if self.loaded is None or not engine.same_path(self.loaded, p.source):
            self.loaded = p.source
            self.pending = (target, play)
            self.player.setSource(QUrl.fromLocalFile(p.source))
        elif self.pending is not None or self.player.mediaStatus() in (S.NoMedia, S.LoadingMedia):
            self.pending = (target, play)       # still loading: a jump now would be lost
        else:
            self.pending = None
            self.player.setPosition(int(target * 1000))
            self.set_playing(play)
        self.apply_sound_levels()

    def media_status_changed(self, status):
        S = QMediaPlayer.MediaStatus
        # While switching files the old one still reports "loaded" once: wait for the new one.
        if (status in (S.LoadedMedia, S.BufferedMedia) and self.pending is not None
                and self.is_current(self.player, self.loaded)):
            target, play = self.pending
            self.pending = None
            self.player.setPosition(int(target * 1000))
            self.set_playing(play)
        elif status == S.EndOfMedia and self.playing:
            self.next_piece()

    @staticmethod
    def is_current(player, path):
        source = player.source().toLocalFile()
        return bool(path and source) and engine.same_path(source, path)

    def player_error(self, error, message):
        if error != QMediaPlayer.Error.NoError:
            self.pending = None
            self.say(f"Can't preview {file_name(self.loaded or '')}: {message}")

    def set_playing(self, play):
        self.playing = play
        if play:
            self.player.play()
        else:
            self.player.pause()
        self.update_play_button()
        self.update_added_sound()

    def toggle_play(self):
        if not self.edit.pieces:
            self.add_to_end()
            return
        if self.playing:
            self.set_playing(False)
        else:
            if self.t >= self.edit.duration - 0.05:
                self.t = 0.0
            self.seek(self.t, play=True)

    def update_play_button(self):
        SP = QStyle.StandardPixmap
        self.play_button.setText("Pause" if self.playing else "Play")
        self.play_button.setIcon(self.style().standardIcon(
            SP.SP_MediaPause if self.playing else SP.SP_MediaPlay))

    def position_changed(self, ms):
        if self.pending is not None or not 0 <= self.current < len(self.edit.pieces):
            return
        p = self.edit.pieces[self.current]
        here = ms / 1000
        self.t = self.edit.starts()[self.current] + min(max(here - p.start, 0.0), p.length)
        if self.playing and here >= p.end - 0.02:
            self.next_piece()
            return
        self.apply_sound_levels()
        self.update_added_sound()
        self.update_position_display()

    def next_piece(self):
        i = self.current + 1
        if i >= len(self.edit.pieces):
            self.set_playing(False)
            self.t = self.edit.duration
            self.update_position_display()
            return
        p, q = self.edit.pieces[self.current], self.edit.pieces[i]
        if engine.same_path(p.source, q.source) and abs(q.start - p.end) < 0.05:
            self.current = i          # the next piece carries straight on: no jump needed
            return
        self.load(i, 0.0, True)

    def step_frame(self, direction):
        if not self.edit.pieces:
            return
        i, _ = self.edit.locate(self.t)
        fps = self.edit.pieces[i].fps or 25
        self.seek(self.t + direction / fps, play=False)

    def apply_sound_levels(self):
        master = self.preview_volume.value() / 100
        gain = 0.0
        if 0 <= self.current < len(self.edit.pieces):
            p = self.edit.pieces[self.current]
            inside = self.t - self.edit.starts()[self.current]
            if not p.silent:
                gain = min(1.0, p.volume) * fade_gain(inside, p.length, p.fade_in, p.fade_out)
        if any(s.replace and s.at <= self.t < s.finish for s in self.edit.sounds):
            gain = 0.0
        self.audio.setVolume(master * gain)

    def update_added_sound(self):
        end = self.edit.duration
        active = next((s for s in self.edit.sounds if s.at <= self.t < min(s.finish, end)), None)
        playing = self.sound_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        if active is None or not self.playing:
            if playing:
                self.sound_player.pause()
            return
        inside = self.t - active.at
        target = active.start + inside
        self.sound_audio.setVolume(self.preview_volume.value() / 100 * min(1.0, active.volume)
                                   * fade_gain(inside, active.length, active.fade_in,
                                               active.fade_out))
        if self.sound_loaded is None or not engine.same_path(self.sound_loaded, active.source):
            self.sound_loaded = active.source
            self.sound_pending = target
            self.sound_player.setSource(QUrl.fromLocalFile(active.source))
            return
        if self.sound_pending is not None:
            return
        if not playing or abs(self.sound_player.position() / 1000 - target) > 0.3:
            self.sound_player.setPosition(int(target * 1000))
            self.sound_player.play()

    def sound_status_changed(self, status):
        S = QMediaPlayer.MediaStatus
        if (status in (S.LoadedMedia, S.BufferedMedia) and self.sound_pending is not None
                and self.is_current(self.sound_player, self.sound_loaded)):
            self.sound_player.setPosition(int(self.sound_pending * 1000))
            self.sound_pending = None
            if self.playing:
                self.sound_player.play()

    # ---------------------------------------------------------------- saving
    def save_as(self):
        if not self.edit.pieces:
            self.say("Add some videos or sound first.")
            return
        if self.job is not None:
            return
        video = self.edit.has_video()
        formats = (list(engine.VIDEO_FORMATS.items()) if video else []) + list(
            engine.SOUND_FORMATS.items())
        sound_note = " - the sound only" if video else ""
        filters = [f"{label}{sound_note if ext in engine.SOUND_FORMATS else ''} (*{ext})"
                   for ext, (label, _) in formats]
        default_ext = ".mp4" if video else ".mp3"
        folder = self.folder("save_folder", QStandardPaths.StandardLocation.MoviesLocation
                             if video else QStandardPaths.StandardLocation.MusicLocation)
        stem = os.path.splitext(file_name(self.edit.pieces[0].source))[0] + " (edited)"
        path, chosen = QFileDialog.getSaveFileName(
            self, "Save the edit as a new file", os.path.join(folder, stem + default_ext),
            ";;".join(filters), filters[0])
        if not path:
            return
        path = pick_extension(path, chosen, filters, [ext for ext, _ in formats], default_ext)
        try:
            plan = engine.plan_export(self.edit, path)
        except ValueError as err:
            QMessageBox.warning(self, TITLE, str(err))
            return
        self.settings.setValue("save_folder", os.path.dirname(path))
        self.start_export(plan, path)

    def start_export(self, plan, path):
        self.set_playing(False)
        self.progress = QProgressDialog(f"Saving {file_name(path)}...", "Stop", 0, 1000, self)
        self.progress.setWindowTitle(TITLE)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setMinimumDuration(0)
        self.progress.setValue(0)
        self.job = ExportJob(plan, path, self)
        self.job.progress.connect(lambda share: self.progress.setValue(int(share * 1000)))
        self.progress.canceled.connect(self.job.stop)
        self.job.done.connect(self.export_done)

    def export_done(self, ok, message):
        self.progress.canceled.disconnect()        # closing the dialog would signal Stop
        self.progress.reset()
        self.progress.close()
        self.job.deleteLater()
        self.job = None
        if ok:
            self.dirty = False
            self.refresh()
            self.say(f"Saved {message}")
            self.show_saved(message)
        elif not message:
            self.say("Saving stopped. Nothing was written.")
        else:
            box = QMessageBox(QMessageBox.Icon.Warning, TITLE,
                              "The file couldn't be saved. ffmpeg said:", parent=self)
            box.setDetailedText(message)
            box.setInformativeText(message.splitlines()[-1] if message else "")
            box.exec()

    def show_saved(self, path):
        box = QMessageBox(QMessageBox.Icon.Information, TITLE, f"Saved:\n{path}", parent=self)
        play = box.addButton("Play it", QMessageBox.ButtonRole.ActionRole)
        show = box.addButton("Show in folder", QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Ok)
        box.exec()
        if box.clickedButton() is play:
            quiet = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            python = quiet if os.path.isfile(quiet) else sys.executable
            subprocess.Popen([python, vp.__file__, path])
        elif box.clickedButton() is show:
            subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')


def pick_extension(path, chosen, filters, extensions, suggested):
    """The file name to save to. A typed extension the editor knows is kept,
    unless it is just the suggested one left in place while a different format
    was picked in the Save window; then the picked format wins."""
    picked = extensions[filters.index(chosen)] if chosen in filters else None
    stem, typed = os.path.splitext(path)
    if typed.lower() in extensions:
        if picked and typed.lower() == suggested and picked != suggested:
            return stem + picked
        return path
    return path + (picked or suggested)


def settings_store():
    return QSettings(QSettings.Format.IniFormat, QSettings.Scope.UserScope, TITLE, "settings")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(TITLE)
    app.setFont(QFont("Segoe UI", 10))
    window = EditorWindow(settings_store())
    window.show()
    files = [p for p in sys.argv[1:] if os.path.exists(p)]
    if files:
        window.add_to_end(files)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
