"""
video_player.py

A simple video player for the common video types: MP4, MKV, WebM, AVI, MOV,
WMV, FLV, MPG, M4V, 3GP, TS and OGV, with their sound.

Opening one video lines up the other videos in the same folder after it, so
Next always has somewhere to go. Opening several videos or a whole folder
plays just those, in name order. Videos can also be dragged onto the window or
onto Run_VideoPlayer.bat.

Buttons and keys:
  Play / Pause ........ Space (or click the button)
  Back 10 s ........... Left arrow      Forward 10 s ... Right arrow
  Previous video ...... P or Page Up    Next video ..... N or Page Down
                        (Previous goes back to the start of the video first,
                        if more than 3 seconds have played)
  Volume .............. Up / Down arrows, M to mute
  Full screen ......... F or double-click the picture, Esc to leave
  Save frame .......... S (pauses and saves the picture on screen as a PNG
                        file, in a folder you pick; the folder is remembered)
  Open videos ......... Ctrl+O          Open a folder .. Ctrl+Shift+O

When a video ends, the next one starts. A video that can't be played is
marked in the list and skipped.

Requirements (one-time): PySide6, which brings its own copy of FFmpeg for
decoding. If the Python running this lacks it, it offers to install it.
"""

import importlib.util
import os
import re
import site
import subprocess
import sys
import traceback

TITLE = "Video Player"
PYSIDE = "pyside6>=6.10"        # 6.10 and later bundle FFmpeg on Windows

VIDEO_EXTENSIONS = {
    ".3g2", ".3gp", ".asf", ".avi", ".divx", ".f4v", ".flv", ".m2ts", ".m2v",
    ".m4v", ".mkv", ".mov", ".mp4", ".mpeg", ".mpg", ".mts", ".mxf", ".ogv",
    ".qt", ".ts", ".vob", ".webm", ".wmv",
}
# Played when picked by hand, but not gathered from folders, which are for videos.
AUDIO_EXTENSIONS = {".aac", ".flac", ".m4a", ".mp3", ".ogg", ".opus", ".wav", ".wma"}

SKIP_SECONDS = 10
RESTART_IF_PAST_MS = 3000       # Previous restarts the video first, past this point
SKIP_BROKEN_AFTER_MS = 1500     # time to read the error before moving on
HIDE_CONTROLS_AFTER_MS = 2500   # in full screen


# ------------------------------------------------------------- playlist logic
# Kept free of Qt so it can be tested on its own.

def natural_key(path):
    """Sort "clip 2" before "clip 10", ignoring case."""
    return [int(part) if part.isdigit() else part.casefold()
            for part in re.split(r"(\d+)", os.path.normpath(path))]


def is_video(path):
    return os.path.splitext(path)[1].lower() in VIDEO_EXTENSIONS


def is_playable(path):
    ext = os.path.splitext(path)[1].lower()
    return ext in VIDEO_EXTENSIONS or ext in AUDIO_EXTENSIONS


def videos_in_folder(folder):
    """The videos directly inside a folder (not its subfolders), in name order."""
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    paths = (os.path.join(folder, name) for name in names)
    return sorted((p for p in paths if is_video(p) and os.path.isfile(p)), key=natural_key)


def same_file(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def playlist_from(paths):
    """Turn files and folders that were opened or dropped into a play list.

    Returns (videos, start index, skipped). One file on its own brings along the
    other videos in its folder and starts at that file. Anything else plays just
    what was given, folders expanded, in name order. Files that aren't videos or
    sound are skipped; start is -1 when nothing is left to play.
    """
    paths = [os.path.abspath(p) for p in paths]
    if len(paths) == 1 and os.path.isfile(paths[0]):
        only = paths[0]
        if not is_playable(only):
            return [], -1, [only]
        videos = videos_in_folder(os.path.dirname(only))
        if not any(same_file(only, v) for v in videos):
            videos = sorted(videos + [only], key=natural_key)
        start = next(i for i, v in enumerate(videos) if same_file(only, v))
        return videos, start, []

    found, skipped = [], []
    for path in paths:
        if os.path.isdir(path):
            found.extend(videos_in_folder(path))
        elif os.path.isfile(path) and is_playable(path):
            found.append(path)
        else:
            skipped.append(path)
    unique = {os.path.normcase(p): p for p in found}
    videos = sorted(unique.values(), key=natural_key)
    return videos, (0 if videos else -1), skipped


def format_time(ms):
    seconds = max(0, int(ms) // 1000)
    hours, rest = divmod(seconds, 3600)
    minutes, seconds = divmod(rest, 60)
    if hours:
        return f"{hours}:{minutes:02}:{seconds:02}"
    return f"{minutes}:{seconds:02}"


def frame_file_name(video_path, ms):
    """A name for a saved frame: the video's name and where in it the frame is,
    to a hundredth of a second so frames close together don't share a name."""
    ms = max(0, int(ms))
    hours, rest = divmod(ms, 3_600_000)
    minutes, rest = divmod(rest, 60_000)
    seconds = f"{rest // 1000:02}.{rest % 1000 // 10:02}s"
    when = f"{hours}h{minutes:02}m{seconds}" if hours else f"{minutes}m{seconds}"
    stem = os.path.splitext(os.path.basename(video_path))[0]
    return f"{stem} {when}.png"


class Playlist:
    """The list of videos and which one is playing."""

    def __init__(self):
        self.items = []
        self.index = -1

    def replace(self, items, start=0):
        self.items = list(items)
        self.index = start if 0 <= start < len(self.items) else -1

    @property
    def current(self):
        return self.items[self.index] if 0 <= self.index < len(self.items) else None

    def has_next(self):
        return 0 <= self.index < len(self.items) - 1

    def has_previous(self):
        return self.index > 0

    def select(self, index):
        if not 0 <= index < len(self.items):
            raise IndexError(index)
        self.index = index
        return self.current

    def step(self, offset):
        """Move by offset and return the new video, or None at either end."""
        target = self.index + offset
        if not 0 <= target < len(self.items):
            return None
        return self.select(target)


# ------------------------------------------------------- start-up and crashes

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
    """A window started without a console closes the moment it fails, taking
    the error with it, so show anything unexpected in a dialog."""
    text = "".join(traceback.format_exception(kind, value, tb))
    print(text, file=sys.stderr)
    message = "Something went wrong:\n\n" + text[-1500:]
    qt_widgets = sys.modules.get("PySide6.QtWidgets")
    if qt_widgets and qt_widgets.QApplication.instance():
        qt_widgets.QMessageBox.critical(None, TITLE, message)
    else:
        tk_message("showerror", message)


def has_pyside():
    try:
        return importlib.util.find_spec("PySide6.QtMultimedia") is not None
    except ModuleNotFoundError:
        return False


def install_missing():
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add PySide6 if this one lacks it."""
    if has_pyside():
        return
    if not tk_message("askyesno",
                      f"This copy of Python ({sys.version.split()[0]}) still needs PySide6, "
                      "which plays the videos.\n\nDownload and install it now? This only "
                      "happens once. It is about 150 MB, so it can take a few minutes; "
                      "a window shows the progress."):
        sys.exit()
    command = [sys.executable, "-m", "pip", "install", PYSIDE]
    flags = subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0   # show pip's progress
    if subprocess.run(command, creationflags=flags).returncode:
        raise RuntimeError("Could not install PySide6. Check the internet connection "
                           "and try again.")
    site.addsitedir(site.getusersitepackages())   # a first install may have just created it
    importlib.invalidate_caches()


sys.excepthook = show_crash
install_missing()

from PySide6.QtCore import QEvent, QSettings, QSize, QStandardPaths, Qt, QTimer, QUrl
from PySide6.QtGui import QAction, QFont, QIcon, QKeySequence, QShortcut, QWindow
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer, QtAudio
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QMainWindow, QSlider, QSplitter, QStackedWidget, QStyle, QToolButton,
    QVBoxLayout, QWidget,
)


# ------------------------------------------------------------------ window

def file_filter():
    videos = " ".join("*" + ext for ext in sorted(VIDEO_EXTENSIONS))
    sounds = " ".join("*" + ext for ext in sorted(AUDIO_EXTENSIONS))
    return f"Videos ({videos});;Sound files ({sounds});;All files (*)"


class SeekSlider(QSlider):
    """A slider that jumps straight to where it is clicked, instead of creeping
    toward it a page at a time."""

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.maximum() > self.minimum():
            value = QStyle.sliderValueFromPosition(
                self.minimum(), self.maximum(), int(event.position().x()), self.width())
            self.setValue(value)
            self.sliderMoved.emit(value)
        super().mousePressEvent(event)


class PlayerWindow(QMainWindow):

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.playlist = Playlist()
        self.broken = set()             # list rows that failed to play
        self.list_was_visible = True
        self.was_maximized = False

        self.player = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.player.setAudioOutput(self.audio)
        self.video = QVideoWidget()
        self.player.setVideoOutput(self.video)

        self.build_window()
        self.build_shortcuts()
        self.connect_player()

        self.hide_timer = QTimer(self, singleShot=True, interval=HIDE_CONTROLS_AFTER_MS)
        self.hide_timer.timeout.connect(self.hide_controls)
        self.skip_timer = QTimer(self, singleShot=True, interval=SKIP_BROKEN_AFTER_MS)
        self.skip_timer.timeout.connect(self.skip_broken)
        QApplication.instance().installEventFilter(self)   # mouse moves and double-clicks on the picture

        self.restore_settings()
        self.update_buttons()

    # ---------------------------------------------------------------- layout
    def icon(self, picture):
        """A theme icon (Windows 11 draws these) or one of the style's own pictures."""
        if isinstance(picture, QIcon.ThemeIcon):
            return QIcon.fromTheme(picture, self.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarMaxButton))
        return self.style().standardIcon(picture)

    def button(self, text, icon, tip, slot):
        b = QToolButton()
        b.setText(text)
        b.setIcon(self.icon(icon))
        b.setToolTip(tip)
        b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)      # leave Space and the arrows to the shortcuts
        b.setIconSize(QSize(20, 20))
        b.clicked.connect(slot)
        return b

    def build_window(self):
        self.setWindowTitle(TITLE)
        self.setAcceptDrops(True)
        SP = QStyle.StandardPixmap

        self.hint = QLabel("Drag videos or a folder here,\nor click Open videos or Open folder.")
        self.hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint.setStyleSheet("background: black; color: #bbbbbb; font-size: 16pt;")
        self.screen_stack = QStackedWidget()
        self.screen_stack.addWidget(self.hint)
        self.screen_stack.addWidget(self.video)
        self.screen_stack.setMinimumSize(320, 180)

        self.list = QListWidget()
        self.list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.list.setToolTip("Double-click a video to play it")
        self.list.itemActivated.connect(lambda item: self.play_index(self.list.row(item)))
        self.splitter = QSplitter()
        self.splitter.addWidget(self.screen_stack)
        self.splitter.addWidget(self.list)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setSizes([900, 260])

        self.seek = SeekSlider(Qt.Orientation.Horizontal)
        self.seek.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.seek.setToolTip("Drag or click to move through the video")
        self.seek.sliderMoved.connect(self.seek_to)
        self.time_label = QLabel("0:00 / 0:00")
        self.time_label.setMinimumWidth(130)
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.open_button = self.button("Open videos", SP.SP_DialogOpenButton,
                                       "Open one or more videos (Ctrl+O)", self.choose_files)
        self.folder_button = self.button("Open folder", SP.SP_DirOpenIcon,
                                         "Play every video in a folder (Ctrl+Shift+O)",
                                         self.choose_folder)
        self.prev_button = self.button("Previous", SP.SP_MediaSkipBackward,
                                       "Previous video (P or Page Up)", self.previous_video)
        self.back_button = self.button(f"{SKIP_SECONDS} s", SP.SP_MediaSeekBackward,
                                       f"Rewind {SKIP_SECONDS} seconds (Left arrow)",
                                       lambda: self.skip(-SKIP_SECONDS))
        self.play_button = self.button("Play", SP.SP_MediaPlay, "Play or pause (Space)",
                                       self.toggle_play)
        self.play_button.setMinimumWidth(90)
        self.forward_button = self.button(f"{SKIP_SECONDS} s", SP.SP_MediaSeekForward,
                                          f"Forward {SKIP_SECONDS} seconds (Right arrow)",
                                          lambda: self.skip(SKIP_SECONDS))
        self.next_button = self.button("Next", SP.SP_MediaSkipForward,
                                       "Next video (N or Page Down)", self.next_video)
        self.frame_button = self.button("Save frame", SP.SP_DialogSaveButton,
                                        "Pause and save the picture on screen as a PNG file (S)",
                                        self.save_frame)
        self.mute_button = self.button("", SP.SP_MediaVolume, "Mute or unmute (M)",
                                       self.toggle_mute)
        self.volume = QSlider(Qt.Orientation.Horizontal)
        self.volume.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.volume.setRange(0, 100)
        self.volume.setFixedWidth(110)
        self.volume.setToolTip("Volume (Up and Down arrows)")
        self.volume.valueChanged.connect(self.set_volume)
        self.full_button = self.button("Full screen", QIcon.ThemeIcon.ViewFullscreen,
                                       "Full screen (F or double-click the picture, Esc to leave)",
                                       self.toggle_full_screen)
        self.list_button = self.button("List", SP.SP_FileDialogListView,
                                       "Show or hide the list of videos", self.toggle_list)

        seek_row = QHBoxLayout()
        seek_row.addWidget(self.seek, 1)
        seek_row.addWidget(self.time_label)
        button_row = QHBoxLayout()
        for widget in (self.open_button, self.folder_button):
            button_row.addWidget(widget)
        button_row.addStretch(1)
        for widget in (self.prev_button, self.back_button, self.play_button,
                       self.forward_button, self.next_button):
            button_row.addWidget(widget)
        button_row.addSpacing(12)
        button_row.addWidget(self.frame_button)
        button_row.addStretch(1)
        for widget in (self.mute_button, self.volume, self.full_button, self.list_button):
            button_row.addWidget(widget)

        self.controls = QWidget()
        controls_layout = QVBoxLayout(self.controls)
        controls_layout.setContentsMargins(8, 4, 8, 6)
        controls_layout.addLayout(seek_row)
        controls_layout.addLayout(button_row)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.splitter, 1)
        layout.addWidget(self.controls)
        self.setCentralWidget(central)
        self.statusBar().showMessage("Open some videos to start.")

    def build_shortcuts(self):
        keys = {
            "Space": self.toggle_play, "K": self.toggle_play,
            "Left": lambda: self.skip(-SKIP_SECONDS), "Right": lambda: self.skip(SKIP_SECONDS),
            "N": self.next_video, "PgDown": self.next_video,
            "P": self.previous_video, "PgUp": self.previous_video,
            "Up": lambda: self.volume.setValue(self.volume.value() + 5),
            "Down": lambda: self.volume.setValue(self.volume.value() - 5),
            "M": self.toggle_mute, "F": self.toggle_full_screen, "F11": self.toggle_full_screen,
            "S": self.save_frame,
            "Esc": self.leave_full_screen,
            "Ctrl+O": self.choose_files, "Ctrl+Shift+O": self.choose_folder,
        }
        for key, slot in keys.items():
            QShortcut(QKeySequence(key), self, slot)
        # Shortcuts belong to the window, which goes on receiving them in full screen.
        quit_action = QAction(self, shortcut=QKeySequence("Ctrl+Q"), triggered=self.close)
        self.addAction(quit_action)

    def connect_player(self):
        p = self.player
        p.positionChanged.connect(self.position_changed)
        p.durationChanged.connect(self.duration_changed)
        p.playbackStateChanged.connect(lambda _: self.update_buttons())
        p.mediaStatusChanged.connect(self.media_status_changed)
        p.errorOccurred.connect(self.player_error)

    # -------------------------------------------------------------- settings
    def restore_settings(self):
        s = self.settings
        self.volume.setValue(int(s.value("volume", 80)))
        self.set_volume(self.volume.value())    # setValue stays quiet when nothing changed
        self.audio.setMuted(s.value("muted", "false") == "true")
        self.update_mute_button()
        geometry = s.value("geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)
        else:
            self.resize(1180, 720)

    def closeEvent(self, event):
        QApplication.instance().removeEventFilter(self)
        self.leave_full_screen()
        s = self.settings
        s.setValue("volume", self.volume.value())
        s.setValue("muted", "true" if self.audio.isMuted() else "false")
        s.setValue("geometry", self.saveGeometry())
        self.player.stop()
        super().closeEvent(event)

    # ------------------------------------------------------- opening videos
    def choose_files(self):
        start = self.settings.value("folder", os.path.expanduser("~/Videos"))
        paths, _ = QFileDialog.getOpenFileNames(self, "Open videos", start, file_filter())
        if paths:
            self.open_paths(paths)

    def choose_folder(self):
        start = self.settings.value("folder", os.path.expanduser("~/Videos"))
        folder = QFileDialog.getExistingDirectory(self, "Play every video in this folder", start)
        if folder:
            self.open_paths([folder])

    def open_paths(self, paths):
        videos, start, skipped = playlist_from(paths)
        if not videos:
            names = ", ".join(os.path.basename(p) or p for p in (skipped or paths)[:3])
            self.statusBar().showMessage(f"No videos to play in {names}.")
            return
        first = paths[0]
        self.settings.setValue("folder", first if os.path.isdir(first) else os.path.dirname(first))
        self.playlist.replace(videos, start)
        self.broken.clear()
        self.list.clear()
        for path in videos:
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            self.list.addItem(item)
        self.play_index(start)
        if skipped:
            self.statusBar().showMessage(
                f"Skipped {len(skipped)} file(s) that aren't videos.", 6000)

    def play_index(self, index):
        self.skip_timer.stop()
        path = self.playlist.select(index)
        self.screen_stack.setCurrentWidget(self.video)
        self.player.setSource(QUrl.fromLocalFile(path))
        self.player.play()
        self.mark_current()
        self.setWindowTitle(f"{os.path.basename(path)} - {TITLE}")
        self.statusBar().showMessage(
            f"Video {index + 1} of {len(self.playlist.items)}: {path}")
        self.update_buttons()

    def mark_current(self):
        for row in range(self.list.count()):
            item = self.list.item(row)
            font = item.font()
            font.setBold(row == self.playlist.index)
            item.setFont(font)
            if row in self.broken:
                item.setForeground(Qt.GlobalColor.gray)
        if self.playlist.index >= 0:
            self.list.setCurrentRow(self.playlist.index)
            self.list.scrollToItem(self.list.currentItem())

    # ------------------------------------------------------------- controls
    def toggle_play(self):
        if self.playlist.current is None:
            self.choose_files()
        elif self.player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.player.pause()
        else:
            self.player.play()

    def skip(self, seconds):
        if self.playlist.current is None:
            return
        target = self.player.position() + seconds * 1000
        duration = self.player.duration()
        if duration > 0:
            target = min(target, duration)
        self.player.setPosition(max(0, target))

    def seek_to(self, ms):
        self.player.setPosition(ms)
        self.show_time(ms, self.player.duration())

    def next_video(self):
        if self.playlist.has_next():
            self.play_index(self.playlist.index + 1)
        elif self.playlist.current is not None:
            self.statusBar().showMessage("That was the last video in the list.", 4000)

    def previous_video(self):
        if self.playlist.current is None:
            return
        if self.player.position() > RESTART_IF_PAST_MS or not self.playlist.has_previous():
            self.player.setPosition(0)
            self.player.play()
        else:
            self.play_index(self.playlist.index - 1)

    def save_frame(self):
        """Pause, then save the picture on screen as a PNG, at the video's own size."""
        if self.playlist.current is None:
            return
        self.player.pause()
        frame = self.video.videoSink().videoFrame()
        image = frame.toImage() if frame.isValid() else None
        if image is None or image.isNull():
            self.statusBar().showMessage("There is no picture on screen to save.", 5000)
            return
        pictures = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.PicturesLocation)
        folder = self.settings.value("frame_folder", pictures)
        if not os.path.isdir(folder):
            folder = pictures
        # The frame's own time; the player's clock can run a frame or two ahead of the picture.
        shown_at = frame.startTime() // 1000 if frame.startTime() >= 0 else self.player.position()
        suggested = os.path.join(folder, frame_file_name(self.playlist.current, shown_at))
        path, _ = QFileDialog.getSaveFileName(self, "Save this frame as a picture", suggested,
                                              "PNG pictures (*.png)")
        if not path:
            return
        if not path.lower().endswith(".png"):
            path += ".png"
        if image.save(path, "PNG"):
            self.settings.setValue("frame_folder", os.path.dirname(path))
            self.statusBar().showMessage(
                f"Saved the frame ({image.width()} x {image.height()}) to {path}", 8000)
        else:
            self.statusBar().showMessage(f"Couldn't save the frame to {path}.", 8000)

    def set_volume(self, value):
        linear = QtAudio.convertVolume(value / 100, QtAudio.VolumeScale.LogarithmicVolumeScale,
                                       QtAudio.VolumeScale.LinearVolumeScale)
        self.audio.setVolume(linear)
        if value and self.audio.isMuted():
            self.audio.setMuted(False)
        self.update_mute_button()

    def toggle_mute(self):
        self.audio.setMuted(not self.audio.isMuted())
        self.update_mute_button()

    def update_mute_button(self):
        silent = self.audio.isMuted() or self.volume.value() == 0
        icon = QStyle.StandardPixmap.SP_MediaVolumeMuted if silent else QStyle.StandardPixmap.SP_MediaVolume
        self.mute_button.setIcon(self.style().standardIcon(icon))

    def update_buttons(self):
        playing = self.player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        icon = QStyle.StandardPixmap.SP_MediaPause if playing else QStyle.StandardPixmap.SP_MediaPlay
        self.play_button.setIcon(self.style().standardIcon(icon))
        self.play_button.setText("Pause" if playing else "Play")
        loaded = self.playlist.current is not None
        for b in (self.back_button, self.forward_button, self.prev_button, self.frame_button):
            b.setEnabled(loaded)
        self.next_button.setEnabled(self.playlist.has_next())

    # ------------------------------------------------------- player events
    def position_changed(self, ms):
        if not self.seek.isSliderDown():
            self.seek.setValue(ms)
            self.show_time(ms, self.player.duration())

    def duration_changed(self, ms):
        self.seek.setRange(0, max(0, ms))
        self.show_time(self.player.position(), ms)

    def show_time(self, position, duration):
        self.time_label.setText(f"{format_time(position)} / {format_time(duration)}")

    def media_status_changed(self, status):
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            if self.playlist.has_next():
                self.play_index(self.playlist.index + 1)
            else:
                self.statusBar().showMessage("Finished the last video in the list.")
                self.update_buttons()

    def player_error(self, error, message):
        if error == QMediaPlayer.Error.NoError or self.playlist.current is None:
            return
        self.broken.add(self.playlist.index)
        self.mark_current()
        name = os.path.basename(self.playlist.current)
        more = " Moving on to the next one." if self.playlist.has_next() else ""
        self.statusBar().showMessage(f"Can't play {name}: {message or 'unknown problem'}.{more}")
        if self.playlist.has_next():
            self.skip_timer.setProperty("index", self.playlist.index)
            self.skip_timer.start()

    def skip_broken(self):
        # Only move on if nobody picked another video in the meantime.
        if self.skip_timer.property("index") == self.playlist.index and self.playlist.has_next():
            self.play_index(self.playlist.index + 1)

    # --------------------------------------------------------- full screen
    def toggle_full_screen(self):
        if self.isFullScreen():
            self.leave_full_screen()
        else:
            self.list_was_visible = self.list.isVisible()
            self.was_maximized = self.isMaximized()
            self.list.hide()
            self.statusBar().hide()
            self.showFullScreen()
            self.full_button.setText("Leave full screen")
            self.hide_timer.start()

    def leave_full_screen(self):
        if not self.isFullScreen():
            return
        self.hide_timer.stop()
        if self.was_maximized:
            self.showMaximized()
        else:
            self.showNormal()
        self.controls.show()
        self.statusBar().show()
        self.list.setVisible(self.list_was_visible)
        self.unsetCursor()
        self.video.unsetCursor()
        self.full_button.setText("Full screen")

    def hide_controls(self):
        if self.isFullScreen() and not self.controls.underMouse():
            self.controls.hide()
            self.video.setCursor(Qt.CursorShape.BlankCursor)

    def toggle_list(self):
        self.list.setVisible(not self.list.isVisible())

    def over_picture(self, event):
        """True when a mouse event happened over the video picture."""
        pos = self.video.mapFromGlobal(event.globalPosition().toPoint())
        return self.video.isVisible() and self.video.rect().contains(pos)

    def eventFilter(self, obj, event):
        # Watch mouse events for the whole window here, because the video widget
        # and its inner parts would otherwise each need handlers. Each event
        # arrives once at the window (QWindow) and again at every widget it passes
        # through, so only the window copy is used.
        kind = event.type()
        if isinstance(obj, QWindow):
            if kind == QEvent.Type.MouseMove and self.isFullScreen():
                self.controls.show()
                self.video.unsetCursor()
                self.hide_timer.start()
            elif kind == QEvent.Type.MouseButtonDblClick and self.over_picture(event):
                self.toggle_full_screen()
                return True
        return False

    # -------------------------------------------------------- drag and drop
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.open_paths(paths)
            self.activateWindow()


def settings_store():
    # An .ini file in %APPDATA%\Video Player, rather than the registry, so it is easy to find.
    return QSettings(QSettings.Format.IniFormat, QSettings.Scope.UserScope, TITLE, "settings")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(TITLE)
    app.setFont(QFont("Segoe UI", 10))
    window = PlayerWindow(settings_store())
    window.show()
    if len(sys.argv) > 1:
        window.open_paths(sys.argv[1:])
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
