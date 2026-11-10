"""
music_player.py

A music player for the music on this PC, internet radio and Spotify.

My Music
  Open a folder plays every song in it and in all its subfolders. Add songs
  and Add a folder put more on the end of the list. Songs, folders and
  playlists can also be dragged onto the window or onto Run_MusicPlayer.bat.
  The last 4 folders opened stay as buttons, each with its own Remove button.
  Search the list as you type; click a column heading to sort by it. Shuffle,
  Repeat (whole list or one song), playback speed, Play next and Up next,
  saving and opening playlists. The list and the place in the song are
  remembered for next time.

Radio
  Search about 50,000 free stations by name, type of music or country, keep
  favourites in My stations, or add one by its web address. The last hour of
  a station is kept, so live radio can be paused, rewound (Back 10 s / 30 s),
  sent forward, taken back to the start of the song playing, or a marked part
  repeated over and over; Go live catches up again.

Spotify
  Open the Spotify app starts Spotify from here (a box offers to, when it
  isn't open); start any playlist there. Or paste a playlist's link once (in
  Spotify: Share > Copy link to playlist) and Play in Spotify opens it. The
  buttons at the top, the volume and mute then work Spotify, and the song
  playing shows here.

Sleep timer
  Stops the music after 5 minutes to 3 hours (or any number of minutes
  typed in Other), or at the end of the song,
  fading out over the last minute; then it can also close the player, or put
  the computer to sleep or shut it down after a 30-second warning that can
  be cancelled.

Keys
  Play / Pause .......... Space            Stop ............... Ctrl+.
  Back / Forward 10 s ... Left / Right     Previous / Next .... P / N
  Volume ................ + and -          Mute ............... M
  Open a folder ......... Ctrl+O           Add songs .......... Ctrl+Shift+O
  Find in the list ...... Ctrl+F           Go live (radio) .... L
  The keyboard's own media keys work too while this window is in front.

Requirements (one-time): PySide6 (which plays the music), mutagen (song
names and pictures), imageio-ffmpeg (radio), winrt-Windows.Media.Control and
pycaw (Spotify's buttons and volume). If the Python running this lacks any of
them, it offers to install them.
"""

import importlib.util
import json
import os
import re
import site
import subprocess
import sys
import threading
import time
import traceback

TITLE = "Music Player"

NEEDS = {
    # module to look for: (what pip installs, what it is for, download size in MB)
    "PySide6.QtMultimedia": ("pyside6>=6.10", "PySide6, which plays the music", 150),
    "mutagen": ("mutagen", "Mutagen, which reads song names and pictures", 1),
    "imageio_ffmpeg": ("imageio-ffmpeg", "FFmpeg, which brings in radio stations", 30),
    "winrt.windows.media.control": ("winrt-Windows.Media.Control",
                                    "WinRT, which works Spotify's buttons", 2),
    "winrt.windows.media": ("winrt-Windows.Media", "WinRT media, for Spotify's Repeat", 2),
    "pycaw": ("pycaw", "pycaw, which sets Spotify's volume", 2),
}


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
        # Shown over the player's window, so it gets the same big words.
        owner = next((w for w in qt_widgets.QApplication.topLevelWidgets()
                      if type(w).__name__ == "MusicWindow" and w.isVisible()), None)
        qt_widgets.QMessageBox.critical(owner, TITLE, message)
    else:
        tk_message("showerror", message)


def has_module(name):
    """True if the module is really there. A folder that only holds another
    add-on's subfolder (winrt.windows.media before its own add-on is in)
    doesn't count."""
    try:
        spec = importlib.util.find_spec(name)
    except ModuleNotFoundError:
        return False
    return spec is not None and spec.origin not in (None, "namespace")


def install_missing(needs=NEEDS):
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add whatever this one lacks."""
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
        raise RuntimeError("Could not install the add-ons. Check the internet connection "
                           "and try again.")
    site.addsitedir(site.getusersitepackages())   # a first install may have just created it
    importlib.invalidate_caches()


if __name__ == "__main__":
    sys.excepthook = show_crash
    install_missing()

from PySide6.QtCore import (
    QAbstractTableModel, QByteArray, QEvent, QLocale, QModelIndex, QObject, QSettings, QSize,
    QSortFilterProxyModel, Qt, QTimer, QUrl, Signal,
)
from PySide6.QtGui import (
    QAction, QColor, QFont, QFontMetrics, QGuiApplication, QIcon, QKeySequence, QPainter, QPixmap,
    QShortcut,
)
from PySide6.QtMultimedia import (
    QAudioFormat, QAudioOutput, QAudioSink, QMediaDevices, QMediaPlayer, QtAudio,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket, QNetworkAccessManager, QNetworkRequest
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QButtonGroup, QComboBox, QDialog, QFileDialog,
    QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPushButton,
    QFrame, QRadioButton, QScrollArea, QSizePolicy, QSlider, QSpinBox, QSplitter, QStyle,
    QTableView,
    QTableWidget, QTableWidgetItem, QTabWidget, QTextBrowser, QToolButton, QToolTip, QVBoxLayout,
    QWidget,
)

import music_player_engine as engine
import music_player_radio as radio
import music_player_spotify as spotify
from music_player_engine import PlayList, RecentFolders, SleepTimer, Track, format_time

BLUE = "#2f6fde"                 # buttons only (white words on it); all other words are black
WARNING_RED = "#b3261e"          # Unmute, and songs that couldn't be played
SKIP_SECONDS = 10
LONG_SKIP_SECONDS = 30
RESTART_IF_PAST = 3              # Previous restarts the song first, past this many seconds
SKIP_BROKEN_AFTER_MS = 1500
PREBUFFER_SECONDS = 1.5          # radio sound gathered before playing starts
LIVE_GAP_SECONDS = 1.0           # Go live plays this far behind the newest sound
LIVE_SLACK_SECONDS = 4           # closer than this to the newest sound counts as live
SPEEDS = (0.5, 0.75, 1.0, 1.25, 1.5, 2.0)
TRANSPORT_HEIGHT = 96            # the top row of buttons: picture above a big word
SECOND_ROW_HEIGHT = 64           # the row under it, and radio's own buttons
TIMER_CHOICES = ((5, "5 minutes"), (10, "10 minutes"), (15, "15 minutes"), (30, "30 minutes"),
                 (45, "45 minutes"), (60, "1 hour"), (90, "1½ hours"), (120, "2 hours"),
                 (180, "3 hours"))
# Word sizes, as times the usual Windows size (at least double, to be easy to read).
WINDOW_TEXT = 2.0                # the player's own window
POPUP_TEXT = 2.0                 # pop-up boxes
LIST_TEXT = 2.2                  # the song and station lists
ROW_PADDING = 10                 # pixels around the words in each list row (one line per row)
MUSIC, RADIO, SPOTIFY = "music", "radio", "spotify"


# ------------------------------------------------------------ small pieces

def file_filter():
    music = " ".join("*" + ext for ext in sorted(engine.MUSIC_EXTENSIONS))
    lists = " ".join("*" + ext for ext in sorted(engine.PLAYLIST_EXTENSIONS))
    return f"Music ({music});;Playlists ({lists});;All files (*)"


def plural(count, word):
    return f"{count:,} {word}" if count == 1 else f"{count:,} {word}s"


def long_time(seconds):
    hours, rest = divmod(int(seconds), 3600)
    minutes = rest // 60
    if not hours and not minutes:
        return "under a minute" if seconds else "length not known"
    if hours:
        return f"{plural(hours, 'hour')} {plural(minutes, 'minute')}"
    return plural(minutes, "minute")


def base_points():
    """The usual size of words in Windows (the program's own setting), in
    points; the window, lists and boxes are made bigger from this."""
    size = QApplication.font().pointSizeF()
    return size if size > 0 else 10


def words(scale=1.0):
    """The size of the window's words, in points, times scale."""
    return base_points() * WINDOW_TEXT * scale


def black_words(app):
    """Every word black, not Windows' soft grey-black, while the window is in
    front or not. Greyed-out (disabled) buttons stay grey: that shows they
    can't be pressed just now."""
    from PySide6.QtGui import QPalette
    palette = app.palette()
    roles = (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text,
             QPalette.ColorRole.ButtonText, QPalette.ColorRole.PlaceholderText,
             QPalette.ColorRole.ToolTipText, QPalette.ColorRole.BrightText)
    for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive):
        for role in roles:
            palette.setColor(group, role, QColor("black"))
    app.setPalette(palette)


def popup_style(width):
    """For the message and question boxes: big words, big buttons, and room
    to use (about half the screen) rather than a narrow column of text."""
    big = base_points() * POPUP_TEXT
    return (f"QMessageBox QLabel, QInputDialog QLabel, QInputDialog QLineEdit,"
            f" QMessageBox QPushButton, QInputDialog QPushButton {{ font-size: {big:g}pt; }}"
            f" QMessageBox QLabel#qt_msgbox_label {{ min-width: {width}px; }}"
            f" QInputDialog QLineEdit {{ min-width: {width}px; padding: 4px; }}"
            " QMessageBox QPushButton, QInputDialog QPushButton {"
            " min-width: 150px; min-height: 52px; padding: 4px 24px; }")


class MinutesBox(QSpinBox):
    """The Other: minutes box. Clicking it picks Other and selects the number
    already there, so typing 1 makes it 1. (Before, a click could leave the
    typing out of the box, or mix it in with the 20 there into 120.)"""

    picked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        # Clicks on the number itself go to the box's text part, not to the box.
        self.lineEdit().installEventFilter(self)

    def clicked_in(self):
        self.picked.emit()
        QTimer.singleShot(0, self.selectAll)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            self.clicked_in()
        return False

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.clicked_in()

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        self.clicked_in()


class ColumnSharer(QObject):
    """Keeps a list's columns in proportion as it changes size: the fixed ones
    (a number, a length) keep their width and the rest share what is left, so
    the song names don't take all the room."""

    def __init__(self, table, fixed, shares):
        super().__init__(table)
        self.table, self.fixed, self.shares = table, fixed, shares
        table.horizontalHeader().setStretchLastSection(False)
        table.viewport().installEventFilter(self)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.Resize:
            self.share()
        return False

    def share(self):
        header = self.table.horizontalHeader()
        for column, width in self.fixed.items():
            header.resizeSection(column, width)
        room = max(0, self.table.viewport().width() - sum(self.fixed.values()))
        total = sum(self.shares.values())
        widths = {c: max(60, int(room * s / total)) for c, s in self.shares.items()}
        first = next(iter(widths))
        widths[first] += max(0, room - sum(widths.values()))    # the odd pixels left over
        for column, width in widths.items():
            header.resizeSection(column, width)


def list_font():
    font = QFont(QApplication.font())
    font.setPointSizeF(base_points() * LIST_TEXT)
    return font


def make_big_list(table, fixed, shares):
    """Bigger words, rows just tall enough for them (one line each, so as
    many songs show as can; a name too long ends in ... and shows whole when
    pointed at), and the columns sharing the width."""
    font = list_font()
    table.setFont(font)
    table.horizontalHeader().setFont(font)
    table.setStyleSheet(f"QTableView, QHeaderView {{ font-size: {font.pointSizeF():g}pt; }}")
    table.setWordWrap(False)
    table.setTextElideMode(Qt.TextElideMode.ElideRight)
    table.setIconSize(QSize(28, 28))
    table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
    table.verticalHeader().setMinimumSectionSize(1)
    table.verticalHeader().setDefaultSectionSize(QFontMetrics(font).lineSpacing() + ROW_PADDING)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    return ColumnSharer(table, fixed, shares)


def placeholder_picture(size, symbol="♫"):
    picture = QPixmap(size, size)
    picture.fill(QColor("#e8eaed"))
    painter = QPainter(picture)
    font = QFont("Segoe UI Symbol", int(size * 0.42))
    painter.setFont(font)
    painter.setPen(QColor("#9aa0a6"))
    painter.drawText(picture.rect(), Qt.AlignmentFlag.AlignCenter, symbol)
    painter.end()
    return picture


def glyph_icon(symbol, size=64, colour="black"):
    """A button picture drawn from a character, for pictures Windows doesn't have."""
    picture = QPixmap(size, size)
    picture.fill(Qt.GlobalColor.transparent)
    painter = QPainter(picture)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    painter.setFont(QFont("Segoe UI Symbol", int(size * 0.55)))
    painter.setPen(QColor(colour))
    painter.drawText(picture.rect(), Qt.AlignmentFlag.AlignCenter, symbol)
    painter.end()
    return QIcon(picture)


class SeekSlider(QSlider):
    """A slider that jumps straight to where it is clicked."""

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.maximum() > self.minimum():
            share = event.position().x() / max(1, self.width())
            value = self.minimum() + round(share * (self.maximum() - self.minimum()))
            self.setValue(value)
            self.sliderMoved.emit(value)
        super().mousePressEvent(event)


class ElidedButton(QToolButton):
    """A button whose words are cut short with "..." to fit however wide it
    is, so a long folder name never widens the side panel."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.full_text = ""

    def set_full_text(self, text):
        self.full_text = text
        self.elide()

    def elide(self):
        room = max(30, self.width() - self.iconSize().width() - 24)
        shown = self.fontMetrics().elidedText(self.full_text, Qt.TextElideMode.ElideRight, room)
        self.setText(shown.replace("&", "&&"))      # a lone & would underline the next letter

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.elide()

    def sizeHint(self):
        hint = super().sizeHint()
        return QSize(min(hint.width(), 16 * self.fontMetrics().averageCharWidth()),
                     hint.height())

    def minimumSizeHint(self):
        hint = super().minimumSizeHint()
        return QSize(min(hint.width(), 8 * self.fontMetrics().averageCharWidth()),
                     hint.height())


class FitLabel(QLabel):
    """One line of words that ends in ... when it doesn't fit, rather than
    being cut off mid-letter. text() gives all of it."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.full = ""
        self.setTextFormat(Qt.TextFormat.PlainText)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

    def setText(self, text):
        if (text or "") == self.full:
            return
        self.full = text or ""
        self.setToolTip(self.full)
        self.updateGeometry()
        self.fit()

    def text(self):
        return self.full

    def sizeHint(self):
        return QSize(self.fontMetrics().horizontalAdvance(self.full) + 8,
                     super().sizeHint().height())

    def minimumSizeHint(self):
        return QSize(min(self.sizeHint().width(), 6 * self.fontMetrics().averageCharWidth()),
                     super().minimumSizeHint().height())

    def fit(self):
        QLabel.setText(self, self.fontMetrics().elidedText(
            self.full, Qt.TextElideMode.ElideRight, max(0, self.width() - 4)))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit()

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() in (QEvent.Type.FontChange, QEvent.Type.StyleChange):
            self.updateGeometry()
            self.fit()


class Bridge(QObject):
    """Brings news from worker threads onto the window's thread."""
    found = Signal(int, object)                # scan job, list of Track
    scanned = Signal(int, int, object)         # scan job, songs found, web addresses
    radio_event = Signal(object, str, object)  # recorder, kind, value
    spotify_now = Signal(object)
    done = Signal(object, object, object)      # callback, result, error


def in_background(bridge, work, callback):
    """Run work() on a thread; callback(result, error) on the window's thread."""
    def run():
        try:
            result, error = work(), None
        except Exception as problem:
            result, error = None, problem
        bridge.done.emit(callback, result, error)
    threading.Thread(target=run, daemon=True).start()


# ------------------------------------------------------------ the song list

class SongModel(QAbstractTableModel):
    COLUMNS = (("#", "path"), ("Title", "title"), ("Artist", "artist"),
               ("Album", "album"), ("Length", "length"))

    def __init__(self, playlist, parent=None):
        super().__init__(parent)
        self.playlist = playlist
        self.playing = None
        self.broken = set()
        self.play_icon = None
        bold = list_font()               # the list's own size: a plain QFont() would be tiny
        bold.setBold(True)
        self.bold = bold

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.playlist.tracks)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self.COLUMNS[section][0]
        return None

    def track(self, row):
        return self.playlist.tracks[row]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        track = self.playlist.tracks[index.row()]
        column = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            if column == 0:
                return str(index.row() + 1)
            if column == 1:
                return track.name
            if column == 2:
                return track.artist or track.album_artist
            if column == 3:
                return track.album
            if column == 4:
                return format_time(track.length) if track.length else ""
        elif role == Qt.ItemDataRole.FontRole and track is self.playing:
            return self.bold
        elif role == Qt.ItemDataRole.DecorationRole and column == 1 and track is self.playing:
            return self.play_icon
        elif role == Qt.ItemDataRole.ForegroundRole and id(track) in self.broken:
            return QColor(WARNING_RED)            # couldn't be played (red, not faint grey)
        elif role == Qt.ItemDataRole.ToolTipRole:
            note = "\n(This file couldn't be played.)" if id(track) in self.broken else ""
            who = " – ".join(part for part in (track.artist or track.album_artist, track.album)
                             if part)
            return "\n".join(part for part in (track.name, who, track.path) if part) + note
        elif role == Qt.ItemDataRole.TextAlignmentRole and column in (0, 4):
            return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return None

    def reset(self):
        self.beginResetModel()
        self.endResetModel()

    def appended(self, count):
        if count:
            total = len(self.playlist.tracks)
            self.beginInsertRows(QModelIndex(), total - count, total - 1)
            self.endInsertRows()

    def refresh_track(self, track):
        row = self.playlist.find(track)
        if row is not None:
            self.dataChanged.emit(self.index(row, 0), self.index(row, len(self.COLUMNS) - 1))


class SongFilter(QSortFilterProxyModel):
    """Shows only songs with every searched word in the title, artist, album
    or file name."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.words = []

    def set_text(self, text):
        self.words = [w.casefold() for w in text.split()]
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        if not self.words:
            return True
        track = self.sourceModel().track(row)
        folder = os.path.dirname(track.path)
        # The file's own name and the two folders above it count too ("Disc 2",
        # an album folder), for songs without proper tags.
        haystack = " ".join((track.name, track.artist, track.album_artist, track.album,
                             track.genre, os.path.basename(track.path),
                             os.path.basename(folder),
                             os.path.basename(os.path.dirname(folder)))).casefold()
        return all(word in haystack for word in self.words)


class Scanner:
    """Finds the music in picked files and folders, and reads each song's
    details, on a thread of its own; songs arrive in batches as they're read."""

    def __init__(self, bridge, cache):
        self.bridge = bridge
        self.cache = cache
        self.job = 0
        self.stop_flag = threading.Event()

    def start(self, paths):
        self.stop_flag.set()               # a newer job replaces any older one
        self.stop_flag = threading.Event()
        self.job += 1
        job, stop = self.job, self.stop_flag
        threading.Thread(target=self._run, args=(job, list(paths), stop), daemon=True).start()
        return job

    def cancel(self):
        self.stop_flag.set()
        self.job += 1

    def _run(self, job, paths, stop):
        batch, sent, addresses = [], 0, []
        try:
            files, addresses = engine.expand_paths(paths, stop=stop.is_set)
            last = time.monotonic()
            for path in files:
                if stop.is_set():
                    return
                try:
                    track = self.cache.track(path)
                except Exception:              # one odd file mustn't stop the rest
                    track = Track(path)
                batch.append(track)
                if len(batch) >= 400 or time.monotonic() - last > 0.4:
                    self.bridge.found.emit(job, batch)
                    sent += len(batch)
                    batch, last = [], time.monotonic()
        except Exception:
            pass                               # say what was found so far
        finally:
            if not stop.is_set():
                if batch:
                    self.bridge.found.emit(job, batch)
                    sent += len(batch)
                self.bridge.scanned.emit(job, sent, addresses)


# ------------------------------------------------------- radio through speakers

class RadioOutput(QObject):
    """Plays a station's TimeshiftBuffer through the speakers from any point
    kept, while the station goes on arriving. Pausing just stops reading."""

    loop_lost = Signal()               # the part being repeated got older than what's kept

    def __init__(self, parent=None):
        super().__init__(parent)
        self.buffer = None
        self.sink = None
        self.device = None
        self.volume = 1.0
        self.loop = None                 # (start, end) frames to repeat
        self.next_frame = 0
        self.paused_at = 0
        self.timer = QTimer(self)
        self.timer.setInterval(15)
        self.timer.timeout.connect(self.feed)

    @staticmethod
    def best_rate():
        device = QMediaDevices.defaultAudioOutput()
        rate = device.preferredFormat().sampleRate() if not device.isNull() else 0
        return rate if 8000 <= rate <= 192000 else 48000

    @property
    def playing(self):
        return self.sink is not None

    def attach(self, buffer):
        self.stop()
        self.buffer = buffer
        self.loop = None
        self.next_frame = self.paused_at = 0

    def play(self, frame):
        """Start (or move) playing at frame."""
        self.stop()
        if self.buffer is None:
            return False
        device = QMediaDevices.defaultAudioOutput()
        if device.isNull():
            raise RuntimeError("No speakers or headphones were found.")
        fmt = QAudioFormat()
        fmt.setSampleRate(self.buffer.rate)
        fmt.setChannelCount(self.buffer.channels)
        fmt.setSampleFormat(QAudioFormat.SampleFormat.Int16)
        self.as_float = False
        if not device.isFormatSupported(fmt):
            fmt.setSampleFormat(QAudioFormat.SampleFormat.Float)
            self.as_float = True
        self.out_bytes = self.buffer.channels * (4 if self.as_float else 2)
        frame = max(int(frame), self.buffer.oldest)
        self.next_frame = frame
        self.written = 0
        self.segments = [(0, frame)]          # (frames written when, frame then)
        self.sink = QAudioSink(device, fmt, self)
        self.sink.setBufferSize(int(self.buffer.rate * 0.3) * self.out_bytes)
        self.sink.setVolume(self.volume)
        self.output = self.sink.start()
        if self.output is None:
            self.sink.deleteLater()
            self.sink = None
            raise RuntimeError("The speakers or headphones couldn't be opened.")
        self.feed()
        self.timer.start()
        return True

    def feed(self):
        if self.sink is None or self.buffer is None:
            return
        free = self.sink.bytesFree() // self.out_bytes
        while free > 0:
            oldest = self.buffer.oldest
            if self.next_frame < oldest:                 # written over: skip ahead
                self.next_frame = oldest
                self.segments.append((self.written, oldest))
            count = free
            if self.loop and self.loop[1] <= oldest:
                self.loop = None                         # nothing of the part is left
                self.loop_lost.emit()
            if self.loop:
                start, end = self.loop
                if self.next_frame >= end:
                    self.next_frame = max(start, oldest)
                    self.segments.append((self.written, self.next_frame))
                count = min(count, end - self.next_frame)
            data = self.buffer.read(self.next_frame, count)
            if not data:
                break
            frames = len(data) // self.buffer.frame_bytes
            if self.as_float:
                import array
                samples = array.array("h")
                samples.frombytes(data)
                data = array.array("f", (s / 32768 for s in samples)).tobytes()
            sent = self.output.write(data) // self.out_bytes
            if sent <= 0:
                break
            self.written += sent
            self.next_frame += sent
            free -= sent
            if sent < frames:
                break

    def heard(self):
        """The frame coming out of the speakers now."""
        if self.sink is None:
            return self.paused_at
        played = min(self.sink.processedUSecs() * self.buffer.rate / 1e6, self.written)
        while len(self.segments) > 1 and self.segments[1][0] <= played:
            self.segments.pop(0)               # repeating for hours mustn't pile these up
        out_start, source_start = self.segments[0]
        return int(source_start + played - out_start)

    def pause(self):
        self.paused_at = self.heard()
        self.stop()

    def set_volume(self, volume):
        self.volume = volume
        if self.sink:
            self.sink.setVolume(volume)

    def stop(self):
        self.timer.stop()
        if self.sink is not None:
            self.paused_at = self.heard()
            sink, self.sink, self.output = self.sink, None, None
            sink.stop()
            sink.deleteLater()


# ---------------------------------------------------------------- dialogs

class SleepDialog(QDialog):
    def __init__(self, parent, timer, minutes, action):
        super().__init__(parent)
        self.setWindowTitle("Sleep timer")
        self.choice = None
        big = base_points() * POPUP_TEXT
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        if timer.active:
            now = QLabel("Now: " + timer.describe() + ".")
            now.setStyleSheet("font-weight: bold;")
            layout.addWidget(now)

        # Twice the window's word size, like every pop-up box.
        self.setStyleSheet(f"* {{ font-size: {big:g}pt; }}"
                           f" QPushButton {{ padding: 6px 14px; }}"
                           f" QPushButton:checked {{ background: {BLUE}; color: white;"
                           " border-radius: 4px; }")
        tall = int(big * 2.6)
        words = QFont(QApplication.font())
        words.setPointSizeF(big)
        fits = QFontMetrics(words)          # buttons wide enough for their words at this size
        when = QGroupBox("1. Stop the music after")
        grid = QGridLayout(when)
        self.when_group = QButtonGroup(self)
        self.when_buttons = {}
        for i, (value, text) in enumerate(TIMER_CHOICES):
            b = QPushButton(text)
            b.setCheckable(True)
            b.setMinimumHeight(tall)
            b.setMinimumWidth(fits.horizontalAdvance(text) + 44)
            self.when_group.addButton(b)
            self.when_buttons[value] = b
            grid.addWidget(b, i // 5, i % 5)
        song_end = QPushButton("The end of this song")
        song_end.setCheckable(True)
        song_end.setMinimumHeight(tall)
        self.when_group.addButton(song_end)
        self.when_buttons[0] = song_end
        grid.addWidget(song_end, 2, 0, 1, 2)
        other = QHBoxLayout()
        self.other_button = QPushButton("Other:")
        self.other_button.setCheckable(True)
        self.other_button.setMinimumHeight(tall)
        self.when_group.addButton(self.other_button)
        self.other_minutes = MinutesBox()
        # No little arrows: at this size they crowd the number out. Click and type.
        self.other_minutes.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.other_minutes.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.other_minutes.setRange(1, 600)
        self.other_minutes.setValue(20)
        self.other_minutes.setMinimumHeight(tall)
        self.other_minutes.setMinimumWidth(int(big * 5))
        self.other_minutes.setToolTip("Click here and type how many minutes")
        self.other_minutes.picked.connect(lambda: self.other_button.setChecked(True))
        self.other_minutes.valueChanged.connect(lambda _: self.other_button.setChecked(True))
        other.addWidget(self.other_button)
        other.addWidget(self.other_minutes)
        other.addWidget(QLabel("minutes"))
        other.addStretch(1)
        grid.addLayout(other, 2, 2, 1, 3)
        layout.addWidget(when)
        if minutes in self.when_buttons:
            self.when_buttons[minutes].setChecked(True)
        else:
            self.other_minutes.setValue(max(1, minutes))
            self.other_button.setChecked(True)

        then = QGroupBox("2. Then")
        then_layout = QVBoxLayout(then)
        self.action_buttons = {}
        self.hibernates = engine.sleep_means_hibernate()
        for key, text in engine.SLEEP_ACTIONS.items():
            if key == engine.SLEEP_PC and self.hibernates:
                text += " (on this computer: hibernate)"
            b = QRadioButton(text)
            self.action_buttons[key] = b
            then_layout.addWidget(b)
        self.action_buttons.get(action, self.action_buttons[engine.STOP_MUSIC]).setChecked(True)
        note = QLabel(f"The music fades out over the last minute. Before the computer sleeps or "
                      f"shuts down, a {engine.WARNING_SECONDS}-second warning lets you cancel."
                      + (" Hibernate turns the computer off with everything kept; the power "
                         "button brings it back as it was." if self.hibernates else ""))
        note.setWordWrap(True)
        then_layout.addWidget(note)
        layout.addWidget(then)

        # What Start will do, in words, so a wrong number shows before it's started.
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.summary)
        self.when_group.buttonToggled.connect(lambda *_: self.update_summary())
        self.other_minutes.valueChanged.connect(lambda _: self.update_summary())
        for b in self.action_buttons.values():
            b.toggled.connect(lambda _: self.update_summary())

        buttons = QHBoxLayout()
        start = QPushButton("Start the timer")
        start.setStyleSheet(f"QPushButton {{ background: {BLUE}; color: white; font-weight: bold;"
                            " padding: 6px 18px; border-radius: 4px; }")
        start.clicked.connect(self.start)
        off = QPushButton("Turn the timer off")
        off.setEnabled(timer.active)
        off.clicked.connect(self.turn_off)
        close = QPushButton("Close")
        close.clicked.connect(self.reject)
        for b in (start, off, close):
            b.setMinimumHeight(tall)
            buttons.addWidget(b)
        layout.addLayout(buttons)
        self.start_button, self.off_button = start, off
        self.update_summary()

    def update_summary(self):
        minutes = self.minutes()
        when = ("at the end of this song" if minutes == 0
                else f"in {long_time(minutes * 60)}")
        then = {engine.STOP_MUSIC: "",
                engine.CLOSE_PLAYER: ", then the Music Player closes",
                engine.SLEEP_PC: (", then the computer hibernates" if self.hibernates
                                  else ", then the computer goes to sleep"),
                engine.SHUT_DOWN_PC: ", then the computer shuts down"}[self.action()]
        self.summary.setText(f"Start the timer and the music stops {when}{then}.")

    def minutes(self):
        if self.other_button.isChecked():
            return self.other_minutes.value()
        for value, b in self.when_buttons.items():
            if b.isChecked():
                return value
        return 30

    def action(self):
        return next((k for k, b in self.action_buttons.items() if b.isChecked()),
                    engine.STOP_MUSIC)

    def start(self):
        self.choice = ("start", self.minutes(), self.action())
        self.accept()

    def turn_off(self):
        self.choice = ("off", 0, None)
        self.accept()


class WarningDialog(QDialog):
    """The last chance to stop the computer sleeping or shutting down."""

    def __init__(self, parent, action, seconds, act):
        super().__init__(parent)
        self.setWindowTitle(TITLE)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.action, self.left, self.act = action, seconds, act
        hibernates = action == engine.SLEEP_PC and engine.sleep_means_hibernate()
        self.word = ("shut down" if action == engine.SHUT_DOWN_PC
                     else "hibernate (turn off, keeping everything)" if hibernates
                     else "go to sleep")
        big = base_points() * POPUP_TEXT
        self.setStyleSheet(f"* {{ font-size: {big:g}pt; }}")
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        self.label = QLabel()
        self.label.setWordWrap(True)
        self.label.setMinimumWidth(int(big * 30))
        self.label.setStyleSheet(f"font-size: {big * 1.4:g}pt;")
        layout.addWidget(self.label)
        row = QHBoxLayout()
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setStyleSheet(f"QPushButton {{ background: {BLUE}; color: white;"
                                         " font-weight: bold; padding: 10px 40px;"
                                         " border-radius: 4px; }")
        self.cancel_button.setMinimumHeight(int(big * 3))
        self.cancel_button.clicked.connect(self.reject)
        self.now_button = QPushButton("Shut down now" if action == engine.SHUT_DOWN_PC
                                      else "Hibernate now" if hibernates else "Sleep now")
        self.now_button.setMinimumHeight(int(big * 3))
        self.now_button.clicked.connect(self.finish)
        row.addWidget(self.cancel_button)
        row.addWidget(self.now_button)
        layout.addLayout(row)
        self.timer = QTimer(self, interval=1000)
        self.timer.timeout.connect(self.tick)
        self.timer.start()
        self.show_left()

    def show_left(self):
        self.label.setText(f"The computer will {self.word} in {self.left} seconds.")

    def tick(self):
        self.left -= 1
        if self.left <= 0:
            self.finish()
        else:
            self.show_left()

    def finish(self):
        self.timer.stop()
        self.accept()
        self.act()


def help_html():
    """Keys and tips, from the top of this file, as a page: each part with
    its heading, and the keys in a table (their dotted lines only line up in
    a typewriter font)."""
    from html import escape
    text = __doc__.split("Requirements")[0].strip()
    parts = []
    for block in text.split("\n\n"):
        lines = block.splitlines()
        head, body = lines[0].strip(), [line.strip() for line in lines[1:]]
        if head.endswith(".py"):                  # the file's own name
            continue
        if not body:
            parts.append(f"<p>{escape(head)}</p>")
            continue
        parts.append(f"<h3>{escape(head)}</h3>")
        keys, words = [], []
        for line in body:
            pairs = [re.match(r"(.+?)\s*\.{3,}\s*(.+)", p) for p in re.split(r"\s{2,}", line)]
            if all(pairs):
                keys += [(m.group(1), m.group(2)) for m in pairs]
            else:
                words.append(line)
        if keys:
            rows = "".join(f"<tr><td>{escape(n)}</td><td><b>{escape(k)}</b></td></tr>"
                           for n, k in keys)
            parts.append(f"<table cellpadding='4'>{rows}</table>")
        if words:
            parts.append(f"<p>{escape(' '.join(words))}</p>")
    return "".join(parts)


class HelpDialog(QDialog):
    """Keys and tips, in big words, scrolling if it's taller than the screen.
    It can stay open beside the player."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(f"{TITLE}: keys and tips")
        big = base_points() * POPUP_TEXT
        self.setStyleSheet(f"* {{ font-size: {big:g}pt; }}")
        layout = QVBoxLayout(self)
        self.text = QTextBrowser()
        self.text.setHtml(help_html())
        self.text.setOpenLinks(False)
        layout.addWidget(self.text, 1)
        close = QPushButton("Close")
        close.setMinimumHeight(int(big * 2.6))
        close.setMinimumWidth(int(big * 8))
        close.clicked.connect(self.close)
        layout.addWidget(close, 0, Qt.AlignmentFlag.AlignRight)
        screen = parent.screen().availableGeometry() if parent.screen() else None
        if screen is not None:
            self.resize(int(screen.width() * 0.55), int(screen.height() * 0.85))


# ---------------------------------------------------------------- the window

class MusicWindow(QMainWindow):

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.data_folder = os.path.dirname(os.path.abspath(settings.fileName()))
        os.makedirs(self.data_folder, exist_ok=True)
        self.bridge = Bridge()
        self.bridge.found.connect(self.songs_found)
        self.bridge.scanned.connect(self.scan_finished)
        self.bridge.radio_event.connect(self.radio_event)
        self.bridge.spotify_now.connect(self.spotify_news)
        self.bridge.done.connect(lambda callback, result, error: callback(result, error))

        self.cache = engine.TagCache(os.path.join(self.data_folder, "song details.json"))
        self.scanner = Scanner(self.bridge, self.cache)
        self.scan_job = None
        self.scan_play = False           # start playing when the first songs arrive
        self.scan_restore = None         # (path, seconds) to come back to
        self.restore_job = None          # the scan reading back last time's list
        self.scan_paths, self.scan_base = [], 0
        self.scan_waiting = []           # songs added while others are read wait their turn
        self.playlist = PlayList()
        self.recent = RecentFolders()
        self.timer = SleepTimer()
        self.mode = MUSIC
        self.failures = 0
        self.pending_position = None     # (QUrl, ms) to jump to once a song has loaded
        self.local_volume = 80
        self.muted = False
        self.sort_key, self.sort_descending = None, False
        self.system_actions = {engine.SLEEP_PC: engine.put_computer_to_sleep,
                               engine.SHUT_DOWN_PC: engine.shut_down_computer}
        self.warning = None

        self.stations = []
        self.station = None
        self.radio_buffer = None
        self.recorder = None
        self.radio_waiting = False       # gathering sound before playing starts
        self.radio_marks = [None, None]
        self.radio_song = None
        self.radio_message = ""
        self.directory = radio.RadioDirectory()
        self.radio_output = RadioOutput(self)
        self.radio_output.loop_lost.connect(self.radio_loop_lost)
        self.countries_loaded = False
        self.radio_searched = False

        self.spotify_links = []
        self.remote = spotify.SpotifyRemote()
        self.spotify_volume = spotify.SpotifyVolume()
        self.spotify_now = None
        self.spotify_asking = False
        self.spotify_base_volume = None  # Spotify's own volume before a sleep fade
        self.spotify_running = spotify.is_running
        self.spotify_offered = False     # the "open Spotify?" box, once a run

        self.player = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.player.setAudioOutput(self.audio)
        if hasattr(self.player, "setPitchCompensation"):    # Qt 6.10+: speed without chipmunks
            self.player.setPitchCompensation(True)
        self.player.positionChanged.connect(self.position_changed)
        self.player.durationChanged.connect(lambda _: self.update_position())
        self.player.playbackStateChanged.connect(lambda _: self.update_buttons())
        self.player.mediaStatusChanged.connect(self.media_status_changed)
        self.player.errorOccurred.connect(self.player_error)
        self.network = QNetworkAccessManager(self)
        self.media_devices = QMediaDevices(self)
        self.speaker_id = QMediaDevices.defaultAudioOutput().id()
        self.media_devices.audioOutputsChanged.connect(self.speakers_changed)

        # Words twice the usual size all through this window, and black.
        big = QFont(QApplication.font())
        big.setPointSizeF(words())
        self.setFont(big)
        QToolTip.setFont(big)
        black_words(QApplication.instance())
        self.build_window()
        self.build_menus()
        self.build_shortcuts()
        self.skip_timer = QTimer(self, singleShot=True, interval=SKIP_BROKEN_AFTER_MS)
        self.skip_timer.timeout.connect(self.skip_broken)
        self.ticker = QTimer(self, interval=250)
        self.ticker.timeout.connect(self.tick)
        self.ticker.start()
        self.spotify_timer = QTimer(self, interval=1000)
        self.spotify_timer.timeout.connect(self.ask_spotify)
        self.spotify_timer.start()
        self.restore_settings()
        self.show_now_playing()
        self.update_buttons()

    # ================================================================ layout
    def icon(self, picture):
        """A Windows 11 theme picture (none if missing) or one of the style's own."""
        if isinstance(picture, QIcon.ThemeIcon):
            return QIcon.fromTheme(picture)
        return self.style().standardIcon(picture)

    def button(self, text, icon=None, tip="", slot=None, blue=False, big=False):
        b = QToolButton()
        b.setText(text)
        if icon is not None:
            b.setIcon(self.icon(icon))
            b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        else:
            b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        b.setToolTip(tip)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)        # leave Space and the arrows to the keys
        b.setIconSize(QSize(28, 28))
        if big:
            b.setMinimumHeight(52)
        if blue:
            b.setStyleSheet(f"QToolButton {{ background: {BLUE}; color: white; padding: 5px 12px;"
                            " border-radius: 4px; font-weight: bold; }")
        if slot:
            b.clicked.connect(slot)
        return b

    def build_window(self):
        self.setWindowTitle(TITLE)
        # Set before anything is made, so every part takes it. The word size goes in the
        # style sheet: once a window has one, Qt stops handing setFont() down to its parts.
        screen = self.screen().availableGeometry().width() if self.screen() else 1600
        self.setStyleSheet(f"QWidget, QMenuBar, QMenu {{ font-size: {words():g}pt; }} "
                           f"QMenuBar::item, QMenu::item {{ font-size: {words():g}pt; }} "
                           # Drawn here, not by Windows 11's own menu bar, which keeps its words small.
                           "QMenuBar::item { padding: 6px 16px; background: transparent; color: black; }"
                           " QMenuBar::item:selected { background: #dce6f7; }"
                           " QMenu::item { padding: 8px 30px; } "
                           + popup_style(max(420, screen // 2 - 140)))
        self.setWindowIcon(self.icon(QStyle.StandardPixmap.SP_MediaVolume))
        self.setAcceptDrops(True)
        SP = QStyle.StandardPixmap
        TI = QIcon.ThemeIcon

        # --- now playing
        self.cover = QLabel()
        # The song's picture: as tall as the two lines of words beside it.
        title_font, info_font = QFont(self.font()), QFont(self.font())
        title_font.setPointSizeF(words(1.3))
        info_font.setPointSizeF(words())
        self.cover_size = (QFontMetrics(title_font).height() + QFontMetrics(info_font).height())
        self.cover.setFixedSize(self.cover_size, self.cover_size)
        self.cover.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Two lines across the top, so the list below gets the height:
        #   the song (bold)   its artist
        #   where it's from   album · year · song 3 of 30
        self.title_label = FitLabel()
        self.title_label.setStyleSheet(f"font-size: {words(1.3):g}pt; font-weight: bold;")
        self.artist_label = FitLabel()
        self.artist_label.setStyleSheet(f"font-size: {words(1.15):g}pt;")
        self.info_label = FitLabel()
        self.source_label = QLabel()
        self.source_label.setStyleSheet("font-weight: bold;")
        gap = int(words() * 1.5)
        first_line = QHBoxLayout()
        first_line.setSpacing(gap)
        # The song's name gets its room first; the artist (and station) and the
        # details take what is left, ending in ... if they must.
        for label in (self.artist_label, self.info_label):
            label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        first_line.addWidget(self.title_label)
        first_line.addWidget(self.artist_label, 1)
        second_line = QHBoxLayout()
        second_line.setSpacing(gap)
        second_line.addWidget(self.source_label)
        second_line.addWidget(self.info_label, 1)
        lines = QVBoxLayout()
        lines.setSpacing(0)
        lines.addLayout(first_line)
        lines.addLayout(second_line)
        top = QHBoxLayout()
        top.addWidget(self.cover)
        top.addSpacing(12)
        top.addLayout(lines, 1)

        # --- seek bar
        self.seek = SeekSlider(Qt.Orientation.Horizontal)
        self.seek.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.seek.setToolTip("Click or drag to move through the song")
        self.seek.sliderMoved.connect(self.seek_to)
        self.time_label = QLabel("0:00")
        self.length_label = QLabel("0:00")
        for label in (self.time_label, self.length_label):
            label.setMinimumWidth(self.fontMetrics().horizontalAdvance("-00:00:00 behind live") // 2)
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        seek_row = QHBoxLayout()
        seek_row.addWidget(self.time_label)
        seek_row.addWidget(self.seek, 1)
        seek_row.addWidget(self.length_label)

        # --- buttons
        self.prev_button = self.button("Previous", SP.SP_MediaSkipBackward, "Previous song (P)",
                                       self.previous, big=True)
        self.back_button = self.button(f"Back {SKIP_SECONDS} s", SP.SP_MediaSeekBackward,
                                       f"Back {SKIP_SECONDS} seconds (Left arrow)",
                                       lambda: self.skip(-SKIP_SECONDS), big=True)
        self.play_button = self.button("Play", SP.SP_MediaPlay, "Play or pause (Space)",
                                       self.toggle_play, big=True)
        self.play_button.setMinimumWidth(self.fontMetrics().horizontalAdvance("Pause") + 60)
        self.stop_button = self.button("Stop", SP.SP_MediaStop, "Stop (Ctrl+.)", self.stop, big=True)
        self.forward_button = self.button(f"Forward {SKIP_SECONDS} s", SP.SP_MediaSeekForward,
                                          f"Forward {SKIP_SECONDS} seconds (Right arrow)",
                                          lambda: self.skip(SKIP_SECONDS), big=True)
        self.next_button = self.button("Next", SP.SP_MediaSkipForward, "Next song (N)",
                                       self.next, big=True)
        self.shuffle_button = self.button("Shuffle: Off", TI.MediaPlaylistShuffle,
                                          "Play the list in a random order", self.toggle_shuffle,
                                          big=True)
        self.shuffle_button.setCheckable(True)
        self.repeat_button = self.button("Repeat: Off", TI.MediaPlaylistRepeat,
                                         "Repeat off, the whole list, or one song",
                                         self.cycle_repeat, big=True)
        self.speed_box = QComboBox()
        self.speed_box.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.speed_box.setToolTip("How fast the music plays (your own music only)")
        for speed in SPEEDS:
            self.speed_box.addItem(f"Speed {speed:g}x", speed)
        self.speed_box.setCurrentIndex(SPEEDS.index(1.0))
        self.speed_box.currentIndexChanged.connect(self.speed_changed)
        self.mute_button = self.button("Mute", SP.SP_MediaVolume, "Mute or unmute (M)",
                                       self.toggle_mute, big=True)
        self.volume = QSlider(Qt.Orientation.Horizontal)
        self.volume.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.volume.setRange(0, 100)
        self.volume.setMinimumWidth(140)
        self.volume.setMaximumWidth(260)
        self.volume.setToolTip("Volume (+ and -)")
        self.volume.valueChanged.connect(self.volume_changed)
        self.volume_label = QLabel("80%")
        self.volume_label.setMinimumWidth(self.fontMetrics().horizontalAdvance("100%") + 4)
        self.sleep_button = self.button("Sleep timer", None,
                                        "Stop the music (or more) after a while", self.sleep_dialog,
                                        big=True)
        self.sleep_button.setIcon(glyph_icon("☾"))         # a crescent moon

        # One row of tall buttons, the picture above the word, easy to hit: the playing
        # buttons, then Shuffle, Repeat and Speed (up here, so the list
        # below gets the height a second row took), then the volume. The Sleep timer sits
        # at the right of the song's details above, where there's room.
        self.transport = QWidget()
        tall = TRANSPORT_HEIGHT
        for b in (self.prev_button, self.back_button, self.play_button, self.stop_button,
                  self.forward_button, self.next_button, self.shuffle_button,
                  self.repeat_button, self.mute_button):
            b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            b.setIconSize(QSize(36, 36))
            b.setMinimumHeight(tall)
            b.setMinimumWidth(100)
        fits = self.fontMetrics().horizontalAdvance
        heavy = QFont(self.font())
        heavy.setBold(True)
        fits_bold = QFontMetrics(heavy).horizontalAdvance
        # Room for the whole of every word each button can show, so a narrow window never
        # cuts them to "Pr...us" (it scrolls, or opens full screen, instead), and a button
        # doesn't change width as its words change (Play/Pause, Repeat's three settings).
        for b, texts, bold in (
                (self.prev_button, ["Previous"], False),
                (self.back_button, [f"Back {SKIP_SECONDS} s"], False),
                (self.play_button, ["Play", "Pause"], False),
                (self.stop_button, ["Stop"], False),
                (self.forward_button, [f"Forward {SKIP_SECONDS} s"], False),
                (self.next_button, ["Next"], False),
                (self.shuffle_button, ["Shuffle: Off", "Shuffle: On"], False),
                (self.repeat_button, ["Repeat: Off", "Repeat: All", "Repeat: One"], True),
                (self.mute_button, ["Mute", "Unmute"], True)):
            measure = fits_bold if bold else fits
            b.setMinimumWidth(max(100, max(measure(text) for text in texts) + 30))
        self.volume.setMinimumWidth(110)
        self.speed_box.setMinimumHeight(SECOND_ROW_HEIGHT - 4)
        row = QHBoxLayout(self.transport)
        row.setContentsMargins(0, 2, 0, 2)
        row.setSpacing(6)
        for b in (self.prev_button, self.back_button, self.play_button, self.stop_button,
                  self.forward_button, self.next_button):
            row.addWidget(b)
        row.addSpacing(14)
        for b in (self.shuffle_button, self.repeat_button, self.speed_box):
            row.addWidget(b)
        row.addStretch(1)
        row.addWidget(QLabel("Volume"))
        row.addWidget(self.volume)
        row.addWidget(self.volume_label)
        row.addWidget(self.mute_button)
        self.sleep_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.sleep_button.setIconSize(QSize(32, 32))
        self.sleep_button.setMinimumHeight(min(self.cover_size, SECOND_ROW_HEIGHT + 8))
        self.sleep_button.setMinimumWidth(fits_bold("Sleep: 0:00:00") + 70)   # steady, even bold
        top.addWidget(self.sleep_button, 0, Qt.AlignmentFlag.AlignVCenter)

        # --- radio's own buttons, shown while a station plays
        self.back30_button = self.button(f"Back {LONG_SKIP_SECONDS} s", SP.SP_MediaSeekBackward,
                                         f"Back {LONG_SKIP_SECONDS} seconds",
                                         lambda: self.skip(-LONG_SKIP_SECONDS), big=True)
        self.live_button = self.button("Go live", TI.MediaPlaybackStart,
                                       "Catch up with the station (L)", self.go_live, big=True)
        self.song_start_button = self.button("Start of this song", SP.SP_MediaSkipBackward,
                                             "Back to where the song playing began",
                                             self.radio_song_start, big=True)
        self.mark_start_button = self.button("1. Mark start", None,
                                             "Mark where the part to repeat begins",
                                             self.radio_mark_start, big=True)
        self.mark_end_button = self.button("2. Mark end", None,
                                           "Mark where the part to repeat ends",
                                           self.radio_mark_end, big=True)
        self.repeat_part_button = self.button("3. Repeat marked part", TI.MediaPlaylistRepeat,
                                              "Play the marked part over and over",
                                              self.radio_repeat_part, big=True)
        self.repeat_part_button.setCheckable(True)
        self.clear_marks_button = self.button("Clear marks", None, "Forget the marks",
                                              self.radio_clear_marks, big=True)
        self.marks_label = QLabel()
        self.marks_label.setWordWrap(True)
        self.radio_bar = QGroupBox("Radio: the last hour is kept, so you can pause, go back "
                                   "and repeat")
        for b in (self.back30_button, self.song_start_button, self.live_button,
                  self.mark_start_button, self.mark_end_button, self.repeat_part_button,
                  self.clear_marks_button):
            b.setMinimumHeight(SECOND_ROW_HEIGHT)
            b.setIconSize(QSize(30, 30))
        radio_rows = QVBoxLayout(self.radio_bar)            # two rows, so the big words fit
        radio_rows.setSpacing(6)
        going = QHBoxLayout()
        for b in (self.back30_button, self.song_start_button, self.live_button):
            going.addWidget(b)
        going.addStretch(1)
        marking = QHBoxLayout()
        for b in (self.mark_start_button, self.mark_end_button, self.repeat_part_button,
                  self.clear_marks_button):
            marking.addWidget(b)
        marking.addWidget(self.marks_label, 1)
        radio_rows.addLayout(going)
        radio_rows.addLayout(marking)
        self.radio_bar.hide()

        # --- the three pages
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"QTabBar::tab {{ font-size: {words(1.15):g}pt;"
                                " padding: 6px 34px; }")
        self.tabs.addTab(self.build_music_page(), "My Music")
        self.tabs.addTab(self.build_radio_page(), "Radio")
        self.tabs.addTab(self.build_spotify_page(), "Spotify")
        self.tabs.currentChanged.connect(self.tab_changed)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(6, 2, 6, 0)
        layout.setSpacing(4)
        layout.addLayout(top)
        layout.addLayout(seek_row)
        layout.addWidget(self.transport)
        layout.addWidget(self.radio_bar)
        layout.addWidget(self.tabs, 1)
        # On a small screen (or big text, or 150% scaling) the whole window can
        # scroll rather than hang off the edge of the screen.
        self.page_scroll = QScrollArea()
        self.page_scroll.setWidget(central)
        self.page_scroll.setWidgetResizable(True)
        self.page_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.setCentralWidget(self.page_scroll)
        self.statusBar()                 # made now, so it takes the big words with the rest
        self.fit_side_panels()

    def build_music_page(self):
        SP = QStyle.StandardPixmap
        page = QWidget()
        side = QVBoxLayout()
        self.open_folder_button = self.button("Open a folder...", SP.SP_DirOpenIcon,
                                              "Play every song in a folder and its subfolders "
                                              "(Ctrl+O)", self.choose_folder, blue=True, big=True)
        self.add_songs_button = self.button("Add songs...", SP.SP_FileIcon,
                                            "Add songs to the end of the list (Ctrl+Shift+O)",
                                            self.choose_songs, big=True)
        self.add_folder_button = self.button("Add a folder...", SP.SP_DirIcon,
                                             "Add a folder's songs, subfolders too, to the end "
                                             "of the list", lambda: self.choose_folder(add=True),
                                             big=True)
        for b in (self.open_folder_button, self.add_songs_button, self.add_folder_button):
            b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        side.addWidget(self.open_folder_button)
        adds = QHBoxLayout()                    # side by side, to keep the panel short
        adds.addWidget(self.add_songs_button)
        adds.addWidget(self.add_folder_button)
        side.addLayout(adds)

        recent_box = QGroupBox("Your last folders")
        self.recent_layout = QVBoxLayout(recent_box)
        self.recent_rows = []
        for place in range(engine.RECENT_FOLDER_COUNT):
            row = QWidget()
            line = QHBoxLayout(row)
            line.setContentsMargins(0, 0, 0, 0)
            open_button = ElidedButton()
            open_button.setIcon(self.icon(SP.SP_DirIcon))
            open_button.setIconSize(QSize(28, 28))
            open_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
            open_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            open_button.setMinimumHeight(52)
            open_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            remove = self.button("Remove", None, "Take this folder off the list", big=True)
            open_button.clicked.connect(lambda _=False, p=place: self.open_recent(p))
            remove.clicked.connect(lambda _=False, p=place: self.remove_recent(p))
            line.addWidget(open_button, 1)
            line.addWidget(remove)
            self.recent_layout.addWidget(row)
            self.recent_rows.append((row, open_button, remove))
        self.no_recent = QLabel("Folders you open show here, so you can open them again "
                                "with one click.")
        self.no_recent.setWordWrap(True)
        self.recent_layout.addWidget(self.no_recent)
        side.addWidget(recent_box)

        lists_box = QGroupBox("Playlists")
        lists = QGridLayout(lists_box)
        self.open_list_button = self.button("Open a playlist...", SP.SP_DialogOpenButton,
                                            "Play an .m3u, .m3u8 or .pls playlist",
                                            self.choose_playlist, big=True)
        self.save_list_button = self.button("Save this list as a playlist...",
                                            SP.SP_DialogSaveButton,
                                            "Save the list so you can open it again (Ctrl+S)",
                                            self.save_playlist, big=True)
        self.clear_list_button = self.button("Clear the list", SP.SP_DialogResetButton,
                                             "Take every song off the list (the files stay)",
                                             self.clear_list, big=True)
        for b in (self.open_list_button, self.save_list_button, self.clear_list_button):
            b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        lists.addWidget(self.open_list_button, 0, 0)        # one to a row: whole words
        lists.addWidget(self.save_list_button, 1, 0)
        lists.addWidget(self.clear_list_button, 2, 0)
        side.addWidget(lists_box)
        side.addStretch(1)

        # Up next sits under the song list, and shows only while songs are waiting in it.
        self.queue_box = QGroupBox("Up next: these play before the list carries on")
        queue = QHBoxLayout(self.queue_box)
        self.queue_list = QListWidget()
        self.queue_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.queue_list.setToolTip("Songs that play next, before the list carries on")
        self.queue_list.setFixedHeight(self.fontMetrics().height() * 3 + 16)
        queue.addWidget(self.queue_list, 1)
        queue_buttons = QVBoxLayout()
        self.unqueue_button = self.button("Take off", None, "Take the selected song off Up next",
                                          self.unqueue)
        self.clear_queue_button = self.button("Clear", None, "Empty Up next", self.clear_queue)
        for b in (self.unqueue_button, self.clear_queue_button):
            b.setMinimumHeight(44)
            b.setMinimumWidth(120)
            queue_buttons.addWidget(b)
        queue.addLayout(queue_buttons)
        self.queue_box.hide()

        side_inside = QWidget()
        side_inside.setLayout(side)
        # As wide as its biggest button needs, now the words are big.
        side_widget = QScrollArea()             # scrolls if the window is very short
        side_widget.setWidget(side_inside)
        side_widget.setWidgetResizable(True)
        side_widget.setFrameShape(QFrame.Shape.NoFrame)
        side_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.side_inside, self.side_widget = side_inside, side_widget

        # the list itself
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search this list: song, artist or album")
        self.search.setClearButtonEnabled(True)
        self.search.setMinimumHeight(48)
        self.search.textChanged.connect(self.search_changed)
        self.model = SongModel(self.playlist, self)
        self.model.play_icon = self.icon(QStyle.StandardPixmap.SP_MediaPlay)
        self.filter = SongFilter(self)
        self.filter.setSourceModel(self.model)
        self.table = QTableView()
        self.table.setModel(self.filter)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().hide()
        big = QFontMetrics(list_font())
        # Title, Artist and Album share the width; # and Length keep just enough room.
        self.song_columns = make_big_list(
            self.table, {0: big.horizontalAdvance("00000") + 16,
                         4: big.horizontalAdvance("Length") + 28},
            {1: 40, 2: 30, 3: 30})
        header = self.table.horizontalHeader()
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self.sort_by)
        header.setToolTip("Click a heading to sort by it; click # to go back to folder order")
        self.table.doubleClicked.connect(lambda index: self.play_row(self.filter.mapToSource(index).row()))
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.song_menu)
        self.table.setToolTip("Double-click a song to play it")

        self.list_label = QLabel("No songs yet.")
        self.list_label.setWordWrap(True)            # two lines rather than push buttons off
        self.list_label.setMinimumWidth(120)
        self.play_selected_button = self.button("Play", QStyle.StandardPixmap.SP_MediaPlay,
                                                "Play the selected song (Enter)",
                                                self.play_selected)
        self.play_next_button = self.button("Play next", None,
                                            "Play the selected songs after this one",
                                            lambda: self.queue_selected(first=True))
        self.queue_button = self.button("Add to Up next", None,
                                        "Add the selected songs to Up next",
                                        lambda: self.queue_selected(first=False))
        self.remove_button = self.button("Remove from list", None,
                                         "Take the selected songs off the list (Delete); "
                                         "the files stay", self.remove_selected)
        self.show_file_button = self.button("Show in folder", None,
                                            "Open the selected song's folder", self.show_in_folder)
        under = QHBoxLayout()
        under.addWidget(self.list_label, 1)
        for b in (self.play_selected_button, self.play_next_button, self.queue_button,
                  self.remove_button, self.show_file_button):
            b.setMinimumHeight(52)
            under.addWidget(b)

        self.empty_hint = QPushButton("Open a folder of music...")
        self.empty_hint.setStyleSheet(f"QPushButton {{ background: {BLUE}; color: white;"
                                      f" font-size: {words(1.5):g}pt; padding: 14px 30px;"
                                      " border-radius: 8px; }")
        self.empty_hint.clicked.connect(self.choose_folder)
        self.empty_note = QLabel("or drag songs or folders onto this window")
        self.empty_note.setStyleSheet(f"font-size: {words(1.1):g}pt;")
        self.empty_note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_page = QWidget()            # shown instead of the list while it's empty
        empty = QVBoxLayout(self.empty_page)
        empty.addStretch(2)
        empty.addWidget(self.empty_hint, 0, Qt.AlignmentFlag.AlignCenter)
        empty.addSpacing(8)
        empty.addWidget(self.empty_note)
        empty.addStretch(3)

        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(4)
        right.addWidget(self.search)
        right.addWidget(self.table, 1)
        right.addWidget(self.empty_page, 1)
        right.addWidget(self.queue_box)
        right.addLayout(under)
        right_widget = QWidget()
        right_widget.setLayout(right)

        self.music_splitter = QSplitter()
        self.music_splitter.addWidget(side_widget)
        self.music_splitter.addWidget(right_widget)
        self.music_splitter.setStretchFactor(1, 1)
        layout = QHBoxLayout(page)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.addWidget(self.music_splitter)
        return page

    def build_radio_page(self):
        SP = QStyle.StandardPixmap
        page = QWidget()
        self.station_search = QLineEdit()
        self.station_search.setPlaceholderText("A station's name or a type of music, e.g. "
                                               "country, jazz, news, KOA")
        self.station_search.setMinimumHeight(48)
        self.station_search.returnPressed.connect(self.search_stations)
        self.country_box = QComboBox()
        self.country_box.setMinimumHeight(48)
        self.country_box.setMinimumContentsLength(13)
        self.country_box.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.country_box.setToolTip("Only stations from this country")
        self.country_box.addItem("All countries", "")
        code = QLocale.system().name().split("_")[-1].upper()
        if len(code) == 2:
            self.country_box.addItem(QLocale.territoryToString(QLocale.system().territory()), code)
            self.country_box.setCurrentIndex(1)
        search_button = self.button("Search", SP.SP_FileDialogContentsView,
                                    "Find stations", self.search_stations, blue=True, big=True)
        popular_button = self.button("Most listened", None,
                                     "The stations most people play, in the country picked",
                                     self.popular_stations, big=True)
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Find a station:"))
        search_row.addWidget(self.station_search, 1)
        search_row.addWidget(QLabel("in"))
        search_row.addWidget(self.country_box)
        search_row.addWidget(search_button)
        search_row.addWidget(popular_button)

        self.results = QTableWidget(0, 4)
        self.results.setHorizontalHeaderLabels(["Station", "Type of music", "Country", "Sound"])
        self.results.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.results.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.results.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.results.verticalHeader().hide()
        self.results.setAlternatingRowColors(True)
        self.station_columns = make_big_list(self.results, {}, {0: 36, 1: 30, 2: 18, 3: 16})
        self.results.cellDoubleClicked.connect(lambda row, _: self.play_result(row))
        self.results.setToolTip("Double-click a station to play it")
        self.results_label = QLabel("Type a name or a kind of music and press Search, or press "
                                    "Most listened.")
        self.results_label.setWordWrap(True)
        self.results_label.setMinimumWidth(120)
        play_result = self.button("Play this station", SP.SP_MediaPlay,
                                  "Play the station picked above", self.play_result, big=True)
        keep_result = self.button("Add to My stations", None,
                                  "Keep the station picked above in My stations",
                                  self.keep_result, big=True)
        result_buttons = QHBoxLayout()
        result_buttons.addWidget(self.results_label, 1)
        result_buttons.addWidget(play_result)
        result_buttons.addWidget(keep_result)
        left = QVBoxLayout()
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(4)
        left.addLayout(search_row)
        left.addWidget(self.results, 1)
        left.addLayout(result_buttons)
        left_widget = QWidget()
        left_widget.setLayout(left)

        mine_box = QGroupBox("My stations")
        mine = QVBoxLayout(mine_box)
        self.my_stations = QListWidget()
        self.my_stations.setToolTip("Double-click a station to play it")
        self.my_stations.itemDoubleClicked.connect(lambda item: self.play_my_station())
        mine.addWidget(self.my_stations, 1)
        mine_buttons = QGridLayout()
        b = self.button("Play", SP.SP_MediaPlay, "Play the station picked", self.play_my_station,
                        big=True)
        mine_buttons.addWidget(b, 0, 0)
        b = self.button("Remove", None, "Take the station picked off My stations",
                        self.remove_my_station, big=True)
        mine_buttons.addWidget(b, 0, 1)
        b = self.button("Add by web address...", None,
                        "Add a station from its stream address (from the station's own web site)",
                        self.add_station_address, big=True)
        mine_buttons.addWidget(b, 1, 0, 1, 2)
        for i in range(mine_buttons.count()):
            mine_buttons.itemAt(i).widget().setSizePolicy(QSizePolicy.Policy.Expanding,
                                                          QSizePolicy.Policy.Fixed)
        mine.addLayout(mine_buttons)

        heard_box = QGroupBox("Songs heard on this station")
        heard = QVBoxLayout(heard_box)
        self.heard_list = QListWidget()
        self.heard_list.setToolTip("Double-click a song to go back to it, while it's still kept")
        self.heard_list.itemDoubleClicked.connect(self.play_heard_song)
        heard.addWidget(self.heard_list, 1)
        b = self.button("Go back to this song", SP.SP_MediaSkipBackward,
                        "Play the song picked above again from its start", self.play_heard_song,
                        big=True)
        b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        heard.addWidget(b)

        right = QVBoxLayout()
        right.addWidget(mine_box, 3)
        right.addWidget(heard_box, 2)
        right_widget = QWidget()
        right_widget.setLayout(right)
        self.radio_side = right_widget

        splitter = QSplitter()
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        self.radio_splitter = splitter
        layout = QHBoxLayout(page)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.addWidget(splitter)
        return page

    def fit_side_panels(self):
        """The folder panel and My stations: as wide as their biggest button
        needs, measured once the window's big words have been put on them."""
        self.ensurePolished()
        for part in (self.side_inside, self.radio_side):
            part.ensurePolished()
            for child in part.findChildren(QWidget):
                child.ensurePolished()
        wide = max(270, self.side_inside.sizeHint().width())
        self.side_inside.setMinimumWidth(wide)
        self.side_widget.setMinimumWidth(wide + self.side_widget.verticalScrollBar().sizeHint().width())
        self.music_splitter.setSizes([self.side_widget.minimumWidth(), 1400])
        self.radio_side.setMinimumWidth(max(260, self.radio_side.sizeHint().width()))
        self.radio_splitter.setSizes([1300, self.radio_side.minimumWidth() + 40])

    def build_spotify_page(self):
        SP = QStyle.StandardPixmap
        page = QWidget()
        how = QLabel(
            "Spotify doesn't let other programs play its music, so the Music Player works the "
            "Spotify app for you. To add a playlist once: in Spotify, right-click the "
            "playlist's name in Your Library on the left, then Share, then Copy link to "
            "playlist. (The three dots at the very top corner of Spotify's window are a "
            "different menu.) Then press Add a Spotify link below.\n\nPlay in Spotify opens "
            "the playlist in the Spotify app. If it doesn't start by itself, press the green "
            "Play button there once. No link? Press Open the Spotify app, start any playlist "
            "there, and come back. Either way, the buttons at the top of this window (Play, "
            "Pause, Next, Previous, volume and Mute) then work Spotify, and the song playing "
            "shows above.")
        how.setWordWrap(True)
        self.spotify_list = QListWidget()
        self.spotify_list.setStyleSheet(f"QListWidget {{ font-size: {words(1.2):g}pt; }}")
        self.spotify_list.setToolTip("Double-click to play it in Spotify")
        self.spotify_list.itemDoubleClicked.connect(lambda item: self.play_spotify())
        buttons = QHBoxLayout()
        add = self.button("Add a Spotify link...", None, "Paste a playlist (or album) link from "
                          "Spotify", self.add_spotify_link, big=True)
        play = self.button("Play in Spotify", SP.SP_MediaPlay, "Open the picked playlist in the "
                           "Spotify app", self.play_spotify, big=True)
        rename = self.button("Rename...", None, "Change the name shown here",
                             self.rename_spotify, big=True)
        remove = self.button("Remove", None, "Take the picked playlist off this list",
                             self.remove_spotify, big=True)
        open_app = self.button("Open the Spotify app", SP.SP_MediaPlay,
                               "Start Spotify (or bring it to the front)", self.open_spotify,
                               blue=True, big=True)
        open_app.setMinimumHeight(SECOND_ROW_HEIGHT)
        open_app.setStyleSheet(open_app.styleSheet().replace(
            "font-weight: bold;", f"font-weight: bold; font-size: {words(1.2):g}pt;"))
        self.open_spotify_button = open_app
        buttons.addWidget(open_app)
        buttons.addSpacing(16)
        for b in (add, play, rename, remove):
            buttons.addWidget(b)
        buttons.addStretch(1)
        self.spotify_status = QLabel("")
        self.spotify_status.setWordWrap(True)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 8, 10, 6)
        layout.addWidget(how)
        layout.addLayout(buttons)
        layout.addWidget(self.spotify_list, 1)
        layout.addWidget(self.spotify_status)
        return page

    def build_menus(self):
        menu = self.menuBar().addMenu("&File")
        for text, key, slot in (
                ("&Open a folder...", "Ctrl+O", self.choose_folder),
                ("Add &songs...", "Ctrl+Shift+O", self.choose_songs),
                ("Add a &folder...", None, lambda: self.choose_folder(add=True)),
                (None, None, None),
                ("Open a &playlist...", None, self.choose_playlist),
                ("Sa&ve this list as a playlist...", "Ctrl+S", self.save_playlist),
                ("&Clear the list", None, self.clear_list),
                (None, None, None),
                ("E&xit", "Ctrl+Q", self.close)):
            if text is None:
                menu.addSeparator()
                continue
            action = QAction(text, self)
            if key:
                action.setShortcut(QKeySequence(key))
            action.triggered.connect(slot)
            menu.addAction(action)
        menu = self.menuBar().addMenu("&Play")
        for text, slot in (("Play / &Pause\tSpace", self.toggle_play), ("&Stop\tCtrl+.", self.stop),
                           ("&Next\tN", self.next), ("P&revious\tP", self.previous),
                           (None, None), ("S&huffle", self.toggle_shuffle),
                           ("Re&peat", self.cycle_repeat), (None, None),
                           ("Sleep &timer...", self.sleep_dialog)):
            if text is None:
                menu.addSeparator()
            else:
                menu.addAction(text, slot)
        menu = self.menuBar().addMenu("&Help")
        menu.addAction("&Keys and tips", self.show_help)

    def build_shortcuts(self):
        keys = {
            "Space": self.toggle_play, "K": self.toggle_play, "Ctrl+.": self.stop,
            "Left": lambda: self.skip(-SKIP_SECONDS), "Right": lambda: self.skip(SKIP_SECONDS),
            "N": self.next, "P": self.previous, "L": self.go_live,
            "+": lambda: self.volume.setValue(self.volume.value() + 5),
            "=": lambda: self.volume.setValue(self.volume.value() + 5),
            "-": lambda: self.volume.setValue(self.volume.value() - 5),
            "Ctrl+Up": lambda: self.volume.setValue(self.volume.value() + 5),
            "Ctrl+Down": lambda: self.volume.setValue(self.volume.value() - 5),
            "M": self.toggle_mute, "Ctrl+F": self.find_in_list,
        }
        for key, slot in keys.items():
            QShortcut(QKeySequence(key), self, slot)
        for key, slot in ((Qt.Key.Key_MediaTogglePlayPause, self.media_key_play),
                          (Qt.Key.Key_MediaPlay, self.media_key_play),
                          (Qt.Key.Key_MediaPause, self.media_key_play),
                          (Qt.Key.Key_MediaStop, lambda: self.mode != SPOTIFY and self.stop()),
                          (Qt.Key.Key_MediaNext, lambda: self.mode != SPOTIFY and self.next()),
                          (Qt.Key.Key_MediaPrevious, lambda: self.mode != SPOTIFY and self.previous())):
            QShortcut(QKeySequence(key), self, slot)
        # These keys belong to one list each, and only while that list has the focus
        # (PySide ignores a context= argument, so it is set afterwards).
        for widget, key, slot in ((self.table, "Return", self.play_selected),
                                  (self.table, "Enter", self.play_selected),
                                  (self.table, "Delete", self.remove_selected),
                                  (self.my_stations, "Delete", self.remove_my_station),
                                  (self.my_stations, "Return", self.play_my_station),
                                  (self.spotify_list, "Delete", self.remove_spotify),
                                  (self.spotify_list, "Return", self.play_spotify),
                                  (self.results, "Return", self.play_result),
                                  (self.results, "Enter", self.play_result)):
            shortcut = QShortcut(QKeySequence(key), widget, slot)
            shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)

    def media_key_play(self):
        # Windows already hands media keys to Spotify itself; doing it here too would undo it.
        if self.mode != SPOTIFY:
            self.toggle_play()

    def message(self, text, ms=8000):
        self.statusBar().showMessage(text, ms)

    def show_help(self):
        if getattr(self, "help_dialog", None) is None:
            self.help_dialog = HelpDialog(self)
        self.help_dialog.show()
        self.help_dialog.raise_()
        self.help_dialog.activateWindow()

    # ============================================================= settings
    def restore_settings(self):
        s = self.settings
        self.local_volume = max(0, min(100, self.int_setting("volume", 80)))
        self.muted = s.value("muted", "false") == "true"
        self.volume.blockSignals(True)
        self.volume.setValue(self.local_volume)
        self.volume.blockSignals(False)
        self.playlist.set_shuffle(s.value("shuffle", "false") == "true")
        repeat = s.value("repeat", engine.REPEAT_OFF)
        self.playlist.repeat = repeat if repeat in engine.REPEAT_MODES else engine.REPEAT_OFF
        try:
            speed = float(s.value("speed", 1.0))
        except (TypeError, ValueError):
            speed = 1.0
        if speed in SPEEDS:
            self.speed_box.setCurrentIndex(SPEEDS.index(speed))
        # A settings file damaged by hand or by a crash mustn't stop the window opening.
        self.recent = RecentFolders(f for f in self.json_setting("recent folders", [])
                                    if isinstance(f, str) and f.strip())
        self.stations = self.saved_items("my stations", radio.Station.from_dict)
        self.spotify_links = self.saved_items("spotify links", spotify.SpotifyLink.from_dict)
        self.sleep_minutes = self.int_setting("sleep minutes", 30)
        action = self.str_setting("sleep action", engine.STOP_MUSIC)
        self.sleep_action = action if action in engine.SLEEP_ACTIONS else engine.STOP_MUSIC
        geometry = s.value("window place")
        if isinstance(geometry, (QByteArray, bytes)) and self.restoreGeometry(geometry):
            self.first_size(restored=True)          # unless last time's size is now too small
        else:
            self.first_size()
        self.tabs.setCurrentIndex(self.int_setting("tab", 0))
        self.apply_volume()
        self.show_recent()
        self.show_my_stations()
        self.show_spotify_links()
        self.show_queue()
        self.update_option_buttons()
        # The list and the place in the song from last time.
        last = os.path.join(self.data_folder, "last list.m3u8")
        try:
            entries = engine.read_playlist(last) if os.path.exists(last) else []
        except OSError:
            entries = []
        if entries:
            try:
                seconds = float(s.value("resume seconds", 0))
            except (TypeError, ValueError):
                seconds = 0
            if not 0 <= seconds < 1e7:                       # also not a number at all
                seconds = 0
            self.scan_restore = (self.str_setting("resume path", ""), seconds)
            # Read as the songs it names (not as the file), so that closing
            # before they're all read writes them all back.
            self.scan_play = False
            self.start_scan(entries)
            self.restore_job = self.scan_job
        self.update_list_label()

    def needed_size(self):
        """Width and height the window needs for everything in it, the radio's
        own rows too, measured with the big words already on (before it is
        shown, a part only has them once polished)."""
        for part in [self] + self.findChildren(QWidget):
            part.ensurePolished()
        content = self.page_scroll.widget().minimumSizeHint()
        extra = 0 if self.radio_bar.isVisible() else self.radio_bar.sizeHint().height()
        return (content.width() + 40,
                content.height() + extra + self.menuBar().sizeHint().height()
                + self.statusBar().sizeHint().height() + 40)

    def first_size(self, restored=False):
        """Big enough for everything, never bigger than the screen. When that is
        most of the screen anyway (as with the big words), it opens full screen:
        a window a little smaller would cut button words or need scrolling."""
        if restored and self.isMaximized():
            return
        wanted_w, wanted_h = self.needed_size()
        if restored and self.width() >= wanted_w and self.height() >= wanted_h:
            return                                   # the size kept from last time is fine
        wanted_w, wanted_h = max(1280, wanted_w), max(900, wanted_h)
        screen = self.screen().availableGeometry() if self.screen() else None
        if screen is None:
            self.resize(wanted_w, wanted_h)
        elif wanted_w > screen.width() * 0.85 or wanted_h > screen.height() * 0.85:
            self.resize(screen.width(), screen.height())
            self.setWindowState(Qt.WindowState.WindowMaximized)
        else:
            self.resize(wanted_w, wanted_h)

    def int_setting(self, key, default):
        try:
            return int(float(self.settings.value(key, default)))
        except (TypeError, ValueError, OverflowError):
            return default

    def str_setting(self, key, default):
        # An .ini line with a comma in it comes back as a list.
        value = self.settings.value(key, default)
        return value if isinstance(value, str) else default

    def saved_items(self, key, make):
        """Saved stations or links; a damaged one is left out, not the lot."""
        items = []
        for data in self.json_setting(key, []):
            try:
                items.append(make(data))
            except (ValueError, TypeError, AttributeError):
                continue
        return items

    def json_setting(self, key, default):
        try:
            value = json.loads(self.settings.value(key, ""))
            return value if isinstance(value, type(default)) else default
        except (TypeError, ValueError):
            return default

    def save_json_setting(self, key, value):
        self.settings.setValue(key, json.dumps(value, ensure_ascii=False))

    def save_settings(self):
        s = self.settings
        s.setValue("volume", self.local_volume)
        s.setValue("muted", "true" if self.muted else "false")
        s.setValue("shuffle", "true" if self.playlist.shuffle else "false")
        s.setValue("repeat", self.playlist.repeat)
        s.setValue("speed", self.speed_box.currentData())
        s.setValue("window place", self.saveGeometry())
        s.setValue("tab", self.tabs.currentIndex())
        s.setValue("sleep minutes", self.sleep_minutes)
        s.setValue("sleep action", self.sleep_action)
        self.save_json_setting("recent folders", list(self.recent))
        self.save_json_setting("my stations", [st.to_dict() for st in self.stations])
        self.save_json_setting("spotify links", [link.to_dict() for link in self.spotify_links])
        current = self.playlist.current
        if current is not None and self.playlist.find(current) is not None:
            s.setValue("resume path", current.path)
            ms = self.player.position() if self.player.source().isValid() else 0
            if self.player.mediaStatus() == QMediaPlayer.MediaStatus.EndOfMedia:
                ms = 0               # it had finished: next time it starts again, not at its end
            s.setValue("resume seconds", ms / 1000)
        else:
            s.setValue("resume path", "")
            s.setValue("resume seconds", 0)
        s.sync()

    def closeEvent(self, event):
        # The every-second timers stop first: one firing while the window
        # closes would ask the Spotify link, already shut, and fail.
        for timer in (self.ticker, self.spotify_timer, self.skip_timer):
            timer.stop()
        for extra in (getattr(self, "help_dialog", None), self.warning):
            if extra is not None:
                extra.close()
        reading = self.scan_job
        self.scanner.cancel()
        self.save_settings()
        entries = list(self.playlist.tracks)
        if reading is not None:
            # Still reading what was picked (or last time's list): keep what is
            # still to be read too, so next time it's all read rather than cut short.
            entries = entries[:self.scan_base] + [Track(p) for p in self.scan_paths]
            entries += [Track(p) for paths in self.scan_waiting for p in paths]
        for save in (lambda: engine.write_playlist(os.path.join(self.data_folder,
                                                                "last list.m3u8"), entries),
                     self.cache.save):
            try:
                save()
            except (OSError, ValueError):          # disk full, say: the other may still work
                pass
        self.stop_radio()
        self.player.stop()
        self.remote.close()
        super().closeEvent(event)

    # =========================================================== the modes
    def set_mode(self, mode):
        """Only one thing plays at a time: starting one quiets the others."""
        if mode != MUSIC and self.player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.player.pause()
        if mode != RADIO:
            self.stop_radio()
        if mode != SPOTIFY and self.spotify_now is not None and self.spotify_now.playing:
            self.remote.command("pause")
            self.spotify_now.playing = False
        if mode != self.mode:
            self.mode = mode
            self.volume.blockSignals(True)
            if mode == SPOTIFY:
                found = self.spotify_volume.get()
                if found is not None:
                    self.volume.setValue(round(found[0] * 100))
            else:
                self.volume.setValue(self.local_volume)
            self.volume.blockSignals(False)
            self.volume_label.setText(f"{self.volume.value()}%")
        self.radio_bar.setVisible(mode == RADIO)
        tips = {
            MUSIC: ("Previous song (P)", "Next song (N)", "Click or drag to move through the song"),
            RADIO: ("The station before this one in My stations (P)",
                    "The next station in My stations (N)",
                    "Click or drag to move through the last hour of the station"),
            SPOTIFY: ("Spotify's previous song", "Spotify's next song",
                      "Click or drag to move through Spotify's song"),
        }[mode]
        self.prev_button.setToolTip(tips[0])
        self.next_button.setToolTip(tips[1])
        self.seek.setToolTip(tips[2])
        self.update_option_buttons()
        self.show_now_playing()
        self.update_buttons()

    def toggle_play(self):
        if self.mode == RADIO:
            return self.radio_toggle()
        if self.mode == SPOTIFY:
            if self.spotify_now is None:
                if self.offer_spotify():
                    return
                self.message("Spotify isn't playing anything yet. Press Play in Spotify once, "
                             "or open a playlist from the Spotify page.")
            self.remote.command("toggle")
            if self.spotify_now is not None:
                self.spotify_now.playing = not self.spotify_now.playing
                self.spotify_now.position = self.spotify_now.position_now()
                self.spotify_now.read_at = time.monotonic()
            self.update_buttons()
            return
        state = self.player.playbackState()
        if state == QMediaPlayer.PlaybackState.PlayingState:
            self.player.pause()
        elif self.player.source().isValid() and self.playlist.current is not None:
            self.player.play()
        else:
            self.play_selected() if self.selected_tracks() else self.play_first()

    def stop(self):
        if self.mode == RADIO:
            self.stop_radio()
            self.message("Stopped the station.")
        elif self.mode == SPOTIFY:
            self.remote.command("pause")
        else:
            self.player.stop()
        self.update_buttons()
        self.show_now_playing()

    def skip(self, seconds):
        if self.mode == RADIO:
            return self.radio_seek_by(seconds)
        if self.mode == SPOTIFY:
            if self.spotify_now is not None:
                target = max(0, self.spotify_now.position_now() + seconds)
                self.remote.command("seek", target)
                self.spotify_now.position, self.spotify_now.read_at = target, time.monotonic()
            return
        if self.player.source().isValid():
            target = max(0, self.player.position() + seconds * 1000)
            duration = self.player.duration()
            if duration > 0 and target >= duration:
                return self.next()
            self.player.setPosition(int(target))

    def seek_to(self, value):
        if self.mode == RADIO:
            return self.radio_play_at(value * self.radio_buffer.rate // 1000
                                      if self.radio_buffer else 0)
        if self.mode == SPOTIFY:
            self.remote.command("seek", value / 1000)
            if self.spotify_now is not None:
                self.spotify_now.position, self.spotify_now.read_at = value / 1000, time.monotonic()
            return
        self.player.setPosition(value)

    def next(self):
        if self.mode == RADIO:
            return self.radio_step(+1)
        if self.mode == SPOTIFY:
            return self.remote.command("next")
        track = self.playlist.next()
        if track is None:
            self.message("That was the last song in the list. Turn on Repeat to go round again.")
            return
        self.play_track(track)

    def previous(self):
        if self.mode == RADIO:
            return self.radio_step(-1)
        if self.mode == SPOTIFY:
            return self.remote.command("previous")
        if self.player.position() > RESTART_IF_PAST * 1000 and self.player.source().isValid():
            self.player.setPosition(0)
            return
        track = self.playlist.previous()
        if track is None:
            self.player.setPosition(0)
            return
        self.play_track(track)

    def toggle_shuffle(self):
        if self.mode == SPOTIFY:
            on = not (self.spotify_now.shuffle if self.spotify_now and self.spotify_now.shuffle
                      is not None else False)
            self.remote.command("shuffle", on)
            if self.spotify_now is not None:
                self.spotify_now.shuffle = on
        elif self.mode == RADIO:
            self.message("Shuffle is for your own music. For radio, use Mark start, Mark end "
                         "and Repeat marked part.")
        else:
            self.playlist.set_shuffle(not self.playlist.shuffle)
            self.message("Shuffle is on: the list plays in a random order." if self.playlist.shuffle
                         else "Shuffle is off: the list plays in order.")
        self.update_option_buttons()
        self.update_buttons()

    def cycle_repeat(self):
        if self.mode == SPOTIFY:
            current = (self.spotify_now.repeat if self.spotify_now and self.spotify_now.repeat
                       else engine.REPEAT_OFF)
            following = engine.REPEAT_MODES[(engine.REPEAT_MODES.index(current) + 1) % 3]
            self.remote.command("repeat", following)
            if self.spotify_now is not None:
                self.spotify_now.repeat = following
        elif self.mode == RADIO:
            self.message("Repeat is for your own music. For radio, use Mark start, Mark end and "
                         "Repeat marked part.")
        else:
            mode = self.playlist.cycle_repeat()
            self.message({engine.REPEAT_OFF: "Repeat is off: playing stops after the last song.",
                          engine.REPEAT_ALL: "Repeat is on: the whole list plays round again.",
                          engine.REPEAT_ONE: "Repeat one song: the song playing plays over and "
                                             "over."}[mode])
        self.update_option_buttons()
        self.update_buttons()

    def update_option_buttons(self):
        if self.mode == SPOTIFY and self.spotify_now is not None:
            shuffle, repeat = bool(self.spotify_now.shuffle), self.spotify_now.repeat or "off"
        else:
            shuffle, repeat = self.playlist.shuffle, self.playlist.repeat
        self.shuffle_button.setChecked(shuffle)
        self.shuffle_button.setText("Shuffle: On" if shuffle else "Shuffle: Off")
        self.repeat_button.setText({"off": "Repeat: Off", "all": "Repeat: All",
                                    "one": "Repeat: One"}[repeat])
        self.repeat_button.setStyleSheet("" if repeat == "off" else "QToolButton { font-weight: bold; }")
        self.speed_box.setEnabled(self.mode == MUSIC)

    def speed_changed(self, index):
        self.player.setPlaybackRate(SPEEDS[index])

    # --- volume
    def volume_changed(self, value):
        self.volume_label.setText(f"{value}%")
        if self.mode == SPOTIFY:
            self.spotify_base_volume = None
            if not self.spotify_volume.set(value / 100, muted=False if value else None):
                self.message("Spotify's volume can be set once it has started playing.")
        else:
            self.local_volume = value
            if value and self.muted:
                self.muted = False
        self.apply_volume()

    def toggle_mute(self):
        if self.mode == SPOTIFY:
            found = self.spotify_volume.get()
            if found is None:
                self.message("Spotify's sound can be muted once it has started playing.")
                return
            self.spotify_volume.set(muted=not found[1])
        else:
            self.muted = not self.muted
        self.apply_volume()

    def apply_volume(self):
        factor = self.timer.volume_factor()
        linear = QtAudio.convertVolume(self.local_volume / 100,
                                       QtAudio.VolumeScale.LogarithmicVolumeScale,
                                       QtAudio.VolumeScale.LinearVolumeScale)
        self.audio.setVolume(linear * factor)
        self.audio.setMuted(self.muted)
        self.radio_output.set_volume(0.0 if self.muted else linear * factor)
        if self.mode == SPOTIFY:
            found = self.spotify_volume.get()
            muted = bool(found and found[1])
            if factor < 1 and found is not None:
                if self.spotify_base_volume is None:
                    self.spotify_base_volume = found[0]
                self.spotify_volume.set(self.spotify_base_volume * factor)
        else:
            muted = self.muted
        silent = muted or self.volume.value() == 0
        self.mute_button.setIcon(self.style().standardIcon(
            QStyle.StandardPixmap.SP_MediaVolumeMuted if silent else QStyle.StandardPixmap.SP_MediaVolume))
        self.mute_button.setText("Unmute" if muted else "Mute")
        self.mute_button.setStyleSheet("QToolButton { color: #b3261e; font-weight: bold; }"
                                       if muted else "")

    # ============================================================ My Music
    def start_folder(self):
        start = self.str_setting("last folder", "")
        if not start or not os.path.isdir(start):
            start = os.path.join(os.path.expanduser("~"), "Music")
        return start

    def choose_folder(self, add=False):
        folder = QFileDialog.getExistingDirectory(
            self, "Add a folder of music" if add else "Open a folder of music",
            self.start_folder())
        if folder:
            self.settings.setValue("last folder", os.path.dirname(folder) or folder)
            self.open_folder(folder, add=add)

    def open_folder(self, folder, add=False):
        if not os.path.isdir(folder):
            answer = QMessageBox.question(
                self, TITLE, f"The folder\n\n{folder}\n\ncan't be found. It may have been moved, "
                "renamed or be on a drive that isn't plugged in.\n\nTake it off your last folders?")
            if answer == QMessageBox.StandardButton.Yes:
                self.recent.remove(folder)
                self.show_recent()
            return
        self.recent.add(folder)
        self.show_recent()
        self.add_paths([folder], replace=not add, play=not add or not self.is_playing())

    def choose_songs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Add songs", self.start_folder(),
                                                file_filter())
        if paths:
            self.settings.setValue("last folder", os.path.dirname(paths[0]))
            self.add_paths(paths, replace=False, play=not self.is_playing())

    def choose_playlist(self):
        lists = " ".join("*" + ext for ext in sorted(engine.PLAYLIST_EXTENSIONS))
        path, _ = QFileDialog.getOpenFileName(self, "Open a playlist", self.start_folder(),
                                              f"Playlists ({lists});;All files (*)")
        if path:
            self.settings.setValue("last folder", os.path.dirname(path))
            self.add_paths([path], replace=True, play=True)

    def save_playlist(self):
        if not self.playlist.tracks:
            self.message("There are no songs in the list to save yet.")
            return
        start = os.path.join(self.start_folder(), "My playlist.m3u8")
        path, _ = QFileDialog.getSaveFileName(self, "Save this list as a playlist", start,
                                              "Playlist (*.m3u8);;Older playlist (*.m3u)")
        if not path:
            return
        if not os.path.splitext(path)[1]:
            path += ".m3u8"
        try:
            engine.write_playlist(path, self.playlist.tracks)
        except OSError as error:
            QMessageBox.warning(self, TITLE, f"The playlist couldn't be saved there.\n\n{error}")
            return
        self.message(f"Saved {plural(len(self.playlist.tracks), 'song')} to {path}")

    def open_paths(self, paths):
        """Files or folders given when starting (dragged onto the .bat)."""
        paths = [p for p in paths if os.path.exists(p)]
        if paths:
            for p in paths:
                if os.path.isdir(p):
                    self.recent.add(p)
            self.show_recent()
            self.add_paths(paths, replace=True, play=True)

    def add_paths(self, paths, replace, play):
        if replace:
            self.player.stop()
            self.player.setSource(QUrl())
            self.playlist.replace([])
            self.model.broken.clear()
            self.model.playing = None
            self.model.reset()
            self.show_queue()
            self.sort_key = None
            self.table.horizontalHeader().setSortIndicatorShown(False)
            self.scan_restore = None
            self.scan_waiting = []
        elif self.scan_job is not None:
            # Still reading what was picked before: these wait their turn, so
            # those songs all still arrive, and in order.
            self.scan_waiting.append([os.path.abspath(p) for p in paths])
            self.scan_play = self.scan_play or play
            self.tabs.setCurrentIndex(0)
            return
        self.scan_play = play
        self.start_scan(paths)
        self.tabs.setCurrentIndex(0)

    def start_scan(self, paths):
        self.scan_paths = [os.path.abspath(p) for p in paths]
        self.scan_base = len(self.playlist.tracks)
        self.scan_job = self.scanner.start(self.scan_paths)
        self.update_list_label(still_reading=True)

    def songs_found(self, job, batch):
        if job != self.scan_job:
            return
        self.playlist.add(batch)
        self.model.appended(len(batch))
        self.update_list_label(still_reading=True)
        if self.scan_play and self.playlist.current is None and self.playlist.tracks:
            self.scan_play = False
            first = (self.playlist.shuffled[0] if self.playlist.shuffle
                     else self.playlist.tracks[0])
            self.play_track(self.playlist.choose(first))
        if self.scan_restore:
            path, seconds = self.scan_restore
            for track in batch:
                if path and engine.same_path(track.path, path):
                    self.scan_restore = None
                    self.playlist.choose(track)
                    self.play_track(track, position=seconds, start=False)
                    self.message("Press Play to carry on where you left off.")
                    break

    def scan_finished(self, job, count, addresses):
        if job != self.scan_job:
            return
        quiet = job == self.restore_job       # last time's list: nothing to announce
        self.scan_job = self.restore_job = None
        self.scan_restore = None
        if addresses and not quiet:
            names = [st.url for st in self.stations]
            added = [a for a in addresses if a not in names]
            for address in added:
                self.stations.append(radio.Station(address, address))
            if added:
                self.show_my_stations()
                self.message(f"The playlist also had {plural(len(added), 'web address')} "
                             "(radio, probably). They're in My stations on the Radio page.")
        elif count == 0 and not self.playlist.tracks and not quiet and not self.scan_waiting:
            self.message("No music was found there.")
        if self.scan_play and self.playlist.tracks and self.playlist.current is None:
            self.scan_play = False
            self.play_track(self.playlist.choose(self.playlist.tracks[0]))
        try:
            self.cache.save()
        except (OSError, ValueError):
            pass
        if self.scan_waiting:                 # songs added meanwhile: their turn now
            self.start_scan(self.scan_waiting.pop(0))
        else:
            self.update_list_label()

    def update_list_label(self, still_reading=False):
        tracks = self.playlist.tracks
        empty = not tracks and not still_reading and self.scan_job is None
        self.table.setVisible(not empty)
        self.empty_page.setVisible(empty)
        if not tracks:
            self.list_label.setText("Finding music..." if still_reading or self.scan_job
                                    else "No songs yet.")
            return
        total = sum(t.length for t in tracks)
        text = f"{plural(len(tracks), 'song')}, {long_time(total)}"
        shown = self.filter.rowCount()
        if self.filter.words:
            text = f"{shown:,} of {text} match the search"
        if still_reading:
            text += " (still finding more...)"
        self.list_label.setText(text)

    def search_changed(self, text):
        self.filter.set_text(text)
        self.update_list_label(self.scan_job is not None)

    def find_in_list(self):
        self.tabs.setCurrentIndex(0)
        self.search.setFocus()
        self.search.selectAll()

    def sort_by(self, column):
        key = SongModel.COLUMNS[column][1]
        if key == self.sort_key:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_key, self.sort_descending = key, False
        self.playlist.sort(key, self.sort_descending)
        self.model.reset()
        header = self.table.horizontalHeader()
        header.setSortIndicatorShown(column != 0)
        header.setSortIndicator(column, Qt.SortOrder.DescendingOrder if self.sort_descending
                                else Qt.SortOrder.AscendingOrder)
        self.scroll_to_playing()

    def selected_tracks(self):
        rows = sorted({self.filter.mapToSource(i).row()
                       for i in self.table.selectionModel().selectedRows()})
        return [self.playlist.tracks[r] for r in rows if 0 <= r < len(self.playlist.tracks)]

    def play_row(self, row):
        if 0 <= row < len(self.playlist.tracks):
            self.play_track(self.playlist.choose(self.playlist.tracks[row]))

    def play_selected(self):
        chosen = self.selected_tracks()
        if not chosen:
            if not self.playlist.tracks:
                self.message("Open a folder of music first.")
            else:
                self.message("Click a song in the list first, then press Play.")
            return
        self.play_track(self.playlist.choose(chosen[0]))

    def play_first(self):
        if not self.playlist.tracks:
            self.message("Open a folder of music first, or pick a station on the Radio page.")
            return
        first = self.playlist.shuffled[0] if self.playlist.shuffle else self.playlist.tracks[0]
        self.play_track(self.playlist.choose(first))

    def queue_selected(self, first):
        chosen = self.selected_tracks()
        if not chosen:
            self.message("Click one or more songs in the list first.")
            return
        if first:
            self.playlist.play_next(chosen)
            self.message(f"{plural(len(chosen), 'song')} will play after this one.")
        else:
            self.playlist.add_to_queue(chosen)
            self.message(f"Added {plural(len(chosen), 'song')} to Up next.")
        self.show_queue()
        self.update_buttons()

    def show_queue(self):
        self.queue_list.clear()
        for track in self.playlist.queue:
            item = QListWidgetItem(f"{track.name} – {track.artist}" if track.artist else track.name)
            item.setToolTip(track.path)
            self.queue_list.addItem(item)
        self.queue_box.setVisible(bool(self.playlist.queue))

    def unqueue(self):
        row = self.queue_list.currentRow()
        if 0 <= row < len(self.playlist.queue):
            del self.playlist.queue[row]
            self.show_queue()
        else:
            self.message("Click a song in Up next first.")

    def clear_queue(self):
        self.playlist.queue.clear()
        self.show_queue()

    def remove_selected(self):
        chosen = self.selected_tracks()
        if not chosen:
            self.message("Click one or more songs in the list first.")
            return
        self.playlist.remove(chosen)
        self.model.reset()
        self.show_queue()
        self.update_list_label()
        self.update_buttons()
        self.message(f"Took {plural(len(chosen), 'song')} off the list. The files are still there.")

    def clear_list(self):
        if not self.playlist.tracks:
            return
        if QMessageBox.question(self, TITLE, "Take every song off the list? (The music files "
                                "themselves stay where they are.)") != QMessageBox.StandardButton.Yes:
            return
        self.scanner.cancel()
        self.scan_job = None
        self.scan_waiting = []
        self.player.stop()
        self.player.setSource(QUrl())
        self.playlist.replace([])
        self.model.playing = None
        self.model.reset()
        self.show_queue()
        self.update_list_label()
        self.show_now_playing()
        self.update_buttons()

    def show_in_folder(self):
        chosen = self.selected_tracks() or ([self.playlist.current] if self.playlist.current else [])
        if not chosen:
            self.message("Click a song in the list first.")
            return
        subprocess.Popen(["explorer", "/select,", os.path.normpath(chosen[0].path)])

    def song_menu(self, point):
        if not self.selected_tracks():
            return
        menu = QMenu(self)
        menu.addAction("Play", self.play_selected)
        menu.addAction("Play next", lambda: self.queue_selected(first=True))
        menu.addAction("Add to Up next", lambda: self.queue_selected(first=False))
        menu.addSeparator()
        menu.addAction("Remove from list", self.remove_selected)
        menu.addAction("Show in folder", self.show_in_folder)
        menu.exec(self.table.viewport().mapToGlobal(point))

    def show_recent(self):
        folders = list(self.recent)
        for i, (row, open_button, remove) in enumerate(self.recent_rows):
            if i < len(folders):
                folder = folders[i]
                name = os.path.basename(folder.rstrip("\\/")) or folder
                open_button.set_full_text(name)
                open_button.setToolTip(f"Open {folder}\n(its subfolders too)")
                row.show()
            else:
                row.hide()
        self.no_recent.setVisible(not folders)

    def open_recent(self, place):
        folders = list(self.recent)
        if place < len(folders):
            self.open_folder(folders[place])

    def remove_recent(self, place):
        folders = list(self.recent)
        if place >= len(folders):
            return
        folder = folders[place]
        self.recent.remove(folder)
        self.show_recent()
        self.message(f"Took {os.path.basename(folder)} off your last folders. "
                     "The folder itself is still there.")

    def scroll_to_playing(self):
        row = self.playlist.find(self.playlist.current) if self.playlist.current else None
        if row is not None:
            index = self.filter.mapFromSource(self.model.index(row, 1))
            if index.isValid():
                self.table.scrollTo(index, QAbstractItemView.ScrollHint.EnsureVisible)

    # --- playing a song
    def is_playing(self):
        if self.mode == MUSIC:
            return self.player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        if self.mode == RADIO:
            return self.radio_output.playing or self.radio_waiting
        return bool(self.spotify_now and self.spotify_now.playing)

    def play_track(self, track, position=0.0, start=True):
        if track is None:
            return
        self.set_mode(MUSIC)
        self.skip_timer.stop()
        url = QUrl.fromLocalFile(track.path)
        if self.player.source() == url and self.player.mediaStatus() not in (
                QMediaPlayer.MediaStatus.NoMedia, QMediaPlayer.MediaStatus.InvalidMedia):
            self.player.setPosition(int(position * 1000))
        else:
            self.pending_position = (url, int(position * 1000)) if position else None
            self.player.setSource(url)
        self.player.setPlaybackRate(self.speed_box.currentData())
        if start:
            self.player.play()
        old = self.model.playing
        self.model.playing = track
        if old is not None:
            self.model.refresh_track(old)
        self.model.refresh_track(track)
        self.scroll_to_playing()
        self.show_queue()
        self.show_now_playing()
        self.update_buttons()

    def media_status_changed(self, status):
        Status = QMediaPlayer.MediaStatus
        if status in (Status.LoadedMedia, Status.BufferedMedia) and self.pending_position:
            url, ms = self.pending_position
            if self.player.source() == url:        # not the last song's news
                self.pending_position = None
                self.player.setPosition(ms)
        if status in (Status.LoadedMedia, Status.BufferedMedia):
            self.failures = 0
            self.update_position()
        elif status == Status.EndOfMedia:
            self.song_finished()

    def song_finished(self):
        if self.timer.at_song_end:
            return self.sleep_now()
        track = self.playlist.next(automatic=True)
        if track is None:
            self.message("That was the last song in the list.")
            self.update_buttons()
            return
        if track is self.model.playing and self.playlist.repeat == engine.REPEAT_ONE:
            self.player.setPosition(0)
            self.player.play()
            return
        self.play_track(track)

    def player_error(self, error, text):
        if error == QMediaPlayer.Error.NoError or self.mode != MUSIC:
            return
        track = self.model.playing
        if track is not None:
            self.model.broken.add(id(track))
            self.model.refresh_track(track)
        name = track.name if track else "This song"
        self.message(f"“{name}” couldn't be played ({text}). Skipping it.")
        self.failures += 1
        if self.failures < max(3, len(self.playlist.tracks)) and self.playlist.has_next():
            self.skip_timer.start()

    def skip_broken(self):
        if self.mode == MUSIC:
            self.next()

    def position_changed(self, ms):
        if self.mode == MUSIC:
            self.update_position()

    # ============================================================== Radio
    def tab_changed(self, index):
        if index == 1 and not self.radio_searched:
            self.radio_searched = True
            self.popular_stations()
        if index == 1 and not self.countries_loaded:
            self.countries_loaded = True
            in_background(self.bridge, self.directory.countries, self.countries_found)
        if index == 2:
            self.ask_spotify()
            if not self.spotify_offered:
                QTimer.singleShot(300, self.offer_spotify_here)     # once the page shows

    def offer_spotify_here(self):
        if self.tabs.currentIndex() == 2 and self.isVisible():
            self.offer_spotify()

    def countries_found(self, countries, error):
        if error or not countries:
            self.countries_loaded = False
            return
        chosen = self.country_box.currentData()
        self.country_box.blockSignals(True)
        self.country_box.clear()
        self.country_box.addItem("All countries", "")
        for name, code in countries:
            self.country_box.addItem(name, code)
        index = self.country_box.findData(chosen)
        self.country_box.setCurrentIndex(max(0, index))
        self.country_box.blockSignals(False)

    def search_stations(self):
        text = self.station_search.text().strip()
        if not text:
            return self.popular_stations()
        country = self.country_box.currentData()
        self.results_label.setText(f"Looking for “{text}”...")
        asked = self.results_asked = getattr(self, "results_asked", 0) + 1
        in_background(self.bridge, lambda: self.directory.search(text, country),
                      lambda found, error: self.show_results(found, error, text, asked))

    def popular_stations(self):
        country = self.country_box.currentData()
        self.results_label.setText("Looking up the most listened stations...")
        asked = self.results_asked = getattr(self, "results_asked", 0) + 1
        in_background(self.bridge, lambda: self.directory.popular(country),
                      lambda found, error: self.show_results(found, error, None, asked))

    def show_results(self, found, error, text, asked=None):
        if asked is not None and asked != self.results_asked:
            return            # an older, slower search: the newest one decides what shows
        if error:
            self.results_label.setText(str(error) if isinstance(error, ConnectionError) else
                                       "The station list couldn't be reached. Check the "
                                       "internet connection and try again.")
            return
        self.found_stations = found
        self.results.setRowCount(0)
        for station in found:
            row = self.results.rowCount()
            self.results.insertRow(row)
            for column, value in enumerate((station.name, station.tags, station.country,
                                            station.quality)):
                item = QTableWidgetItem(value)
                item.setToolTip(f"{station.name}\n{station.url}" if column == 0 else value)
                self.results.setItem(row, column, item)
        where = self.country_box.currentText()
        where = "" if not self.country_box.currentData() else f" in {where}"
        if not found:
            self.results_label.setText(f"No stations found for “{text}”{where}. Try "
                                       "another word, or All countries." if text else
                                       "No stations found.")
        elif text:
            self.results_label.setText(f"{plural(len(found), 'station')} for “{text}”"
                                       f"{where}, most listened first.")
        else:
            self.results_label.setText(f"The {len(found)} most listened stations{where}.")
        if found:
            self.results.selectRow(0)

    def picked_result(self):
        row = self.results.currentRow()
        stations = getattr(self, "found_stations", [])
        return stations[row] if 0 <= row < len(stations) else None

    def play_result(self, row=None):
        if isinstance(row, int) and row >= 0:
            self.results.selectRow(row)
        station = self.picked_result()
        if station is None:
            self.message("Click a station in the list first.")
            return
        self.play_station(station)

    def keep_result(self):
        station = self.picked_result()
        if station is None:
            self.message("Click a station in the list first.")
            return
        self.keep_station(station)

    def keep_station(self, station):
        if any(s.url == station.url for s in self.stations):
            self.message(f"{station.name} is already in My stations.")
            return
        self.stations.append(station)
        self.show_my_stations()
        self.save_json_setting("my stations", [st.to_dict() for st in self.stations])
        self.message(f"Added {station.name} to My stations.")

    def show_my_stations(self):
        self.my_stations.clear()
        for station in self.stations:
            item = QListWidgetItem(station.name)
            item.setToolTip(f"{station.tags}\n{station.url}".strip())
            if self.station is not None and station.url == self.station.url:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
            self.my_stations.addItem(item)
        if not self.stations:
            item = QListWidgetItem("Stations you add show here.")
            item.setFlags(Qt.ItemFlag.ItemIsEnabled)     # a black hint, not a choice
            self.my_stations.addItem(item)

    def play_my_station(self):
        row = self.my_stations.currentRow()
        if not 0 <= row < len(self.stations):
            self.message("Click a station in My stations first.")
            return
        self.play_station(self.stations[row])

    def remove_my_station(self):
        row = self.my_stations.currentRow()
        if not 0 <= row < len(self.stations):
            self.message("Click a station in My stations first.")
            return
        station = self.stations.pop(row)
        self.show_my_stations()
        self.save_json_setting("my stations", [st.to_dict() for st in self.stations])
        self.message(f"Took {station.name} off My stations.")

    def add_station_address(self):
        text, ok = QInputDialog.getText(
            self, "Add a station by web address",
            "The station's stream address\n(it often ends in .mp3, .aac, .pls or .m3u):",
            text=QGuiApplication.clipboard().text().strip()
            if engine.is_web_address(QGuiApplication.clipboard().text()) else "")
        text = text.strip()
        if not ok or not text:
            return
        if not engine.is_web_address(text):
            text = "http://" + text
        name, ok = QInputDialog.getText(self, "Add a station by web address",
                                        "A name for it\n(you can leave this as it is):",
                                        text=text)
        station = radio.Station(name.strip() if ok and name.strip() else text, text)
        self.keep_station(station)
        self.play_station(station)

    def play_station(self, station):
        self.set_mode(RADIO)
        self.stop_radio()
        self.station = station
        self.radio_message = "Connecting..."
        self.radio_buffer = radio.TimeshiftBuffer(RadioOutput.best_rate(), 2, radio.BUFFER_MINUTES)
        self.radio_output.attach(self.radio_buffer)
        self.radio_marks = [None, None]
        self.repeat_part_button.setChecked(False)
        try:
            recorder = radio.RadioRecorder(
                station.url, self.radio_buffer,
                lambda kind, value: self.bridge.radio_event.emit(recorder, kind, value))
        except Exception as error:          # FFmpeg missing, say
            self.radio_message = ""
            QMessageBox.warning(self, TITLE, f"Radio can't be played: {error}")
            return
        self.recorder = recorder
        self.radio_waiting = True
        recorder.start()
        in_background(self.bridge, lambda: self.directory.count_click(station), lambda *_: None)
        self.show_my_stations()
        self.show_heard_songs()
        self.show_marks()
        self.show_now_playing()
        self.update_buttons()
        self.load_station_picture(station)

    def stop_radio(self):
        recorder, self.recorder = self.recorder, None
        if recorder is not None:
            recorder.stop()
        self.radio_output.stop()
        if self.radio_buffer is not None:
            self.radio_output.buffer = None
            self.radio_buffer.close()
            self.radio_buffer = None
        self.radio_waiting = False
        self.radio_song = None
        self.show_heard_songs()

    def radio_event(self, recorder, kind, value):
        if recorder is not self.recorder:
            return                              # news from a station already left
        if kind == "name" and value and self.station and self.station.name == self.station.url:
            self.station.name = value
            self.show_my_stations()
            self.save_json_setting("my stations", [st.to_dict() for st in self.stations])
        elif kind == "started":
            self.radio_message = ""
        elif kind == "song":
            self.show_heard_songs()
        elif kind == "reconnecting":
            self.radio_message = "The station dropped out; trying again..."
            self.message(self.radio_message)
        elif kind == "error":
            had_sound = self.radio_buffer is not None and self.radio_buffer.written > 0
            self.radio_message = value.split("\n")[0]
            if had_sound:
                self.message(value.split("\n")[0] + " What was already kept can still be played.",
                             15000)
            else:
                self.radio_waiting = False
                QMessageBox.information(self, TITLE, f"{self.station.name if self.station else 'The station'}"
                                        f" can't be played right now.\n\n{value}")
        elif kind == "ended":
            self.recorder = None
            if self.radio_buffer is not None and self.radio_buffer.written == 0:
                self.radio_waiting = False
        self.show_now_playing()
        self.update_buttons()

    def load_station_picture(self, station):
        self.station_picture = None
        if not station.favicon.startswith(("http://", "https://")):
            return
        reply = self.network.get(QNetworkRequest(QUrl(station.favicon)))

        def arrived():
            data = bytes(reply.readAll())
            reply.deleteLater()
            if self.station is station and data:
                picture = QPixmap()
                if picture.loadFromData(data):
                    self.station_picture = picture
                    self.show_now_playing()
        reply.finished.connect(arrived)

    def radio_heard(self):
        return self.radio_output.heard() if self.radio_buffer else 0

    def radio_start(self, frame):
        """Play the kept station sound from frame. False, having said why in
        plain words, if the speakers can't be used."""
        try:
            self.radio_output.play(frame)
            return True
        except RuntimeError as error:
            self.radio_waiting = False
            QMessageBox.warning(self, TITLE, f"{error}\n\nCheck that the speakers or headphones "
                                "are plugged in and switched on, then press Play again.")
            return False

    def radio_play_at(self, frame):
        buffer = self.radio_buffer
        if buffer is None or buffer.written == 0:
            self.message("Nothing has come in from the station yet.")
            return
        newest = buffer.written - int(LIVE_GAP_SECONDS * buffer.rate)
        frame = int(min(max(frame, buffer.oldest), max(buffer.oldest, newest)))
        if self.radio_output.playing or self.radio_waiting:
            self.radio_waiting = False
            self.radio_start(frame)
        else:
            self.radio_output.paused_at = frame      # paused: stay paused, at the new place
        self.show_now_playing()
        self.update_buttons()

    def radio_seek_by(self, seconds):
        if self.radio_buffer is None:
            return
        self.radio_play_at(self.radio_heard() + seconds * self.radio_buffer.rate)

    def radio_toggle(self):
        nothing_came = (self.radio_buffer is not None and self.radio_buffer.written == 0
                        and self.recorder is None)
        if self.radio_buffer is None or nothing_came:
            if self.station is not None:
                self.play_station(self.station)
            else:
                self.message("Pick a station on the Radio page first.")
            return
        if self.radio_output.playing:
            self.radio_output.pause()
            self.message("Paused. The station is still being kept, so you can carry on from "
                         "here or press Go live.")
        elif self.radio_waiting:
            self.radio_waiting = False
            self.radio_output.paused_at = 0
        else:
            buffer = self.radio_buffer
            frame = self.radio_output.paused_at
            if frame < buffer.oldest:
                self.message("Paused for longer than the hour kept: playing from the oldest part.")
            self.radio_start(frame)
        self.show_now_playing()
        self.update_buttons()

    def go_live(self):
        if self.mode != RADIO or self.radio_buffer is None:
            return
        buffer = self.radio_buffer
        self.repeat_part_button.setChecked(False)
        self.radio_output.loop = None
        self.radio_waiting = False
        self.radio_start(max(buffer.oldest, buffer.written - int(LIVE_GAP_SECONDS * buffer.rate)))
        self.show_now_playing()
        self.update_buttons()

    def radio_step(self, step):
        if not self.stations:
            self.message("Add stations to My stations to step through them.")
            return
        urls = [s.url for s in self.stations]
        here = urls.index(self.station.url) if self.station and self.station.url in urls else -1
        target = (here + step) % len(self.stations) if here >= 0 else (0 if step > 0 else -1)
        self.play_station(self.stations[target])

    def radio_song_start(self):
        if self.radio_buffer is None:
            self.message("Play a station first.")
            return
        song = self.radio_buffer.song_at(self.radio_heard())
        if song is None:
            self.message("This station doesn't say when songs change. Use Back 30 s instead.")
            return
        self.radio_play_at(song[0])
        self.message(f"Back to the start of “{song[1]}” (as near as the station says).")

    def radio_mark_start(self):
        if self.radio_buffer is None:
            self.message("Play a station first.")
            return
        self.radio_marks[0] = self.radio_heard()
        if self.radio_marks[1] is not None and self.radio_marks[1] <= self.radio_marks[0]:
            self.radio_marks[1] = None
        self.show_marks()
        self.message("Marked the start. Now press 2. Mark end where the part should end.")

    def radio_mark_end(self):
        if self.radio_buffer is None:
            self.message("Play a station first.")
            return
        here = self.radio_heard()
        if self.radio_marks[0] is None or here <= self.radio_marks[0]:
            self.message("Press 1. Mark start first, where the part to repeat begins, then let "
                         "it play to where it should end and press 2. Mark end.")
            return
        self.radio_marks[1] = here
        self.show_marks()
        self.message("Marked the end. Press 3. Repeat marked part to hear it over and over.")

    def radio_repeat_part(self):
        start, end = self.radio_marks
        if self.radio_buffer is None or start is None or end is None:
            self.repeat_part_button.setChecked(False)
            self.message("First press 1. Mark start where the part begins and 2. Mark end where "
                         "it ends; then this repeats it.")
            return
        if self.repeat_part_button.isChecked():
            if start < self.radio_buffer.oldest:
                self.repeat_part_button.setChecked(False)
                self.message("That part is older than the hour kept, so it's gone.")
                return
            self.radio_output.loop = (start, end)
            self.radio_waiting = False
            if not self.radio_start(start):
                self.radio_output.loop = None
                self.repeat_part_button.setChecked(False)
                return
            self.message("Repeating the marked part. Press it again (or Go live) to stop repeating.")
        else:
            self.radio_output.loop = None
            self.message("Stopped repeating; the station plays on from here.")
        self.show_marks()
        self.update_buttons()

    def radio_loop_lost(self):
        self.repeat_part_button.setChecked(False)
        self.show_marks()
        self.message("The marked part is now older than the hour kept, so repeating stopped; "
                     "the station plays on from the oldest part kept.", 15000)

    def radio_clear_marks(self):
        self.radio_marks = [None, None]
        self.radio_output.loop = None
        self.repeat_part_button.setChecked(False)
        self.show_marks()

    def frame_clock(self, frame):
        """The time of day a frame of the station was heard live."""
        buffer = self.radio_buffer
        seconds_ago = (buffer.written - frame) / buffer.rate
        return time.strftime("%I:%M:%S %p", time.localtime(time.time() - seconds_ago)).lstrip("0")

    def show_marks(self):
        start, end = self.radio_marks
        if self.radio_buffer is None or (start is None and end is None):
            self.marks_label.setText("")
            return
        text = f"Marked from {self.frame_clock(start)}" if start is not None else ""
        if end is not None:
            text += f" to {self.frame_clock(end)} ({format_time((end - start) / self.radio_buffer.rate)})"
        if self.radio_output.loop:
            text += ", repeating"
        self.marks_label.setText(text)

    def show_heard_songs(self):
        self.heard_list.clear()
        self.heard_songs = []
        if self.radio_buffer is None:
            return
        for frame, title, when in reversed(self.radio_buffer.songs()):
            self.heard_songs.append(frame)
            self.heard_list.addItem(f"{time.strftime('%I:%M %p', time.localtime(when)).lstrip('0')}"
                                    f"   {title}")

    def play_heard_song(self, item=None):
        row = self.heard_list.currentRow()
        if not 0 <= row < len(getattr(self, "heard_songs", [])) or self.radio_buffer is None:
            self.message("Click a song in Songs heard first.")
            return
        frame = self.heard_songs[row]
        self.set_mode(RADIO) if self.mode != RADIO else None
        if not self.radio_output.playing:
            self.radio_output.paused_at = frame
            self.radio_toggle()
        else:
            self.radio_play_at(frame)

    # ============================================================= Spotify
    def show_spotify_links(self):
        self.spotify_list.clear()
        for link in self.spotify_links:
            item = QListWidgetItem(link.label)
            item.setToolTip(link.web_url)
            self.spotify_list.addItem(item)
        if not self.spotify_links:
            item = QListWidgetItem("Your Spotify playlists show here once you add their links.")
            item.setFlags(Qt.ItemFlag.ItemIsEnabled)     # a black hint, not a choice
            self.spotify_list.addItem(item)

    def save_spotify_links(self):
        self.save_json_setting("spotify links", [link.to_dict() for link in self.spotify_links])

    def add_spotify_link(self):
        pasted = QGuiApplication.clipboard().text().strip()
        start = pasted if spotify.parse_link(pasted) or spotify.SHORT_LINK.search(pasted) else ""
        text, ok = QInputDialog.getText(
            self, "Add a Spotify link",
            "Paste the link here.\n(In Spotify: right-click the playlist in Your Library,\n"
            "then Share, then Copy link to playlist.)",
            text=start)
        if not ok or not text.strip():
            return
        link = spotify.resolve_link(text.strip())
        if link is None:
            QMessageBox.information(self, TITLE, "That doesn't look like a Spotify link. It should "
                                    "start with https://open.spotify.com/ (for example "
                                    "https://open.spotify.com/playlist/...).")
            return
        if any(l.kind == link.kind and l.id == link.id for l in self.spotify_links):
            self.message("That one is already on the list.")
            return
        self.spotify_links.append(link)
        self.show_spotify_links()
        self.save_spotify_links()
        self.spotify_list.setCurrentRow(len(self.spotify_links) - 1)
        self.message("Added. Looking up its name...")

        def named(name, error):
            if name:
                link.name = name
                self.show_spotify_links()
                self.save_spotify_links()
                self.message(f"Added “{name}”.")
        in_background(self.bridge, lambda: spotify.look_up_name(link), named)

    def picked_spotify(self):
        row = self.spotify_list.currentRow()
        return self.spotify_links[row] if 0 <= row < len(self.spotify_links) else None

    def play_spotify(self):
        link = self.picked_spotify()
        if link is None:
            self.message("Click a playlist in the list first." if self.spotify_links else
                         "Add a Spotify link first.")
            return
        self.set_mode(SPOTIFY)
        where = spotify.open_in_spotify(link)
        if where == "web":
            self.message("The Spotify app wasn't found, so the playlist opened in the web "
                         "browser instead.", 15000)
        else:
            self.message(f"Opened “{link.label}” in Spotify. If it doesn't start by "
                         "itself, press the green Play button there once.", 15000)
        self.ask_spotify()

    def rename_spotify(self):
        link = self.picked_spotify()
        if link is None:
            self.message("Click a playlist in the list first.")
            return
        name, ok = QInputDialog.getText(self, "Rename", "Name shown here:", text=link.label)
        if ok and name.strip():
            link.name = name.strip()
            self.show_spotify_links()
            self.save_spotify_links()

    def remove_spotify(self):
        link = self.picked_spotify()
        if link is None:
            self.message("Click a playlist in the list first.")
            return
        self.spotify_links.remove(link)
        self.show_spotify_links()
        self.save_spotify_links()
        self.message(f"Took “{link.label}” off the list (it's still in Spotify).")

    def open_spotify(self):
        if not spotify.open_spotify_app():
            QMessageBox.information(self, TITLE, "The Spotify app wasn't found on this computer. "
                                    "It can be installed free from the Microsoft Store.")
            return
        self.set_mode(SPOTIFY)
        self.message("Spotify is opening. Start any playlist there; the buttons at the top "
                     "of this window then work it.", 15000)

    def offer_spotify(self):
        """The box that opens Spotify from here when it isn't open yet. Asked
        once a run when going to the Spotify page, and whenever Play is
        pressed for Spotify with Spotify closed. True if it opened it."""
        self.spotify_offered = True
        if self.spotify_running():
            return False
        answer = QMessageBox.question(self, TITLE, "The Spotify app isn't open.\n\n"
                                      "Open it now?")
        if answer != QMessageBox.StandardButton.Yes:
            return False
        self.open_spotify()
        return True

    def ask_spotify(self):
        wanted = self.mode == SPOTIFY or self.tabs.currentIndex() == 2
        if wanted and not self.spotify_asking:
            self.spotify_asking = True
            self.remote.request_now(self.bridge.spotify_now.emit)

    def spotify_news(self, now):
        self.spotify_asking = False
        before = self.spotify_now
        self.spotify_now = now
        if now is None:
            self.spotify_status.setText("Spotify isn't open, or hasn't played anything since it "
                                        "started.")
        else:
            state = "Playing" if now.playing else "Paused"
            song = f"“{now.title}”" + (f" by {now.artist}" if now.artist else "")
            self.spotify_status.setText(f"Spotify: {state} {song}" if now.title else
                                        f"Spotify is open ({state.lower()}).")
        if self.mode == SPOTIFY:
            changed = (before is not None and now is not None and before.song_key != now.song_key)
            if changed and self.timer.at_song_end:
                self.sleep_now()
            self.update_option_buttons()
            self.show_now_playing()
            self.update_buttons()

    # ======================================================== now playing
    def show_now_playing(self):
        picture = None
        source = ""
        if self.mode == RADIO and self.station is not None:
            source = "Internet radio"
            song = self.radio_buffer.song_at(self.radio_heard()) if self.radio_buffer else None
            if song:
                parts = song[1].split(" - ", 1)
                title, artist = (parts[1], parts[0]) if len(parts) == 2 else (song[1], "")
                self.title_label.setText(title)
                self.artist_label.setText(f"{artist}   ·   {self.station.name}" if artist
                                          else self.station.name)
            else:
                self.title_label.setText(self.station.name)
                self.artist_label.setText(self.station.tags)
            picture = getattr(self, "station_picture", None)
            self.info_label.setText(self.radio_state_text())
            placeholder = "\U0001F4FB"
        elif self.mode == SPOTIFY:
            source = "Spotify"
            now = self.spotify_now
            if now is None or not now.title:
                self.title_label.setText("Spotify")
                self.artist_label.setText("Pick a playlist on the Spotify page, or press Play in "
                                          "Spotify.")
                self.info_label.setText("")
            else:
                self.title_label.setText(now.title)
                self.artist_label.setText(now.artist)
                self.info_label.setText(now.album)
                if now.picture:
                    picture = QPixmap()
                    if not picture.loadFromData(now.picture):
                        picture = None
            placeholder = "♫"
        else:
            track = self.model.playing
            placeholder = "♫"
            if track is None:
                source = ""
                self.title_label.setText("Nothing playing")
                self.artist_label.setText("Open a folder of music, pick a radio station, or play "
                                          "a Spotify playlist.")
                self.info_label.setText("")
            else:
                source = "My Music"
                self.title_label.setText(track.name)
                self.artist_label.setText(track.artist or track.album_artist)
                details = [track.album] if track.album else []
                if track.year:
                    details.append(track.year)
                row = self.playlist.find(track)
                if row is not None:
                    details.append(f"song {row + 1:,} of {len(self.playlist.tracks):,}")
                self.info_label.setText("   ·   ".join(details))
                data = engine.cover_art(track.path) if getattr(self, "cover_for", None) != track.path \
                    else getattr(self, "cover_data", None)
                self.cover_for, self.cover_data = track.path, data
                if data:
                    picture = QPixmap()
                    if not picture.loadFromData(data):
                        picture = None
        self.source_label.setText(source)
        if picture is None or picture.isNull():
            key = ("placeholder", placeholder)
            if key != getattr(self, "cover_key", None):
                picture = placeholder_picture(self.cover_size, placeholder)
        else:
            if self.mode == MUSIC:
                key = ("file", self.cover_for)
            elif self.mode == SPOTIFY:
                key = ("spotify", id(self.spotify_now.picture))
            else:
                key = picture.cacheKey()
        if key != getattr(self, "cover_key", None):
            self.cover_key = key
            self.cover.setPixmap(picture.scaled(self.cover_size, self.cover_size, Qt.AspectRatioMode.KeepAspectRatio,
                                                Qt.TransformationMode.SmoothTransformation))
        for label in (self.title_label, self.artist_label, self.info_label):
            label.setToolTip(label.text())
        name = self.title_label.text() if source else ""
        self.setWindowTitle(f"{name} – {TITLE}" if name else TITLE)
        self.update_position()

    def radio_state_text(self):
        if self.radio_buffer is None:
            return "Stopped. Press Play to listen again."
        if self.radio_waiting:
            return self.radio_message or "Connecting..."
        behind = (self.radio_buffer.written - self.radio_heard()) / self.radio_buffer.rate
        if not self.radio_output.playing:
            state = "Paused"
        elif self.radio_output.loop:
            state = "Repeating the marked part"
        elif behind < LIVE_SLACK_SECONDS:
            state = "Live"
        else:
            state = f"{format_time(behind)} behind live"
        if self.radio_message and self.recorder is None:
            state += f"   ·   {self.radio_message}"
        return state

    def update_position(self):
        if self.mode == RADIO:
            buffer = self.radio_buffer
            if buffer is None or buffer.written == 0:
                self.seek.setRange(0, 0)
                self.time_label.setText("")
                self.length_label.setText("")
                return
            rate = buffer.rate
            heard = self.radio_heard()
            if not self.seek.isSliderDown():
                self.seek.setRange(buffer.oldest * 1000 // rate, buffer.written * 1000 // rate)
                self.seek.setValue(int(heard * 1000 // rate))
            behind = (buffer.written - heard) / rate
            self.time_label.setText(self.frame_clock(heard))
            self.length_label.setText("Live" if behind < LIVE_SLACK_SECONDS and self.radio_output.playing
                                      else f"-{format_time(behind)}")
            return
        if self.mode == SPOTIFY:
            now = self.spotify_now
            position, length = (now.position_now(), now.length) if now else (0, 0)
        else:
            position, length = self.player.position() / 1000, self.player.duration() / 1000
        if not self.seek.isSliderDown():
            self.seek.setRange(0, int(max(0, length) * 1000))
            self.seek.setValue(int(position * 1000))
        self.time_label.setText(format_time(position))
        self.length_label.setText(format_time(length))

    def update_buttons(self):
        playing = self.is_playing()
        icon = QStyle.StandardPixmap.SP_MediaPause if playing else QStyle.StandardPixmap.SP_MediaPlay
        self.play_button.setIcon(self.style().standardIcon(icon))
        self.play_button.setText("Pause" if playing else "Play")
        self.repeat_part_button.setChecked(bool(self.radio_output.loop))
        self.update_sleep_button()

    # ========================================================= the timer
    def sleep_dialog(self):
        dialog = SleepDialog(self, self.timer, self.sleep_minutes, self.sleep_action)
        dialog.exec()
        if not dialog.choice:
            return
        what, minutes, action = dialog.choice
        if what == "off":
            self.cancel_sleep()
            self.message("The sleep timer is off.")
            return
        self.sleep_minutes, self.sleep_action = minutes, action
        self.start_sleep(minutes, action)

    def song_end_problem(self):
        """Why the end of this song can't be waited for now, or None if it can.
        Waiting with nothing playing would let the timer go off at the end of
        whatever is played next, maybe the next day."""
        nothing = ("Nothing is playing right now, so there's no song end to wait for. "
                   "Press Play first, or pick a time instead.")
        if self.mode == RADIO:
            if self.radio_buffer is None or not (self.radio_output.playing or self.radio_waiting):
                return nothing
            if self.radio_buffer.song_at(self.radio_heard()) is None:
                return ("This station hasn't said which song is playing, so the end of the song "
                        "can't be told. Pick a time instead.")
            return None
        return None if self.is_playing() else nothing

    def start_sleep(self, minutes, action):
        if minutes == 0:
            problem = self.song_end_problem()
            if problem:
                QMessageBox.information(self, TITLE, problem)
                return
        self.restore_spotify_volume()
        if minutes == 0:
            self.timer.start_at_song_end(action)
        else:
            self.timer.start(minutes, action)
        self.message(self.timer.describe() + ".")
        self.update_sleep_button()

    def cancel_sleep(self):
        self.timer.cancel()
        self.restore_spotify_volume()
        self.apply_volume()
        self.update_sleep_button()

    def restore_spotify_volume(self):
        if self.spotify_base_volume is not None:
            self.spotify_volume.set(self.spotify_base_volume)
            self.spotify_base_volume = None

    def update_sleep_button(self):
        if self.timer.active:
            left = self.timer.remaining()
            text = "Sleep: after song" if left is None else f"Sleep: {format_time(left + 0.999)}"
            self.sleep_button.setText(text)
            self.sleep_button.setToolTip(self.timer.describe() + ". Click to change or turn off.")
            self.sleep_button.setStyleSheet("QToolButton { font-weight: bold; }")
        else:
            self.sleep_button.setText("Sleep timer")
            self.sleep_button.setToolTip("Stop the music (or more) after a while")
            self.sleep_button.setStyleSheet("")

    def sleep_now(self):
        """Time's up: stop everything, then do what was asked."""
        action = self.timer.action
        self.timer.cancel()
        self.player.pause() if self.mode == MUSIC else None
        if self.mode == RADIO:
            # Stopped, not paused: a paused station goes on coming in, all night.
            self.stop_radio()
            self.show_now_playing()
        if self.mode == SPOTIFY:
            self.remote.command("pause")
            QTimer.singleShot(1500, self.restore_spotify_volume)
        self.apply_volume()
        self.update_buttons()
        self.update_sleep_button()
        if action == engine.CLOSE_PLAYER:
            self.close()
        elif action in (engine.SLEEP_PC, engine.SHUT_DOWN_PC):
            self.warning = WarningDialog(self, action, engine.WARNING_SECONDS,
                                         lambda: self.do_system_action(action))
            self.warning.show()
            self.warning.raise_()
            self.warning.activateWindow()
        else:
            self.message("The sleep timer stopped the music.", 0)

    def do_system_action(self, action):
        self.save_settings()
        try:
            self.system_actions[action]()
        except Exception as error:
            QMessageBox.warning(self, TITLE, f"Windows didn't allow it: {error}")

    # ========================================================== every tick
    def tick(self):
        if self.timer.active:
            if self.timer.due():
                self.sleep_now()
            else:
                if self.timer.volume_factor() < 1:
                    self.apply_volume()
                self.update_sleep_button()
        if self.mode == RADIO and self.radio_buffer is not None:
            buffer = self.radio_buffer
            if self.radio_waiting and buffer.written >= PREBUFFER_SECONDS * buffer.rate:
                self.radio_waiting = False
                self.radio_start(0)
                self.update_buttons()
            song = buffer.song_at(self.radio_heard())
            name = song[1] if song else None
            if name != self.radio_song:
                if self.radio_song is not None and name is not None and self.timer.at_song_end:
                    self.sleep_now()
                self.radio_song = name
            self.show_now_playing()
            self.show_marks()
        elif self.mode == SPOTIFY:
            self.update_position()

    def come_forward(self, paths):
        """Started again (songs dragged onto Run_MusicPlayer.bat, say) while
        already open: show this window, playing what was given."""
        if self.isMinimized():
            self.showNormal()
        self.raise_()
        self.activateWindow()
        if paths:
            self.open_paths(paths)

    def speakers_changed(self):
        """Headphones plugged in or out: play through whatever Windows uses now."""
        device = QMediaDevices.defaultAudioOutput()
        if device.id() == self.speaker_id:
            return
        self.speaker_id = device.id()
        self.audio.setDevice(device)
        if self.radio_output.playing:
            try:
                self.radio_output.play(self.radio_output.heard())
            except RuntimeError as error:
                self.message(str(error))

    # ======================================================= drag and drop
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.add_paths(paths, replace=False, play=not self.is_playing())
            for path in paths:
                if os.path.isdir(path):
                    self.recent.add(path)
            self.show_recent()


def settings_store():
    # An .ini file in %APPDATA%\Music Player, rather than the registry, so it is easy to find.
    return QSettings(QSettings.Format.IniFormat, QSettings.Scope.UserScope, TITLE, "settings")


# One window at a time: a second start hands its songs to the first and stops.
INSTANCE_NAME = "MusicPlayer-" + re.sub(r"\W", "", os.environ.get("USERNAME", "user"))


def hand_over(paths, name=INSTANCE_NAME):
    """If the Music Player is already open, give it these paths and bring it
    forward. Returns True if it was open."""
    socket = QLocalSocket()
    socket.connectToServer(name)
    if not socket.waitForConnected(500):
        return False
    socket.write(json.dumps([os.path.abspath(p) for p in paths]).encode("utf-8"))
    socket.flush()
    socket.waitForBytesWritten(2000)
    socket.disconnectFromServer()
    if socket.state() != QLocalSocket.LocalSocketState.UnconnectedState:
        socket.waitForDisconnected(1000)
    return True


def listen_for_others(window, name=INSTANCE_NAME):
    """Wait for later starts to hand over their songs."""
    server = QLocalServer(window)
    if not server.listen(name):
        return None

    def arrived():
        while server.hasPendingConnections():
            socket = server.nextPendingConnection()
            chunks = []
            finished = []

            def read(s=socket, c=chunks):
                c.append(bytes(s.readAll()))

            def done(s=socket, c=chunks, f=finished):
                if f:
                    return
                f.append(True)
                read(s, c)
                try:
                    paths = json.loads(b"".join(c).decode("utf-8") or "[]")
                except ValueError:
                    paths = []
                window.come_forward([p for p in paths if isinstance(p, str)]
                                    if isinstance(paths, list) else [])
                s.deleteLater()

            socket.readyRead.connect(read)
            socket.disconnected.connect(done)
            if socket.state() == QLocalSocket.LocalSocketState.UnconnectedState:
                done()
    server.newConnection.connect(arrived)
    return server


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(TITLE)
    app.setFont(QFont("Segoe UI", 10))
    paths = sys.argv[1:]
    if hand_over(paths):
        return
    window = MusicWindow(settings_store())
    window.instance_server = listen_for_others(window)
    window.show()
    if paths:
        window.open_paths(paths)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
