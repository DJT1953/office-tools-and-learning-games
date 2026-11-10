"""
music_player_engine.py

The Music Player's rules for music on this PC, kept free of Qt so they can be
tested on their own: finding music in folders and subfolders, reading song
names and cover pictures, the playing order (shuffle, repeat, queue), the last
folders picked, playlist files and the sleep timer.
"""

import base64
import json
import os
import random
import re
import struct
import subprocess
import threading
import time
from dataclasses import asdict, dataclass, fields

MUSIC_EXTENSIONS = {
    ".aac", ".aif", ".aifc", ".aiff", ".ape", ".flac", ".m4a", ".m4b", ".mka",
    ".mp2", ".mp3", ".mpc", ".oga", ".ogg", ".opus", ".wav", ".wma", ".wv",
}
PLAYLIST_EXTENSIONS = {".m3u", ".m3u8", ".pls"}
PICTURE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}
COVER_NAMES = ("cover", "folder", "front", "albumart", "album")
SKIP_FOLDERS = {"$recycle.bin", "system volume information"}
RECENT_FOLDER_COUNT = 4
BIGGEST_PICTURE = 20 * 1024 * 1024

REPEAT_OFF, REPEAT_ALL, REPEAT_ONE = "off", "all", "one"
REPEAT_MODES = (REPEAT_OFF, REPEAT_ALL, REPEAT_ONE)


# ------------------------------------------------------------ small helpers

def natural_key(text):
    """Sort "Track 2" before "Track 10", ignoring case."""
    return [int(part) if part.isdigit() else part.casefold()
            for part in re.split(r"(\d+)", text)]


def same_path(a, b):
    return os.path.normcase(os.path.normpath(a)) == os.path.normcase(os.path.normpath(b))


def is_music(path):
    name = os.path.basename(path)
    return not name.startswith("._") and os.path.splitext(name)[1].lower() in MUSIC_EXTENSIONS


def is_playlist(path):
    return os.path.splitext(path)[1].lower() in PLAYLIST_EXTENSIONS


def is_web_address(text):
    return re.match(r"(?i)^(https?|icy|mms|rtmp|rtsp)://", text.strip()) is not None


def format_time(seconds):
    """3:05, or 1:02:03 past an hour."""
    if seconds is None or seconds != seconds or seconds < 0:   # None, NaN or negative
        seconds = 0
    seconds = int(seconds)
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    return f"{hours}:{minutes:02d}:{secs:02d}" if hours else f"{minutes}:{secs:02d}"


# ------------------------------------------------------------ finding music

FILE_ATTRIBUTE_HIDDEN = 0x2


def is_hidden_folder(entry):
    """Folders Windows hides, such as AppData: Explorer doesn't show them, and
    the sounds in them are programs' own noises, not music."""
    try:
        attributes = entry.stat(follow_symlinks=False).st_file_attributes
    except (OSError, AttributeError):         # gone, or not Windows
        return False
    return bool(attributes & FILE_ATTRIBUTE_HIDDEN)


def find_music(folder, include_subfolders=True, stop=lambda: False):
    """Every music file in a folder, then in each subfolder in turn, in name
    order, so albums come out in track order."""
    found = []
    waiting = [folder]
    seen = set()          # folders read already: a folder link can lead back up the tree
    while waiting:
        if stop():
            break
        current = waiting.pop()
        try:
            info = os.stat(current)
            if info.st_ino:                   # some drives don't number their folders
                if (info.st_dev, info.st_ino) in seen:
                    continue
                seen.add((info.st_dev, info.st_ino))
            entries = list(os.scandir(current))
        except OSError:                       # no permission, or the drive went away
            continue
        files, folders = [], []
        for entry in entries:
            try:
                if entry.is_file() and is_music(entry.name):
                    files.append(entry)
                elif (include_subfolders and entry.is_dir(follow_symlinks=False)
                      and entry.name.lower() not in SKIP_FOLDERS
                      and not entry.name.startswith(".")
                      and not is_hidden_folder(entry)):
                    folders.append(entry)
            except OSError:
                continue
        found.extend(e.path for e in sorted(files, key=lambda e: natural_key(e.name)))
        # The stack is last-in first-out, so push subfolders in reverse.
        waiting.extend(e.path for e in sorted(folders, key=lambda e: natural_key(e.name),
                                              reverse=True))
    return found


def expand_paths(paths, stop=lambda: False, depth=0):
    """Turn picked or dropped files, folders and playlists into music files.
    Returns (music files, web addresses found in playlists). Playlists may
    list folders and other playlists (a few levels deep, so two playlists
    naming each other can't go round for ever)."""
    music, addresses = [], []
    for path in paths:
        if stop():
            break
        if is_web_address(path):
            addresses.append(path)
        elif os.path.isdir(path):
            music.extend(find_music(path, True, stop))
        elif is_playlist(path):
            if depth >= 3:
                continue
            try:
                entries = read_playlist(path)
            except OSError:                   # gone, or no permission
                entries = []
            found, more = expand_paths(entries, stop, depth + 1)
            music.extend(found)
            addresses.extend(more)
        elif is_music(path):
            music.append(path)
    return music, addresses


# ---------------------------------------------------------- song details

@dataclass(eq=False)          # compared by identity: the same file can be in a list twice
class Track:
    path: str
    title: str = ""
    artist: str = ""
    album: str = ""
    album_artist: str = ""
    genre: str = ""
    year: str = ""
    track: int = 0
    disc: int = 0
    length: float = 0.0       # seconds; 0 when not known

    @property
    def name(self):
        return self.title or os.path.splitext(os.path.basename(self.path))[0]

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, data):
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in data.items() if k in known})

    def copy(self):
        return Track.from_dict(self.to_dict())


# Tag names differ by file type: plain names for FLAC, Ogg and APE, ID3 frames
# for MP3/WAV/AIFF, iTunes atoms for M4A and Windows names for WMA.
TAG_NAMES = {
    "title": ("title", "TIT2", "\xa9nam", "Title"),
    "artist": ("artist", "TPE1", "\xa9ART", "Author", "Artist"),
    "album": ("album", "TALB", "\xa9alb", "WM/AlbumTitle", "Album"),
    "album_artist": ("albumartist", "album artist", "TPE2", "aART", "WM/AlbumArtist"),
    "genre": ("genre", "TCON", "\xa9gen", "WM/Genre", "Genre"),
    "year": ("date", "year", "TDRC", "TYER", "\xa9day", "WM/Year", "Year"),
    "track": ("tracknumber", "TRCK", "trkn", "WM/TrackNumber", "Track"),
    "disc": ("discnumber", "TPOS", "disk", "WM/PartOfSet", "Disc"),
}


def tag_text(value):
    """The first piece of text in a tag, whatever kind of file it came from."""
    if isinstance(value, (list, tuple)):
        if not value:
            return ""
        value = value[0]
    if isinstance(value, tuple):                      # M4A track numbers: (3, 12)
        return str(value[0]) if value and value[0] else ""
    for attribute in ("text", "value"):               # ID3 frames, WMA attributes
        inner = getattr(value, attribute, None)
        if inner is not None and inner is not value:
            return tag_text(inner)
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace").strip("\x00 ")
    return str(value).strip("\x00 ").strip()


def leading_number(text):
    match = re.match(r"\s*(\d+)", text or "")
    return int(match.group(1)) if match else 0


def guess_from_file_name(track):
    """Songs without tags: "03 - Artist - Song.mp3" or "Artist - Song.mp3"."""
    stem = os.path.splitext(os.path.basename(track.path))[0]
    match = re.match(r"^\s*(\d{1,3})\s*[-._)]?\s+(.*)$", stem)
    if match and match.group(2):
        if not track.track:
            track.track = int(match.group(1))
        stem = match.group(2)
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]
    if len(parts) >= 2 and not track.artist:
        track.artist = parts[0]
        track.title = " - ".join(parts[1:])
    else:
        track.title = stem.strip() or stem


RIFF_INFO = {b"INAM": "title", b"IART": "artist", b"IPRD": "album", b"IGNR": "genre",
             b"ICRD": "year", b"ITRK": "track", b"IPRT": "track"}


def wav_info_tags(path):
    """WAV files usually keep their tags in a RIFF "LIST INFO" block, which
    the tag reader doesn't look at."""
    found = {}
    try:
        with open(path, "rb") as f:
            header = f.read(12)
            if header[:4] != b"RIFF" or header[8:12] != b"WAVE":
                return found
            while True:
                chunk = f.read(8)
                if len(chunk) < 8:
                    break
                kind, size = chunk[:4], struct.unpack("<I", chunk[4:])[0]
                if kind == b"LIST" and size < 1 << 20:
                    body = f.read(size)
                    if body[:4] == b"INFO":
                        i = 4
                        while i + 8 <= len(body):
                            name = body[i:i + 4]
                            length = struct.unpack("<I", body[i + 4:i + 8])[0]
                            text = body[i + 8:i + 8 + length].split(b"\x00")[0]
                            if name in RIFF_INFO and text:
                                for encoding in ("utf-8", "cp1252"):
                                    try:
                                        found.setdefault(RIFF_INFO[name], text.decode(encoding).strip())
                                        break
                                    except UnicodeDecodeError:
                                        continue
                            i += 8 + length + (length & 1)
                    if size & 1:
                        f.read(1)
                else:
                    f.seek(size + (size & 1), 1)
    except (OSError, struct.error):
        pass
    return found


def read_track(path):
    """A song's name, artist, album and length from its tags."""
    track = Track(path)
    try:
        import mutagen
        audio = mutagen.File(path)
    except Exception:                       # unreadable or damaged: show the file name
        audio = None
    if audio is not None:
        track.length = float(getattr(getattr(audio, "info", None), "length", 0) or 0)
        tags = audio.tags
        if tags is not None:
            for field_name, names in TAG_NAMES.items():
                for name in names:
                    try:
                        value = tags.get(name)
                    except Exception:
                        value = None
                    text = tag_text(value) if value is not None else ""
                    if text:
                        if field_name in ("track", "disc"):
                            setattr(track, field_name, leading_number(text))
                        elif field_name == "year":
                            track.year = text[:4]
                        else:
                            setattr(track, field_name, text)
                        break
    if path.lower().endswith(".wav"):
        for field_name, text in wav_info_tags(path).items():
            if not getattr(track, field_name):
                if field_name == "track":
                    track.track = leading_number(text)
                else:
                    setattr(track, field_name, text[:4] if field_name == "year" else text)
    if not track.title:
        guess_from_file_name(track)
    return track


class TagCache:
    """Song details remembered between runs, so a big folder opens quickly.
    An entry counts only while the file's size and date are unchanged."""

    def __init__(self, path=None):
        self.path = path
        self.entries = {}
        self.changed = False
        self.lock = threading.Lock()      # folders are read on a thread of their own
        if path and os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    self.entries = json.load(f)
            except (OSError, ValueError):
                self.entries = {}
            if not isinstance(self.entries, dict):
                self.entries = {}

    @staticmethod
    def _stamp(path):
        info = os.stat(path)
        return [info.st_size, int(info.st_mtime)]

    def track(self, path):
        try:
            stamp = self._stamp(path)
        except OSError:
            return read_track(path)
        key = os.path.normcase(path)
        with self.lock:
            entry = self.entries.get(key)
        if isinstance(entry, dict) and entry.get("stamp") == stamp:
            try:
                track = Track.from_dict(entry["track"])
                track.path = path
                return track
            except (KeyError, TypeError, AttributeError):
                pass                          # damaged entry: read the file again
        track = read_track(path)
        with self.lock:
            self.entries[key] = {"stamp": stamp, "track": track.to_dict()}
            self.changed = True
        return track

    def save(self):
        if not self.path or not self.changed:
            return
        with self.lock:
            # Plain ASCII, so even a file name Windows lets through half-made
            # (an odd character from another program) can be written.
            text = json.dumps(self.entries)
            self.changed = False
        temporary = self.path + ".saving"
        with open(temporary, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(temporary, self.path)


# ---------------------------------------------------------- cover pictures

def asf_picture(data):
    """A WMA picture: type, size, then a picture type and description in UTF-16."""
    try:
        size = struct.unpack_from("<I", data, 1)[0]
        position = 5
        for _ in range(2):                                  # mime type, description
            end = data.index(b"\x00\x00", position)
            while (end - position) % 2:
                end = data.index(b"\x00\x00", end + 1)
            position = end + 2
        return data[position:position + size] or None
    except (ValueError, struct.error):
        return None


def embedded_picture(path):
    """The cover picture stored inside a music file, if any."""
    try:
        import mutagen
        audio = mutagen.File(path)
    except Exception:
        return None
    if audio is None:
        return None

    def front_first(pictures):
        pictures = [p for p in pictures if getattr(p, "data", None)]
        pictures.sort(key=lambda p: getattr(p, "type", 0) != 3)   # 3 is the front cover
        return pictures[0].data if pictures else None

    found = front_first(getattr(audio, "pictures", []) or [])      # FLAC
    if found:
        return found
    tags = audio.tags
    if tags is None:
        return None
    try:
        if hasattr(tags, "getall"):                                 # ID3: MP3, WAV, AIFF
            found = front_first(tags.getall("APIC"))
            if found:
                return found
        covers = tags.get("covr")                                   # M4A
        if covers:
            return bytes(covers[0])
        blocks = tags.get("metadata_block_picture")                 # Ogg Vorbis and Opus
        if blocks:
            from mutagen.flac import Picture
            return front_first([Picture(base64.b64decode(b)) for b in blocks])
        pictures = tags.get("WM/Picture")                           # WMA
        if pictures:
            return asf_picture(bytes(pictures[0].value))
        for key in ("Cover Art (Front)", "Cover Art (Back)"):       # APE, WavPack, Musepack
            item = tags.get(key)
            if item is not None:
                data = bytes(item.value)
                return data[data.index(b"\x00") + 1:] if b"\x00" in data else data
    except Exception:
        return None
    return None


def folder_picture(folder):
    """cover.jpg, folder.jpg and the like beside the music, or else the only
    picture there."""
    try:
        names = [n for n in os.listdir(folder)
                 if os.path.splitext(n)[1].lower() in PICTURE_EXTENSIONS]
    except OSError:
        return None
    names.sort(key=natural_key)
    chosen = None
    for wanted in COVER_NAMES:
        chosen = next((n for n in names if os.path.splitext(n)[0].lower().startswith(wanted)), None)
        if chosen:
            break
    if chosen is None and len(names) == 1:
        chosen = names[0]
    if chosen is None:
        return None
    path = os.path.join(folder, chosen)
    try:
        if os.path.getsize(path) > BIGGEST_PICTURE:
            return None
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


def cover_art(path):
    return embedded_picture(path) or folder_picture(os.path.dirname(path))


# ------------------------------------------------------- the playing order

class PlayList:
    """The songs in the list, which one is playing, and which comes next.

    Next plays any queued songs first, then carries on down the list, or in
    shuffled order. The "anchor" is the last song played from the list itself
    (not from the queue), so the list carries on from there afterwards."""

    def __init__(self, rng=None):
        self.rng = rng or random.Random()
        self.tracks = []
        self.shuffled = []
        self.queue = []
        self.current = None
        self.anchor = None
        self.anchor_pos = 0       # where the anchor was, in case it is removed
        self.shuffle = False
        self.repeat = REPEAT_OFF

    # --- what's in the list
    def __len__(self):
        return len(self.tracks)

    def find(self, track, sequence=None):
        sequence = self.tracks if sequence is None else sequence
        for i, t in enumerate(sequence):
            if t is track:
                return i
        return None

    def sequence(self):
        return self.shuffled if self.shuffle else self.tracks

    def replace(self, tracks):
        self.tracks = list(tracks)
        self.queue = []
        self.current = self.anchor = None
        self.anchor_pos = 0
        self.shuffled = list(self.tracks)
        self.rng.shuffle(self.shuffled)

    def add(self, tracks):
        tracks = list(tracks)
        self.tracks.extend(tracks)
        played = (self.find(self.anchor, self.shuffled) if self.anchor is not None else None)
        first_free = 0 if played is None else played + 1
        for track in tracks:      # new songs go somewhere among those not yet played
            self.shuffled.insert(self.rng.randint(first_free, len(self.shuffled)), track)

    def remove(self, tracks):
        gone = {id(t) for t in tracks}
        sequence = self.sequence()
        if self.anchor is not None and id(self.anchor) in gone:
            # The song after it moves up into its place, so it plays next.
            position = self.find(self.anchor, sequence)
            if position is not None:
                self.anchor_pos = sum(1 for t in sequence[:position] if id(t) not in gone)
            self.anchor = None
        elif self.anchor is None:
            self.anchor_pos -= sum(1 for t in sequence[:self.anchor_pos] if id(t) in gone)
        self.tracks = [t for t in self.tracks if id(t) not in gone]
        self.shuffled = [t for t in self.shuffled if id(t) not in gone]
        self.queue = [t for t in self.queue if id(t) not in gone]

    def sort(self, key, descending=False):
        """Sort the list by "title", "artist", "album", "length", "number" or
        "path". Albums sort by disc and track number within them."""
        def text(value):
            return natural_key(value or "￿")    # blanks last
        keys = {
            "title": lambda t: (text(t.name), text(t.artist)),
            "artist": lambda t: (text(t.artist or t.album_artist), text(t.album),
                                 t.disc, t.track, text(t.name)),
            "album": lambda t: (text(t.album), t.disc, t.track, text(t.name)),
            "length": lambda t: t.length,
            "number": lambda t: (t.disc, t.track),
            "path": lambda t: natural_key(t.path),
        }
        self.tracks.sort(key=keys[key], reverse=descending)

    # --- moving through it
    def _target(self, step, peek):
        sequence = self.sequence()
        if not sequence:
            return None
        position = self.find(self.anchor, sequence) if self.anchor is not None else None
        if position is not None:
            target = position + step
        else:
            target = self.anchor_pos if step > 0 else self.anchor_pos - 1
        if 0 <= target < len(sequence):
            return target
        if self.repeat == REPEAT_OFF:
            return None
        if step > 0:
            if self.shuffle and not peek:
                self._reshuffle()
            return 0
        return len(sequence) - 1

    def _reshuffle(self):
        """A new random order for the next time round, not starting with the
        song that just finished."""
        last = self.current
        self.rng.shuffle(self.shuffled)
        if len(self.shuffled) > 1 and self.shuffled[0] is last:
            swap = self.rng.randint(1, len(self.shuffled) - 1)
            self.shuffled[0], self.shuffled[swap] = self.shuffled[swap], self.shuffled[0]

    def _go(self, position):
        self.anchor = self.current = self.sequence()[position]
        self.anchor_pos = position
        return self.current

    def choose(self, track):
        """Play this song now (double-clicked, say); the list carries on after it."""
        if self.shuffle and track is not self.anchor:
            # Played songs stay before it; the rest of the shuffle follows.
            if self.find(track, self.shuffled) is not None:
                self.shuffled.remove(track)
            played = self.find(self.anchor, self.shuffled) if self.anchor is not None else None
            self.shuffled.insert(0 if played is None else played + 1, track)
        position = self.find(track, self.sequence())
        if position is None:
            return None
        return self._go(position)

    def next(self, automatic=False):
        """The next song, or None at the end. automatic means the last song
        just finished by itself, which is when Repeat one plays it again."""
        if automatic and self.repeat == REPEAT_ONE and self.current is not None:
            return self.current
        if self.queue:
            self.current = self.queue.pop(0)
            return self.current
        target = self._target(+1, peek=False)
        return None if target is None else self._go(target)

    def previous(self):
        if self.current is not None and self.current is not self.anchor and self.anchor is not None:
            self.current = self.anchor     # leave a queued song: back to where the list was
            return self.current
        target = self._target(-1, peek=False)
        return None if target is None else self._go(target)

    def has_next(self):
        return bool(self.queue) or self._target(+1, peek=True) is not None

    def has_previous(self):
        if self.current is not None and self.current is not self.anchor and self.anchor is not None:
            return True
        return self._target(-1, peek=True) is not None

    # --- options
    def set_shuffle(self, on):
        if on == self.shuffle:
            return
        self.shuffle = on
        if on:
            others = [t for t in self.tracks if t is not self.anchor]
            self.rng.shuffle(others)
            self.shuffled = ([self.anchor] if self.anchor is not None
                             and self.find(self.anchor) is not None else []) + others
            self.anchor_pos = 0
        elif self.anchor is not None:
            position = self.find(self.anchor)
            self.anchor_pos = position if position is not None else 0

    def cycle_repeat(self):
        self.repeat = REPEAT_MODES[(REPEAT_MODES.index(self.repeat) + 1) % len(REPEAT_MODES)]
        return self.repeat

    # --- the queue
    def play_next(self, tracks):
        self.queue[0:0] = list(tracks)

    def add_to_queue(self, tracks):
        self.queue.extend(tracks)


# --------------------------------------------------------- recent folders

class RecentFolders:
    """The last few folders picked, newest first. Picking one again moves it
    to the top; Remove takes one off."""

    def __init__(self, folders=(), limit=RECENT_FOLDER_COUNT):
        self.limit = limit
        self.folders = []
        for folder in reversed(list(folders)):
            self.add(folder)

    def add(self, folder):
        self.folders = [f for f in self.folders if not same_path(f, folder)]
        self.folders.insert(0, os.path.normpath(folder))
        del self.folders[self.limit:]

    def remove(self, folder):
        self.folders = [f for f in self.folders if not same_path(f, folder)]

    def __iter__(self):
        return iter(self.folders)

    def __len__(self):
        return len(self.folders)


# ---------------------------------------------------------- playlist files

def _read_text(path):
    with open(path, "rb") as f:
        raw = f.read()
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def _entry_path(entry, folder):
    """The whole path for one playlist entry, or None for an empty one."""
    entry = entry.strip().strip('"').strip()
    if not entry:              # made whole it would be the playlist's own folder
        return None
    if is_web_address(entry):
        return entry
    if entry.lower().startswith("file:"):
        from urllib.parse import unquote, urlparse
        parsed = urlparse(entry)
        entry = unquote(parsed.path)
        if re.match(r"^/[A-Za-z]:", entry):
            entry = entry[1:]
        if parsed.netloc:                                   # file://server/share
            entry = "//" + parsed.netloc + entry
    entry = entry.replace("/", os.sep) if os.sep == "\\" else entry
    if not os.path.isabs(entry):
        entry = os.path.join(folder, entry)
    return os.path.normpath(entry)


def read_playlist(path):
    """The entries in an .m3u, .m3u8 or .pls file: file paths (made whole)
    and web addresses, in order."""
    text = _read_text(path)
    folder = os.path.dirname(os.path.abspath(path))
    if path.lower().endswith(".pls") or text.lstrip().lower().startswith("[playlist]"):
        numbered = []
        for line in text.splitlines():
            match = re.match(r"\s*file(\d+)\s*=\s*(.+)", line, re.IGNORECASE)
            if match:
                numbered.append((int(match.group(1)), match.group(2)))
        entries = [_entry_path(entry, folder) for _, entry in sorted(numbered)]
    else:
        entries = [_entry_path(line, folder) for line in text.splitlines()
                   if line.strip() and not line.strip().startswith("#")]
    return [entry for entry in entries if entry]


def write_playlist(path, tracks):
    """Save the list as an .m3u8 (or .m3u) playlist that other players can open
    too. Songs inside the playlist's folder are written relative to it, so the
    folder can be moved together."""
    folder = os.path.dirname(os.path.abspath(path))
    lines = ["#EXTM3U"]
    for track in tracks:
        seconds = int(round(track.length)) if track.length else -1
        label = f"{track.artist} - {track.name}" if track.artist else track.name
        label = re.sub(r"[\r\n]+", " ", label)     # a song name on two lines would split the entry
        lines.append(f"#EXTINF:{seconds},{label}")
        entry = os.path.abspath(track.path)
        try:
            relative = os.path.relpath(entry, folder)
            if not relative.startswith(("..", "#")):
                entry = relative
        except ValueError:                     # another drive
            pass
        lines.append(entry)
    temporary = path + ".saving"
    # errors="replace": a half-made file name loses one character, not the whole list.
    with open(temporary, "w", encoding="utf-8", errors="replace", newline="\r\n") as f:
        f.write("\n".join(lines) + "\n")
    os.replace(temporary, path)


# ------------------------------------------------------------ sleep timer

STOP_MUSIC, CLOSE_PLAYER, SLEEP_PC, SHUT_DOWN_PC = "stop", "close", "sleep", "shutdown"
SLEEP_ACTIONS = {
    STOP_MUSIC: "Stop the music",
    CLOSE_PLAYER: "Close the Music Player",
    SLEEP_PC: "Put the computer to sleep",
    SHUT_DOWN_PC: "Shut down the computer",
}
FADE_SECONDS = 60
WARNING_SECONDS = 30          # before the computer sleeps or shuts down


class SleepTimer:
    """Counts down to when the music stops, fading it out over the last
    minute, or waits for the end of the song playing."""

    def __init__(self, clock=time.monotonic):
        self.clock = clock
        self.cancel()

    def cancel(self):
        self.end = None
        self.length = 0
        self.at_song_end = False
        self.action = STOP_MUSIC

    def start(self, minutes, action=STOP_MUSIC):
        self.cancel()
        self.length = max(1.0, minutes * 60)
        self.end = self.clock() + self.length
        self.action = action

    def start_at_song_end(self, action=STOP_MUSIC):
        self.cancel()
        self.at_song_end = True
        self.action = action

    @property
    def active(self):
        return self.end is not None or self.at_song_end

    def remaining(self):
        """Seconds left, or None when not counting down."""
        if self.end is None:
            return None
        return max(0.0, self.end - self.clock())

    def due(self):
        left = self.remaining()
        return left is not None and left <= 0

    def volume_factor(self):
        """1 until the fade starts, then down to 0 when the time is up."""
        left = self.remaining()
        if left is None:
            return 1.0
        fade = min(FADE_SECONDS, self.length)
        return 1.0 if left >= fade else max(0.0, left / fade)

    def describe(self):
        what = {STOP_MUSIC: "Music stops", CLOSE_PLAYER: "Player closes",
                SLEEP_PC: "Computer sleeps", SHUT_DOWN_PC: "Computer shuts down"}[self.action]
        if self.at_song_end:
            return f"{what} after this song"
        if self.end is None:
            return ""
        return f"{what} in {format_time(self.remaining() + 0.999)}"


def shut_down_computer():
    subprocess.run(["shutdown", "/s", "/t", "0"],
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def sleep_means_hibernate():
    """True on computers with no ordinary sleep (only Modern Standby and
    hibernate, as on many newer laptops): there, asking Windows for sleep
    hibernates, which is off with everything kept."""
    try:
        import ctypes
        capabilities = (ctypes.c_ubyte * 256)()      # SYSTEM_POWER_CAPABILITIES
        if not ctypes.windll.powrprof.GetPwrCapabilities(ctypes.byref(capabilities)):
            return False
    except (AttributeError, OSError):                # not Windows
        return False
    # One byte each: power button, sleep button, lid, S1, S2, S3, S4, S5, hiberfile...
    s1, s2, s3, s4, hiberfile = (capabilities[i] for i in (3, 4, 5, 6, 8))
    return not (s1 or s2 or s3) and bool(s4 and hiberfile)


def put_computer_to_sleep():
    """Sleep (not hibernate), the same as Start > Power > Sleep."""
    import ctypes
    powrprof = ctypes.WinDLL("powrprof")
    if not powrprof.SetSuspendState(False, False, False):
        raise OSError("Windows did not let the Music Player put the computer to sleep.")
