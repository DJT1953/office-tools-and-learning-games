"""Song Cleaner: reading songs.

Decodes each MP3 once (through FFmpeg), and keeps three small things about its
sound: a fingerprint for telling whether two files are the same recording, how
loud it is moment by moment, and what the sound is like moment by moment (for
finding where talk ends and music starts). Results are remembered per file, so
a song is only ever read once. Nothing here changes a file, and nothing here
needs a window.
"""

import concurrent.futures
import dataclasses
import hashlib
import json
import os
import subprocess

import numpy as np

RATE = 11025
WINDOW = 4096
HOP = 1024
FRAME_S = HOP / RATE          # about 0.093 s between frames
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
CACHE_VERSION = 1
TEMP_PREFIX = "~song cleaner"     # files this app is in the middle of making

_FP_EDGES = np.geomspace(300.0, 2000.0, 34)    # 33 bands give 32 bits
_BAND_EDGES = np.geomspace(100.0, 5000.0, 17)  # 16 bands of "what it sounds like"


class SongError(Exception):
    """Something wrong with one song, said in plain words."""


def ffmpeg_exe():
    import imageio_ffmpeg
    return imageio_ffmpeg.get_ffmpeg_exe()


def app_folder():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "Song Cleaner")
    os.makedirs(path, exist_ok=True)
    return path


def decode_mono(path):
    """The whole song as one channel of float32 samples at 11,025 a second."""
    try:
        done = subprocess.run(
            [ffmpeg_exe(), "-v", "error", "-nostdin", "-i", path, "-vn", "-ac", "1",
             "-ar", str(RATE), "-f", "f32le", "-"],
            capture_output=True, creationflags=NO_WINDOW)
    except OSError as error:
        raise SongError(f"FFmpeg could not be started ({error}).") from error
    samples = np.frombuffer(done.stdout[:len(done.stdout) // 4 * 4], dtype=np.float32)
    if len(samples) < WINDOW * 2:
        raise SongError("This file could not be read as sound.")
    return samples


def _band_sums(power, freqs, edges):
    """Energy in each band between neighbouring edges; one row per frame."""
    cum = np.concatenate([np.zeros((power.shape[0], 1), dtype=np.float64),
                          np.cumsum(power, axis=1, dtype=np.float64)], axis=1)
    idx = np.searchsorted(freqs, edges)
    return cum[:, idx[1:]] - cum[:, idx[:-1]]


def analyse(samples):
    """(fingerprint, level, bands) for a song.

    fingerprint: one 32-bit word per frame. Each bit says whether the energy
    step between two neighbouring bands grew or shrank since the frame before,
    which survives re-encoding, a different volume and a different station.
    level: loudness of each frame in dB (0 = full scale).
    bands: 16 band levels per frame in half-dB steps (uint8).
    """
    samples = np.asarray(samples, dtype=np.float32)
    frames = 1 + (len(samples) - WINDOW) // HOP
    if frames < 2:
        raise SongError("This file is too short to be a song.")
    window = np.hanning(WINDOW).astype(np.float32)
    freqs = np.fft.rfftfreq(WINDOW, 1 / RATE)
    fp_energy = np.empty((frames, 33))
    bands = np.empty((frames, 16))
    level = np.empty(frames, dtype=np.float32)
    for first in range(0, frames, 1024):
        last = min(frames, first + 1024)
        idx = np.arange(WINDOW)[None, :] + HOP * np.arange(first, last)[:, None]
        chunk = samples[idx]
        level[first:last] = 10 * np.log10((chunk * chunk).mean(axis=1) + 1e-12)
        power = np.abs(np.fft.rfft(chunk * window, axis=1)) ** 2
        fp_energy[first:last] = _band_sums(power, freqs, _FP_EDGES)
        bands[first:last] = _band_sums(power, freqs, _BAND_EDGES)
    step = fp_energy[:, :-1] - fp_energy[:, 1:]
    bits = (step[1:] - step[:-1]) > 0
    bits = np.concatenate([bits[:1], bits])            # frame 0 copies frame 1
    weights = (1 << np.arange(32, dtype=np.uint64))
    fingerprint = (bits.astype(np.uint64) * weights).sum(axis=1).astype(np.uint32)
    band_db = 10 * np.log10(bands / WINDOW + 1e-12)
    return fingerprint, level, np.clip((band_db + 110) * 2, 0, 255).astype(np.uint8)


def fingerprint(samples):
    return analyse(samples)[0]


def sound_profile(samples):
    return analyse(samples)[1:]


@dataclasses.dataclass
class Song:
    path: str
    size: int = 0
    mtime: float = 0.0
    artist: str = ""
    title: str = ""
    album: str = ""
    bitrate: int = 0          # kbps
    channels: int = 0
    length: float = 0.0       # seconds
    fp: np.ndarray = None
    level: np.ndarray = None
    bands: np.ndarray = None
    error: str = ""

    @property
    def name(self):
        return os.path.basename(self.path)

    @property
    def stem(self):
        return os.path.splitext(self.name)[0]


def _text(tags, key):
    try:
        frame = tags.get(key)
        return str(frame.text[0]).strip() if frame is not None and frame.text else ""
    except Exception:  # noqa: BLE001 - an odd tag is the same as no tag
        return ""


def read_tags(song):
    """Fill artist, title, album, bitrate, channels and length from the file."""
    import mutagen
    import mutagen.mp3
    try:
        sound = mutagen.mp3.MP3(song.path)
    except Exception as error:  # noqa: BLE001 - mutagen raises many kinds
        raise SongError("This file is not an MP3 that can be read.") from error
    song.bitrate = int(round((sound.info.bitrate or 0) / 1000))
    song.channels = int(sound.info.channels or 0)
    song.length = float(sound.info.length or 0.0)
    if sound.tags is not None:
        song.artist = _text(sound.tags, "TPE1")
        song.title = _text(sound.tags, "TIT2")
        song.album = _text(sound.tags, "TALB")
    if not song.artist and not song.title:
        left, dash, right = song.stem.partition(" - ")
        song.artist, song.title = (left.strip(), right.strip()) if dash else ("", song.stem.strip())


def read_song(path):
    """Everything about one song. Trouble goes in song.error, never an exception."""
    song = Song(path=path)
    try:
        stat = os.stat(path)
        song.size, song.mtime = stat.st_size, stat.st_mtime
        read_tags(song)
        song.fp, song.level, song.bands = analyse(decode_mono(path))
        song.length = len(song.fp) * FRAME_S + (WINDOW - HOP) / RATE
    except SongError as error:
        song.error = str(error)
    except OSError as error:
        song.error = f"Windows would not let this file be read ({error.strerror or error})."
    except Exception as error:  # noqa: BLE001 - whatever is odd about one file, the scan goes on
        song.error = f"This file could not be read ({type(error).__name__})."
    return song


class Cache:
    """Remembers what was learnt about each song, until the file changes."""

    def __init__(self, folder=None):
        self.folder = folder or os.path.join(app_folder(), "songs")
        os.makedirs(self.folder, exist_ok=True)

    def _file(self, path):
        key = hashlib.sha1(os.path.normcase(os.path.abspath(path)).encode("utf-8")).hexdigest()
        return os.path.join(self.folder, key + ".npz")

    def get(self, path):
        try:
            stat = os.stat(path)
            with np.load(self._file(path), allow_pickle=False) as data:
                facts = json.loads(str(data["facts"]))
                if (facts.get("version") != CACHE_VERSION or facts["size"] != stat.st_size
                        or abs(facts["mtime"] - stat.st_mtime) > 0.001):
                    return None
                song = Song(path=path, **{k: facts[k] for k in (
                    "size", "mtime", "artist", "title", "album", "bitrate", "channels",
                    "length", "error")})
                if not song.error:
                    song.fp, song.level, song.bands = data["fp"], data["level"], data["bands"]
                return song
        except Exception:  # noqa: BLE001 - a missing or damaged note just means "read it again"
            return None

    def put(self, song):
        facts = {k: getattr(song, k) for k in (
            "size", "mtime", "artist", "title", "album", "bitrate", "channels", "length", "error")}
        facts["version"] = CACHE_VERSION
        arrays = {} if song.error else {"fp": song.fp, "level": song.level, "bands": song.bands}
        target = self._file(song.path)
        try:
            # np.savez adds ".npz" itself, so the temporary name must end that way
            temp = target[:-4] + ".part.npz"
            np.savez_compressed(temp, facts=np.array(json.dumps(facts)), **arrays)
            os.replace(temp, target)
        except OSError:
            pass                                        # no room to remember: read again next time


def list_songs(folder):
    """The MP3s directly in a folder, in name order."""
    try:
        names = os.listdir(folder)
    except OSError as error:
        raise SongError(f"That folder could not be opened ({error.strerror or error}).") from error
    paths = [os.path.join(folder, n) for n in names
             if n.lower().endswith(".mp3") and not n.lower().startswith(TEMP_PREFIX)]
    return sorted((p for p in paths if os.path.isfile(p)), key=lambda p: os.path.basename(p).lower())


OTHER_SOUNDS = (".wma", ".wav", ".aac", ".m4a", ".flac", ".ogg", ".opus", ".mp4", ".mpg", ".aiff")


def other_sounds(folder):
    """Names of sound files in the folder that are not MP3 (this app leaves them alone)."""
    try:
        return sorted((n for n in os.listdir(folder) if n.lower().endswith(OTHER_SOUNDS)), key=str.lower)
    except OSError:
        return []


def scan(folder, cache=None, progress=None, should_stop=None, workers=4, reader=read_song):
    """Read every MP3 in the folder. Returns the songs read so far if stopped.

    progress(done, total, name) is called from this thread after each song.
    """
    paths = list_songs(folder)
    songs = {}
    waiting = []
    for path in paths:
        known = cache.get(path) if cache else None
        if known is not None:
            songs[path] = known
        else:
            waiting.append(path)
    done = len(songs)
    if progress:
        progress(done, len(paths), "")
    stopped = False
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        queue = iter(waiting)
        running = set()

        def top_up():
            while len(running) < max(1, workers) * 2:
                path = next(queue, None)
                if path is None:
                    return
                running.add(pool.submit(reader, path))

        top_up()
        while running:
            finished, running = concurrent.futures.wait(
                running, return_when=concurrent.futures.FIRST_COMPLETED)
            for future in finished:
                song = future.result()
                songs[song.path] = song
                if cache:
                    cache.put(song)
                done += 1
                if progress:
                    progress(done, len(paths), song.name)
            if should_stop and should_stop():
                stopped = True
            if not stopped:
                top_up()
    return [songs[p] for p in paths if p in songs]
