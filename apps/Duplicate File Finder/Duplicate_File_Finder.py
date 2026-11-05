#!/usr/bin/env python3
"""
Duplicate File Finder
=====================
Pick a folder. The program looks for duplicates in one of four ways:

  1. Exact copies        - any kind of file, identical byte for byte.
  2. Documents and text  - Word (.docx), PDF, RTF, OpenDocument (.odt) and plain text
                           files that contain the same WORDS. Formatting (fonts,
                           colours, layout) is ignored.
  3. Look-alike photos   - pictures (.jpg, .png ...) that look the same even after
                           being resized, re-saved or compressed differently.
  4. Look-alike videos   - videos (.mp4, .mov, .mkv, .avi ...) that look the same even
                           after being resized, re-encoded, converted to another format,
                           given black bars, or shortened at the start or end.

Nothing is ever deleted. Copies are MOVED into a folder called
"Duplicates_To_Review" inside the folder you picked. Every move is written to a
log, so "Undo last move" can put them back. You see the full list (with pictures
or text previews) BEFORE anything moves, and you can take items off the list.

Speed:
  * A file whose size is unique cannot have a copy, so it is never opened.
  * Possible copies get a quick fingerprint (first and last 64 KB); only files
    that still match are read in full.
  * Several files are read at the same time.
  * Results are remembered in a small cache, so re-scanning the same folder only
    reads files that are new or changed.

Windows 11:  double-click this file, or type in Command Prompt:
                 py Duplicate_File_Finder.py
             Self-test (checks the logic on sample files):
                 py Duplicate_File_Finder.py --selftest
The first run installs the free Pillow picture library (needs internet).
Only the photo mode needs Pillow; the other modes work without it.
The first video scan installs the free ffmpeg video reader (needs internet, one time only).
"""

import csv
import errno
import hashlib
import io
import math
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import xml.etree.ElementTree as ET
import zipfile
import zlib
from array import array
from bisect import bisect_left, bisect_right
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from operator import mul


# ---------------------------------------------------------------------------
# Picture library (Pillow). Optional: only the photo mode needs it.
# ---------------------------------------------------------------------------
def _ensure_pillow():
    import importlib.util
    if importlib.util.find_spec("PIL") is not None:
        return True
    print("First-time setup: installing the Pillow picture library (needs internet)...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "pillow"])
    except Exception as exc:
        print("Could not install Pillow automatically: %s" % exc)
        print("Photo mode is off until you run:  py -m pip install pillow")
        return False
    importlib.invalidate_caches()
    return importlib.util.find_spec("PIL") is not None


PIL_OK = _ensure_pillow()
HEIC_OK = False
if PIL_OK:
    from PIL import Image, ImageOps
    Image.MAX_IMAGE_PIXELS = None
    _BOX = getattr(getattr(Image, "Resampling", Image), "BOX")
    try:
        import pillow_heif  # optional add-on for iPhone HEIC photos
        pillow_heif.register_heif_opener()
        HEIC_OK = True
    except Exception:
        HEIC_OK = False

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter import font as tkfont
    HAVE_TK = True
except Exception:
    HAVE_TK = False


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
MODE_EXACT = "exact"
MODE_TEXT = "text"
MODE_PHOTO = "photo"
MODE_VIDEO = "video"
MODE_NAMES = {
    MODE_EXACT: "Exact copies (any kind of file)",
    MODE_TEXT: "Documents and text (same words, any formatting)",
    MODE_PHOTO: "Look-alike photos",
    MODE_VIDEO: "Look-alike videos",
}
MODE_HELP = {
    MODE_EXACT: "Any file type. Only files identical byte for byte.",
    MODE_TEXT: "Word, PDF, RTF, ODT, text, CSV, HTML. Compares the words only.",
    MODE_PHOTO: "JPG, PNG, etc. Finds pictures that look the same after resizing or re-saving.",
    MODE_VIDEO: "MP4, MOV, MKV, AVI, etc. Finds resized, re-encoded, letterboxed and trimmed copies.",
}
CHUNK = 1024 * 1024          # read big files 1 MB at a time
EDGE = 64 * 1024             # quick check: first and last 64 KB
REVIEW_NAME = "Duplicates_To_Review"
LOG_NAME = "move_log.csv"
LOG_FIELDS = ["run_id", "when", "mode", "moved_from", "moved_to", "kept_file", "match", "undone"]
TEXT_EXTS = {".txt", ".md", ".csv", ".tsv", ".log", ".htm", ".html", ".xml", ".json", ".ini",
             ".cfg", ".yaml", ".yml", ".py", ".js", ".css", ".bat", ".ps1", ".sql", ".srt", ".vtt"}
PHOTO_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
HEIC_EXTS = {".heic", ".heif"}
VIDEO_EXTS = {".mp4", ".m4v", ".mov", ".mkv", ".avi", ".wmv", ".asf", ".flv", ".f4v", ".webm",
              ".mpg", ".mpeg", ".m2ts", ".mts", ".vob", ".3gp", ".ogv"}
JUNK_NAMES = {"desktop.ini", "thumbs.db", ".ds_store"}
SKIP_DIR_NAMES = {"$recycle.bin", "system volume information"}
HASH_SIZE = 16
HASH_BITS = HASH_SIZE * HASH_SIZE
BRIGHTNESS_TOLERANCE = 25
DOC_EXTS = {".docx", ".odt", ".rtf", ".pdf"}
MATCH_CHOICES = ["95%", "90%", "85%", "80%"]
SHINGLE = 5                  # words per phrase used to compare documents
VIDEO_MATCH_CHOICES = ["95%", "90%", "85%"]
VIDEO_FRAMES = 10            # moments compared per video, evenly spread along its length
VIDEO_MIN_FRAMES = 3         # fewer usable moments than this is too little to judge by
VIDEO_W, VIDEO_H = 32, 18    # each moment is compared as a small grey picture of this size
VIDEO_GRAB_W, VIDEO_GRAB_H = 160, 90   # moments are read at this size, to find black bars
VIDEO_GRAB_SIZE = VIDEO_GRAB_W * VIDEO_GRAB_H
VIDEO_FULL_BOX = (0, 0, VIDEO_GRAB_W, VIDEO_GRAB_H)
VIDEO_BAR_LEVEL = 24         # a line of the picture no brighter than this can be part of a black bar
VIDEO_ANCHORS = (10, 30)     # long videos: also read the moments this many seconds from each end...
VIDEO_ANCHOR_MIN = 60        # ...when they are at least this many seconds long
VIDEO_CUT_LOWEST = 0.85      # the lowest level on offer: nothing below it is ever worth checking
VIDEO_CUT_SMALLEST = 0.25    # a difference in length smaller than this many seconds is not a cut
VIDEO_CUT_SEARCH = 60        # seconds of the longer video searched, at most, to line a copy up
VIDEO_CUT_NUDGE = 0.12       # each moment is compared with the frames this many seconds either side
VIDEO_CUT_LINE_UP = 0.6      # lined up as well as possible, the moments must be at least this alike
VIDEO_CUT_HINT = 0.5         # one stored moment this alike is worth checking properly
VIDEO_CUT_SIGNS = 0.70       # quick check before that: this share of the outline must agree
VIDEO_CUT_FPS = 25           # frames per second read when lining a shortened copy up
VIDEO_EDGE_LIMIT = 16        # a strong edge (a logo, a title bar) counts for no more than this
VIDEO_BLANK_RANGE = 10       # a frame with less contrast than this is blank (black, white...)
VIDEO_KEEP_SECONDS = 0.5     # a shorter copy with a bigger picture stays if it is only this much shorter...
VIDEO_KEEP_SHARE = 0.01      # ...or this share of the length shorter, whichever is more
VIDEO_TIMEOUT = 60           # give up on a damaged video after this many seconds
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # no black window flashing up


def photo_exts():
    return PHOTO_EXTS | (HEIC_EXTS if HEIC_OK else set())


if hasattr(int, "bit_count"):
    def popcount(x):
        return x.bit_count()
else:
    def popcount(x):
        return bin(x).count("1")


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
@dataclass
class FileInfo:
    path: str
    size: int
    mtime: float
    width: int = 0
    height: int = 0

    @property
    def pixels(self):
        return self.width * self.height


@dataclass
class Group:
    keeper: FileInfo
    copies: list = field(default_factory=list)   # [(FileInfo, "how it matches")]

    def total_size(self):
        return sum(f.size for f, _ in self.copies)


@dataclass
class PhotoPrint:
    width: int
    height: int
    dhash: int
    mean: float


@dataclass
class VideoPrint:
    width: int
    height: int
    seconds: float
    mean: float
    frames: list                 # one small grey picture per moment; None = blank or unreadable
    heads: list = field(default_factory=list)    # the moments VIDEO_ANCHORS seconds from the start
    tails: list = field(default_factory=list)    # the moments VIDEO_ANCHORS seconds from the end
    box: tuple = VIDEO_FULL_BOX                  # the picture without its black bars


def keep_key(f):
    """Sort order for choosing which copy STAYS (first = keeper).
    More pixels first (photos), then names that do not look like 'copy' or '(2)',
    then the oldest date, then the shorter path."""
    stem = os.path.splitext(os.path.basename(f.path))[0]
    looks_like_copy = 1 if re.search(r"copy|\(\d+\)\s*$", stem, re.IGNORECASE) else 0
    return (-f.pixels, looks_like_copy, f.mtime, len(f.path), f.path.lower())


def video_keep_key(f):
    """Like keep_key, but between pictures of the same size the bigger file stays:
    it has kept more of the detail."""
    return (-f.pixels, -f.size) + keep_key(f)[1:]


def make_groups(clusters, how):
    groups = []
    for cluster in clusters:
        ordered = sorted(cluster, key=keep_key)
        groups.append(Group(ordered[0], [(f, how) for f in ordered[1:]]))
    groups.sort(key=lambda g: -g.total_size())
    return groups


# ---------------------------------------------------------------------------
# Finding files
# ---------------------------------------------------------------------------
def _skip_dir(full, name, review_path):
    low = name.lower()
    if name.startswith(".") or name.startswith("@") or low in SKIP_DIR_NAMES:
        return False
    return os.path.normcase(os.path.abspath(full)) != review_path


def collect_files(root, include_sub, mode):
    """Return (list of FileInfo, number of HEIC photos that cannot be read)."""
    review_path = os.path.normcase(os.path.abspath(os.path.join(root, REVIEW_NAME)))
    photo_ok = photo_exts()
    out, heic_skipped = [], 0
    for cur, dirs, files in os.walk(root):
        if include_sub:
            dirs[:] = [d for d in dirs if _skip_dir(os.path.join(cur, d), d, review_path)]
        else:
            dirs[:] = []
        for name in files:
            low = name.lower()
            if low in JUNK_NAMES or low.startswith("._"):
                continue
            ext = os.path.splitext(low)[1]
            if mode == MODE_TEXT and ext not in TEXT_EXTS | DOC_EXTS:
                continue
            if mode == MODE_PHOTO and ext not in photo_ok:
                if ext in HEIC_EXTS:
                    heic_skipped += 1
                continue
            if mode == MODE_VIDEO and ext not in VIDEO_EXTS:
                continue
            full = os.path.join(cur, name)
            try:
                st = os.stat(full)
            except OSError:
                continue
            if st.st_size == 0:
                continue
            out.append(FileInfo(full, st.st_size, st.st_mtime))
    return out, heic_skipped


def run_parallel(label, func, items, progress, cancel, every=20):
    """Call func(item) for every item on several threads. func must not raise.
    Returns the list of results, or None if the user pressed Stop."""
    if not items:
        return []
    workers = min(8, (os.cpu_count() or 2) + 2)
    results, total = [], len(items)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(func, it) for it in items]
        for n, fut in enumerate(as_completed(futures), 1):
            if cancel.is_set():
                for f in futures:
                    f.cancel()
                return None
            results.append(fut.result())
            if n % every == 0 or n == total:
                progress(label, n, total)
    return results


# ---------------------------------------------------------------------------
# Fingerprints
# ---------------------------------------------------------------------------
def full_hash(path):
    h = hashlib.blake2b(digest_size=32)
    with open(path, "rb") as f:
        while True:
            block = f.read(CHUNK)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def quick_hash(path, size):
    """Fingerprint of the first and last EDGE bytes (for files bigger than 2*EDGE)."""
    h = hashlib.blake2b(digest_size=32)
    with open(path, "rb") as f:
        h.update(f.read(EDGE))
        f.seek(max(0, size - EDGE))
        h.update(f.read(EDGE))
    return h.hexdigest()


def photo_print(path):
    """256-bit 'look' fingerprint of a picture. Raises if it cannot be read."""
    with Image.open(path) as im:
        width, height = im.size
        im.draft("L", (128, 128))           # JPEGs: decode at reduced size (much faster)
        im = ImageOps.exif_transpose(im)    # respect the rotation stored by phones
        if im.mode.startswith("I"):         # 16-bit pictures -> 8-bit
            im = im.point(lambda v: v * (1.0 / 256.0))
        small = im.convert("L").resize((HASH_SIZE + 1, HASH_SIZE), _BOX)
    data = small.tobytes()
    stride = HASH_SIZE + 1
    bits = 0
    for row in range(HASH_SIZE):
        base = row * stride
        for col in range(HASH_SIZE):
            bits = (bits << 1) | (1 if data[base + col] > data[base + col + 1] else 0)
    return PhotoPrint(width, height, bits, sum(data) / len(data))


def _encode_photo(pp):
    return "%d,%d,%x,%.4f" % (pp.width, pp.height, pp.dhash, pp.mean)


def _decode_photo(s):
    try:
        w, h, bits, mean = s.split(",")
        return PhotoPrint(int(w), int(h), int(bits, 16), float(mean))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Video reader (ffmpeg). Optional: only the video mode needs it.
# ---------------------------------------------------------------------------
def find_ffmpeg():
    """Path of ffmpeg on this PC, or None."""
    exe = shutil.which("ffmpeg")
    if exe:
        return exe
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return None


def install_ffmpeg():
    """Install the free package that contains ffmpeg. Returns its path, or None."""
    import importlib
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "imageio-ffmpeg"], check=True,
                       stdin=subprocess.DEVNULL, capture_output=True, creationflags=NO_WINDOW)
    except Exception:
        return None
    importlib.invalidate_caches()
    return find_ffmpeg()


def _run_ffmpeg(exe, args, timeout=VIDEO_TIMEOUT):
    return subprocess.run([exe, "-hide_banner", "-nostdin"] + args, stdin=subprocess.DEVNULL,
                          capture_output=True, timeout=timeout, creationflags=NO_WINDOW)


def video_info(exe, path):
    """(seconds, width, height) as reported by ffmpeg. Raises if the file has no picture."""
    text = _run_ffmpeg(exe, ["-i", path]).stderr.decode("utf-8", "replace")
    length = re.search(r"(?m)^\s*Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", text)
    size = re.search(r"(?m)^\s*Stream #.*?Video:.*?\b(\d{2,5})x(\d{2,5})\b", text)
    if not length or not size:
        raise ValueError("no video in this file")
    seconds = int(length[1]) * 3600 + int(length[2]) * 60 + float(length[3])
    if seconds <= 0:
        raise ValueError("this video has no length")
    return seconds, int(size[1]), int(size[2])


def _grab(exe, path, at):
    """The moment `at` seconds into a video, as a grey picture of the size VIDEO_GRAB_W x VIDEO_GRAB_H."""
    return _run_ffmpeg(exe, ["-loglevel", "error", "-threads", "1", "-ss", "%.3f" % at, "-i", path,
                             "-an", "-sn", "-frames:v", "1",
                             "-vf", "scale=%d:%d:flags=area,format=gray" % (VIDEO_GRAB_W, VIDEO_GRAB_H),
                             "-f", "rawvideo", "-"]).stdout


def find_bars(pictures):
    """Black bars along the edges of grabbed pictures. Returns the box (left, top, right, bottom)
    of what is inside them. Real bars come in equal pairs, so the smaller of the two is taken as
    their size: on the other side, a dark part of the scene only looks like more bar. A line
    counts as black if it is black in half of the pictures, so that subtitles shown inside a bar
    do not hide it."""
    w, h = VIDEO_GRAB_W, VIDEO_GRAB_H
    pictures = [p for p in pictures if max(p) - min(p) >= VIDEO_BLANK_RANGE]
    if not pictures:
        return VIDEO_FULL_BOX

    def black(line_of):
        return sum(max(line_of(p)) <= VIDEO_BAR_LEVEL for p in pictures) * 2 >= len(pictures)

    def bars(size, line_of):
        near = far = 0
        while near < size and black(lambda p: line_of(p, near)):
            near += 1
        while far < size - near and black(lambda p: line_of(p, size - 1 - far)):
            far += 1
        bar = min(near, far)
        return bar if bar >= 2 and (size - 2 * bar) * 5 >= size * 2 else 0
    top = bars(h, lambda p, y: p[y * w:(y + 1) * w])
    left = bars(w, lambda p, x: p[x::w])
    return left, top, w - left, h - top


def video_small(data, box):
    """Cut a grabbed picture down to `box` and shrink that to the size VIDEO_W x VIDEO_H."""
    left, top, right, bottom = box
    xs = [left + (right - left) * k // VIDEO_W for k in range(VIDEO_W + 1)]
    ys = [top + (bottom - top) * k // VIDEO_H for k in range(VIDEO_H + 1)]
    gw = VIDEO_GRAB_W
    out = bytearray()
    for cy in range(VIDEO_H):
        y0, y1 = ys[cy], ys[cy + 1]
        for cx in range(VIDEO_W):
            x0, x1 = xs[cx], xs[cx + 1]
            total = sum(sum(data[y * gw + x0:y * gw + x1]) for y in range(y0, y1))
            out.append(total // ((x1 - x0) * (y1 - y0)))
    return bytes(out)


def video_times(seconds):
    """The evenly spread moments of a video, in seconds from its start."""
    return [seconds * (k + 0.5) / VIDEO_FRAMES for k in range(VIDEO_FRAMES)]


def video_print(exe, path, cancel):
    """Small grey pictures of VIDEO_FRAMES moments spread along a video. Blank moments do not
    show which video this is, so they are left empty. Returns None if the user pressed Stop.
    Raises if the video cannot be read."""
    seconds, width, height = video_info(exe, path)
    anchors = list(VIDEO_ANCHORS) if seconds >= VIDEO_ANCHOR_MIN else []
    grabbed = []
    for at in video_times(seconds) + anchors + [seconds - t for t in anchors]:
        if cancel.is_set():
            return None
        grabbed.append(_grab(exe, path, at))
    readable = [d for d in grabbed if len(d) == VIDEO_GRAB_SIZE]
    if not readable:
        raise ValueError("no pictures could be read")
    box = find_bars(readable)
    smalls = [video_small(d, box) if len(d) == VIDEO_GRAB_SIZE else None for d in grabbed]
    levels = [sum(d) / len(d) for d in smalls[:VIDEO_FRAMES] if d]
    pics = [d if d and max(d) - min(d) >= VIDEO_BLANK_RANGE else None for d in smalls]
    n = VIDEO_FRAMES + len(anchors)
    return VideoPrint(width, height, seconds, sum(levels) / len(levels) if levels else 0.0,
                      pics[:VIDEO_FRAMES], pics[VIDEO_FRAMES:n], pics[n:], box)


def video_outline(data):
    """The outlines in a small grey picture: the brightness step from each pixel to the next one
    across and down, with strong edges capped. Scaled to length 1, so that multiplying two
    outlines gives their likeness (1 = the same picture, 0 = nothing in common).
    None if the picture has no outlines at all."""
    w = VIDEO_W
    steps = [data[i] - data[i + 1] for i in range(len(data)) if (i + 1) % w]
    steps += [data[i] - data[i + w] for i in range(len(data) - w)]
    steps = [max(-VIDEO_EDGE_LIMIT, min(VIDEO_EDGE_LIMIT, v)) for v in steps]
    length = math.sqrt(sum(v * v for v in steps))
    return array("f", [v / length for v in steps]) if length else None


def video_still(path, maxsize):
    """A moment from the middle of a video, as PNG bytes, for the preview."""
    exe = find_ffmpeg()
    seconds = video_info(exe, path)[0]
    return _run_ffmpeg(exe, ["-loglevel", "error", "-ss", "%.3f" % (seconds / 2), "-i", path, "-an", "-sn",
                             "-frames:v", "1", "-vf", "scale=%d:%d:force_original_aspect_ratio=decrease" % maxsize,
                             "-f", "image2pipe", "-c:v", "png", "-"], timeout=15).stdout


def _encode_video(vp):
    def pics(lst):
        return ";".join(d.hex() if d else "" for d in lst)
    return "%d,%d,%.3f,%.4f,%d:%d:%d:%d,%s|%s|%s" % ((vp.width, vp.height, vp.seconds, vp.mean) + tuple(vp.box)
                                                     + (pics(vp.frames), pics(vp.heads), pics(vp.tails)))


def _decode_video(s):
    def pics(part):
        return [bytes.fromhex(x) if x else None for x in part.split(";")] if part else []
    try:
        w, h, seconds, mean, box, rest = s.split(",")
        frames, heads, tails = rest.split("|")
        vp = VideoPrint(int(w), int(h), float(seconds), float(mean), pics(frames), pics(heads), pics(tails),
                        tuple(int(v) for v in box.split(":")))
    except ValueError:
        return None
    good = (len(vp.frames) == VIDEO_FRAMES and len(vp.box) == 4
            and len(vp.heads) == len(vp.tails) and len(vp.heads) in (0, len(VIDEO_ANCHORS))
            and all(d is None or len(d) == VIDEO_W * VIDEO_H for d in vp.frames + vp.heads + vp.tails))
    return vp if good else None


# ---------------------------------------------------------------------------
# Remembered fingerprints (cache). Only the scan thread touches the database.
# ---------------------------------------------------------------------------
class FingerprintCache:
    def __init__(self, folder=None):
        self.con = None
        try:
            if folder is None:
                base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".cache")
                folder = os.path.join(base, "DuplicateFileFinder")
            os.makedirs(folder, exist_ok=True)
            self.con = sqlite3.connect(os.path.join(folder, "fingerprints.db"))
            self.con.execute("CREATE TABLE IF NOT EXISTS fp (kind TEXT, path TEXT, size INTEGER, "
                             "mtime REAL, value TEXT, PRIMARY KEY (kind, path))")
            self.con.commit()
        except Exception:
            self.con = None

    def get(self, kind, f):
        if self.con is None:
            return None
        try:
            row = self.con.execute("SELECT value FROM fp WHERE kind=? AND path=? AND size=? AND mtime=?",
                                   (kind, f.path, f.size, f.mtime)).fetchone()
        except Exception:
            return None
        return row[0] if row else None

    def put_many(self, rows):
        if self.con is None or not rows:
            return
        try:
            self.con.executemany("INSERT OR REPLACE INTO fp (kind, path, size, mtime, value) VALUES (?,?,?,?,?)",
                                 [(k, f.path, f.size, f.mtime, v) for k, f, v in rows])
            self.con.commit()
        except Exception:
            pass

    def close(self):
        if self.con is not None:
            try:
                self.con.close()
            except Exception:
                pass


class NoCache:
    def get(self, kind, f):
        return None

    def put_many(self, rows):
        pass

    def close(self):
        pass


# ---------------------------------------------------------------------------
# The finders. Each returns (clusters_or_groups, unreadable_count),
# or None if the user pressed Stop.
# ---------------------------------------------------------------------------
def _quick_for(f):
    try:
        if f.size <= 2 * EDGE:
            return f, full_hash(f.path), True
        return f, quick_hash(f.path, f.size), False
    except Exception:
        return f, None, False


def _full_for(f):
    try:
        return f, full_hash(f.path)
    except Exception:
        return f, None


def find_exact(files, cache, progress, cancel):
    by_size = defaultdict(list)
    for f in files:
        by_size[f.size].append(f)
    cands = [f for g in by_size.values() if len(g) > 1 for f in g]
    if not cands:
        return [], 0

    # 1) one quick fingerprint per candidate (from memory when possible)
    quick, todo, fresh = {}, [], []
    for f in cands:
        small = f.size <= 2 * EDGE
        v = cache.get("exact" if small else "quick", f)
        if v is not None:
            quick[f.path] = (v, small)
        else:
            todo.append(f)
    res = run_parallel("Reading files", _quick_for, todo, progress, cancel)
    if res is None:
        return None
    failed = 0
    for f, v, is_full in res:
        if v is None:
            failed += 1
            continue
        quick[f.path] = (v, is_full)
        fresh.append(("exact" if is_full else "quick", f, v))

    # 2) files sharing a quick fingerprint need a full fingerprint
    full, partial = {}, defaultdict(list)
    for f in cands:
        if f.path not in quick:
            continue
        v, is_full = quick[f.path]
        if is_full:
            full[f.path] = v
        else:
            partial[(f.size, v)].append(f)
    need = []
    for group in partial.values():
        if len(group) > 1:
            for f in group:
                v = cache.get("exact", f)
                if v is not None:
                    full[f.path] = v
                else:
                    need.append(f)
    res2 = run_parallel("Comparing files", _full_for, need, progress, cancel)
    if res2 is None:
        return None
    for f, v in res2:
        if v is None:
            failed += 1
            continue
        full[f.path] = v
        fresh.append(("exact", f, v))
    cache.put_many(fresh)

    by_hash = defaultdict(list)
    for f in cands:
        if f.path in full:
            by_hash[full[f.path]].append(f)
    return [g for g in by_hash.values() if len(g) > 1], failed


class PdfNeedsPypdf(Exception):
    pass


def _read_text_file(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace")


def extract_text(path):
    """Plain words of a document. Formatting is dropped. Raises on failure."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("word/document.xml"))
        paras = ["".join(t.text or "" for t in p.iter() if t.tag.endswith("}t"))
                 for p in root.iter() if p.tag.endswith("}p")]
        return "\n".join(paras)
    if ext == ".odt":
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("content.xml"))
        return " ".join(root.itertext())
    if ext == ".rtf":
        text = _read_text_file(path)
        text = re.sub(r"\\'[0-9a-fA-F]{2}", " ", text)
        text = re.sub(r"\\[a-zA-Z]+-?\d* ?", " ", text)
        return re.sub(r"[{}\\]", " ", text)
    if ext == ".pdf":
        try:
            from pypdf import PdfReader
        except ImportError:
            raise PdfNeedsPypdf()
        return " ".join((page.extract_text() or "") for page in PdfReader(path).pages)
    text = _read_text_file(path)
    if ext in (".htm", ".html", ".xml"):
        text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", text)
        text = re.sub(r"<[^>]+>", " ", text)
    return text


def words_of(text):
    return re.findall(r"\w+", text.lower())


def _words_for(f):
    try:
        return f, words_of(extract_text(f.path)), None
    except PdfNeedsPypdf:
        return f, None, "pdf"
    except Exception:
        return f, None, "other"


def shingles(words):
    """Set of SHINGLE-word phrases (as numbers). Short texts count as one phrase."""
    if not words:
        return set()
    if len(words) < SHINGLE:
        return {zlib.crc32(" ".join(words).encode("utf-8"))}
    return {zlib.crc32(" ".join(words[i:i + SHINGLE]).encode("utf-8"))
            for i in range(len(words) - SHINGLE + 1)}


def group_text(items, threshold_pct):
    """items: list of (FileInfo, words). A pair matches when the share of phrases they
    have in common (Jaccard overlap) is at least threshold_pct. Whole-number arithmetic,
    so exactly 90% counts as a 90% match."""
    ordered = sorted(items, key=lambda it: keep_key(it[0]))
    files = [f for f, _w in ordered]
    sets = [shingles(w) for _f, w in ordered]
    postings = defaultdict(list)
    for pos, s in enumerate(sets):
        for sh in s:
            postings[sh].append(pos)
    shared = defaultdict(int)
    for lst in postings.values():
        for a in range(len(lst)):
            for b in range(a + 1, len(lst)):
                shared[(lst[a], lst[b])] += 1
    neighbours = defaultdict(list)
    for (a, b), inter in shared.items():
        neighbours[a].append((b, inter))
    assigned = [False] * len(ordered)
    groups = []
    for i in range(len(ordered)):
        if assigned[i] or not sets[i]:
            continue
        assigned[i] = True
        copies = []
        for j, inter in sorted(neighbours.get(i, [])):
            if assigned[j] or not sets[j]:
                continue
            union = len(sets[i]) + len(sets[j]) - inter
            if inter * 100 >= threshold_pct * union:
                assigned[j] = True
                copies.append((files[j], "%d%% same words" % round(100.0 * inter / union)))
        if copies:
            groups.append(Group(files[i], copies))
    groups.sort(key=lambda g: -g.total_size())
    return groups


def find_text(files, pct, progress, cancel, notes):
    items, failed, pdf = [], 0, 0
    res = run_parallel("Reading documents", _words_for, files, progress, cancel)
    if res is None:
        return None
    for f, words, err in res:
        if err == "pdf":
            pdf += 1
        elif err is not None:
            failed += 1
        elif words:
            items.append((f, words))
    if pdf:
        notes.append("%d PDF(s) skipped - install the free pypdf library: py -m pip install pypdf" % pdf)
    return group_text(items, pct), failed


def _photo_for(f):
    try:
        return f, photo_print(f.path)
    except Exception:
        return f, None


def group_similar(usable, threshold_pct):
    """usable: list of (FileInfo, PhotoPrint). Returns a list of Group."""
    threshold = max(80.0, min(100.0, float(threshold_pct)))
    max_dist = int(HASH_BITS * (100.0 - threshold) / 100.0 + 1e-9)
    ordered = sorted(usable, key=lambda fp: keep_key(fp[0]))
    n = len(ordered)
    # Pigeonhole: two prints within max_dist differing bits must match exactly
    # in at least one of (max_dist + 1) slices, so only those need comparing.
    nb = max_dist + 1
    bounds = [(k * HASH_BITS) // nb for k in range(nb + 1)]
    bands = [(bounds[k], (1 << (bounds[k + 1] - bounds[k])) - 1) for k in range(nb)]
    hashes = [pp.dhash for _, pp in ordered]
    means = [pp.mean for _, pp in ordered]
    tables = [defaultdict(list) for _ in range(nb)]
    for idx, h in enumerate(hashes):
        for k, (shift, mask) in enumerate(bands):
            tables[k][(h >> shift) & mask].append(idx)

    assigned = [False] * n
    groups = []
    for i in range(n):
        if assigned[i]:
            continue
        assigned[i] = True
        hi, mi = hashes[i], means[i]
        cand = set()
        for k, (shift, mask) in enumerate(bands):
            cand.update(tables[k].get((hi >> shift) & mask, ()))
        copies = []
        for j in sorted(cand):
            if j <= i or assigned[j]:
                continue
            d = popcount(hi ^ hashes[j])
            if d <= max_dist and abs(mi - means[j]) <= BRIGHTNESS_TOLERANCE:
                assigned[j] = True
                copies.append((ordered[j][0], "%.1f%% alike" % (100.0 * (1 - d / HASH_BITS))))
        if copies:
            groups.append(Group(ordered[i][0], copies))
    groups.sort(key=lambda g: -g.total_size())
    return groups


def find_similar_photos(files, cache, pct, progress, cancel):
    prints, todo = {}, []
    for f in files:
        pp = _decode_photo(cache.get("photo", f) or "")
        if pp is not None:
            prints[f.path] = pp
        else:
            todo.append(f)
    res = run_parallel("Reading pictures", _photo_for, todo, progress, cancel)
    if res is None:
        return None
    failed, fresh = 0, []
    for f, pp in res:
        if pp is None:
            failed += 1
            continue
        prints[f.path] = pp
        fresh.append(("photo", f, _encode_photo(pp)))
    cache.put_many(fresh)

    usable = []
    for f in files:
        pp = prints.get(f.path)
        if pp is None:
            continue
        f.width, f.height = pp.width, pp.height
        usable.append((f, pp))
    progress("Comparing pictures", 0, 0)
    return group_similar(usable, pct), failed


def _video_for(exe, f, cancel):
    try:
        return f, video_print(exe, f.path, cancel)
    except Exception:
        return f, None


def video_likeness(outlines_a, outlines_b, floor):
    """Average likeness (0 to 1) of two videos, moment by moment. Moments that are blank in
    either video are left out. Returns 0 as soon as `floor` can no longer be reached."""
    pairs = [(a, b) for a, b in zip(outlines_a, outlines_b) if a is not None and b is not None]
    if len(pairs) < VIDEO_MIN_FRAMES:
        return 0.0
    allowed = len(pairs) * (1.0 - floor)
    total = 0.0
    for k, (a, b) in enumerate(pairs, 1):
        total += max(0.0, sum(map(mul, a, b)))
        if k - total > allowed:
            return 0.0
    return total / len(pairs)


def group_videos(usable, threshold_pct):
    """usable: list of (FileInfo, VideoPrint). A pair matches when the two videos are the same
    length and about the same brightness, and their outlines are at least threshold_pct alike.
    Videos that differ in length are left to find_cut_copies."""
    floor = threshold_pct / 100.0 - 1e-6     # allow for rounding: exactly 90% counts at 90%
    ordered = sorted(usable, key=lambda fp: video_keep_key(fp[0]))
    prints = [vp for _f, vp in ordered]
    outlines = [[video_outline(d) if d else None for d in vp.frames] for vp in prints]
    n = len(ordered)
    by_length = sorted(range(n), key=lambda k: prints[k].seconds)
    lengths = [prints[k].seconds for k in by_length]
    assigned = [False] * n
    groups = []
    for i in range(n):
        if assigned[i]:
            continue
        assigned[i] = True
        a = prints[i]
        lo = bisect_right(lengths, a.seconds - VIDEO_CUT_SMALLEST)
        hi = bisect_left(lengths, a.seconds + VIDEO_CUT_SMALLEST)
        copies = []
        for j in sorted(by_length[lo:hi]):
            if j <= i or assigned[j] or abs(a.mean - prints[j].mean) > BRIGHTNESS_TOLERANCE:
                continue
            alike = video_likeness(outlines[i], outlines[j], floor)
            if alike >= floor:
                assigned[j] = True
                copies.append((ordered[j][0], "%.1f%% alike" % (100.0 * alike)))
        if copies:
            groups.append(Group(ordered[i][0], copies))
    groups.sort(key=lambda g: -g.total_size())
    return groups


def fmt_length(seconds):
    if seconds < 60:
        return "%.1fs" % seconds
    return "%dm %02ds" % divmod(round(seconds), 60)


def _with_signs(pics):
    """For each small picture: (its outline, the same as one big number with one bit per step)."""
    out = []
    for d in pics:
        o = video_outline(d) if d else None
        out.append((o, int("".join("1" if v > 0 else "0" for v in o), 2)) if o is not None else None)
    return out


def _moment_alike(a, b):
    """How alike one stored moment is in two videos. The big numbers rule out most pairs very cheaply."""
    if a is None or b is None or popcount(a[1] ^ b[1]) > len(a[0]) * (1.0 - VIDEO_CUT_SIGNS):
        return 0.0
    return sum(map(mul, a[0], b[0]))


def cut_ranges(long_look, short_look, missing):
    """Where might a shorter video start inside a longer one? A list of (earliest, latest) in
    seconds, judged only from the stored moments, so that no time goes on unrelated videos.
      - same moment counted from the start: nothing was cut from the start
      - same moment counted from the end: everything was cut from the start
      - cut from both ends: the evenly spread moments slide past each other along the length, and
        line up where the share of the length matches the share cut from the start. Moments in the
        same scene look alike too, so everything around the moments that match is searched."""
    ranges = []
    if any(_moment_alike(a, b) >= VIDEO_CUT_HINT for a, b in zip(long_look["heads"], short_look["heads"])):
        ranges.append((0.0, 0.0))
    if any(_moment_alike(a, b) >= VIDEO_CUT_HINT for a, b in zip(long_look["tails"], short_look["tails"])):
        ranges.append((missing, missing))
    hits = [k for k, (a, b) in enumerate(zip(long_look["frames"], short_look["frames"]))
            if _moment_alike(a, b) >= VIDEO_CUT_HINT]
    if hits:
        step = missing / len(long_look["frames"])
        middle = (hits[len(hits) // 2] + 0.5) * step
        ranges.append((max(0.0, (hits[0] - 1) * step, middle - VIDEO_CUT_SEARCH / 2),
                       min(missing, (hits[-1] + 2) * step, middle + VIDEO_CUT_SEARCH / 2)))
    return ranges


def _read_stretch(exe, path, vp, start, seconds, fps):
    """A stretch of a video, frame by frame, as a list of outlines (None for a frame without any)."""
    raw = _run_ffmpeg(exe, ["-loglevel", "error", "-threads", "1", "-ss", "%.3f" % start, "-t", "%.3f" % seconds,
                            "-i", path, "-an", "-sn", "-vf",
                            "fps=%d,scale=%d:%d:flags=area,format=gray" % (fps, VIDEO_GRAB_W, VIDEO_GRAB_H),
                            "-f", "rawvideo", "-"]).stdout
    return [video_outline(video_small(raw[i * VIDEO_GRAB_SIZE:(i + 1) * VIDEO_GRAB_SIZE], vp.box))
            for i in range(len(raw) // VIDEO_GRAB_SIZE)]


def _scan_window(exe, path, vp, times, outlines, lo, hi, fps, cancel):
    """Read a stretch of one video around a few moments of the other. Returns the offset (seconds)
    that lines those moments up best and how alike they are there, or (None, 0)."""
    usable = [k for k, o in enumerate(outlines)
              if o is not None and times[k] + lo >= 0 and times[k] + hi <= vp.seconds]
    picked = usable[::max(1, (len(usable) + 2) // 3)][:3]
    scores = defaultdict(float)
    for k in picked:
        if cancel.is_set():
            return None, 0.0
        for i, o in enumerate(_read_stretch(exe, path, vp, times[k] + lo, hi - lo + 1.0 / fps, fps)):
            if o is not None:
                scores[i] += max(0.0, sum(map(mul, o, outlines[k])))
    if not scores:
        return None, 0.0
    # Several frames in a row can fit equally well (a still scene, a low frame rate): take the middle.
    top = max(scores.values())
    near = sorted(i for i, v in scores.items() if v >= top - 0.01)
    last = near[0]
    for i in near[1:]:
        if i != last + 1:
            break
        last = i
    return lo + (near[0] + last) / 2.0 / fps, top / len(picked)


def locate_cut(exe, path, vp, times, outlines, lo, hi, cancel):
    """Exactly how far apart two videos run, searching around (lo, hi)."""
    lo, hi = lo - 0.5, hi + 0.5
    if hi - lo > 20:             # a long stretch: look coarsely first, then close in
        at, _alike = _scan_window(exe, path, vp, times, outlines, lo, hi, 5, cancel)
        if at is None:
            return None, 0.0
        lo, hi = at - 0.4, at + 0.4
    return _scan_window(exe, path, vp, times, outlines, lo, hi, VIDEO_CUT_FPS, cancel)


def _cut_likeness(exe, path, vp, times, outlines, offset, floor, cancel):
    """Likeness (0 to 1) of the moments of one video and the same moments of another that runs
    `offset` seconds later. Gives up with 0 as soon as `floor` cannot be reached."""
    usable = [k for k, o in enumerate(outlines) if o is not None and 0 <= times[k] + offset <= vp.seconds]
    if len(usable) < VIDEO_MIN_FRAMES:
        return 0.0
    allowed = len(usable) * (1.0 - floor)
    total = 0.0
    for n, k in enumerate(usable, 1):
        if cancel.is_set():
            return 0.0
        stretch = _read_stretch(exe, path, vp, max(0.0, times[k] + offset - VIDEO_CUT_NUDGE),
                                2 * VIDEO_CUT_NUDGE + 1.0 / VIDEO_CUT_FPS, VIDEO_CUT_FPS)
        total += max([sum(map(mul, o, outlines[k])) for o in stretch if o is not None] + [0.0])
        if n - total > allowed:
            return 0.0
    return total / len(usable)


def _check_cut(exe, job, cancel):
    """Is the shorter video of a pair a cut-down copy of the longer one? Returns (job, likeness of the
    part they share), or (job, None) if a video could not be read. Independent of the chosen level,
    so that the answer can be remembered. Only the file that is quicker to read gets read: it is
    compared with the stored moments of the other."""
    a, va, b, vb, ranges, look_a, look_b = job
    read_long = a.size * vb.seconds <= b.size * va.seconds
    reader, vr, vs, look = (a, va, vb, look_b) if read_long else (b, vb, va, look_a)
    outlines = [o[0] if o else None for o in look["frames"]]
    times = video_times(vs.seconds)
    floor = VIDEO_CUT_LOWEST * va.seconds / vb.seconds
    best = 0.0
    try:
        for lo, hi in ranges:
            if cancel.is_set():
                break
            span = (lo, hi) if read_long else (-hi, -lo)
            offset, alike = locate_cut(exe, reader.path, vr, times, outlines, span[0], span[1], cancel)
            if offset is not None and alike >= VIDEO_CUT_LINE_UP:
                best = max(best, _cut_likeness(exe, reader.path, vr, times, outlines, offset, floor, cancel))
            if best >= floor:
                break
    except Exception:
        return job, None
    return job, best


def _cut_kind(other):
    return "cut|%s|%d|%f" % (other.path, other.size, other.mtime)


def _join_cut_copies(a, va, matched, by_keeper, seconds, gone):
    """Put a video and the cut-down copies found for it in one group. matched: list of
    (FileInfo, VideoPrint, score). The longest copy stays, unless another is hardly shorter
    and has a bigger picture."""
    keeper = a
    allowance = max(VIDEO_KEEP_SECONDS, va.seconds * VIDEO_KEEP_SHARE)
    for b, vb, _score in matched:
        if va.seconds - vb.seconds <= allowance and b.pixels > keeper.pixels:
            keeper = b

    def how(score, f):
        gap = seconds[keeper.path] - seconds[f.path]
        text = "%.1f%% alike" % (100.0 * score)
        if abs(gap) >= VIDEO_CUT_SMALLEST:
            text += ", %s %s" % (fmt_length(abs(gap)), "shorter" if gap > 0 else "longer")
        return text
    members = []
    own = by_keeper.pop(a.path).copies if a.path in by_keeper else []
    if keeper is a:
        members += own
    else:
        kept_score = next(score for b, _vb, score in matched if b is keeper)
        members += [(f, how(kept_score, f)) for f in [a] + [f for f, _h in own]]
    for b, _vb, score in matched:
        own = by_keeper.pop(b.path).copies if b.path in by_keeper else []
        if b is keeper:
            members += own
        else:
            members += [(f, how(score, f)) for f in [b] + [f for f, _h in own]]
    by_keeper[keeper.path] = Group(keeper, members)
    gone.update(f.path for f, _h in members)


def find_cut_copies(groups, usable, exe, cache, pct, progress, cancel):
    """Second pass, over the videos group_videos left as keepers or alone: is a shorter one a
    cut-down copy of a longer one? The score is the likeness of the part they share times the
    share of the longer video that is still there, so a copy with 10% cut off can reach 90% at
    most. The longest video goes first, and what it takes in needs no checking of its own.
    Returns the updated groups, or None if the user pressed Stop."""
    floor = pct / 100.0 - 1e-6
    copies = {f.path for g in groups for f, _h in g.copies}
    leaders = sorted(((f, vp) for f, vp in usable if f.path not in copies),
                     key=lambda fp: (-fp[1].seconds, video_keep_key(fp[0])))
    looks = {f.path: {"frames": _with_signs(vp.frames), "heads": _with_signs(vp.heads), "tails": _with_signs(vp.tails)}
             for f, vp in leaders}
    worth = defaultdict(list)
    for i, (a, va) in enumerate(leaders):
        for b, vb in leaders[i + 1:]:
            if vb.seconds < va.seconds * floor:
                break
            missing = va.seconds - vb.seconds
            ranges = cut_ranges(looks[a.path], looks[b.path], missing) if missing >= VIDEO_CUT_SMALLEST else None
            if ranges:
                worth[a.path].append((b, vb, ranges))
    total = sum(len(v) for v in worth.values())
    done = 0
    by_keeper = {g.keeper.path: g for g in groups}
    seconds = {f.path: vp.seconds for f, vp in usable}
    gone = set()
    for a, va in leaders:
        todo = [t for t in worth.get(a.path, ()) if t[0].path not in gone]
        if a.path in gone or not todo:
            continue
        likeness, jobs = {}, []
        for b, vb, ranges in todo:
            known = cache.get(_cut_kind(b), a)
            if known is not None:
                likeness[b.path] = float(known)
            else:
                jobs.append((a, va, b, vb, ranges, looks[a.path], looks[b.path]))
        res = run_parallel("Checking for shortened copies", lambda job: _check_cut(exe, job, cancel), jobs,
                           lambda label, n, _t: progress(label, done + n, total), cancel, every=1)
        if res is None:
            return None
        done += len(todo)
        fresh = []
        for job, look in res:
            if look is not None:
                likeness[job[2].path] = look
                fresh.append((_cut_kind(job[2]), a, "%.5f" % look))
        cache.put_many(fresh)
        scored = [(b, vb, likeness.get(b.path, 0.0) * vb.seconds / va.seconds) for b, vb, _ranges in todo]
        matched = [m for m in scored if m[2] >= floor]
        if matched:
            _join_cut_copies(a, va, matched, by_keeper, seconds, gone)
    return sorted((g for g in by_keeper.values() if g.copies), key=lambda g: -g.total_size())


def find_similar_videos(files, cache, pct, progress, cancel, notes):
    prints, todo = {}, []
    for f in files:
        vp = _decode_video(cache.get("video", f) or "")
        if vp is not None:
            prints[f.path] = vp
        else:
            todo.append(f)
    exe = find_ffmpeg() if todo else None
    if todo and exe is None:
        progress("First-time setup: installing the free video reader (needs internet)", 0, 0)
        exe = install_ffmpeg()
    if todo and exe is None:
        notes.append("%d video(s) skipped - could not install the video reader; check the internet, "
                     "or run: py -m pip install imageio-ffmpeg" % len(todo))
        todo = []
    res = run_parallel("Reading videos", lambda f: _video_for(exe, f, cancel), todo, progress, cancel, every=1)
    if res is None:
        return None
    failed, fresh = 0, []
    for f, vp in res:
        if vp is None:
            failed += 1
            continue
        prints[f.path] = vp
        fresh.append(("video", f, _encode_video(vp)))
    cache.put_many(fresh)

    usable = []
    for f in files:
        vp = prints.get(f.path)
        if vp is None:
            continue
        left, top, right, bottom = vp.box        # black bars do not count as picture
        f.width = vp.width * (right - left) // VIDEO_GRAB_W
        f.height = vp.height * (bottom - top) // VIDEO_GRAB_H
        usable.append((f, vp))
    progress("Comparing videos", 0, 0)
    groups = group_videos(usable, pct)
    exe = exe or find_ffmpeg()
    if exe:
        groups = find_cut_copies(groups, usable, exe, cache, pct, progress, cancel)
        if groups is None:
            return None
    return groups, failed


def find_groups(mode, files, cache, pct, progress, cancel, notes=None):
    """Returns (list of Group, unreadable_count) or None if stopped."""
    if notes is None:
        notes = []
    if mode == MODE_PHOTO:
        return find_similar_photos(files, cache, pct, progress, cancel)
    if mode == MODE_VIDEO:
        return find_similar_videos(files, cache, pct, progress, cancel, notes)
    if mode == MODE_TEXT:
        return find_text(files, pct, progress, cancel, notes)
    out = find_exact(files, cache, progress, cancel)
    if out is None:
        return None
    return make_groups(out[0], "identical"), out[1]


def scan_job(root, include_sub, mode, pct, use_cache, cancel, out_q):
    """Runs in a background thread. Reports through out_q."""
    def progress(label, done, total):
        out_q.put(("progress", label, done, total))
    try:
        files, heic_skipped = collect_files(root, include_sub, mode)
        notes = []
        stats = {"files": len(files), "failed": 0, "heic": heic_skipped, "notes": notes}
        if cancel.is_set():
            out_q.put(("cancelled",))
            return
        if not files:
            out_q.put(("done", [], stats))
            return
        cache = FingerprintCache() if use_cache else NoCache()
        try:
            result = find_groups(mode, files, cache, pct, progress, cancel, notes)
        finally:
            cache.close()
        if result is None:
            out_q.put(("cancelled",))
            return
        stats["failed"] = result[1]
        out_q.put(("done", result[0], stats))
    except Exception as exc:
        out_q.put(("error", "%s: %s" % (type(exc).__name__, exc)))


# ---------------------------------------------------------------------------
# Moving, logging and undo. Nothing is deleted.
# ---------------------------------------------------------------------------
def unique_name(folder, filename, tag=None):
    base, ext = os.path.splitext(filename)
    if tag:
        base += " " + tag
    dest = os.path.join(folder, base + ext)
    n = 2
    while os.path.exists(dest):
        dest = os.path.join(folder, "%s (%d)%s" % (base, n, ext))
        n += 1
    return dest


def explain(exc):
    if isinstance(exc, PermissionError) or getattr(exc, "winerror", None) in (5, 32, 33):
        return "another program has it open, or it is write-protected"
    if getattr(exc, "errno", None) == errno.ENOSPC:
        return "the drive is full"
    if isinstance(exc, FileNotFoundError):
        return "the file disappeared"
    return "it could not be moved (%s)" % (getattr(exc, "strerror", None) or exc)


def read_log(log_path):
    if not os.path.isfile(log_path):
        return []
    with open(log_path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_log(log_path, rows):
    tmp = log_path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, log_path)


def move_copies(root, moves, mode, progress=None):
    """moves: list of (copy_path, keeper_path, how). Each copy goes into the review
    folder and is logged. Returns a dict with moved, skipped, failed, run_id."""
    review = os.path.join(root, REVIEW_NAME)
    os.makedirs(review, exist_ok=True)
    log_path = os.path.join(review, LOG_NAME)
    now = datetime.now()
    base_id = now.strftime("%Y%m%d-%H%M%S")
    run_id, used, k = base_id, {r.get("run_id") for r in read_log(log_path)}, 2
    while run_id in used:
        run_id = "%s-%d" % (base_id, k)
        k += 1
    is_new = not os.path.isfile(log_path)
    moved, skipped, failed = 0, 0, []
    with open(log_path, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=LOG_FIELDS)
        if is_new:
            w.writeheader()
        for i, (src, keeper, how) in enumerate(moves, 1):
            if progress and i % 25 == 0:
                progress(i, len(moves))
            if not os.path.isfile(src):
                skipped += 1
                continue
            dest = unique_name(review, os.path.basename(src))
            try:
                shutil.move(src, dest)
            except Exception as exc:
                failed.append((src, explain(exc)))
                continue
            try:
                w.writerow({"run_id": run_id, "when": now.strftime("%Y-%m-%d %H:%M:%S"), "mode": mode,
                            "moved_from": os.path.abspath(src), "moved_to": os.path.abspath(dest),
                            "kept_file": os.path.abspath(keeper), "match": how, "undone": "0"})
                fh.flush()
                moved += 1
            except Exception as exc:
                try:                       # cannot record it -> put it back, leave no unlogged move
                    shutil.move(dest, src)
                except Exception:
                    pass
                failed.append((src, "could not write the log file (%s)" % exc))
    return {"moved": moved, "skipped": skipped, "failed": failed, "run_id": run_id, "review": review}


def undo_last_move(root):
    """Put back everything from the most recent move not yet undone. None if nothing to undo."""
    log_path = os.path.join(root, REVIEW_NAME, LOG_NAME)
    rows = read_log(log_path)
    pending = [r for r in rows if r.get("undone") != "1"]
    if not pending:
        return None
    run_id = pending[-1]["run_id"]
    restored = missing = renamed = 0
    failed = []
    try:
        for r in rows:
            if r["run_id"] != run_id or r.get("undone") == "1":
                continue
            src, dst = r["moved_to"], r["moved_from"]
            if not os.path.isfile(src):
                r["undone"] = "1"          # someone already removed it from the review folder
                missing += 1
                continue
            target = dst
            if os.path.exists(target):     # name taken again: restore beside it, never overwrite
                target = unique_name(os.path.dirname(dst), os.path.basename(dst), "(restored)")
                renamed += 1
            try:
                os.makedirs(os.path.dirname(target), exist_ok=True)
                shutil.move(src, target)
                r["undone"] = "1"
                restored += 1
            except Exception as exc:
                failed.append((src, explain(exc)))
    finally:
        write_log(log_path, rows)
    return {"restored": restored, "missing": missing, "renamed": renamed,
            "failed": failed, "run_id": run_id}


# ---------------------------------------------------------------------------
# Small helpers for the window
# ---------------------------------------------------------------------------
def fmt_size(nbytes):
    if nbytes >= 1024 * 1024 * 1024:
        return "%.1f GB" % (nbytes / (1024 * 1024 * 1024))
    if nbytes >= 1024 * 1024:
        return "%.1f MB" % (nbytes / (1024 * 1024))
    return "%d KB" % max(1, round(nbytes / 1024))


def fmt_date(ts):
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


def text_snippet(path, lines=14):
    try:
        text = extract_text(path)
    except PdfNeedsPypdf:
        return "(PDF: install pypdf to preview)"
    except Exception:
        return "(cannot read this file)"
    shown = "\n".join(text.splitlines()[:lines])[:1200]
    return shown or "(no text found in this file)"


# ---------------------------------------------------------------------------
# The window
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Duplicate File Finder")
        root.geometry("1180x%d" % max(640, min(840, root.winfo_screenheight() - 100)))
        root.minsize(980, 640)
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=11)
        tkfont.nametofont("TkTextFont").configure(size=11)
        style = ttk.Style()
        style.configure("Treeview", rowheight=26, font=base)
        style.configure("Treeview.Heading", font=(base.actual("family"), 11, "bold"))

        self.root_dir = None          # folder the user picked
        self.scan_root = None         # folder that was scanned
        self.scan_mode = MODE_EXACT
        self.rows = {}                # tree id -> (copy FileInfo, keeper FileInfo, how)
        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self._img_refs = []

        self.mode_var = tk.StringVar(value=MODE_EXACT)
        self.dir_var = tk.StringVar(value="(no folder chosen yet)")
        self.sub_var = tk.BooleanVar(value=True)
        self.cache_var = tk.BooleanVar(value=True)
        self.pct_var = tk.StringVar(value="95%")
        self.status_var = tk.StringVar(value="Step 1: choose the folder to check.")
        self.summary_var = tk.StringVar(value="")

        self._build()
        self._refresh()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build(self):
        pad = {"padx": 10, "pady": 4}
        top = ttk.Frame(self.root)
        top.pack(side="top", fill="x", **pad)

        r1 = ttk.Frame(top)
        r1.pack(fill="x", pady=2)
        self.btn_choose = ttk.Button(r1, text="1. Choose folder...", command=self.choose_folder)
        self.btn_choose.pack(side="left")
        ttk.Label(r1, textvariable=self.dir_var).pack(side="left", padx=10)

        box = ttk.LabelFrame(top, text="2. What kind of duplicates to look for")
        box.pack(fill="x", pady=4)
        for i, mode in enumerate((MODE_EXACT, MODE_TEXT, MODE_PHOTO, MODE_VIDEO)):
            rb = ttk.Radiobutton(box, text=MODE_NAMES[mode], value=mode,
                                 variable=self.mode_var, command=self._on_mode)
            rb.grid(row=i, column=0, sticky="w", padx=8)
            ttk.Label(box, text=MODE_HELP[mode], foreground="#555555").grid(row=i, column=1, sticky="w", padx=8)
            if mode == MODE_PHOTO and not PIL_OK:
                rb.state(["disabled"])
                ttk.Label(box, text="(off: Pillow library not installed - see the top of the program)",
                          foreground="#aa3300").grid(row=i, column=2, sticky="w")

        r3 = ttk.Frame(top)
        r3.pack(fill="x", pady=2)
        ttk.Checkbutton(r3, text="Include subfolders", variable=self.sub_var).pack(side="left")
        ttk.Checkbutton(r3, text="Remember results (faster re-scans)",
                        variable=self.cache_var).pack(side="left", padx=14)
        ttk.Label(r3, text="Documents, photos, videos: match at least").pack(side="left")
        self.pct_box = ttk.Combobox(r3, values=MATCH_CHOICES, state="readonly", width=6,
                                    textvariable=self.pct_var)
        self.pct_box.pack(side="left", padx=4)
        ttk.Label(r3, text="(95% = very close, lower = looser)").pack(side="left")
        self._on_mode()

        r4 = ttk.Frame(top)
        r4.pack(fill="x", pady=2)
        self.btn_scan = ttk.Button(r4, text="3. Scan", command=self.start_scan)
        self.btn_scan.pack(side="left")
        self.btn_stop = ttk.Button(r4, text="Stop", command=self.stop_scan)
        self.btn_stop.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(r4, mode="determinate", length=320)
        self.progress.pack(side="left", padx=10)
        self.status_lbl = ttk.Label(top, textvariable=self.status_var, anchor="w", justify="left")
        self.status_lbl.pack(fill="x", pady=(2, 0))
        top.bind("<Configure>", lambda e: self.status_lbl.configure(wraplength=max(300, e.width - 20)))

        # Bottom blocks first, so a short window squeezes the list, not the buttons.
        bot = ttk.Frame(self.root)
        bot.pack(side="bottom", fill="x", **pad)
        self.btn_leave = ttk.Button(bot, text="Leave selected rows alone", command=self.leave_selected)
        self.btn_leave.pack(side="left")
        self.btn_move = ttk.Button(bot, text="4. Move all listed copies to the review folder",
                                   command=self.move_all)
        self.btn_move.pack(side="left", padx=12)
        self.btn_open = ttk.Button(bot, text="Open review folder", command=self.open_review)
        self.btn_open.pack(side="right")
        self.btn_undo = ttk.Button(bot, text="Undo last move", command=self.undo)
        self.btn_undo.pack(side="right", padx=8)

        sumrow = ttk.Frame(self.root)
        sumrow.pack(side="bottom", fill="x", padx=10)
        ttk.Label(sumrow, textvariable=self.summary_var, anchor="w").pack(fill="x")

        prev = ttk.Frame(self.root)
        prev.pack(side="bottom", fill="x", **pad)
        left = ttk.LabelFrame(prev, text="WILL BE MOVED  (click a row above to preview)")
        right = ttk.LabelFrame(prev, text="STAYS PUT  (the better copy)")
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))
        self.img_left = self._preview_slot(left)
        self.img_right = self._preview_slot(right)

        mid = ttk.Frame(self.root)
        mid.pack(side="top", fill="both", expand=True, **pad)
        cols = ("grp", "dup", "size", "date", "how", "keep")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended", height=6)
        heads = [("grp", "Group", 60, "center", False), ("dup", "Copy (will be moved)", 330, "w", True),
                 ("size", "Size", 80, "e", False), ("date", "Date", 90, "center", False),
                 ("how", "How it matches", 190, "w", False), ("keep", "Better copy (stays)", 330, "w", True)]
        for cid, text, width, anc, stretch in heads:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=width, anchor=anc, stretch=stretch)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_preview)

    @staticmethod
    def _preview_slot(parent):
        box = ttk.Frame(parent, width=370, height=200)
        box.pack(padx=6, pady=6)
        box.pack_propagate(False)
        lbl = ttk.Label(box, anchor="center", justify="left")
        lbl.place(relx=0.5, rely=0.5, anchor="center")
        return lbl

    # ----- helpers -----
    def _rel(self, path):
        base = self.scan_root or self.root_dir
        try:
            return os.path.relpath(path, base)
        except ValueError:
            return path

    def _refresh(self):
        def st(on):
            return "normal" if on else "disabled"
        idle = not self.busy
        review = os.path.join(self.root_dir, REVIEW_NAME) if self.root_dir else None
        has_log = bool(review) and os.path.isfile(os.path.join(review, LOG_NAME))
        self.btn_choose.configure(state=st(idle))
        self.btn_scan.configure(state=st(idle and bool(self.root_dir)))
        self.btn_stop.configure(state=st(self.busy))
        self.btn_move.configure(state=st(idle and bool(self.rows)))
        self.btn_leave.configure(state=st(idle and bool(self.rows)))
        self.btn_undo.configure(state=st(idle and has_log))
        self.btn_open.configure(state=st(bool(review) and os.path.isdir(review)))

    def _on_mode(self):
        mode = self.mode_var.get()
        on = mode in (MODE_PHOTO, MODE_TEXT, MODE_VIDEO)
        self.pct_box.configure(state="readonly" if on else "disabled",
                               values=VIDEO_MATCH_CHOICES if mode == MODE_VIDEO else MATCH_CHOICES)
        if mode == MODE_VIDEO and self.pct_var.get() not in VIDEO_MATCH_CHOICES:
            self.pct_var.set(VIDEO_MATCH_CHOICES[-1])

    def _update_summary(self):
        if not self.rows:
            self.summary_var.set("")
            return
        total = sum(d.size for d, _k, _w in self.rows.values())
        self.summary_var.set("%d copies listed - moving them frees %s in this folder"
                             % (len(self.rows), fmt_size(total)))

    def _clear_results(self):
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.img_left.configure(image="", text="")
        self.img_right.configure(image="", text="")
        self._img_refs = []
        self.summary_var.set("")

    def _keep_only(self, paths):
        for iid in list(self.rows):
            if self.rows[iid][0].path not in paths:
                del self.rows[iid]
                self.tree.delete(iid)
        self.img_left.configure(image="", text="")
        self.img_right.configure(image="", text="")
        self._img_refs = []
        self._update_summary()

    def on_close(self):
        self.cancel.set()
        self.root.destroy()

    # ----- actions -----
    def choose_folder(self):
        d = filedialog.askdirectory(title="Choose the folder to check")
        if not d:
            return
        self.root_dir = os.path.normpath(d)
        self.dir_var.set(self.root_dir)
        self._clear_results()
        self.progress.configure(value=0)
        self.status_var.set("Step 2: choose the kind of duplicates, then click Scan.")
        self._refresh()

    def start_scan(self):
        if not self.root_dir or self.busy:
            return
        pct = float(self.pct_var.get().rstrip("%"))
        self._clear_results()
        self.cancel.clear()
        self.busy = True
        self.scan_root = self.root_dir
        self.scan_mode = self.mode_var.get()
        self.progress.configure(value=0, maximum=100)
        self.status_var.set("Looking through the folder...")
        self._refresh()
        threading.Thread(target=scan_job, daemon=True, args=(
            self.root_dir, self.sub_var.get(), self.scan_mode, pct,
            self.cache_var.get(), self.cancel, self.q)).start()
        self.root.after(100, self._poll)

    def stop_scan(self):
        self.cancel.set()
        self.status_var.set("Stopping...")

    def _poll(self):
        finished, latest = False, None
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    latest = msg
                elif kind == "done":
                    self._finish(msg[1], msg[2])
                    finished = True
                elif kind == "cancelled":
                    self.status_var.set("Stopped. Nothing was moved.")
                    finished = True
                elif kind == "error":
                    self.status_var.set("Something went wrong.")
                    messagebox.showerror("Duplicate File Finder", msg[1])
                    finished = True
        except queue.Empty:
            pass
        if latest is not None and not finished:
            _, label, done, total = latest
            if total:
                self.progress.configure(maximum=total, value=done)
                self.status_var.set("%s... %d of %d" % (label, done, total))
            else:
                self.status_var.set(label + "...")
        if finished:
            self.busy = False
            self._refresh()
        else:
            self.root.after(100, self._poll)

    def _finish(self, groups, stats):
        for gi, g in enumerate(groups, 1):
            for dup, how in g.copies:
                iid = self.tree.insert("", "end", values=(
                    gi, self._rel(dup.path), fmt_size(dup.size), fmt_date(dup.mtime),
                    how, self._rel(g.keeper.path)))
                self.rows[iid] = (dup, g.keeper, how)
        self.progress.configure(value=self.progress.cget("maximum"))
        notes = list(stats.get("notes", []))
        if stats["failed"]:
            notes.append("%d file(s) could not be read and were skipped" % stats["failed"])
        if stats["heic"]:
            notes.append("%d HEIC photo(s) skipped - optional add-on needed" % stats["heic"])
        tail = ("  (" + "; ".join(notes) + ")") if notes else ""
        if stats["files"] == 0:
            self.status_var.set("Nothing to check in that folder for this kind of file." + tail)
        elif self.rows:
            self.status_var.set("Checked %d files: found %d copies in %d group(s). Review, then move them."
                                % (stats["files"], len(self.rows), len(groups)) + tail)
        else:
            self.status_var.set("Checked %d files: no duplicates found." % stats["files"] + tail)
        self._update_summary()

    def _thumb(self, path, maxsize=(350, 180)):
        try:
            from PIL import ImageTk
            with Image.open(path) as im:
                im.draft("RGB", (maxsize[0] * 2, maxsize[1] * 2))
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail(maxsize)
                return ImageTk.PhotoImage(im)
        except Exception:
            return None

    def _video_thumb(self, path, maxsize=(350, 180)):
        try:
            from PIL import ImageTk
            with Image.open(io.BytesIO(video_still(path, maxsize))) as im:
                return ImageTk.PhotoImage(im.convert("RGB"))
        except Exception:
            return None

    def _show(self, label, path):
        ext = os.path.splitext(path)[1].lower()
        img = None
        if PIL_OK and ext in photo_exts():
            img = self._thumb(path)
        elif PIL_OK and ext in VIDEO_EXTS:
            img = self._video_thumb(path)
        if img is not None:
            self._img_refs.append(img)
            label.configure(image=img, text="")
            return
        if ext in TEXT_EXTS | DOC_EXTS:
            label.configure(image="", text=text_snippet(path), wraplength=360)
            return
        label.configure(image="", text="(no picture or text to show - compared by its contents)")

    def show_preview(self, _event=None):
        sel = self.tree.selection()
        if not sel or sel[0] not in self.rows:
            return
        dup, keeper, _how = self.rows[sel[0]]
        self._img_refs = []
        self._show(self.img_left, dup.path)
        self._show(self.img_right, keeper.path)

    def leave_selected(self):
        for iid in self.tree.selection():
            self.rows.pop(iid, None)
            self.tree.delete(iid)
        self._update_summary()
        self._refresh()

    def move_all(self):
        if not self.rows or self.busy:
            return
        review = os.path.join(self.scan_root, REVIEW_NAME)
        n = len(self.rows)
        if not messagebox.askyesno(
                "Move copies?",
                "Move %d copy/copies into this folder:\n\n%s\n\n"
                "Nothing is deleted. The better copy of each group stays where it is.\n"
                "You can reverse this with \"Undo last move\"." % (n, review)):
            return
        moves = [(d.path, k.path, how) for d, k, how in self.rows.values()]

        def prog(i, total):
            self.status_var.set("Moving... %d of %d" % (i, total))
            self.root.update_idletasks()

        res = move_copies(self.scan_root, moves, self.scan_mode, prog)
        left_behind = {p for p, _why in res["failed"]}
        if left_behind:
            self._keep_only(left_behind)
        else:
            self._clear_results()
        msg = "Moved %d copy/copies to the review folder." % res["moved"]
        if res["skipped"]:
            msg += " %d were already gone and were skipped." % res["skipped"]
        self.status_var.set(msg + (" %d could NOT be moved - see the message." % len(left_behind)
                                   if left_behind else ""))
        self._refresh()
        if left_behind:
            detail = msg + "\n\n%d copy/copies were NOT moved and are still in place:" % len(left_behind)
            for p, why in res["failed"][:4]:
                detail += "\n  - %s: %s" % (os.path.basename(p), why)
            if len(res["failed"]) > 4:
                detail += "\n  - ...and %d more" % (len(res["failed"]) - 4)
            detail += ("\n\nUsually another program has the files open. Close photo viewers and video players, "
                       "pause OneDrive/Dropbox syncing, then press step 4 again.")
            messagebox.showwarning("Duplicate File Finder", detail)
        else:
            messagebox.showinfo("Duplicate File Finder", msg + "\n\nLog: " +
                                os.path.join(review, LOG_NAME))

    def undo(self):
        if not self.root_dir or self.busy:
            return
        if not messagebox.askyesno("Undo last move?",
                                   "Put the copies from the most recent move back where they came from?"):
            return
        res = undo_last_move(self.root_dir)
        self._clear_results()
        if res is None:
            msg = "There is nothing left to undo."
        else:
            msg = "Put %d file(s) back." % res["restored"]
            if res["renamed"]:
                msg += (" %d had a file with the same name there now, so they were restored with "
                        "\"(restored)\" added to the name." % res["renamed"])
            if res["missing"]:
                msg += " %d were no longer in the review folder." % res["missing"]
            if res["failed"]:
                msg += " %d could not be put back and are still in the review folder." % len(res["failed"])
                for p, why in res["failed"][:4]:
                    msg += "\n  - %s: %s" % (os.path.basename(p), why)
        self.status_var.set(msg.split("\n")[0])
        self._refresh()
        if res is not None and res["failed"]:
            messagebox.showwarning("Duplicate File Finder", msg)
        else:
            messagebox.showinfo("Duplicate File Finder", msg)

    def open_review(self):
        if not self.root_dir:
            return
        d = os.path.join(self.root_dir, REVIEW_NAME)
        if not os.path.isdir(d):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(d)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", d])
            else:
                subprocess.Popen(["xdg-open", d])
        except Exception as exc:
            messagebox.showerror("Duplicate File Finder", "Could not open the folder:\n%s" % exc)


# ---------------------------------------------------------------------------
# Self-test: builds sample files, runs every mode, checks the answers,
# then moves one copy and undoes it. Run:  py Duplicate_File_Finder.py --selftest
# ---------------------------------------------------------------------------
def _members(g):
    return [g.keeper] + [f for f, _ in g.copies]


def _names(groups):
    return {frozenset(os.path.basename(f.path) for f in _members(g)) for g in groups}


def self_test():
    import random
    failures = []
    passes = [0]

    def check(label, ok):
        print(("PASS  " if ok else "FAIL  ") + label)
        passes[0] += 1
        if not ok:
            failures.append(label)

    quiet = lambda *a: None
    with tempfile.TemporaryDirectory() as tmp, tempfile.TemporaryDirectory() as cache_dir:
        def write(name, data):
            p = os.path.join(tmp, name)
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "wb") as f:
                f.write(data)

        rnd = random.Random(1)
        big = rnd.randbytes(300_000)
        big_changed = big[:150_000] + b"X" + big[150_001:]   # same size, same start and end
        write("d1.bin", big)
        write("sub/d2.bin", big)
        write("d3.bin", big_changed)
        write("e1.bin", b"small same")
        write("e2.bin", b"small same")
        write("e3.bin", b"small diff")
        write("a.txt", b"hello\nworld\n")
        write("sub/b.txt", b"hello\r\nworld   \n\n\n")
        write("c.txt", b"something else\n")
        if PIL_OK:
            from PIL import Image as PImage
            pw, ph = 400, 300
            buf = bytearray()
            for y in range(ph):
                for x in range(pw):
                    buf += bytes((x * 255 // (pw - 1), y * 255 // (ph - 1), ((x + y) // 3) % 256))
            base = PImage.frombytes("RGB", (pw, ph), bytes(buf))
            base.save(os.path.join(tmp, "p1.png"))
            base.resize((200, 150), _BOX).save(os.path.join(tmp, "p2.jpg"), quality=85)
            PImage.frombytes("RGB", (pw, ph), rnd.randbytes(pw * ph * 3)).save(os.path.join(tmp, "p3.png"))

        cancel = threading.Event()
        # EXACT
        files, _ = collect_files(tmp, True, MODE_EXACT)
        out = find_groups(MODE_EXACT, files, NoCache(), 95, quiet, cancel)
        groups = out[0]
        check("exact: finds the two identical pairs and nothing else",
              _names(groups) == {frozenset({"d1.bin", "d2.bin"}), frozenset({"e1.bin", "e2.bin"})})
        check("exact: big file with a changed middle is NOT grouped",
              all("d3.bin" not in n for n in _names(groups)))
        g_d = [g for g in groups if "d1.bin" in {os.path.basename(f.path) for f in _members(g)}][0]
        check("exact: keeper of the pair is d1.bin (oldest, shortest path)",
              os.path.basename(g_d.keeper.path) == "d1.bin")
        # TEXT: a Word file with the same words but bold formatting must match
        with zipfile.ZipFile(os.path.join(tmp, "f.docx"), "w") as z:
            z.writestr("word/document.xml",
                       '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                       '<w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r>'
                       '<w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p></w:body></w:document>')
        files, _ = collect_files(tmp, True, MODE_TEXT)
        groups = find_groups(MODE_TEXT, files, NoCache(), 95, quiet, cancel)[0]
        check("text: a.txt, sub/b.txt and a bold f.docx with the same words are one group",
              _names(groups) == {frozenset({"a.txt", "b.txt", "f.docx"})})
        check("text: documents show 100% same words", bool(groups) and groups[0].copies[0][1] == "100% same words")
        # PHOTO
        if PIL_OK:
            files, _ = collect_files(tmp, True, MODE_PHOTO)
            groups = find_groups(MODE_PHOTO, files, NoCache(), 95, quiet, cancel)[0]
            check("photo: resized+recompressed copy grouped with the original",
                  _names(groups) == {frozenset({"p1.png", "p2.jpg"})})
            check("photo: the bigger picture is the keeper",
                  bool(groups) and os.path.basename(groups[0].keeper.path) == "p1.png")
        else:
            print("SKIP  photo checks (Pillow not installed)")
        # CACHE: second run must give the same answer and the cache must hold entries
        cache = FingerprintCache(folder=cache_dir)
        files, _ = collect_files(tmp, True, MODE_EXACT)
        first = _names(find_groups(MODE_EXACT, files, cache, 95, quiet, cancel)[0])
        second = _names(find_groups(MODE_EXACT, files, cache, 95, quiet, cancel)[0])
        d1 = [f for f in files if os.path.basename(f.path) == "d1.bin"][0]
        check("cache: remembers a fingerprint and re-scan gives the same answer",
              cache.get("exact", d1) is not None and first == second)
        cache.close()
        # MOVE + UNDO on the exact pair
        files, _ = collect_files(tmp, True, MODE_EXACT)
        g = [g for g in find_groups(MODE_EXACT, files, NoCache(), 95, quiet, cancel)[0]
             if "d1.bin" in {os.path.basename(f.path) for f in _members(g)}][0]
        copy_path = g.copies[0][0].path
        res = move_copies(tmp, [(copy_path, g.keeper.path, "identical")], MODE_EXACT)
        check("move: copy leaves its place and lands in the review folder",
              res["moved"] == 1 and not os.path.exists(copy_path)
              and os.path.exists(os.path.join(tmp, REVIEW_NAME, os.path.basename(copy_path))))
        und = undo_last_move(tmp)
        check("undo: copy is put back exactly where it was",
              und is not None and und["restored"] == 1 and os.path.exists(copy_path))

    # VIDEO: a smaller copy in another format must match; other videos and blank ones must not
    exe = find_ffmpeg()
    if exe:
        with tempfile.TemporaryDirectory() as vdir, tempfile.TemporaryDirectory() as vcache:
            def clip(name, source):
                _run_ffmpeg(exe, ["-loglevel", "error", "-f", "lavfi", "-i", source, "-t", "5", "-c:v", "mpeg4",
                                  os.path.join(vdir, name)])
            clip("v1.mp4", "testsrc2=size=320x240:rate=10")
            _run_ffmpeg(exe, ["-loglevel", "error", "-i", os.path.join(vdir, "v1.mp4"), "-vf", "scale=160:120",
                              "-c:v", "mpeg4", "-q:v", "6", os.path.join(vdir, "v2.mkv")])
            clip("v3.avi", "mandelbrot=size=320x240:rate=10")
            clip("blank1.mp4", "color=c=black:size=320x240:rate=10")
            clip("blank2.avi", "color=c=black:size=320x240:rate=10")
            with open(os.path.join(vdir, "broken.mp4"), "wb") as fh:
                fh.write(b"this is not a video")
            with open(os.path.join(vdir, "notes.txt"), "wb") as fh:
                fh.write(b"not a video either")
            files, _ = collect_files(vdir, True, MODE_VIDEO)
            check("video: only video files are picked up",
                  {os.path.basename(f.path) for f in files}
                  == {"v1.mp4", "v2.mkv", "v3.avi", "blank1.mp4", "blank2.avi", "broken.mp4"})
            groups, failed = find_groups(MODE_VIDEO, files, NoCache(), 90, quiet, cancel)
            check("video: smaller re-encoded copy grouped; other and blank videos are not",
                  _names(groups) == {frozenset({"v1.mp4", "v2.mkv"})})
            check("video: the bigger picture is the keeper",
                  bool(groups) and os.path.basename(groups[0].keeper.path) == "v1.mp4")
            check("video: a damaged file is counted as unreadable, not as a match", failed == 1)
            cache = FingerprintCache(folder=vcache)
            first = _names(find_groups(MODE_VIDEO, files, cache, 90, quiet, cancel)[0])
            v1 = [f for f in files if os.path.basename(f.path) == "v1.mp4"][0]
            remembered = _decode_video(cache.get("video", v1) or "")
            second = _names(find_groups(MODE_VIDEO, files, cache, 90, quiet, cancel)[0])
            check("video: fingerprints are remembered and a re-scan gives the same answer",
                  remembered is not None and first == second == _names(groups))
            cache.close()
        # SHORTENED copies and BLACK BARS
        with tempfile.TemporaryDirectory() as vdir:
            def vp_path(name):
                return os.path.join(vdir, name)

            def encode(name, *args):
                _run_ffmpeg(exe, ["-loglevel", "error"] + list(args) + ["-c:v", "mpeg4", "-q:v", "4", vp_path(name)])
            encode("long.mp4", "-f", "lavfi", "-i", "testsrc2=size=320x240:rate=10", "-t", "20")
            encode("start.mkv", "-ss", "1", "-i", vp_path("long.mp4"))
            encode("end.avi", "-i", vp_path("long.mp4"), "-t", "19")
            encode("both.mp4", "-ss", "0.5", "-i", vp_path("long.mp4"), "-t", "19")
            encode("bars.mp4", "-i", vp_path("long.mp4"), "-vf", "scale=320:160,pad=320:240:0:40")
            encode("sides.avi", "-i", vp_path("long.mp4"), "-vf", "scale=240:240,pad=320:240:40:0")
            encode("toomuch.mp4", "-ss", "5", "-i", vp_path("long.mp4"))
            encode("other.mp4", "-f", "lavfi", "-i", "mandelbrot=size=320x240:rate=10", "-t", "19")
            files, _ = collect_files(vdir, True, MODE_VIDEO)
            groups = find_groups(MODE_VIDEO, files, NoCache(), 85, quiet, cancel)[0]
            check("video: copies shortened at either end or both, and with black bars, are one group",
                  _names(groups) == {frozenset({"long.mp4", "start.mkv", "end.avi", "both.mp4", "bars.mp4", "sides.avi"})})
            check("video: the full-length copy without bars is the keeper",
                  bool(groups) and os.path.basename(groups[0].keeper.path) == "long.mp4")
            bars_only = [f for f in files if os.path.basename(f.path) in ("long.mp4", "bars.mp4")]
            groups_bars = find_groups(MODE_VIDEO, bars_only, NoCache(), 85, quiet, cancel)[0]
            check("video: black bars do not count as picture, so the copy without them stays",
                  len(groups_bars) == 1 and os.path.basename(groups_bars[0].keeper.path) == "long.mp4"
                  and groups_bars[0].keeper.pixels > groups_bars[0].copies[0][0].pixels)
            how = {os.path.basename(f.path): h for f, h in groups[0].copies} if groups else {}
            check("video: shortened copies say how much shorter, the others do not",
                  all("1.0s shorter" in how.get(n, "") for n in ("start.mkv", "end.avi", "both.mp4"))
                  and all("shorter" not in how.get(n, "shorter") for n in ("bars.mp4", "sides.avi")))
            check("video: with 5% cut off, a copy does not count at 95%",
                  not any("start.mkv" in n or "end.avi" in n or "both.mp4" in n
                          for n in _names(find_groups(MODE_VIDEO, files, NoCache(), 95, quiet, cancel)[0])))
    else:
        print("SKIP  video file checks (ffmpeg not installed yet: run one video scan first)")

    # BOUNDARY tests: thresholds are "at least", checked exactly on the boundary
    base_words = ["w%d" % i for i in range(194)]                  # 190 five-word phrases
    changed = ["zz%d" % i for i in range(10)] + base_words[10:]   # overlap is exactly 180/200 = 90%
    fa, fb = FileInfo("A.txt", 1, 0.0), FileInfo("B.txt", 1, 0.0)
    check("text boundary: exactly 90% overlap counts at 90%",
          len(group_text([(fa, base_words), (fb, changed)], 90)) == 1)
    check("text boundary: 90% overlap does not count at 91%",
          group_text([(fa, base_words), (fb, changed)], 91) == [])

    def pp(bits):
        return PhotoPrint(10, 10, bits, 128.0)
    pa = FileInfo("pa.png", 1, 0.0, 10, 10)
    pb = FileInfo("pb.png", 1, 0.0, 10, 10)
    check("photo boundary: 12 of 256 bits differ (95.3% alike) counts at 95%",
          len(group_similar([(pa, pp(0)), (pb, pp((1 << 12) - 1))], 95)) == 1)
    check("photo boundary: 13 bits differ (94.9% alike) does not count at 95%",
          group_similar([(pa, pp(0)), (pb, pp((1 << 13) - 1))], 95) == [])
    check("photo boundary: 51 bits differ (80.1% alike) counts at 80%",
          len(group_similar([(pa, pp(0)), (pb, pp((1 << 51) - 1))], 80)) == 1)
    check("photo boundary: 52 bits differ (79.7% alike) does not count at 80%",
          group_similar([(pa, pp(0)), (pb, pp((1 << 52) - 1))], 80) == [])


    def vp(frames, seconds=100.0, mean=128.0):
        return VideoPrint(10, 10, seconds, mean, frames)
    rnd = random.Random(2)
    pics = [rnd.randbytes(VIDEO_W * VIDEO_H) for _ in range(20)]
    flip = [bytes(255 - v for v in d) for d in pics]    # the opposite picture: nothing in common

    def match(same, total, pct):
        va = FileInfo("va.mp4", 1, 0.0, 10, 10)
        vb = FileInfo("vb.mp4", 1, 0.0, 10, 10)
        return len(group_videos([(va, vp(pics[:total])), (vb, vp(pics[:same] + flip[same:total]))], pct)) == 1
    for same, total, pct in ((19, 20, 95), (9, 10, 90), (17, 20, 85)):
        check("video boundary: %d of %d moments the same counts at %d%%" % (same, total, pct), match(same, total, pct))
    for same, total, pct in ((18, 19, 95), (17, 19, 90), (16, 19, 85)):
        check("video boundary: %d of %d moments the same does not count at %d%%" % (same, total, pct),
              not match(same, total, pct))
    va = FileInfo("va.mp4", 1, 0.0, 10, 10)
    vb = FileInfo("vb.mp4", 1, 0.0, 10, 10)
    check("video length: 0.2 seconds apart counts as the same length",
          len(group_videos([(va, vp(pics, 100.0)), (vb, vp(pics, 100.2))], 95)) == 1)
    check("video length: 0.3 seconds apart is left to the check for shortened copies",
          group_videos([(va, vp(pics, 100.0)), (vb, vp(pics, 100.3))], 95) == [])
    check("video blank: too few moments with a picture does not count",
          group_videos([(va, vp(pics[:2] + [None] * 8)), (vb, vp(pics[:2] + [None] * 8))], 85) == [])
    check("video brightness: same outlines but far brighter does not count",
          group_videos([(va, vp(pics, mean=60.0)), (vb, vp(pics, mean=160.0))], 85) == [])
    big = FileInfo("big.mp4", 900, 5.0, 10, 10)
    small = FileInfo("small.mp4", 100, 1.0, 10, 10)
    sharp = FileInfo("sharp.mp4", 100, 9.0, 20, 20)
    groups = group_videos([(small, vp(pics)), (big, vp(pics))], 95)
    check("video keeper: same picture size, the bigger file stays",
          bool(groups) and groups[0].keeper.path == "big.mp4")
    groups = group_videos([(big, vp(pics)), (sharp, vp(pics))], 95)
    check("video keeper: the bigger picture stays, even as the smaller file",
          bool(groups) and groups[0].keeper.path == "sharp.mp4")
    sample = VideoPrint(1920, 1080, 12.5, 90.25, pics[:VIDEO_FRAMES - 1] + [None],
                        pics[:2], [None, pics[3]], (0, 11, 160, 79))
    check("video memory: a fingerprint survives being stored and read back",
          _decode_video(_encode_video(sample)) == sample)
    plain = VideoPrint(1920, 1080, 12.5, 90.25, pics[:VIDEO_FRAMES])
    check("video memory: so does one of a short video without black bars",
          _decode_video(_encode_video(plain)) == plain)
    check("video memory: a stored value of the wrong size or an older kind is ignored",
          _decode_video("1,1,1.0,1.0,abcd") is None and _decode_video("") is None
          and _decode_video("1,1,1.0,1.0,0:0:160:90,abcd||") is None)

    # BLACK BARS, from made-up pictures: bright in the middle, black outside the given box
    def picture(left, top, right, bottom, seed=3):
        r = random.Random(seed)
        out = bytearray(VIDEO_GRAB_SIZE)
        for y in range(top, bottom):
            for x in range(left, right):
                out[y * VIDEO_GRAB_W + x] = r.randint(60, 255)
        return bytes(out)
    full = VIDEO_FULL_BOX
    check("bars: none found in a picture without any", find_bars([picture(*full)] * 4) == full)
    check("bars: top and bottom", find_bars([picture(0, 11, 160, 79)] * 4) == (0, 11, 160, 79))
    check("bars: left and right", find_bars([picture(20, 0, 140, 90)] * 4) == (20, 0, 140, 90))
    check("bars: on all four sides", find_bars([picture(20, 11, 140, 79)] * 4) == (20, 11, 140, 79))
    check("bars: a scene that is dark along one edge only has none",
          find_bars([picture(0, 30, 160, 90)] * 4) == full)
    check("bars: a dark top counts for no more than the true bar at the bottom",
          find_bars([picture(0, 30, 160, 79)] * 4) == (0, 11, 160, 79))
    subtitled = bytearray(picture(0, 11, 160, 79))
    subtitled[82 * VIDEO_GRAB_W + 50:82 * VIDEO_GRAB_W + 110] = b"\xff" * 60
    check("bars: subtitles inside a bar in some of the pictures do not hide it",
          find_bars([bytes(subtitled), bytes(subtitled), picture(0, 11, 160, 79), picture(0, 11, 160, 79)]) == (0, 11, 160, 79))
    check("bars: a blank picture says nothing", find_bars([bytes(VIDEO_GRAB_SIZE)] * 4) == full)
    check("bars: a picture almost all black is left whole", find_bars([picture(70, 40, 90, 50)] * 4) == full)
    inner = picture(20, 11, 140, 79)
    small = video_small(inner, (20, 11, 140, 79))
    check("bars: what is left fills the small picture", len(small) == VIDEO_W * VIDEO_H and min(small) > 24)

    # SHORTENED copies, with the answers of the video check made up (a remembered "this alike")
    class Remembered:
        def __init__(self, value):
            self.value = value

        def get(self, kind, f):
            return self.value

        def put_many(self, rows):
            pass

    def cut(items, pct, value="1.0"):
        return find_cut_copies(group_videos(items, pct), items, None, Remembered(value), pct, quiet, cancel)
    longer = FileInfo("longer.mp4", 500, 3.0, 10, 10)
    shorter = FileInfo("shorter.mp4", 900, 1.0, 10, 10)
    sharper = FileInfo("sharper.mp4", 90, 2.0, 40, 40)
    other = FileInfo("other.mp4", 90, 2.0, 10, 10)
    groups = cut([(shorter, vp(pics, 92.0)), (longer, vp(pics, 100.0))], 90)
    check("cut: 8% shorter counts at 90%, the longer copy stays, even as the smaller file",
          len(groups) == 1 and groups[0].keeper.path == "longer.mp4"
          and groups[0].copies[0][1] == "92.0% alike, 8.0s shorter")
    check("cut: 8% shorter does not count at 95%",
          cut([(shorter, vp(pics, 92.0)), (longer, vp(pics, 100.0))], 95) == [])
    check("cut: the score is the likeness times the share that is left",
          cut([(shorter, vp(pics, 92.0)), (longer, vp(pics, 100.0))], 90, "0.97") == [])
    groups = cut([(sharper, vp(pics, 99.0)), (longer, vp(pics, 100.0))], 90)
    check("cut: 1% shorter but a bigger picture: the sharper copy stays, and the other is called longer",
          groups[0].keeper.path == "sharper.mp4" and groups[0].copies[0][1] == "99.0% alike, 1.0s longer")
    check("cut: 8% shorter with a bigger picture: the longer copy still stays",
          cut([(sharper, vp(pics, 92.0)), (longer, vp(pics, 100.0))], 90)[0].keeper.path == "longer.mp4")
    check("cut: 1 second shorter out of 24 with a bigger picture: the longer copy still stays",
          cut([(sharper, vp(pics, 23.0)), (longer, vp(pics, 24.0))], 90)[0].keeper.path == "longer.mp4")
    groups = cut([(sharper, vp(pics, 99.5)), (longer, vp(pics, 100.0)), (shorter, vp(pics, 92.0))], 90)
    check("cut: when the sharper copy stays, the rest of the group stays with it",
          len(groups) == 1 and groups[0].keeper.path == "sharper.mp4"
          and {f.path: h.split(", ")[1] for f, h in groups[0].copies}
          == {"longer.mp4": "0.5s longer", "shorter.mp4": "7.5s shorter"})
    groups = cut([(longer, vp(pics, 100.0)), (shorter, vp(pics, 92.0)), (other, vp(pics, 92.1))], 90)
    check("cut: copies of the shorter one join the group of the longer one",
          len(groups) == 1 and groups[0].keeper.path == "longer.mp4" and len(groups[0].copies) == 2)
    check("cut: videos with nothing in common are not even looked at",
          cut([(longer, vp(pics[:10], 100.0)), (shorter, vp(flip[:10], 92.0))], 90) == [])
    looks = [{"frames": _with_signs(pics[:10]), "heads": _with_signs(pics[:2]), "tails": _with_signs(pics[2:4])},
             {"frames": _with_signs(flip[:10]), "heads": _with_signs(pics[:2]), "tails": _with_signs(flip[2:4])},
             {"frames": _with_signs(flip[:10]), "heads": _with_signs(flip[:2]), "tails": _with_signs(pics[2:4])},
             {"frames": _with_signs(flip[:4] + pics[4:7] + flip[7:10]), "heads": [], "tails": []}]
    check("cut: same moment counted from the start: nothing cut from the start",
          cut_ranges(looks[0], looks[1], 20.0) == [(0.0, 0.0)])
    check("cut: same moment counted from the end: all cut from the start",
          cut_ranges(looks[0], looks[2], 20.0) == [(20.0, 20.0)])
    check("cut: evenly spread moments that match: the search goes around them",
          cut_ranges(looks[0], looks[3], 20.0) == [(6.0, 16.0)])

    print()
    if failures:
        print("%d of %d checks FAILED: %s" % (len(failures), passes[0], "; ".join(failures)))
        return 1
    print("All %d checks passed." % passes[0])
    return 0


def main():
    if "--selftest" in sys.argv:
        sys.exit(self_test())
    if not HAVE_TK:
        print("This program needs Python's built-in window toolkit (Tkinter). Install Python from\n"
              "https://www.python.org/downloads/ (the Windows installer includes it).")
        sys.exit(1)
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()
    os._exit(0)   # window closed: stop any background scan at once


if __name__ == "__main__":
    main()
