#!/usr/bin/env python3
"""
Cross-Folder Compare
====================
Pick two folders (A and B). The program compares every photo and document in
folder A against every photo and document in folder B, and shows each pair that
matches, side by side, with file information and a picture or text preview.

  * Photos: pictures that look the same, even if one was resized or re-saved.
  * Documents: Word (.docx), OpenDocument (.odt), RTF, PDF, and plain text files
    that contain the same words. Formatting is ignored.
  * Photos are only compared with photos, and documents only with documents.

Match dropdown: 95%, 90%, 85%, 80%.
  The scan is done once at 80%, and the dropdown then shows or hides pairs
  instantly, so you do not have to rescan to change the setting.
  Every matching pair is listed, so a file in A with two copies in B shows twice.

Move the extra copies: pick ONE rule (keep the newer, older, bigger or smaller
  copy, or the sharper photo / longer document). It applies to every pair on the
  list at once, in both folders. The losing copy of each pair is MOVED - never
  deleted - into a folder named "REMOVED BY DUPLICATE SCAN" inside the folder it
  came from (Folder A or Folder B), keeping its subfolders. A tie (same date,
  same size...) moves nothing. A list of every move is saved in that folder, and
  "Undo the last move" puts the files back.

NOTHING IS EVER DELETED. Until you click "Move the extra copies", the program
only reads your files. You can save the list to a CSV file (it opens in Excel)
with the "Save list to CSV" button.

Windows 11:  double-click this file, or type in Command Prompt:
                 py Cross_Folder_Compare.py
             Self-test (checks the logic on sample files):
                 py Cross_Folder_Compare.py --selftest
The first run installs the free Pillow picture library (needs internet).
Without Pillow, documents still work; photos are skipped.
Optional PDF support:  py -m pip install pypdf
"""

import csv
import errno
import importlib
import importlib.util
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import xml.etree.ElementTree as ET
import zipfile
import zlib
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter import font as tkfont
    HAVE_TK = True
except Exception:
    HAVE_TK = False


# ---------------------------------------------------------------------------
# Picture library (Pillow). Only needed for photos.
# ---------------------------------------------------------------------------
def _ensure_pillow():
    if importlib.util.find_spec("PIL") is not None:
        return True
    if getattr(sys, "frozen", False):
        return False    # a built copy (the .exe from Setup) has Pillow inside, and no pip
    print("First-time setup: installing the Pillow picture library (needs internet)...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "pillow"])
    except Exception as exc:
        print("Could not install Pillow automatically: %s" % exc)
        print("Photo comparison is off until you run:  py -m pip install pillow")
        return False
    importlib.invalidate_caches()
    return importlib.util.find_spec("PIL") is not None


PIL_OK = _ensure_pillow()
ImageTk = None
if PIL_OK:
    from PIL import Image, ImageOps
    Image.MAX_IMAGE_PIXELS = None
    _BOX = getattr(getattr(Image, "Resampling", Image), "BOX")
    if HAVE_TK:
        try:
            from PIL import ImageTk  # noqa: F811  (needs tkinter)
        except Exception:
            ImageTk = None
    try:
        import pillow_heif  # optional add-on for iPhone HEIC photos
        pillow_heif.register_heif_opener()
        PHOTO_EXTRA = {".heic", ".heif"}
    except Exception:
        PHOTO_EXTRA = set()
else:
    PHOTO_EXTRA = set()


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
PHOTO_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"} | PHOTO_EXTRA
DOC_EXTS = {".docx", ".odt", ".rtf", ".pdf", ".txt", ".md", ".csv", ".htm", ".html"}
JUNK_NAMES = {"desktop.ini", "thumbs.db", ".ds_store"}
REMOVED_DIR = "REMOVED BY DUPLICATE SCAN"   # extra copies are moved in here
SKIP_DIR_NAMES = {"$recycle.bin", "system volume information", REMOVED_DIR.lower()}
SAME_TIME_SECONDS = 2.0       # dates this close count as the same (USB drives round to 2 s)
MOVE_LOG_FIELDS = ["status", "original place", "moved to", "kept copy", "rule"]
MOVED, PUT_BACK = "moved", "put back"
MATCH_CHOICES = ["95%", "90%", "85%", "80%"]
FLOOR = 80.0                  # scan once at this level; the dropdown filters the list
HASH_SIZE = 16
HASH_BITS = HASH_SIZE * HASH_SIZE
BRIGHTNESS_TOLERANCE = 25     # out of 255: stops "same outline, very different brightness"
SHINGLE = 5                   # words per phrase used to compare documents
LOG_NAME = "cross_folder_matches.csv"


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
class Item:
    path: str
    rel: str            # path relative to the folder it was found in
    folder: str         # the folder the user chose (A or B)
    kind: str           # "photo" or "doc"
    size: int
    mtime: float
    ok: bool = False    # True once it was read successfully
    note: str = ""      # why it was skipped, if it was
    width: int = 0
    height: int = 0
    dhash: int = 0
    mean: float = 0.0
    words: int = 0
    shingles: frozenset = frozenset()


def fmt_size(nbytes):
    if nbytes >= 1024 * 1024:
        return "%.1f MB" % (nbytes / (1024 * 1024))
    return "%d KB" % max(1, round(nbytes / 1024))


def fmt_date(ts):
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


def fmt_when(ts, seconds=False):
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M" + (":%S" if seconds else ""))


def detail(item):
    if item.kind == "photo":
        return "%d x %d" % (item.width, item.height) if item.width else "(not read)"
    return "{:,} words".format(item.words) if item.words else "(not read)"


# ---------------------------------------------------------------------------
# Finding files
# ---------------------------------------------------------------------------
def collect(folder, include_sub, kinds):
    """Every photo/document under folder (only the kinds asked for)."""
    out = []
    for cur, dirs, files in os.walk(folder):
        if include_sub:
            dirs[:] = [d for d in dirs
                       if not d.startswith(".") and not d.startswith("@")
                       and d.lower() not in SKIP_DIR_NAMES]
        else:
            dirs[:] = []
        for name in files:
            low = name.lower()
            if low in JUNK_NAMES or low.startswith("._"):
                continue
            ext = os.path.splitext(low)[1]
            if ext in PHOTO_EXTS:
                kind = "photo"
            elif ext in DOC_EXTS:
                kind = "doc"
            else:
                continue
            if kind not in kinds:
                continue
            full = os.path.join(cur, name)
            try:
                st = os.stat(full)
            except OSError:
                continue
            if st.st_size == 0:
                continue
            out.append(Item(full, os.path.relpath(full, folder), folder, kind,
                            st.st_size, st.st_mtime))
    return out


# ---------------------------------------------------------------------------
# Reading photos and documents
# ---------------------------------------------------------------------------
def photo_print(path):
    """(width, height, 256-bit look fingerprint, average brightness)."""
    with Image.open(path) as im:
        width, height = im.size
        im.draft("L", (128, 128))
        im = ImageOps.exif_transpose(im)
        if im.mode.startswith("I"):
            im = im.point(lambda v: v * (1.0 / 256.0))
        small = im.convert("L").resize((HASH_SIZE + 1, HASH_SIZE), _BOX)
    data = small.tobytes()
    stride = HASH_SIZE + 1
    bits = 0
    for row in range(HASH_SIZE):
        base = row * stride
        for col in range(HASH_SIZE):
            bits = (bits << 1) | (1 if data[base + col] > data[base + col + 1] else 0)
    return width, height, bits, sum(data) / len(data)


class PdfMissing(Exception):
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
    """Plain words of a document, with formatting dropped. Raises on failure."""
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
            raise PdfMissing()
        return " ".join((page.extract_text() or "") for page in PdfReader(path).pages)
    text = _read_text_file(path)
    if ext in (".htm", ".html"):
        text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", text)
        text = re.sub(r"<[^>]+>", " ", text)
    return text


def shingle_set(words):
    """Set of SHINGLE-word phrases (as numbers). Short texts count as one phrase."""
    if not words:
        return frozenset()
    if len(words) < SHINGLE:
        return frozenset([zlib.crc32(" ".join(words).encode("utf-8"))])
    return frozenset(zlib.crc32(" ".join(words[i:i + SHINGLE]).encode("utf-8"))
                     for i in range(len(words) - SHINGLE + 1))


def fingerprint(item):
    """Fill in the comparison data for one item. Never raises; sets item.ok."""
    try:
        if item.kind == "photo":
            w, h, bits, mean = photo_print(item.path)
            item.width, item.height, item.dhash, item.mean = w, h, bits, mean
        else:
            words = re.findall(r"\w+", extract_text(item.path).lower())
            item.words = len(words)
            item.shingles = shingle_set(words)
        item.ok = True
    except PdfMissing:
        item.note = "pdf"
    except Exception:
        item.note = "unreadable"


def run_parallel(items, progress, cancel):
    """Fingerprint all items on several threads. Returns False if Stop was pressed."""
    total = len(items)
    if total == 0:
        return True
    workers = min(8, (os.cpu_count() or 2) + 2)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(fingerprint, it) for it in items]
        for n, fut in enumerate(as_completed(futures), 1):
            if cancel.is_set():
                for f in futures:
                    f.cancel()
                return False
            fut.result()
            if n % 20 == 0 or n == total:
                progress("Reading files", n, total)
    return True


# ---------------------------------------------------------------------------
# Matching. Each returns (a_item, b_item, percent, how) for every A/B pair
# whose percent is at least FLOOR.
# ---------------------------------------------------------------------------
def match_photos(A, B, floor):
    """Look-alike pictures. Percent = share of the 256 fingerprint bits that agree."""
    max_dist = int(HASH_BITS * (100.0 - floor) / 100.0 + 1e-9)
    nb = max_dist + 1
    # Pigeonhole trick: two prints within max_dist differing bits must agree
    # exactly on at least one of (max_dist + 1) slices. So only those need checking.
    bounds = [(k * HASH_BITS) // nb for k in range(nb + 1)]
    bands = [(bounds[k], (1 << (bounds[k + 1] - bounds[k])) - 1) for k in range(nb)]
    tables = [defaultdict(list) for _ in range(nb)]
    for j, b in enumerate(B):
        for k, (shift, mask) in enumerate(bands):
            tables[k][(b.dhash >> shift) & mask].append(j)
    out = []
    for a in A:
        cand = set()
        for k, (shift, mask) in enumerate(bands):
            cand.update(tables[k].get((a.dhash >> shift) & mask, ()))
        for j in sorted(cand):
            b = B[j]
            d = popcount(a.dhash ^ b.dhash)
            if d <= max_dist and abs(a.mean - b.mean) <= BRIGHTNESS_TOLERANCE:
                out.append((a, b, 100.0 * (1 - d / HASH_BITS), "Photo"))
    return out


def match_docs(A, B, floor):
    """Documents. Percent = share of word-phrases the two documents have in common."""
    postings = defaultdict(list)
    for j, b in enumerate(B):
        for sh in b.shingles:
            postings[sh].append(j)
    out = []
    for a in A:
        if not a.shingles:
            continue
        shared = defaultdict(int)
        for sh in a.shingles:
            for j in postings.get(sh, ()):
                shared[j] += 1
        na = len(a.shingles)
        for j in sorted(shared):
            inter = shared[j]
            union = na + len(B[j].shingles) - inter
            if inter * 100 >= floor * union:          # exact whole-number test
                out.append((a, B[j], 100.0 * inter / union, "Document"))
    return out


def compare_job(dir_a, dir_b, include_sub, kinds, cancel, out_q):
    """Runs in a background thread. Reports through out_q."""
    def progress(label, done, total):
        out_q.put(("progress", label, done, total))
    try:
        notes = []
        if "photo" in kinds and not PIL_OK:
            kinds = kinds - {"photo"}
            notes.append("photos skipped: Pillow library not installed")
        A = collect(dir_a, include_sub, kinds)
        B = collect(dir_b, include_sub, kinds)
        stats = {"a": len(A), "b": len(B), "bad": 0, "pdf": 0, "notes": notes}
        if not A or not B:
            out_q.put(("done", [], stats))
            return
        if not run_parallel(A + B, progress, cancel):
            out_q.put(("cancelled",))
            return
        for it in A + B:
            if not it.ok:
                stats["bad"] += 1
                if it.note == "pdf":
                    stats["pdf"] += 1
        ok_a = [i for i in A if i.ok]
        ok_b = [i for i in B if i.ok]
        pairs = []
        pa = [i for i in ok_a if i.kind == "photo"]
        pb = [i for i in ok_b if i.kind == "photo"]
        if pa and pb:
            pairs += match_photos(pa, pb, FLOOR)
        da = [i for i in ok_a if i.kind == "doc"]
        db = [i for i in ok_b if i.kind == "doc"]
        if da and db:
            pairs += match_docs(da, db, FLOOR)
        pairs = [p for p in pairs if p[0].path != p[1].path]   # never pair a file with itself
        pairs.sort(key=lambda p: (-p[2], p[0].rel.lower(), p[1].rel.lower()))
        out_q.put(("done", pairs, stats))
    except Exception as exc:
        out_q.put(("error", "%s: %s" % (type(exc).__name__, exc)))


# ---------------------------------------------------------------------------
# Moving the extra copies. One rule decides, for every pair, which copy is
# moved. Files are only ever moved (renamed into REMOVED_DIR), never deleted.
# ---------------------------------------------------------------------------
KEEP_RULES = [   # (key, words in the dropdown, what a tie means)
    ("newer", "the newer copy (move the older)", "same date and time"),
    ("older", "the older copy (move the newer)", "same date and time"),
    ("bigger", "the bigger file (move the smaller)", "same size"),
    ("smaller", "the smaller file (move the bigger)", "same size"),
    ("sharper", "the sharper photo / longer document", "same pixels or words"),
]
RULE_WORDS = {key: words for key, words, _tie in KEEP_RULES}
RULE_TIE = {key: tie for key, _words, tie in KEEP_RULES}


def _quality(item):
    return item.width * item.height if item.kind == "photo" else item.words


def loser(a, b, rule):
    """The copy that `rule` moves out of the pair (a, b), or None for a tie."""
    if rule in ("newer", "older"):
        if abs(a.mtime - b.mtime) <= SAME_TIME_SECONDS:
            return None
        older, newer = (a, b) if a.mtime < b.mtime else (b, a)
        return older if rule == "newer" else newer
    if rule in ("bigger", "smaller"):
        if a.size == b.size:
            return None
        smaller, bigger = (a, b) if a.size < b.size else (b, a)
        return smaller if rule == "bigger" else bigger
    if rule == "sharper":
        qa, qb = _quality(a), _quality(b)
        if qa == qb:
            return None
        return a if qa < qb else b
    raise ValueError("unknown rule: %r" % (rule,))


def verdict(a, b, rule):
    """Words for the "Will move" column."""
    out = loser(a, b, rule)
    if out is None:
        return "Tie - stays"
    return "A file" if out is a else "B file"


def _key(path):
    return os.path.normcase(os.path.abspath(path))


def n_files(n):
    return "1 file" if n == 1 else "{:,} files".format(n)


def plan_moves(pairs, rule):
    """Which files `rule` moves. Returns (moves, ties).

    moves: list of (item_to_move, kept_item), each file at most once, in list order.
    ties:  number of pairs where the rule could not choose, so nothing moved.
    A file that loses in any pair is moved; the copy it lost to is better by
    the rule, so the best copy of every set of matches always stays.
    """
    moves, seen, ties = [], set(), 0
    for a, b, _pct, _how in pairs:
        out = loser(a, b, rule)
        if out is None:
            ties += 1
            continue
        kept = b if out is a else a
        if _key(out.path) not in seen:
            seen.add(_key(out.path))
            moves.append((out, kept))
    return moves, ties


def removed_root(folder):
    return os.path.join(folder, REMOVED_DIR)


def free_name(path):
    """`path`, or "name (2).ext", "name (3).ext"... if that name is taken."""
    if not os.path.lexists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 2
    while os.path.lexists("%s (%d)%s" % (stem, n, ext)):
        n += 1
    return "%s (%d)%s" % (stem, n, ext)


def _move_one(src, dst):
    """Move one file without ever overwriting another. Raises on failure."""
    if os.path.lexists(dst):
        raise FileExistsError(dst)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        os.rename(src, dst)
    except OSError as exc:
        # A different drive (for example a linked folder): copy, then remove the original.
        if getattr(exc, "winerror", None) == 17 or exc.errno == errno.EXDEV:
            shutil.move(src, dst)
        else:
            raise


def write_move_log(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=MOVE_LOG_FIELDS)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in MOVE_LOG_FIELDS})


def read_move_log(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if rows and not {"status", "original place", "moved to"} <= set(rows[0]):
        raise ValueError("This is not a list of moved files made by Cross-Folder Compare.")
    return rows


def move_extras(moves, rule, when=None):
    """Move each loser into REMOVED_DIR inside its own chosen folder, keeping subfolders.

    Returns (rows, failed, logs): rows = dicts as written to the move lists,
    failed = [(path, reason)], logs = paths of the move lists written (one per
    REMOVED_DIR used).
    """
    when = when or datetime.now()
    rows, failed, by_root = [], [], {}
    for item, kept in moves:
        root = removed_root(item.folder)
        dst = free_name(os.path.join(root, item.rel))
        try:
            _move_one(item.path, dst)
        except Exception as exc:
            failed.append((item.path, _why(exc)))
            continue
        row = {"status": MOVED, "original place": item.path, "moved to": dst,
               "kept copy": kept.path, "rule": "keep " + RULE_WORDS.get(rule, rule)}
        rows.append(row)
        by_root.setdefault(root, []).append(row)
    logs = []
    for root, root_rows in by_root.items():
        log = free_name(os.path.join(root, "moved list %s.csv" % when.strftime("%Y-%m-%d %H-%M-%S")))
        try:
            write_move_log(log, root_rows)
            logs.append(log)
        except Exception as exc:
            failed.append((log, "the list of moves could not be saved: " + _why(exc)))
    return rows, failed, logs


def undo_moves(logs):
    """Put back every file the move lists `logs` say is still moved.

    Never overwrites: a file whose old place is taken again stays where it is.
    Each list is rewritten with "put back" against the files returned.
    Returns (put_back, problems): paths put back, and [(path, reason)].
    """
    put_back, problems = [], []
    for log in logs:
        rows = read_move_log(log)
        changed = False
        for row in rows:
            if row.get("status") != MOVED:
                continue
            src, dst = row["moved to"], row["original place"]
            if not os.path.lexists(src):
                problems.append((dst, "it is no longer in %s" % os.path.dirname(src)))
                continue
            if os.path.lexists(dst):
                problems.append((dst, "another file with that name is in its old place now"))
                continue
            try:
                _move_one(src, dst)
            except Exception as exc:
                problems.append((dst, _why(exc)))
                continue
            row["status"] = PUT_BACK
            put_back.append(dst)
            changed = True
        if changed:
            write_move_log(log, rows)
            _remove_empty_dirs(os.path.dirname(log))
    return put_back, problems


def _remove_empty_dirs(root):
    """Remove empty subfolders left in REMOVED_DIR (never the folder itself)."""
    for cur, _dirs, _files in os.walk(root, topdown=False):
        if os.path.normcase(cur) != os.path.normcase(root):
            try:
                os.rmdir(cur)       # only succeeds when the folder is empty
            except OSError:
                pass


def _why(exc):
    if isinstance(exc, PermissionError):
        return "Windows would not let go of it (is it open in another program?)"
    if isinstance(exc, FileExistsError):
        return "a file with that name is already there"
    if isinstance(exc, FileNotFoundError):
        return "it is not there any more"
    return str(exc) or type(exc).__name__


# ---------------------------------------------------------------------------
# Previews
# ---------------------------------------------------------------------------
def text_snippet(path, max_words=55):
    try:
        text = extract_text(path)
    except PdfMissing:
        return "(PDF preview needs the free pypdf library)"
    except Exception:
        return "(cannot read this file)"
    shown = "\n".join(ln for ln in text.splitlines() if ln.strip())
    shown = " ".join(shown.split()[:max_words])
    if len(text.split()) > max_words:
        shown += " ..."
    return shown or "(no text found in this file)"


def make_thumb(path, maxsize=(380, 200)):
    if ImageTk is None:
        return None
    try:
        with Image.open(path) as im:
            im.draft("RGB", (maxsize[0] * 2, maxsize[1] * 2))
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail(maxsize)
            return ImageTk.PhotoImage(im)
    except Exception:
        return None


def work_area(root):
    """(left, top, width, height) of the screen space not covered by the taskbar."""
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - 60


def info_text(item):
    return ("%s\nFolder: %s\nSize: %s\nModified: %s\n%s"
            % (item.rel, os.path.dirname(item.path), fmt_size(item.size),
               fmt_when(item.mtime, seconds=True), detail(item)))


# ---------------------------------------------------------------------------
# The window
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Cross-Folder Compare")
        # Fit above the taskbar, so the bottom buttons are never hidden on short screens.
        left, top, width, height = work_area(root)
        h = max(480, min(880, height - 50))      # 50 = title bar and window edges
        x = left + max(0, (width - 1240) // 2)
        y = top + max(0, (height - h - 40) // 2)
        root.geometry("1240x%d+%d+%d" % (h, x, y))
        root.minsize(1000, min(640, h))
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=11)
        tkfont.nametofont("TkTextFont").configure(size=11)
        style = ttk.Style()
        style.configure("Treeview", rowheight=26, font=base)
        style.configure("Treeview.Heading", font=(base.actual("family"), 11, "bold"))

        self.dir_a = None
        self.dir_b = None
        self.all_pairs = []     # everything found at the 80% floor
        self.shown = {}         # tree id -> pair currently listed
        self.stats = None
        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self._imgs = []
        self.last_logs = []         # move lists written by the last move in this window
        self.undo_snapshot = None   # (dir_a, dir_b, pairs, stats) from before that move
        # Smaller previews on short screens, so the buttons and the list still fit.
        self.preview_h = 210 if root.winfo_screenheight() >= 1000 else 150

        self.label_a = tk.StringVar(value="(not chosen)")
        self.label_b = tk.StringVar(value="(not chosen)")
        self.sub_var = tk.BooleanVar(value=True)
        self.photo_var = tk.BooleanVar(value=True)
        self.doc_var = tk.BooleanVar(value=True)
        self.pct_var = tk.StringVar(value="95%")
        self.rule_var = tk.StringVar(value=RULE_WORDS["newer"])
        self.status_var = tk.StringVar(value="Step 1: choose Folder A.")
        self.summary_var = tk.StringVar(value="")

        self._build()
        self._refresh()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        self.cancel.set()
        self.root.destroy()

    def _build(self):
        pad = {"padx": 10, "pady": 3}
        top = ttk.Frame(self.root)
        top.pack(side="top", fill="x", **pad)

        r1 = ttk.Frame(top)
        r1.pack(fill="x", pady=2)
        self.btn_a = ttk.Button(r1, text="1. Choose Folder A...", command=lambda: self.choose("A"))
        self.btn_a.pack(side="left")
        ttk.Label(r1, textvariable=self.label_a).pack(side="left", padx=10)

        r2 = ttk.Frame(top)
        r2.pack(fill="x", pady=2)
        self.btn_b = ttk.Button(r2, text="2. Choose Folder B...", command=lambda: self.choose("B"))
        self.btn_b.pack(side="left")
        ttk.Label(r2, textvariable=self.label_b).pack(side="left", padx=10)

        r3 = ttk.Frame(top)
        r3.pack(fill="x", pady=2)
        ttk.Checkbutton(r3, text="Include subfolders", variable=self.sub_var).pack(side="left")
        ttk.Checkbutton(r3, text="Photos", variable=self.photo_var).pack(side="left", padx=12)
        ttk.Checkbutton(r3, text="Documents", variable=self.doc_var).pack(side="left")
        ttk.Label(r3, text="     Show matches at least").pack(side="left")
        self.pct_box = ttk.Combobox(r3, values=MATCH_CHOICES, state="readonly", width=6,
                                    textvariable=self.pct_var)
        self.pct_box.pack(side="left", padx=4)
        self.pct_box.bind("<<ComboboxSelected>>", lambda e: self.on_pct())

        r4 = ttk.Frame(top)
        r4.pack(fill="x", pady=2)
        self.btn_compare = ttk.Button(r4, text="3. Compare the two folders", command=self.start_compare)
        self.btn_compare.pack(side="left")
        self.btn_stop = ttk.Button(r4, text="Stop", command=self.stop_compare)
        self.btn_stop.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(r4, mode="determinate", length=320)
        self.progress.pack(side="left", padx=10)
        self.status_lbl = ttk.Label(top, textvariable=self.status_var, anchor="w", justify="left")
        self.status_lbl.pack(fill="x", pady=(2, 0))
        top.bind("<Configure>", lambda e: self.status_lbl.configure(wraplength=max(300, e.width - 20)))

        # Bottom blocks first, so a short window squeezes the list, not the buttons.
        bot = ttk.Frame(self.root)
        bot.pack(side="bottom", fill="x", **pad)
        self.btn_save = ttk.Button(bot, text="Save list to CSV...", command=self.save_csv)
        self.btn_save.pack(side="right")
        ttk.Label(bot, textvariable=self.summary_var, anchor="w").pack(side="left", fill="x", expand=True)

        movebox = ttk.LabelFrame(self.root, text="Move the extra copies (nothing is deleted)")
        movebox.pack(side="bottom", fill="x", **pad)
        mrow = ttk.Frame(movebox)
        mrow.pack(fill="x", padx=6, pady=4)
        ttk.Label(mrow, text="For every pair on the list, keep:").pack(side="left")
        self.rule_box = ttk.Combobox(mrow, values=[w for _k, w, _t in KEEP_RULES], state="readonly",
                                     width=36, textvariable=self.rule_var)
        self.rule_box.pack(side="left", padx=6)
        self.rule_box.bind("<<ComboboxSelected>>", lambda e: self.on_rule())
        self.btn_move = ttk.Button(mrow, text="4. Move the extra copies...", command=self.move_click)
        self.btn_move.pack(side="left", padx=(10, 6))
        self.btn_undo = ttk.Button(mrow, text="5. Undo the last move...", command=self.undo_click)
        self.btn_undo.pack(side="left")

        prev = ttk.Frame(self.root)
        prev.pack(side="bottom", fill="x", **pad)
        self.box_a = ttk.LabelFrame(prev, text="FOLDER A file (selected row)")
        self.box_b = ttk.LabelFrame(prev, text="FOLDER B file (selected row)")
        self.box_a.pack(side="left", fill="both", expand=True, padx=(0, 6))
        self.box_b.pack(side="left", fill="both", expand=True, padx=(6, 0))
        self.img_a, self.info_a = self._preview_slot(self.box_a)
        self.img_b, self.info_b = self._preview_slot(self.box_b)

        mid = ttk.Frame(self.root)
        mid.pack(side="top", fill="both", expand=True, **pad)
        cols = ("match", "kind", "a_rel", "a_size", "a_date", "a_detail",
                "b_rel", "b_size", "b_date", "b_detail", "move")
        heads = [
            ("match", "Match", 64, "center", False),
            ("kind", "Kind", 85, "w", False),
            ("a_rel", "Folder A file", 165, "w", True),
            ("a_size", "A size", 72, "e", False),
            ("a_date", "A date", 135, "center", False),
            ("a_detail", "A pixels/words", 120, "w", False),
            ("b_rel", "Folder B file", 165, "w", True),
            ("b_size", "B size", 72, "e", False),
            ("b_date", "B date", 135, "center", False),
            ("b_detail", "B pixels/words", 120, "w", False),
            ("move", "Will move", 95, "center", False),
        ]
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse", height=6)
        for cid, text, width, anc, stretch in heads:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=width, anchor=anc, stretch=stretch)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_preview)

    def _preview_slot(self, parent):
        frame = ttk.Frame(parent, width=400, height=self.preview_h)
        frame.pack(padx=6, pady=4)
        frame.pack_propagate(False)
        img = ttk.Label(frame, anchor="center", justify="left")
        img.place(relx=0.5, rely=0.45, anchor="center")
        info = ttk.Label(parent, anchor="w", justify="left", wraplength=470)
        info.pack(fill="x", padx=6, pady=(0, 4))
        return img, info

    # ----- state -----
    def _refresh(self):
        def st(on):
            return "normal" if on else "disabled"
        idle = not self.busy
        ready = bool(self.dir_a and self.dir_b)
        self.btn_a.configure(state=st(idle))
        self.btn_b.configure(state=st(idle))
        self.btn_compare.configure(state=st(idle and ready))
        self.btn_stop.configure(state=st(self.busy))
        self.btn_save.configure(state=st(idle and bool(self.shown)))
        self.btn_move.configure(state=st(idle and bool(self.shown)))
        self.btn_undo.configure(state=st(idle))   # with no move here, it opens a saved list
        self.pct_box.configure(state="readonly" if idle else "disabled")
        self.rule_box.configure(state="readonly" if idle else "disabled")

    def _clear_results(self):
        self.tree.delete(*self.tree.get_children())
        self.shown = {}
        self.all_pairs = []
        self.stats = None
        self._clear_preview()
        self.summary_var.set("")

    def _clear_preview(self):
        self._imgs = []
        self.img_a.configure(image="", text="")
        self.img_b.configure(image="", text="")
        self.info_a.configure(text="")
        self.info_b.configure(text="")

    # ----- actions -----
    def choose(self, which):
        d = filedialog.askdirectory(title="Choose Folder %s" % which)
        if not d:
            return
        d = os.path.normpath(d)
        other = self.dir_b if which == "A" else self.dir_a
        if other and os.path.normcase(os.path.abspath(d)) == os.path.normcase(os.path.abspath(other)):
            messagebox.showwarning("Cross-Folder Compare",
                                   "Folder A and Folder B are the same folder. Choose two different folders.")
            return
        if which == "A":
            self.dir_a = d
            self.label_a.set(d)
        else:
            self.dir_b = d
            self.label_b.set(d)
        self._clear_results()
        self.progress.configure(value=0)
        self.status_var.set("Step 3: click \"Compare the two folders\".")
        self._refresh()

    def start_compare(self):
        if not (self.dir_a and self.dir_b) or self.busy:
            return
        if not (self.photo_var.get() or self.doc_var.get()):
            messagebox.showwarning("Cross-Folder Compare", "Tick Photos, Documents, or both.")
            return
        kinds = set()
        if self.photo_var.get():
            kinds.add("photo")
        if self.doc_var.get():
            kinds.add("doc")
        self._clear_results()
        self.cancel.clear()
        self.busy = True
        self.progress.configure(value=0, maximum=100)
        self.status_var.set("Looking through the two folders...")
        self._refresh()
        threading.Thread(target=compare_job, daemon=True, args=(
            self.dir_a, self.dir_b, self.sub_var.get(), kinds, self.cancel, self.q)).start()
        self.root.after(100, self._poll)

    def stop_compare(self):
        self.cancel.set()
        self.status_var.set("Stopping...")

    def _poll(self):
        finished, latest = False, None
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    latest = msg
                elif msg[0] == "done":
                    self._finish(msg[1], msg[2])
                    finished = True
                elif msg[0] == "cancelled":
                    self.status_var.set("Stopped. Nothing was changed.")
                    finished = True
                elif msg[0] == "error":
                    self.status_var.set("Something went wrong.")
                    messagebox.showerror("Cross-Folder Compare", msg[1])
                    finished = True
        except queue.Empty:
            pass
        if latest is not None and not finished:
            _, label, done, total = latest
            self.progress.configure(maximum=max(1, total), value=done)
            self.status_var.set("%s... %d of %d" % (label, done, total))
        if finished:
            self.busy = False
            self._refresh()
        else:
            self.root.after(100, self._poll)

    def _finish(self, pairs, stats):
        self.all_pairs = pairs
        self.stats = stats
        self.progress.configure(value=self.progress.cget("maximum"))
        self._redraw()
        notes = list(stats["notes"])
        if stats["bad"]:
            notes.append("%d file(s) could not be read and were skipped" % stats["bad"])
        if stats["pdf"]:
            notes.append("%d PDF(s) skipped - install the free pypdf library" % stats["pdf"])
        tail = ("  (" + "; ".join(notes) + ")") if notes else ""
        if stats["a"] == 0 or stats["b"] == 0:
            self.status_var.set("One of the folders has no photos or documents to compare." + tail)
        else:
            self.status_var.set("Compared %d files in A with %d files in B. Pick a match level above."
                                % (stats["a"], stats["b"]) + tail)

    def on_pct(self):
        if self.all_pairs:
            self._redraw()

    def _threshold(self):
        return float(self.pct_var.get().rstrip("%"))

    def _redraw(self):
        self.tree.delete(*self.tree.get_children())
        self.shown = {}
        self._clear_preview()
        t = self._threshold()
        rows = [p for p in self.all_pairs if p[2] >= t - 1e-9]
        rule = self._rule()
        for a, b, pct, how in rows:
            iid = self.tree.insert("", "end", values=(
                "%.1f%%" % pct, how,
                a.rel, fmt_size(a.size), fmt_when(a.mtime), detail(a),
                b.rel, fmt_size(b.size), fmt_when(b.mtime), detail(b), verdict(a, b, rule)))
            self.shown[iid] = (a, b, pct, how)
        if self.all_pairs:
            self.summary_var.set("Showing %d matching pair(s) at %g%% or higher (%d found at 80%% or higher)."
                                 % (len(rows), t, len(self.all_pairs)))
        else:
            self.summary_var.set("")
        if self.stats is not None and self.stats["a"] and self.stats["b"]:
            if not rows:
                self.status_var.set("No pairs at %g%% or higher. Try a lower match level." % t)
        self._refresh()

    def _show(self, img_label, info_label, item):
        if item.kind == "photo" and PIL_OK:
            img = make_thumb(item.path, (380, self.preview_h - 10))
            if img is not None:
                self._imgs.append(img)
                img_label.place_configure(relx=0.5, rely=0.5, anchor="center")
                img_label.configure(image=img, text="")
                info_label.configure(text=info_text(item))
                return
        img_label.place_configure(relx=0.0, rely=0.0, anchor="nw")
        if item.kind == "doc":
            img_label.configure(image="", text=text_snippet(item.path), wraplength=385)
        else:
            img_label.configure(image="", text="(picture cannot be shown)")
        info_label.configure(text=info_text(item))

    def show_preview(self, _event=None):
        sel = self.tree.selection()
        if not sel or sel[0] not in self.shown:
            return
        a, b, _pct, _how = self.shown[sel[0]]
        self._imgs = []
        self._show(self.img_a, self.info_a, a)
        self._show(self.img_b, self.info_b, b)

    # ----- moving the extra copies -----
    def _rule(self):
        words = self.rule_var.get()
        for key, w, _tie in KEEP_RULES:
            if w == words:
                return key
        return "newer"

    def on_rule(self):
        rule = self._rule()
        for iid, (a, b, _pct, _how) in self.shown.items():
            self.tree.set(iid, "move", verdict(a, b, rule))
        if self.shown:
            self.status_var.set("The \"Will move\" column now shows what \"keep %s\" would move."
                                % RULE_WORDS[rule])

    def _ask(self, title, text, yes_text, no_text):
        """A question with worded buttons. True if the first button was clicked."""
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.resizable(False, False)
        answer = [False]

        def close(value):
            answer[0] = value
            win.destroy()
        ttk.Label(win, text=text, justify="left", wraplength=620).pack(padx=18, pady=(16, 10), anchor="w")
        row = ttk.Frame(win)
        row.pack(pady=(0, 14), padx=18, anchor="e")
        yes = ttk.Button(row, text=yes_text, padding=(12, 4), command=lambda: close(True))
        yes.pack(side="left", padx=6)
        ttk.Button(row, text=no_text, padding=(12, 4), command=lambda: close(False)).pack(side="left")
        win.bind("<Escape>", lambda e: close(False))
        win.protocol("WM_DELETE_WINDOW", lambda: close(False))
        win.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - win.winfo_width()) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - win.winfo_height()) // 3)
        win.geometry("+%d+%d" % (x, y))
        win.grab_set()
        yes.focus_set()
        self.root.wait_window(win)
        return answer[0]

    def _busy_cursor(self, on):
        self.root.configure(cursor="watch" if on else "")
        self.root.update_idletasks()

    def _where_lines(self, from_a, from_b):
        out = []
        for n, label, folder in ((from_a, "Folder A", self.dir_a), (from_b, "Folder B", self.dir_b)):
            if n:
                out += ["%s from %s, into:" % (n_files(n), label), "    " + removed_root(folder)]
        return out

    @staticmethod
    def _list_problems(problems, limit=8):
        lines = ["  %s - %s" % (os.path.basename(p), why) for p, why in problems[:limit]]
        if len(problems) > limit:
            lines.append("  ...and %d more" % (len(problems) - limit))
        return "\n".join(lines)

    def move_click(self):
        if self.busy:
            return
        if not self.shown:
            messagebox.showinfo("Cross-Folder Compare",
                                "Compare two folders first (steps 1 to 3). The extra copies are "
                                "picked from the list of matching pairs.")
            return
        rule = self._rule()
        moves, ties = plan_moves(list(self.shown.values()), rule)
        if not moves:
            messagebox.showinfo("Cross-Folder Compare",
                                "Every pair on the list is a tie (%s), so nothing would move.\n\n"
                                "Pick a different \"keep\" choice." % RULE_TIE[rule])
            return
        from_a = sum(1 for item, _k in moves if item.folder == self.dir_a)
        from_b = len(moves) - from_a
        lines = ["Move %s into a folder named %s?" % (n_files(len(moves)), REMOVED_DIR), ""]
        lines += self._where_lines(from_a, from_b)
        if ties:
            lines += ["", ("1 pair is a tie (%s) and stays where it is." if ties == 1 else
                           "%d pairs are ties (%s) and stay where they are.") % (
                               (RULE_TIE[rule],) if ties == 1 else (ties, RULE_TIE[rule]))]
        lines += ["", "Rule: keep %s." % RULE_WORDS[rule]]
        if self._threshold() < 90:
            lines += ["", "Careful: the list shows pairs only %g%% alike. Some may be different "
                          "versions of a document, not copies. Check them first, or pick 95%% in "
                          "\"Show matches at least\"." % self._threshold()]
        lines += ["", "Subfolders are kept. Nothing is deleted. A list of every move is saved in "
                      "that folder, and \"5. Undo the last move\" puts the files back."]
        if not self._ask("Move the extra copies", "\n".join(lines), "Move them", "Cancel"):
            self.status_var.set("Nothing was moved.")
            return
        self.status_var.set("Moving %d file(s)..." % len(moves))
        self._busy_cursor(True)
        try:
            rows, failed, logs = move_extras(moves, rule)
        finally:
            self._busy_cursor(False)
        moved = {_key(r["original place"]) for r in rows}
        if rows:
            self.undo_snapshot = (self.dir_a, self.dir_b, list(self.all_pairs), self.stats)
            self.last_logs = logs
            self.all_pairs = [p for p in self.all_pairs
                              if _key(p[0].path) not in moved and _key(p[1].path) not in moved]
            self._redraw()
        done_a = sum(1 for item, _k in moves if item.folder == self.dir_a and _key(item.path) in moved)
        done_b = len(rows) - done_a
        msg = ["Moved %s." % n_files(len(rows)), ""] + self._where_lines(done_a, done_b)
        if logs:
            msg += ["", "A list of the moves is saved there as \"%s\"." % os.path.basename(logs[0]),
                    "To put the files back, click \"5. Undo the last move\"."]
        if failed:
            msg += ["", "%d could not be moved and stayed where they were:" % len(failed),
                    self._list_problems(failed)]
        self.status_var.set("Moved %s into %s.%s" % (
            n_files(len(rows)), REMOVED_DIR, ("  %d could not be moved." % len(failed)) if failed else ""))
        if not rows:
            messagebox.showerror("Cross-Folder Compare", "\n".join(msg))
            return
        if self._ask("Move the extra copies", "\n".join(msg), "Show the moved files", "OK"):
            try:
                os.startfile(os.path.dirname(logs[0]) if logs else removed_root(moves[0][0].folder))
            except Exception as exc:
                messagebox.showerror("Cross-Folder Compare", "Could not open the folder:\n%s" % exc)

    def undo_click(self):
        if self.busy:
            return
        logs = [log for log in self.last_logs if os.path.exists(log)]
        try:
            if logs:
                waiting = [r for log in logs for r in read_move_log(log) if r.get("status") == MOVED]
                if not waiting:
                    self.last_logs = []
                    messagebox.showinfo("Cross-Folder Compare",
                                        "The files from the last move are already back.\n\n"
                                        "To undo an older move, click \"5. Undo the last move\" "
                                        "again and pick its \"moved list\" file.")
                    return
                text = "Put back the %s moved by the last move?" % n_files(len(waiting))
            else:
                start = next((removed_root(d) for d in (self.dir_a, self.dir_b)
                              if d and os.path.isdir(removed_root(d))), self.dir_a or os.path.expanduser("~"))
                path = filedialog.askopenfilename(
                    title="No move was made in this window - pick a \"moved list\" file to put those files back",
                    initialdir=start,
                    filetypes=[("List of moved files", "moved list*.csv"), ("CSV file", "*.csv")])
                if not path:
                    self.status_var.set("Nothing was put back.")
                    return
                logs = [os.path.normpath(path)]
                waiting = [r for r in read_move_log(logs[0]) if r.get("status") == MOVED]
                if not waiting:
                    messagebox.showinfo("Cross-Folder Compare",
                                        "Everything in that list is already back in its old place.")
                    return
                text = "Put back the %s in this list?\n\n    %s" % (n_files(len(waiting)), logs[0])
        except Exception as exc:
            messagebox.showerror("Cross-Folder Compare", "Could not read the list of moved files:\n%s" % exc)
            return
        text += ("\n\nEach file goes back to the exact place it came from. If another file now has "
                 "its old name there, it is left in %s, so nothing is ever overwritten." % REMOVED_DIR)
        if not self._ask("Undo the move", text, "Put them back", "Cancel"):
            self.status_var.set("Nothing was put back.")
            return
        self._busy_cursor(True)
        try:
            put_back, problems = undo_moves(logs)
        except Exception as exc:
            put_back, problems = [], [("", str(exc))]
        finally:
            self._busy_cursor(False)
        was_last = logs == [log for log in self.last_logs if os.path.exists(log)]
        snap = self.undo_snapshot
        if was_last and not problems and snap and snap[0] == self.dir_a and snap[1] == self.dir_b:
            self.all_pairs, self.stats = snap[2], snap[3]
            self.undo_snapshot = None
            self._redraw()
            tail = ""
        else:
            tail = "  Click \"3. Compare the two folders\" to list them again." if put_back else ""
        if was_last and not problems:
            self.last_logs = []
        msg = "Put back %s where %s came from." % (n_files(len(put_back)),
                                                     "it" if len(put_back) == 1 else "they")
        if problems:
            msg += ("\n\n%d could not be put back and stayed in %s:\n%s"
                    % (len(problems), REMOVED_DIR, self._list_problems(problems)))
        self.status_var.set(msg.split("\n")[0] + tail)
        (messagebox.showwarning if problems else messagebox.showinfo)("Cross-Folder Compare", msg + (
            "\n\n" + tail.strip() if tail else ""))

    def save_csv(self):
        if not self.shown or not self.dir_b:
            return
        path = filedialog.asksaveasfilename(
            title="Save the list", defaultextension=".csv", initialfile=LOG_NAME,
            filetypes=[("CSV file", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                rule = self._rule()
                w.writerow(["match_percent", "kind", "folder_A_file", "A_size_bytes", "A_modified",
                            "A_details", "folder_B_file", "B_size_bytes", "B_modified", "B_details",
                            "will_move (keep %s)" % RULE_WORDS[rule]])
                for a, b, pct, how in self.shown.values():
                    w.writerow(["%.1f" % pct, how, a.path, a.size, fmt_when(a.mtime, True), detail(a),
                                b.path, b.size, fmt_when(b.mtime, True), detail(b), verdict(a, b, rule)])
            messagebox.showinfo("Cross-Folder Compare", "Saved:\n%s" % path)
        except Exception as exc:
            messagebox.showerror("Cross-Folder Compare", "Could not save the file:\n%s" % exc)


# ---------------------------------------------------------------------------
# Self-test: builds sample files, checks every match rule on the exact limits.
#   py Cross_Folder_Compare.py --selftest
# ---------------------------------------------------------------------------
def self_test():
    import random
    fails, count = [], [0]

    def check(label, ok):
        count[0] += 1
        print(("PASS  " if ok else "FAIL  ") + label)
        if not ok:
            fails.append(label)

    with tempfile.TemporaryDirectory() as tmp:
        A_dir = os.path.join(tmp, "A")
        B_dir = os.path.join(tmp, "B")
        os.makedirs(os.path.join(A_dir, "sub"))
        os.makedirs(B_dir)
        rnd = random.Random(7)

        def put(folder, name, data):
            with open(os.path.join(folder, name), "wb") as f:
                f.write(data)

        # Documents: a.txt (A) and b.txt (B) share the same words; c.txt is different
        base_text = " ".join("word%d" % i for i in range(200))
        put(A_dir, "a.txt", base_text.encode())
        put(B_dir, "b.txt", ("Heading\n" + base_text + "\n").encode())
        put(A_dir, "sub/c.txt", b"completely unrelated text about boats and rivers " * 5)
        put(B_dir, "d.txt", b"nothing in common here at all whatsoever " * 5)

        # Photos: p1 (A) original; p2 (B) resized copy; p3 (B) random noise
        if PIL_OK:
            pw, ph = 400, 300
            buf = bytearray()
            for y in range(ph):
                for x in range(pw):
                    buf += bytes((x * 255 // (pw - 1), y * 255 // (ph - 1), ((x + y) // 3) % 256))
            base = Image.frombytes("RGB", (pw, ph), bytes(buf))
            base.save(os.path.join(A_dir, "p1.png"))
            base.resize((200, 150), _BOX).save(os.path.join(B_dir, "p2.jpg"), quality=85)
            Image.frombytes("RGB", (pw, ph), rnd.randbytes(pw * ph * 3)).save(os.path.join(B_dir, "p3.png"))

        files_a = collect(A_dir, True, {"photo", "doc"})
        files_b = collect(B_dir, True, {"photo", "doc"})
        check("collect: finds files in A (including subfolder) and in B",
              len(files_a) == (3 if PIL_OK else 2) and len(files_b) == (4 if PIL_OK else 2))
        for it in files_a + files_b:
            fingerprint(it)
        check("fingerprint: every file read successfully", all(it.ok for it in files_a + files_b))

        A = [i for i in files_a if i.kind == "doc"]
        B = [i for i in files_b if i.kind == "doc"]
        pairs = match_docs(A, B, FLOOR)
        check("docs: a.txt matches b.txt (same words, extra heading)",
              len(pairs) == 1 and os.path.basename(pairs[0][0].path) == "a.txt"
              and os.path.basename(pairs[0][1].path) == "b.txt")
        check("docs: unrelated files do not match",
              all(os.path.basename(p[0].path) != "c.txt" for p in pairs))
        check("docs: the match is shown as very high (over 95%%)", pairs and pairs[0][2] > 95)

        if PIL_OK:
            A = [i for i in files_a if i.kind == "photo"]
            B = [i for i in files_b if i.kind == "photo"]
            pp = match_photos(A, B, FLOOR)
            check("photos: p1.png matches p2.jpg (resized copy)",
                  len(pp) == 1 and os.path.basename(pp[0][1].path) == "p2.jpg")
            check("photos: random noise picture is not matched", all(
                os.path.basename(p[1].path) != "p3.png" for p in pp))

        # Boundary checks with fabricated items, so the exact limits are tested
        def photo(bits):
            return Item("x", "x", "A", "photo", 1, 0.0, ok=True, width=10, height=10,
                        dhash=bits, mean=128.0)
        a = photo(0)
        check("two copies in B of one A photo are both listed",
              len(match_photos([a], [photo(0), photo(1)], 95.0)) == 2)
        check("photo limit: 12 bits differ (95.3%) is listed at 95%",
              len(match_photos([a], [photo((1 << 12) - 1)], 95.0)) == 1)
        check("photo limit: 13 bits differ (94.9%) is not listed at 95%",
              match_photos([a], [photo((1 << 13) - 1)], 95.0) == [])
        check("photo limit: 51 bits differ (80.1%) is listed at 80%",
              len(match_photos([a], [photo((1 << 51) - 1)], 80.0)) == 1)
        check("photo limit: 52 bits differ (79.7%) is not listed at 80%",
              match_photos([a], [photo((1 << 52) - 1)], 80.0) == [])

        def doc(words):
            return Item("y", "y", "A", "doc", 1, 0.0, ok=True, words=len(words),
                        shingles=shingle_set(words))
        base_words = ["w%d" % i for i in range(194)]                  # 190 phrases
        changed = ["zz%d" % i for i in range(10)] + base_words[10:]   # overlap 180 of 200 = 90%
        check("doc limit: exactly 90% shared phrases is listed at 90%",
              len(match_docs([doc(base_words)], [doc(changed)], 90.0)) == 1)
        check("doc limit: exactly 90% shared phrases is not listed at 91%",
              match_docs([doc(base_words)], [doc(changed)], 91.0) == [])

        # Moving the extra copies: rule, tie, move with subfolders kept, and Undo
        M = os.path.join(tmp, "M")
        os.makedirs(os.path.join(M, "A", "sub"))
        os.makedirs(os.path.join(M, "B"))

        def mk(folder, rel, mtime):
            path = os.path.join(M, folder, rel)
            with open(path, "w") as f:
                f.write(rel)
            os.utime(path, (mtime, mtime))
            return Item(path, rel, os.path.join(M, folder), "doc", os.path.getsize(path), mtime, ok=True)
        old = mk("A", os.path.join("sub", "old.txt"), 1e9)
        new = mk("B", "new.txt", 1e9 + 100)
        twin_a, twin_b = mk("A", "twin.txt", 1e9), mk("B", "twin.txt", 1e9 + 1)
        moves, ties = plan_moves([(old, new, 99.0, "Document"), (twin_a, twin_b, 99.0, "Document")], "newer")
        check("move: keep the newer picks the older copy; 1 second apart is a tie",
              [m.path for m, _k in moves] == [old.path] and ties == 1)
        _rows, failed, logs = move_extras(moves, "newer")
        check("move: older copy is in REMOVED BY DUPLICATE SCAN, subfolder kept, list saved",
              not failed and len(logs) == 1 and not os.path.exists(old.path)
              and os.path.isfile(os.path.join(M, "A", REMOVED_DIR, "sub", "old.txt")))
        put_back, problems = undo_moves(logs)
        check("move: Undo puts the file back where it was",
              put_back == [old.path] and not problems and os.path.isfile(old.path))

    print()
    if fails:
        print("%d of %d checks FAILED: %s" % (len(fails), count[0], "; ".join(fails)))
        return 1
    print("All %d checks passed." % count[0])
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
    os._exit(0)   # window closed: stop any background reading at once


if __name__ == "__main__":
    main()
