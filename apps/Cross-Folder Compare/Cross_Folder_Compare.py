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

NOTHING IS MOVED OR DELETED. This program only reads your files. You can save
the list to a CSV file (it opens in Excel) with the "Save list to CSV" button.

Windows 11:  double-click this file, or type in Command Prompt:
                 py Cross_Folder_Compare.py
             Self-test (checks the logic on sample files):
                 py Cross_Folder_Compare.py --selftest
The first run installs the free Pillow picture library (needs internet).
Without Pillow, documents still work; photos are skipped.
Optional PDF support:  py -m pip install pypdf
"""

import csv
import importlib
import importlib.util
import os
import queue
import re
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
SKIP_DIR_NAMES = {"$recycle.bin", "system volume information"}
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


def info_text(item):
    return ("%s\nFolder: %s\nSize: %s\nModified: %s\n%s"
            % (item.rel, os.path.dirname(item.path), fmt_size(item.size),
               fmt_date(item.mtime), detail(item)))


# ---------------------------------------------------------------------------
# The window
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Cross-Folder Compare")
        root.geometry("1240x%d" % max(640, min(880, root.winfo_screenheight() - 100)))
        root.minsize(1000, 640)
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

        self.label_a = tk.StringVar(value="(not chosen)")
        self.label_b = tk.StringVar(value="(not chosen)")
        self.sub_var = tk.BooleanVar(value=True)
        self.photo_var = tk.BooleanVar(value=True)
        self.doc_var = tk.BooleanVar(value=True)
        self.pct_var = tk.StringVar(value="95%")
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
                "b_rel", "b_size", "b_date", "b_detail")
        heads = [
            ("match", "Match", 70, "center", False),
            ("kind", "Kind", 95, "w", False),
            ("a_rel", "Folder A file", 220, "w", True),
            ("a_size", "A size", 75, "e", False),
            ("a_date", "A date", 105, "center", False),
            ("a_detail", "A pixels/words", 125, "w", False),
            ("b_rel", "Folder B file", 220, "w", True),
            ("b_size", "B size", 75, "e", False),
            ("b_date", "B date", 105, "center", False),
            ("b_detail", "B pixels/words", 125, "w", False),
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

    @staticmethod
    def _preview_slot(parent):
        frame = ttk.Frame(parent, width=400, height=210)
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
        self.pct_box.configure(state="readonly" if idle else "disabled")

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
        for a, b, pct, how in rows:
            iid = self.tree.insert("", "end", values=(
                "%.1f%%" % pct, how,
                a.rel, fmt_size(a.size), fmt_date(a.mtime), detail(a),
                b.rel, fmt_size(b.size), fmt_date(b.mtime), detail(b)))
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
            img = make_thumb(item.path)
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
                w.writerow(["match_percent", "kind", "folder_A_file", "A_size_bytes", "A_modified",
                            "A_details", "folder_B_file", "B_size_bytes", "B_modified", "B_details"])
                for a, b, pct, how in self.shown.values():
                    w.writerow(["%.1f" % pct, how, a.path, a.size, fmt_date(a.mtime), detail(a),
                                b.path, b.size, fmt_date(b.mtime), detail(b)])
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
