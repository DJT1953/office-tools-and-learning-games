"""
photos_to_docx.py

Choose with a button whether each photo gets its own Word file or all of them
go into one, then pick a folder in a Windows Explorer dialog. Every picture in
it (any extension Pillow can open: jpg, jpeg, png, tif, bmp, gif, webp,
heic*...) is read in natural filename order and its text is recognized (OCR).

    One Word file per photo:      IMG_0001.jpg  ->  IMG_0001.docx
    All photos in one Word file:  "<folder name> - all photos.docx" in the
                                  folder, a page per photo, each named at the top

Text is placed in a similar position to where it appears in the photo:
left indent, vertical spacing, font size and gaps between columns are all
scaled from the picture onto a Letter page (landscape if the photo is wide).
Columns are held in place by tab stops, so they line up in any font and stay
lined up when the text is corrected. A page photographed sideways or upside
down and a slightly tilted photo are dealt with along the way.

The text is read by RapidOCR (through photo_text_reader.py, shared with Dating
Photographs), a free reader that works on this computer and reads print and
handwriting. Only if RapidOCR isn't installed does Tesseract OCR read the
photos instead (it then also deals with shadows and table borders).

A small window shows which photo is being read, with a Cancel button.

Corrections typed into earlier files aren't lost by accident: a photo's own
Word file is only replaced if you say so, and the all-photos file is never
replaced (a new one is numbered instead).

Requirements (one-time; if the Python running this lacks them, it offers to
install them):
    1. pillow (10.3 or newer), pytesseract and python-docx.
    2. rapidocr and onnxruntime, the reader (about 120 MB with its data).
       Without them, Tesseract OCR must be on the PC instead:
       https://github.com/UB-Mannheim/tesseract/wiki  (default install path is fine)
    3. pillow-heif for iPhone .heic photos (optional; the Setup leaves it out).
"""

import importlib.util
import os
import queue
import re
import shutil
import site
import subprocess
import sys
import threading
import traceback
import webbrowser
from functools import lru_cache
from math import floor, inf
from statistics import median, median_low

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# The "Check for updates" button opens this page, which has the newest version of every program.
UPDATES_PAGE = "https://djt1953.github.io/office-tools-and-learning-games/"

REQUIRED = {"PIL": "pillow", "pytesseract": "pytesseract", "docx": "python-docx"}
# The handwriting reader: offered too, but the program works without it.
HANDWRITING = {"rapidocr": "rapidocr>=3.10,<4", "onnxruntime": "onnxruntime"}
READER_DATA = ("ch_PP-OCRv5_det_mobile.onnx", "en_PP-OCRv5_rec_mobile.onnx", "PP-OCRv6_rec_small.onnx")
HERE = os.path.dirname(os.path.abspath(__file__))
PROGRESS_SEEN = True          # tests set this False so no window flashes on the screen
# Marks that the person said No to the handwriting reader, so it isn't offered every time.
DECLINED = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "Photos to Word",
                        "no handwriting reader")
TESSERACT_PLACES = (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"))
NOT_FILES = ("thumbs.db", "desktop.ini")       # Windows' own hidden files, never pictures
HEIC = (".heic", ".heif")


def dialog_root():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)        # keep dialogs in front of other windows
    return root


def show_crash(kind, value, tb):
    """A double-clicked script's window closes the moment it fails, taking the
    error with it, so show anything unexpected in a dialog."""
    text = "".join(traceback.format_exception(kind, value, tb))
    print(text)
    dialog_root()
    messagebox.showerror("Photos to Word", "Something went wrong:\n\n" + text[-1500:])


def console_python(executable=None):
    """The Python that shows a window. Run_PhotoText.bat starts pythonw.exe, which
    never opens one, so pip's progress would be invisible under it."""
    executable = executable or sys.executable
    folder, name = os.path.split(executable)
    if name.lower() == "pythonw.exe" and os.path.isfile(os.path.join(folder, "python.exe")):
        return os.path.join(folder, "python.exe")
    return executable


def tesseract_on_pc():
    """True if Tesseract OCR is installed where it usually goes (or on the PATH)."""
    return bool(shutil.which("tesseract")) or any(os.path.isfile(p) for p in TESSERACT_PLACES)


def reader_data_missing():
    """True if the handwriting reader is installed but its data isn't fetched yet."""
    spec = importlib.util.find_spec("rapidocr")
    if spec is None or not spec.submodule_search_locations:
        return False
    models = os.path.join(list(spec.submodule_search_locations)[0], "models")
    return not all(os.path.isfile(os.path.join(models, name)) for name in READER_DATA)


def install_missing():
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add whatever this one lacks. A built copy (the .exe from
    Setup) already has everything inside it and has no pip, so it never asks.
    Installing happens in a window of its own, so its progress can be seen."""
    if getattr(sys, "frozen", False):
        return
    missing = [pip for module, pip in REQUIRED.items()
               if importlib.util.find_spec(module) is None]
    handwriting = [pip for module, pip in HANDWRITING.items()
                   if importlib.util.find_spec(module) is None]
    if not missing and not handwriting and not reader_data_missing():
        return
    python = console_python()
    if missing:
        root = dialog_root()
        if not messagebox.askyesno(
                "Photos to Word",
                f"This copy of Python ({sys.version.split()[0]}) still needs: "
                f"{', '.join(missing)}.\n\n"
                "Download and install now? This only happens once and takes about a minute. "
                "A window shows the progress."):
            sys.exit()
        root.destroy()
        result = subprocess.run([python, "-m", "pip", "install", *missing],
                                creationflags=subprocess.CREATE_NEW_CONSOLE)
        if result.returncode:
            raise RuntimeError("Could not install " + ", ".join(missing)
                               + ". Check the internet connection and try again.")
        subprocess.run([python, "-m", "pip", "install", "pillow-heif"],   # optional
                       creationflags=subprocess.CREATE_NEW_CONSOLE)
    # Without Tesseract the reader is the only way to read anything, so it is always offered
    needed = not tesseract_on_pc()
    if (handwriting or reader_data_missing()) and (needed or not os.path.exists(DECLINED)):
        root = dialog_root()
        if needed:
            question = ("Photos to Word needs its text reader, a free one that works on this computer "
                        "and reads print and handwriting (about 120 MB to download, once).\n\nAdd it now?")
        else:
            question = ("Photos to Word can also read handwriting, with a free reader that works on this "
                        "computer (about 120 MB to download, once).\n\nAdd it now? Without it, printed "
                        "words are still read. (If you say No, this isn't asked again.)")
        wanted = messagebox.askyesno("Photos to Word", question)
        root.destroy()
        if not wanted:
            if not needed:          # needed, it is asked again next time
                try:
                    os.makedirs(os.path.dirname(DECLINED), exist_ok=True)
                    open(DECLINED, "w").close()
                except OSError:
                    pass
        else:
            if handwriting:
                subprocess.run([python, "-m", "pip", "install", *handwriting],
                               creationflags=subprocess.CREATE_NEW_CONSOLE)
            fetch = (f"import sys; sys.path.insert(0, {HERE!r}); import photo_text_reader as r; "
                     "print('Fetching the handwriting reader data...'); "
                     "r.Reader(use_tesseract=False).warm_up(); print('Done.')")
            subprocess.run([python, "-c", fetch], creationflags=subprocess.CREATE_NEW_CONSOLE)
    site.addsitedir(site.getusersitepackages())   # a first install may have just created it
    importlib.invalidate_caches()


def no_lost_output():
    """pythonw.exe and the built program have nowhere to print; give them somewhere."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


no_lost_output()
sys.excepthook = show_crash
if __name__ == "__main__":      # (tests import this file and have everything already)
    install_missing()

from PIL import (Image, ImageDraw, ImageFilter, ImageFont, ImageMath, ImageOps,
                 UnidentifiedImageError)
import pytesseract
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

try:
    from photo_text_reader import eight_bit
except ImportError:          # shared on its own, without photo_text_reader.py: Tesseract reads
    def eight_bit(picture):
        return picture

try:  # optional HEIC/HEIF support
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

# ---------------------------------------------------------------- settings
PAGE_W_IN, PAGE_H_IN = 8.5, 11.0     # Letter
MARGIN_IN = 0.5
FONT_NAME = "Times New Roman"
FONT_FILE = "times.ttf"              # the same font's file, used to measure text
MIN_PT, MAX_PT = 6.0, 28.0
MIN_CONFIDENCE = 30                  # drop OCR guesses below this (0-100)
OCR_LANG = "eng"
STRAY_MARKS = ".,;:'\"`~|_"             # alone, these are specks and leader dots, not words
FONT_PER_WORD_H = 1.45               # font size as a multiple of a word's height
LINE_HEIGHT = 1.15                   # line spacing as a multiple of font size
SAME_SIZE_RATIO = 1.6                # rows this close in height share a font size
RULE_LENGTHS = (25, 40, 60)          # table rules are tried at 1/this of the long side
COLUMN_GAP = 1.5                     # a gap this many word heights wide starts a column
MAX_OCR_PX = 8000                    # shrink anything longer than this before OCR
RAPID_LINE_TO_WORD = 1.9             # RapidOCR's line outline is this much taller than a word's letters
RAPID_LONGEST = 3200                 # photos are shrunk to this for RapidOCR; small print stays readable

Image.MAX_IMAGE_PIXELS = None        # or Pillow refuses photos over about 179 MP
METRICS = ImageFont.truetype(FONT_FILE, 1000)


# ---------------------------------------------------------------- helpers
def bundled_tesseract():
    """The copy of Tesseract packed inside the built program (the .exe from
    Setup), or None when running as a plain .py file."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        exe = os.path.join(base, "tesseract", "tesseract.exe")
        if os.path.isfile(exe):
            return exe
    return None


def find_tesseract():
    bundled = bundled_tesseract()
    if bundled:     # always use our own copy, whatever else is on the PC
        pytesseract.pytesseract.tesseract_cmd = bundled
        os.environ["TESSDATA_PREFIX"] = os.path.join(os.path.dirname(bundled), "tessdata")
        return
    if shutil.which("tesseract"):
        return
    for p in TESSERACT_PLACES:
        if os.path.isfile(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return
    raise RuntimeError(
        "Photos to Word can't read anything yet: its text reader isn't installed, and "
        "Tesseract OCR isn't on this PC either.\n\nClose this, start Photos to Word again and "
        "say Yes when it offers to add the reader. (Or install Tesseract OCR from\n"
        "https://github.com/UB-Mannheim/tesseract/wiki)")


@lru_cache(maxsize=1)
def tesseract_ready():
    """True once pytesseract is pointed at a Tesseract on this PC."""
    try:
        find_tesseract()
        return True
    except RuntimeError:
        return False


def natural_key(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


PHOTO_ENDINGS = (".jpg", ".jpeg", ".jfif", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp")


def is_picture(path):
    """False only for files Pillow doesn't recognize. A damaged picture, or an
    empty or broken file named like a photo, still counts, so that it gets
    reported rather than quietly skipped."""
    try:
        with Image.open(path):
            return True
    except UnidentifiedImageError:
        return path.lower().endswith(PHOTO_ENDINGS)
    except Exception:
        return True


def unreadable(path, error):
    """Plain words for why a photo couldn't be read."""
    try:
        if os.path.getsize(path) == 0:
            return "the file is empty"
    except OSError:
        return "it is no longer there"
    if isinstance(error, (UnidentifiedImageError, SyntaxError)) or "truncated" in str(error) \
            or "broken data" in str(error):
        return "the file is damaged, or isn't really a picture"
    return str(error)


def open_image(path):
    """Return an upright RGB image and the number of pages in the file."""
    with Image.open(path) as img:
        pages = img.n_frames if img.format == "TIFF" else 1
        img = eight_bit(ImageOps.exif_transpose(img))   # honor phone rotation; a 16-bit scan stays dark
        if img.mode in ("RGBA", "LA", "PA") or "transparency" in img.info:
            # transparent areas would turn black and hide dark text
            rgba = img.convert("RGBA")
            img = Image.alpha_composite(Image.new("RGBA", rgba.size, "white"), rgba)
        return img.convert("RGB"), pages


def even_out_lighting(gray):
    """Divide out shadows and falloff so the paper is one brightness throughout."""
    w, h = gray.size
    shrink = 1000 / max(w, h)
    small = gray.resize((max(1, round(w * shrink)), max(1, round(h * shrink))), Image.BILINEAR)
    # a max filter erases the dark text, leaving only the paper
    paper = (small.filter(ImageFilter.MaxFilter(15))
                  .filter(ImageFilter.GaussianBlur(12))
                  .resize((w, h), Image.BILINEAR))
    return ImageMath.lambda_eval(
        lambda a: a["convert"](a["g"] * 255 / (a["p"] + 1), "L"), g=gray, p=paper)


def erase_rules(gray, length):
    """White out table borders and ruled lines, meaning dark runs that are both long
    (1/length of the picture's long side) and thin. The strokes of large letters can
    qualify too, so ocr_lines only keeps this version of the page if it reads better."""
    def erase_across(image):
        w, h = image.size
        long_run = re.compile(rb"\x00{%d,}" % (max(w, h) // length))
        max_thick = max(w, h) // 250
        ink = image.point(lambda v: 0 if v < 150 else 255).tobytes()
        out = image.copy()
        draw = ImageDraw.Draw(out)
        for y in range(h):
            for m in long_run.finditer(ink, y * w, (y + 1) * w):
                mid = (m.start() + m.end()) // 2
                up = down = 0
                while up <= max_thick and mid - (up + 1) * w >= 0 and ink[mid - (up + 1) * w] == 0:
                    up += 1
                while down <= max_thick and mid + (down + 1) * w < len(ink) and ink[mid + (down + 1) * w] == 0:
                    down += 1
                if up + down < max_thick:
                    draw.rectangle((m.start() - y * w, y - 2, m.end() - y * w, y + 2), fill=255)
        return out

    flip = Image.Transpose.TRANSPOSE
    return erase_across(erase_across(gray).transpose(flip)).transpose(flip)


def read_words(gray, psm, scale):
    """Run Tesseract and return its confident words, grouped by line."""
    d = pytesseract.image_to_data(gray, lang=OCR_LANG, config=f"--psm {psm}",
                                  output_type=pytesseract.Output.DICT)
    lines = {}
    for i, word in enumerate(d["text"]):
        word = word.strip()
        try:
            conf = float(d["conf"][i])
        except ValueError:
            conf = -1
        if not word.strip(STRAY_MARKS) or conf < MIN_CONFIDENCE:
            continue
        key = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        lines.setdefault(key, []).append({
            "text": word, "conf": conf,
            "l": d["left"][i] / scale, "t": d["top"][i] / scale,
            "r": (d["left"][i] + d["width"][i]) / scale,
            "b": (d["top"][i] + d["height"][i]) / scale,
        })
    return lines


def read_score(lines):
    """Characters read, weighted by Tesseract's confidence, so junk counts for little."""
    return sum(len(w["text"]) * w["conf"] for words in lines.values() for w in words)


def overlap(a, b):
    return a["l"] < b["r"] and b["l"] < a["r"] and a["t"] < b["b"] and b["t"] < a["b"]


def straighten(lines, center_x, center_y, slope=None):
    """Undo a tilted photo so its columns come out upright. The words along each
    line give that line's slope (unless the slope is already known); every word
    is then turned back by the typical slope (for angles this small, shifting it
    is as good as rotating it)."""
    if slope is None:
        slopes = []
        for words in lines:
            first, last = min(words, key=lambda w: w["l"]), max(words, key=lambda w: w["l"])
            span = last["l"] - first["l"]
            if span > 10 * (first["b"] - first["t"]):
                slopes.append(((last["t"] + last["b"]) - (first["t"] + first["b"])) / 2 / span)
        if not slopes:
            return
        slope = median(slopes)
    for words in lines:
        for w in words:
            dx = slope * ((w["t"] + w["b"]) / 2 - center_y)
            dy = -slope * ((w["l"] + w["r"]) / 2 - center_x)
            w["l"], w["r"] = max(0.0, w["l"] + dx), w["r"] + dx
            w["t"], w["b"] = w["t"] + dy, w["b"] + dy


_reader = None                  # the RapidOCR reader, made the first time it's needed


def rapid_reader():
    """RapidOCR through photo_text_reader, or None if it isn't installed (then
    Tesseract reads the photos, as before)."""
    global _reader
    if _reader is None:
        try:
            import photo_text_reader
            if importlib.util.find_spec("rapidocr") is None or importlib.util.find_spec("onnxruntime") is None:
                raise ImportError("rapidocr")
            _reader = photo_text_reader.Reader(use_tesseract=False)
        except Exception:
            _reader = False
    return _reader or None


def rapid_words(pieces):
    """The reader's lines as Tesseract-style words, placed along each line (which
    may be tilted, so straighten() can level the page as it does Tesseract's)."""
    lines = {}
    for number, piece in enumerate(pieces):
        text = " ".join(w for w in piece.text.split() if w.strip(STRAY_MARKS))
        if not text or len(piece.quad) != 4 or piece.confidence < MIN_CONFIDENCE / 100:
            continue
        (tl, tr, br, bl) = piece.quad
        start = ((tl[0] + bl[0]) / 2, (tl[1] + bl[1]) / 2)        # middle of the left end
        end = ((tr[0] + br[0]) / 2, (tr[1] + br[1]) / 2)
        length = max(1e-6, ((end[0] - start[0]) ** 2 + (end[1] - start[1]) ** 2) ** 0.5)
        height = piece.height / RAPID_LINE_TO_WORD
        pad = min(0.25 * length, (piece.height - height) / 2)    # the outline's margin at each end
        step = ((end[0] - start[0]) / length, (end[1] - start[1]) / length)
        start = (start[0] + step[0] * pad, start[1] + step[1] * pad)
        usable = length - 2 * pad

        def at(fraction):
            return start[0] + step[0] * usable * fraction, start[1] + step[1] * usable * fraction

        words, offset = [], 0
        for word in text.split(" "):
            (x0, y0), (x1, y1) = at(offset / len(text)), at((offset + len(word)) / len(text))
            middle = (y0 + y1) / 2
            # a handwritten t is often read as + ("Aun+"): inside a word it can only be a t
            word = re.sub(r"(?<=[A-Za-z])\+|\+(?=[a-z])", "t", word)
            words.append({"text": word, "conf": piece.confidence * 100, "l": x0, "r": x1,
                          "t": middle - height / 2, "b": middle + height / 2})
            offset += len(word) + 1
        lines[("rapid", number)] = words
    return lines


def page_turn(pieces):
    """How far most of the page's writing is turned (0, 90, 180 or 270): a page
    photographed or scanned upside down or sideways."""
    weight = {}
    for piece in pieces:
        weight[piece.turn] = weight.get(piece.turn, 0) + len(piece.text) * piece.confidence
    total = sum(weight.values())
    turn = max(weight, key=weight.get) if weight else 0
    return turn if total and weight[turn] >= 0.6 * total else 0


# turning the picture undoes a turn: text read by turning its strip clockwise (90)
# lies with its top to the left, so the page turns clockwise too
UNDO_TURN = {90: Image.Transpose.ROTATE_270, 180: Image.Transpose.ROTATE_180,
             270: Image.Transpose.ROTATE_90}


def read_page(img):
    """(picture, lines): the photo's text lines with pixel boxes. The new reader is
    used when it is installed, turning the picture first if the page is upside down
    or sideways; otherwise Tesseract."""
    reader = rapid_reader()
    if reader is None:
        return img, tesseract_lines(img)
    try:
        pieces = reader.read_best(img, longest=RAPID_LONGEST)
        turn = page_turn(pieces)
        if turn:
            img = img.transpose(UNDO_TURN[turn])
            pieces = reader.read_best(img, longest=RAPID_LONGEST)
    except Exception as error:
        traceback.print_exc()
        if not tesseract_ready():     # the Setup has no Tesseract to fall back on
            raise RuntimeError(f"the text reader couldn't read it: {error}") from error
        return img, tesseract_lines(img)
    lines = rapid_words(pieces)
    straighten(lines.values(), img.width / 2, img.height / 2, tilt(pieces))
    return img, finish_lines(lines)


def tilt(pieces):
    """The page's slope, from the reader's outlines of its longer patches of writing:
    even one table cell's outline shows the tilt. None if there are none."""
    slopes = []
    for piece in pieces:
        if piece.turn != 0 or len(piece.quad) != 4:
            continue
        (tl, tr, br, bl) = piece.quad
        run = ((tr[0] - tl[0]) + (br[0] - bl[0])) / 2
        if run > 2.5 * max(1.0, piece.height):
            slopes.append(((tr[1] - tl[1]) + (br[1] - bl[1])) / 2 / run)
    return median(slopes) if slopes else None


def ocr_lines(img):
    """OCR the image and return text lines with pixel bounding boxes."""
    return read_page(img)[1]


def finish_lines(lines):
    out = []
    for words in lines.values():
        words.sort(key=lambda w: w["l"])
        out.append({
            "words": words,
            "l": min(w["l"] for w in words), "t": min(w["t"] for w in words),
            "r": max(w["r"] for w in words), "b": max(w["b"] for w in words),
        })
    return out


def tesseract_lines(img):
    """Tesseract's reading of the image: text lines with pixel bounding boxes."""
    center_x, center_y = img.width / 2, img.height / 2
    longest = max(img.size)
    if longest < 1600:                       # small pictures OCR better enlarged
        scale = 2.0
    else:                                    # huge ones only cost time and memory
        scale = min(1.0, MAX_OCR_PX / longest)
    if scale != 1.0:
        img = img.resize((round(img.width * scale), round(img.height * scale)), Image.LANCZOS)
    gray = ImageOps.autocontrast(even_out_lighting(img.convert("L")))

    # table borders can leave Tesseract reading next to nothing, so try the page
    # without them and keep that only if it reads clearly better
    lines = read_words(gray, 6, scale)
    page = gray
    to_beat = 1.1 * read_score(lines)
    for length in RULE_LENGTHS:
        erased = erase_rules(gray, length)
        if erased.tobytes() != gray.tobytes():
            trial = read_words(erased, 6, scale)
            if read_score(trial) > to_beat:
                lines, page, to_beat = trial, erased, read_score(trial)

    # mode 6 reads tables without dropping digits, but can miss large headings on
    # a page of small text; mode 3 finds those, so add whatever only it saw
    seen = [w for words in lines.values() for w in words]
    for key, words in read_words(page, 3, scale).items():
        extra = [w for w in words if not any(overlap(w, s) for s in seen)]
        if extra:
            lines[("mode 3", *key)] = extra
    straighten(lines.values(), center_x, center_y)
    return finish_lines(lines)


def group_rows(lines):
    """Merge lines that sit side-by-side (e.g. two columns) into one row."""
    if not lines:
        return []
    med_h = median(l["b"] - l["t"] for l in lines)
    lines.sort(key=lambda l: (l["t"] + l["b"]) / 2)
    rows, cur = [], [lines[0]]
    for ln in lines[1:]:
        cy = (ln["t"] + ln["b"]) / 2
        row_cy = sum((x["t"] + x["b"]) / 2 for x in cur) / len(cur)
        # lines side by side can sit further out of level than lines above each other
        beside = all(ln["l"] >= x["r"] or ln["r"] <= x["l"] for x in cur)
        if abs(cy - row_cy) < (1.0 if beside else 0.5) * med_h:
            cur.append(ln)
        else:
            rows.append(cur)
            cur = [ln]
    rows.append(cur)

    result = []
    for r in rows:
        words = sorted((w for ln in r for w in ln["words"]), key=lambda w: w["l"])
        h = median(w["b"] - w["t"] for w in words)
        result.append({
            "columns": split_columns(words, h),
            "t": min(w["t"] for w in words),
            "h": h,
        })
    snap_columns(result)
    return result


def split_columns(words, h):
    """Join a row's words into stretches of text, breaking where a gap is too
    wide to be a space."""
    columns = []
    for w in words:
        if columns and w["l"] - columns[-1]["r"] <= COLUMN_GAP * h:
            columns[-1]["text"] += " " + w["text"]
            columns[-1]["r"] = w["r"]
        else:
            columns.append({"text": w["text"], "l": w["l"], "r": w["r"]})
    return columns


def snap_columns(rows):
    """OCR boxes jitter by a few pixels, so make starts that nearly agree agree."""
    tolerance = 0.3 * median(r["h"] for r in rows)
    group = []
    for col in sorted((c for r in rows for c in r["columns"]), key=lambda c: c["l"]) + [None]:
        if group and (col is None or col["l"] - group[0]["l"] > tolerance):
            start = median(c["l"] for c in group)
            for c in group:
                c["l"] = start
            group = []
        if col:
            group.append(col)


@lru_cache(maxsize=None)
def char_ems(ch):
    return METRICS.getlength(ch) / 1000


def text_width(text, size_pt):
    """Width in points from the font's own metrics, with slack for Word's rounding."""
    return sum(char_ems(ch) for ch in text) * size_pt * 1.02


def fits(row, size_pt, pt_per_px, width_pt):
    """True if each column ends before the next starts, and the last before the margin."""
    cols = row["columns"]
    ends = [c["l"] * pt_per_px + text_width(c["text"], size_pt) for c in cols]
    space = text_width(" ", size_pt)
    return (all(end + space <= c["l"] * pt_per_px for end, c in zip(ends, cols[1:]))
            and ends[-1] <= width_pt)


def half_points(size_pt):
    """Clamp to the allowed range and round down; Word only has half points."""
    return floor(max(MIN_PT, min(MAX_PT, size_pt)) * 2) / 2


def largest_fit(row, size_pt, pt_per_px, width_pt):
    """Biggest size, up to size_pt, that keeps the row on one line."""
    size = half_points(size_pt)
    while size > MIN_PT and not fits(row, size, pt_per_px, width_pt):
        size -= 0.5
    return size


def choose_sizes(rows, pt_per_px, width_pt):
    """Pick each row's font size. A word's box changes height with its letters
    ("none" vs "Typing"), so rows of similar height share one body size and only
    real headings and small print get their own."""
    def wanted(h):
        return h * pt_per_px * FONT_PER_WORD_H

    body_h = median_low(r["h"] for r in rows)
    is_body = [1 / SAME_SIZE_RATIO <= r["h"] / body_h <= SAME_SIZE_RATIO for r in rows]

    # the longest rows set the size, but a few odd ones (often garbled OCR) shouldn't
    # shrink the whole page; those get a smaller size of their own below
    body_fits = [largest_fit(r, wanted(body_h), pt_per_px, width_pt)
                 for r, body in zip(rows, is_body) if body]
    body_pt = sorted(body_fits)[len(body_fits) // 10]

    # tightly spaced lines need a smaller size or the page would overflow; limit
    # the cut, since rows split by a tilted photo can look closer than they are
    if len(rows) > 1:
        pitch = median((b["t"] - a["t"]) * pt_per_px for a, b in zip(rows, rows[1:]))
        body_pt = max(body_pt * 0.7, min(body_pt, pitch))

    return [largest_fit(r, body_pt if body else wanted(r["h"]), pt_per_px, width_pt)
            for r, body in zip(rows, is_body)]


def set_font(run, size_pt):
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT_NAME)


def new_document():
    """A document with its default empty paragraph removed, ready for photo pages."""
    doc = Document()
    body = doc.element.body
    for p in list(body.iterchildren(qn("w:p"))):
        body.remove(p)
    return doc


def add_photo_page(doc, img_size, rows, label=None):
    """Lay one photo's text out on a page of its own. Each page is its own
    section, so a wide photo gets a landscape page even among portrait ones.
    A label goes in the page's header, out of the way of the text."""
    img_w, img_h = img_size
    if doc.paragraphs:
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        # the break is an empty paragraph ending the last page; on a full page it
        # could spill over and leave a blank page, so make it 1 point high
        end = doc.paragraphs[-1].paragraph_format
        end.space_before = end.space_after = Pt(0)
        end.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        end.line_spacing = Pt(1)
    else:
        sec = doc.sections[0]
    landscape = img_w > img_h
    pw, ph = (PAGE_H_IN, PAGE_W_IN) if landscape else (PAGE_W_IN, PAGE_H_IN)
    sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Inches(pw), Inches(ph)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(MARGIN_IN))

    if label:
        sec.header.is_linked_to_previous = False
        sec.header_distance = Inches(0.2)     # above the text, which starts at the margin
        run = sec.header.paragraphs[0].add_run(label)
        set_font(run, 8)
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    area_w, area_h = pw - 2 * MARGIN_IN, ph - 2 * MARGIN_IN
    in_per_px = min(area_w / img_w, area_h / img_h)    # keep photo proportions
    pt_per_px = in_per_px * 72

    if not rows:
        p = doc.add_paragraph()
        set_font(p.add_run("(no text found)"), 10)
        return

    width_pt = area_w * 72 - 1      # a point of slack so Word's rounding can't wrap a row
    sizes = choose_sizes(rows, pt_per_px, width_pt)

    prev_bottom_pt = 0.0
    for i, (row, font_pt) in enumerate(zip(rows, sizes)):
        starts = [c["l"] * pt_per_px for c in row["columns"]]
        texts = [c["text"] for c in row["columns"]]

        # too long even at the smallest size: run it together and slide it left
        if not fits(row, font_pt, pt_per_px, width_pt):
            texts = [" ".join(texts)]
            starts = [max(0.0, min(starts[0], width_pt - text_width(texts[0], font_pt)))]

        # tighten the line if the next row follows closely, to hold the row's place
        top_pt = row["t"] * pt_per_px
        next_top_pt = rows[i + 1]["t"] * pt_per_px if i + 1 < len(rows) else inf
        line_pt = max(font_pt, min(font_pt * LINE_HEIGHT, next_top_pt - top_pt))
        space_before = max(0.0, top_pt - prev_bottom_pt)

        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Pt(starts[0])
        for x in starts[1:]:
            pf.tab_stops.add_tab_stop(Pt(x))
        pf.space_before = Pt(min(space_before, 1584))   # Word's max
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
        pf.widow_control = False
        set_font(p.add_run("\t".join(texts)), font_pt)

        prev_bottom_pt = max(prev_bottom_pt, top_pt) + line_pt


def build_docx(img_size, rows, out_path):
    doc = new_document()
    add_photo_page(doc, img_size, rows)
    doc.save(out_path)


def output_name(folder, fname, taken):
    """IMG_1.jpg -> IMG_1.docx; if IMG_1.png also exists -> IMG_1_png.docx."""
    stem, ext = os.path.splitext(fname)
    out = os.path.join(folder, stem + ".docx")
    if out.lower() in taken:
        out = os.path.join(folder, f"{stem}_{ext.lstrip('.').lower()}.docx")
    taken.add(out.lower())
    return out


def combined_name(folder):
    """Folder\\Folder - all photos.docx, numbered rather than replacing an earlier one."""
    base = os.path.basename(os.path.normpath(folder)).rstrip(":") or "Photos"
    out = os.path.join(folder, f"{base} - all photos.docx")
    n = 2
    while os.path.exists(out):
        out = os.path.join(folder, f"{base} - all photos ({n}).docx")
        n += 1
    return out


# ---------------------------------------------------------------- main
def open_updates_page(win):
    """Opens UPDATES_PAGE in the web browser, or says where to go when that can't be done.
    The window stops staying on top of others, so the browser can come in front of it."""
    win.attributes("-topmost", False)
    try:
        opened = webbrowser.open(UPDATES_PAGE)
    except Exception:
        opened = False
    if not opened:
        messagebox.showinfo("Photos to Word", "The web browser could not be opened. To check for "
                            f"a newer version, go to:\n\n{UPDATES_PAGE}", parent=win)


def ask_mode():
    """Ask with two buttons how the text should be saved. Returns "each",
    "one", or None if the window is closed or canceled."""
    choice = []
    win = tk.Toplevel()
    win.title("Photos to Word")
    win.resizable(False, False)
    win.attributes("-topmost", True)

    def pick(mode):
        choice.append(mode)
        win.destroy()

    frame = ttk.Frame(win, padding=18)
    frame.pack()
    bar = ttk.Frame(frame)                               # "Check for updates" in the top right corner
    bar.pack(fill="x", pady=(0, 10))
    ttk.Button(bar, text="Check for updates", command=lambda: open_updates_page(win)).pack(side="right")
    ttk.Label(frame, text="How should the text from the photos be saved?").pack(anchor="w", pady=(0, 12))
    first = ttk.Button(frame, text="One Word file for each photo", command=lambda: pick("each"))
    first.pack(fill="x", ipady=6, pady=3)
    second = ttk.Button(frame, text="All photos in one Word file", command=lambda: pick("one"))
    second.pack(fill="x", ipady=6, pady=3)
    cancel = ttk.Button(frame, text="Cancel", command=win.destroy)
    cancel.pack(pady=(12, 0))

    for button in (first, second, cancel):              # Enter presses the highlighted button
        button.bind("<Return>", lambda e: e.widget.invoke())
    win.bind("<Escape>", lambda e: win.destroy())
    win.update_idletasks()
    win.geometry(f"+{(win.winfo_screenwidth() - win.winfo_width()) // 2}"
                 f"+{(win.winfo_screenheight() - win.winfo_height()) // 3}")
    first.focus_set()
    win.focus_force()
    win.wait_window()
    return choice[0] if choice else None


class Progress:
    """A small window while the photos are read: which one, how far along, and Cancel.
    The reading runs in the background so the window keeps answering."""

    def __init__(self, total):
        self.total = max(1, total)
        self.canceled = threading.Event()
        self.messages = queue.Queue()
        win = self.win = tk.Toplevel()
        win.title("Photos to Word")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        if not PROGRESS_SEEN:
            win.attributes("-alpha", 0.0)
        frame = ttk.Frame(win, padding=18)
        frame.pack()
        self.label = ttk.Label(frame, text="Getting ready...", width=58, font=("Segoe UI", 11))
        self.label.pack(anchor="w")
        self.bar = ttk.Progressbar(frame, maximum=self.total, length=440)
        self.bar.pack(fill="x", pady=(10, 12))
        self.cancel_button = ttk.Button(frame, text="Cancel", command=self.cancel)
        self.cancel_button.pack()
        win.protocol("WM_DELETE_WINDOW", self.cancel)
        win.bind("<Escape>", lambda e: self.cancel())
        win.update_idletasks()
        win.geometry(f"+{(win.winfo_screenwidth() - win.winfo_width()) // 2}"
                     f"+{(win.winfo_screenheight() - win.winfo_height()) // 3}")

    def cancel(self):
        """Stop after the photo being read now (a photo can't be stopped half way)."""
        if not self.canceled.is_set():
            self.canceled.set()
            self.label.configure(text="Stopping after this photo...")
            self.cancel_button.state(["disabled"])

    def step(self, number, total, name):
        """Called from the background: photo `number` of `total` is starting."""
        self.messages.put((number, total, name))

    def run(self, work):
        """Run work(self) in the background, showing its progress, until it ends."""
        outcome = {}

        def go():
            try:
                work(self)
            except BaseException as error:        # shown once the window closes
                outcome["error"] = error
            finally:
                self.messages.put(None)

        def poll():
            try:
                while True:
                    message = self.messages.get_nowait()
                    if message is None:
                        self.win.destroy()
                        return
                    number, total, name = message
                    if not self.canceled.is_set():
                        self.label.configure(text=f"Reading photo {number} of {total}: {name}")
                    self.bar.configure(value=number - 1)
            except queue.Empty:
                pass
            self.win.after(100, poll)

        threading.Thread(target=go, daemon=True).start()
        self.win.after(50, poll)
        self.win.wait_window()
        if "error" in outcome:
            raise outcome["error"]


def list_photos(folder):
    """(pictures in natural name order, iPhone HEIC photos that can't be opened,
    how many other files were skipped). Word files, Windows' hidden files and a
    Mac's "._" notes aren't counted at all."""
    files = sorted((f for f in os.listdir(folder)
                    if os.path.isfile(os.path.join(folder, f))
                    and not f.lower().endswith((".docx", ".doc", ".txt", ".pdf"))
                    and not f.startswith(("._", "~$")) and f.lower() not in NOT_FILES),
                   key=natural_key)
    pictures = [f for f in files if is_picture(os.path.join(folder, f))]
    heic = [f for f in files if f not in pictures and f.lower().endswith(HEIC)]
    return pictures, heic, len(files) - len(pictures) - len(heic)


HEIC_WORDS = ("iPhone HEIC photo(s), which this program can't open. Save them as JPG first, or "
              "set the iPhone to take JPG photos (Settings, Camera, Formats, Most Compatible).")


def main():
    dialog_root()
    # Tesseract is only needed when the reader (RapidOCR) isn't installed
    if rapid_reader() is None and not tesseract_ready():
        try:
            find_tesseract()
        except RuntimeError as e:
            messagebox.showerror("Photos to Word", str(e))
            return

    mode = ask_mode()
    if not mode:
        return
    while True:          # a folder with no photos: say so and ask again
        folder = filedialog.askdirectory(title="Pick the folder of photographs")
        if not folder:
            return
        folder = os.path.normpath(folder)      # C:\Users\..., the way Windows shows it
        pictures, heic, skipped = list_photos(folder)
        if pictures:
            break
        words = f"There are no photos in\n{folder}"
        if heic:
            words += f"\n\nIt has {len(heic)} {HEIC_WORDS}"
        messagebox.showinfo("Photos to Word", words + "\n\nClick OK, then pick the folder that has the "
                            "photos in it (or Cancel to stop).")

    taken = set()
    jobs = [(f, output_name(folder, f, taken) if mode == "each" else None) for f in pictures]
    kept = 0
    existing = sum(1 for _, out in jobs if out and os.path.exists(out))
    if existing:
        answer = messagebox.askyesnocancel(
            "Photos to Word",
            f"{existing} of {len(jobs)} photo(s) already have a Word file in this "
            "folder. Anything you changed in those files would be lost if they "
            "are replaced.\n\n"
            "Yes - replace them\n"
            "No - keep them and only do the new photos\n"
            "Cancel - stop",
            default=messagebox.NO)
        if answer is None:
            return
        if not answer:
            jobs = [j for j in jobs if not os.path.exists(j[1])]
            kept = existing

    combined = new_document() if mode == "one" else None
    done, failed, notes = [], [], []

    def read_all(progress):
        for n, (fname, out) in enumerate(jobs, 1):
            if progress.canceled.is_set():
                return
            progress.step(n, len(jobs), fname)
            print(f"[{n}/{len(jobs)}] {fname} ...", flush=True)
            try:
                img, pages = open_image(os.path.join(folder, fname))
                img, lines = read_page(img)        # turned upright if it was sideways
                rows = group_rows(lines)
                if combined is not None:
                    add_photo_page(combined, img.size, rows, label=fname)
                else:
                    build_docx(img.size, rows, out)
                done.append(fname)
                if pages > 1:
                    notes.append(f"{fname} has {pages} pages; only page 1 was read.")
            except PermissionError:
                failed.append(f"{fname} (access denied - if the .docx is open in Word, "
                              "close it and rerun)")
            except Exception as e:
                traceback.print_exc()
                failed.append(f"{fname} ({unreadable(os.path.join(folder, fname), e)})")

    progress = Progress(len(jobs))
    progress.run(read_all)
    stopped = progress.canceled.is_set() and len(done) + len(failed) < len(jobs)

    if combined is None:
        msg = f"Created {len(done)} Word file(s) in:\n{folder}"
    elif not done:
        msg = "No Word file was made, because no photo could be read."
    else:
        out = combined_name(folder)
        print("Saving", out, flush=True)
        try:
            combined.save(out)
            msg = f"Put {len(done)} photo(s) into one Word file, a page each:\n{out}"
        except Exception as e:
            traceback.print_exc()
            msg = f"The Word file could not be saved:\n{out}\n\n{e}"
    if stopped:
        msg = (f"Stopped after {len(done) + len(failed)} of {len(jobs)} photo(s), as asked. "
               "What was read is saved.\n\n" + msg)
    if kept:
        msg += f"\n\nKept {kept} existing Word file(s)."
    if failed:
        msg += "\n\nProblems:\n" + "\n".join(failed[:15])
        if len(failed) > 15:
            msg += f"\n...and {len(failed) - 15} more"
    if notes:
        msg += "\n\n" + "\n".join(notes[:15])
    if heic:
        msg += f"\n\nSkipped {len(heic)} {HEIC_WORDS}"
    if skipped:
        msg += f"\n\nSkipped {skipped} non-picture file(s)."
    messagebox.showinfo("Photos to Word", msg)


def self_test(folder):
    """For the Setup's test on a clean PC: makes a photo of a page with printed and
    handwritten lines, reads it with only the parts packed into the program, saves
    the Word file, and writes what happened to self-test.txt in the folder."""
    import time
    lines = []
    try:
        os.makedirs(folder, exist_ok=True)
        page = Image.new("RGB", (1700, 1300), (250, 249, 244))
        draw = ImageDraw.Draw(page)
        fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        printed = ImageFont.truetype(os.path.join(fonts, "arial.ttf"), 60)
        hand_file = os.path.join(fonts, "segoepr.ttf")
        hand = ImageFont.truetype(hand_file if os.path.exists(hand_file) else os.path.join(fonts, "arial.ttf"), 70)
        draw.text((150, 200), "The quick brown fox jumps", font=printed, fill=(20, 20, 20))
        draw.text((150, 500), "Call Margaret on Tuesday", font=hand, fill=(25, 40, 120))
        path = os.path.join(folder, "self-test page.jpg")
        page.rotate(1.5, fillcolor=(90, 70, 55), expand=True).save(path, quality=88)
        started = time.time()
        lines.append(f"New reader available: {rapid_reader() is not None}")
        img, _ = open_image(path)
        img, found = read_page(img)
        rows = group_rows(found)
        text = " ".join(c["text"] for r in rows for c in r["columns"])
        lines.append(f"Read in {time.time() - started:.1f} s: {text}")
        lines.append(f"Printed line read: {'quick brown fox' in text}")
        lines.append(f"Handwritten line read: {'Margaret' in text and 'Tuesday' in text}")
        out = os.path.join(folder, "self-test page.docx")
        build_docx(img.size, rows, out)
        lines.append(f"Word file made: {os.path.getsize(out) > 0}")
    except Exception:
        lines.append("Self-test crashed:\n" + traceback.format_exc())
    with open(os.path.join(folder, "self-test.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        self_test(sys.argv[2])
    else:
        main()
