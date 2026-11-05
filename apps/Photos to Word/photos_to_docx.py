"""
photos_to_docx.py

Pick a folder in a Windows Explorer dialog. Every picture in it (any
extension Pillow can open: jpg, jpeg, png, tif, bmp, gif, webp, heic*...)
is read in natural filename order, its text is recognized (OCR), and the
text is written to a like-named .docx in the same folder:

    IMG_0001.jpg  ->  IMG_0001.docx

Text is placed in a similar position to where it appears in the photo:
left indent, vertical spacing, font size and gaps between columns are all
scaled from the picture onto a Letter page (landscape if the photo is wide).
Columns are held in place by tab stops, so they line up in any font and stay
lined up when the text is corrected. Shadows, table borders and a slightly
tilted photo are dealt with along the way.

If a photo already has a Word file you are asked before it is replaced, so
corrections typed into earlier files aren't lost by accident.

Requirements (one-time):
    1. Tesseract OCR for Windows:
       https://github.com/UB-Mannheim/tesseract/wiki  (default install path is fine)
    2. pillow (10.3 or newer), pytesseract and python-docx, plus pillow-heif for
       iPhone .heic photos. If the Python running this lacks them, it offers to
       install them.
"""

import importlib.util
import os
import re
import shutil
import site
import subprocess
import sys
import traceback
from functools import lru_cache
from math import floor, inf
from statistics import median, median_low

import tkinter as tk
from tkinter import filedialog, messagebox

REQUIRED = {"PIL": "pillow", "pytesseract": "pytesseract", "docx": "python-docx"}


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


def install_missing():
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add whatever this one lacks."""
    missing = [pip for module, pip in REQUIRED.items()
               if importlib.util.find_spec(module) is None]
    if not missing:
        return
    root = dialog_root()
    if not messagebox.askyesno(
            "Photos to Word",
            f"This copy of Python ({sys.version.split()[0]}) still needs: "
            f"{', '.join(missing)}.\n\n"
            "Download and install now? This only happens once and takes about a minute."):
        sys.exit()
    root.destroy()

    print("Installing", ", ".join(missing), "...", flush=True)
    pip = [sys.executable, "-m", "pip", "install"]
    result = subprocess.run(pip + missing, stderr=subprocess.PIPE, text=True)
    if result.returncode:
        raise RuntimeError("Could not install " + ", ".join(missing) + ":\n" + result.stderr[-800:])
    subprocess.run(pip + ["pillow-heif"])    # optional, so a failure here doesn't matter
    site.addsitedir(site.getusersitepackages())   # a first install may have just created it
    importlib.invalidate_caches()


sys.excepthook = show_crash
install_missing()

from PIL import (Image, ImageDraw, ImageFilter, ImageFont, ImageMath, ImageOps,
                 UnidentifiedImageError)
import pytesseract
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

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

Image.MAX_IMAGE_PIXELS = None        # or Pillow refuses photos over about 179 MP
METRICS = ImageFont.truetype(FONT_FILE, 1000)


# ---------------------------------------------------------------- helpers
def find_tesseract():
    if shutil.which("tesseract"):
        return
    for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
              os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")):
        if os.path.isfile(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return
    raise RuntimeError(
        "Tesseract OCR was not found.\n\nInstall it from\n"
        "https://github.com/UB-Mannheim/tesseract/wiki\nthen run this again.")


def natural_key(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def is_picture(path):
    """False only for files Pillow doesn't recognize. A damaged picture still
    counts, so that it gets reported rather than quietly skipped."""
    try:
        with Image.open(path):
            return True
    except UnidentifiedImageError:
        return False
    except Exception:
        return True


def open_image(path):
    """Return an upright RGB image and the number of pages in the file."""
    with Image.open(path) as img:
        pages = img.n_frames if img.format == "TIFF" else 1
        img = ImageOps.exif_transpose(img)   # honor phone rotation
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


def straighten(lines, center_x, center_y):
    """Undo a tilted photo so its columns come out upright. The words along each
    line give that line's slope; every word is then turned back by the typical
    slope (for angles this small, shifting it is as good as rotating it)."""
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


def ocr_lines(img):
    """OCR the image and return text lines with pixel bounding boxes."""
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

    out = []
    for words in lines.values():
        words.sort(key=lambda w: w["l"])
        out.append({
            "words": words,
            "l": min(w["l"] for w in words), "t": min(w["t"] for w in words),
            "r": max(w["r"] for w in words), "b": max(w["b"] for w in words),
        })
    return out


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


def build_docx(img_size, rows, out_path):
    img_w, img_h = img_size
    doc = Document()
    sec = doc.sections[0]
    landscape = img_w > img_h
    pw, ph = (PAGE_H_IN, PAGE_W_IN) if landscape else (PAGE_W_IN, PAGE_H_IN)
    sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Inches(pw), Inches(ph)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(MARGIN_IN))

    area_w, area_h = pw - 2 * MARGIN_IN, ph - 2 * MARGIN_IN
    in_per_px = min(area_w / img_w, area_h / img_h)    # keep photo proportions
    pt_per_px = in_per_px * 72

    # remove the default empty paragraph
    body = doc.element.body
    for p in list(body.iterchildren(qn("w:p"))):
        body.remove(p)

    if not rows:
        p = doc.add_paragraph()
        set_font(p.add_run("(no text found)"), 10)
        doc.save(out_path)
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

    doc.save(out_path)


def output_name(folder, fname, taken):
    """IMG_1.jpg -> IMG_1.docx; if IMG_1.png also exists -> IMG_1_png.docx."""
    stem, ext = os.path.splitext(fname)
    out = os.path.join(folder, stem + ".docx")
    if out.lower() in taken:
        out = os.path.join(folder, f"{stem}_{ext.lstrip('.').lower()}.docx")
    taken.add(out.lower())
    return out


# ---------------------------------------------------------------- main
def main():
    dialog_root()
    try:
        find_tesseract()
    except RuntimeError as e:
        messagebox.showerror("Photos to Word", str(e))
        return

    folder = filedialog.askdirectory(title="Pick the folder of photographs")
    if not folder:
        return

    files = sorted((f for f in os.listdir(folder)
                    if os.path.isfile(os.path.join(folder, f))
                    and not f.lower().endswith((".docx", ".doc", ".txt", ".pdf"))),
                   key=natural_key)

    pictures = [f for f in files if is_picture(os.path.join(folder, f))]
    skipped = len(files) - len(pictures)

    taken = set()
    jobs = [(f, output_name(folder, f, taken)) for f in pictures]
    kept = 0
    existing = sum(os.path.exists(out) for _, out in jobs)
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

    done, failed, notes = [], [], []
    for n, (fname, out) in enumerate(jobs, 1):
        print(f"[{n}/{len(jobs)}] {fname} ...", flush=True)
        try:
            img, pages = open_image(os.path.join(folder, fname))
            rows = group_rows(ocr_lines(img))
            build_docx(img.size, rows, out)
            done.append(os.path.basename(out))
            if pages > 1:
                notes.append(f"{fname} has {pages} pages; only page 1 was read.")
        except PermissionError:
            failed.append(f"{fname} (access denied - if the .docx is open in Word, "
                          "close it and rerun)")
        except Exception as e:
            traceback.print_exc()
            failed.append(f"{fname} ({e})")

    msg = f"Created {len(done)} Word file(s) in:\n{folder}"
    if kept:
        msg += f"\n\nKept {kept} existing Word file(s)."
    if failed:
        msg += "\n\nProblems:\n" + "\n".join(failed[:15])
        if len(failed) > 15:
            msg += f"\n...and {len(failed) - 15} more"
    if notes:
        msg += "\n\n" + "\n".join(notes[:15])
    if skipped:
        msg += f"\n\nSkipped {skipped} non-picture file(s)."
    messagebox.showinfo("Photos to Word", msg)


if __name__ == "__main__":
    main()
