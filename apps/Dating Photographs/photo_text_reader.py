"""Reads the words in a photo, handwriting included, and where each one is.

Shared by Dating Photographs and Photos to Word. It uses RapidOCR, a free
reader that runs on this computer (nothing is uploaded); RapidOCR finds each
patch of writing and reads it. Its own "is this upside down?" guess was wrong
on test photos, so each patch is read both ways round instead and the better
reading kept. Tesseract, if it is on the PC, adds a pass for printed words.

    reader = Reader()
    for pieces in reader.passes(picture):   # each pass tries harder than the last
        ...                                   # stop as soon as the words are enough

Requirements: rapidocr (with onnxruntime), numpy, opencv-python-headless and
Pillow. pytesseract and Tesseract OCR are optional. eight_bit() and upright()
need only Pillow, so Photos to Word can use them without the reader.
"""

from __future__ import annotations

import logging
import os
import shutil
import sys
from dataclasses import dataclass

try:
    import numpy as np          # only the reader needs it
except ImportError:
    np = None
from PIL import Image, ImageFilter, ImageOps

LONGEST = 2400          # pictures are shrunk to this before reading; writing stays readable
MIN_CONFIDENCE = 0.5    # readings below this are mostly noise
UPSIDE_DOWN_MARGIN = 0.15   # how much better an upside-down reading must be to be believed
STACKED = 0.75          # trust in two lines read as one ("Easter Sunday" / "1966")


@dataclass
class Piece:
    """Some words read from a picture."""
    text: str
    confidence: float           # 0 to 1
    box: tuple                  # (left, top, right, bottom) in the picture's own pixels
    turn: int = 0               # how the words lie: 0 level, 180 upside down, 90 / 270 sideways
    reader: str = "rapid"
    quad: tuple = ()            # its four corners in the picture (top-left, top-right, bottom-right,
                                # bottom-left as the reader found them), for a tilted line
    height: float = 0.0         # the line's own height across it, in the picture's pixels


# ------------------------------------------------------------------ Tesseract

def bundled_tesseract():
    """The copy of Tesseract packed inside a built program, or None."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        exe = os.path.join(base, "tesseract", "tesseract.exe")
        if os.path.isfile(exe):
            return exe
    return None


def find_tesseract():
    """Point pytesseract at Tesseract; returns False if it isn't on this PC."""
    import pytesseract
    bundled = bundled_tesseract()
    if bundled:
        pytesseract.pytesseract.tesseract_cmd = bundled
        os.environ["TESSDATA_PREFIX"] = os.path.join(os.path.dirname(bundled), "tessdata")
        return True
    if shutil.which("tesseract"):
        return True
    for path in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                 r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                 os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")):
        if os.path.isfile(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    return False


def tesseract_pieces(gray, scale=1.0, psm=11, min_conf=40):
    """Tesseract's words, joined into its own lines. psm 11 finds scattered words."""
    import pytesseract
    data = pytesseract.image_to_data(gray, config=f"--psm {psm}", output_type=pytesseract.Output.DICT)
    lines = {}
    for i, word in enumerate(data["text"]):
        word = (word or "").strip()
        try:
            conf = float(data["conf"][i])
        except ValueError:
            conf = -1
        if not word or conf < min_conf:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        left, top = data["left"][i] / scale, data["top"][i] / scale
        lines.setdefault(key, []).append((word, conf / 100, (
            left, top, left + data["width"][i] / scale, top + data["height"][i] / scale)))
    out = []
    for words in lines.values():
        words.sort(key=lambda w: w[2][0])
        out.append(Piece(" ".join(w[0] for w in words), sum(w[1] for w in words) / len(words),
                         union(w[2] for w in words), 0, "tesseract"))
    return out


# ------------------------------------------------------------------ helpers

def union(boxes):
    boxes = list(boxes)
    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes))


DEEP_MODES = ("I;16", "I;16L", "I;16B", "I;16N", "I", "F")


def eight_bit(picture):
    """A 16-bit or 32-bit picture (some scanners save gray scans as 16-bit TIFFs)
    as an ordinary 8-bit gray one. Converted straight to RGB, Pillow clips every
    value over 255, so the whole scan would turn white. Others come back as they are."""
    if picture.mode not in DEEP_MODES:
        return picture
    deep = picture.convert("F")      # first: a Mac's I;16B can be converted, but not measured
    low, high = deep.getextrema()
    if picture.mode == "F" and high <= 1.0:
        top = 1.0                     # 0 to 1
    elif high <= 255:
        top = 255.0                   # 8-bit values kept in a bigger box
    elif high <= 65535:
        top = 65535.0                 # 16-bit
    else:
        top = float(high)
    return deep.point(lambda v: v * (255.0 / top)).convert("L")


def upright(picture):
    """An RGB copy turned the way the camera or scanner says is up."""
    picture = eight_bit(ImageOps.exif_transpose(picture))
    if picture.mode in ("RGBA", "LA", "PA") or "transparency" in picture.info:
        rgba = picture.convert("RGBA")
        picture = Image.alpha_composite(Image.new("RGBA", rgba.size, "white"), rgba)
    return picture.convert("RGB")


def shrunk(picture, longest=LONGEST):
    """(picture no longer than `longest`, how much smaller it is)."""
    scale = min(1.0, longest / max(picture.size))
    if scale < 1.0:
        picture = picture.resize((max(1, round(picture.width * scale)),
                                  max(1, round(picture.height * scale))), Image.LANCZOS)
    return picture, scale


def crop_quad(image, quad):
    """Straighten the patch inside a four-cornered box into a level strip."""
    import cv2
    quad = np.asarray(quad, dtype=np.float32)
    width = int(max(np.linalg.norm(quad[0] - quad[1]), np.linalg.norm(quad[2] - quad[3])))
    height = int(max(np.linalg.norm(quad[0] - quad[3]), np.linalg.norm(quad[1] - quad[2])))
    if width < 2 or height < 2:
        return None
    target = np.float32([[0, 0], [width, 0], [width, height], [0, height]])
    matrix = cv2.getPerspectiveTransform(quad, target)
    return cv2.warpPerspective(image, matrix, (width, height), borderMode=cv2.BORDER_REPLICATE,
                               flags=cv2.INTER_CUBIC)


def reading_frame(box, turn):
    """The box in coordinates where its words run left to right, so lines can be found
    the same way whichever way the words lie."""
    left, top, right, bottom = box
    return {0: (left, top, right, bottom), 180: (-right, -bottom, -left, -top),
            90: (-bottom, left, -top, right), 270: (top, -right, bottom, -left)}[turn]


def group_lines(pieces):
    """Every line of writing: [(turn, frame, piece, how many pieces it was made of)],
    top to bottom for each way the words lie. Pieces side by side on one line are
    joined into one piece, in reading order."""
    out = []
    groups = {}
    for piece in pieces:
        groups.setdefault(piece.turn, []).append(piece)
    for turn, group in groups.items():
        framed = sorted(((reading_frame(p.box, turn), p) for p in group), key=lambda fp: fp[0][1])
        lines = []
        for frame, piece in framed:
            middle = (frame[1] + frame[3]) / 2
            for line in lines:
                top, bottom = line["top"], line["bottom"]
                if top <= middle <= bottom or abs(middle - (top + bottom) / 2) < 0.5 * (bottom - top):
                    line["pieces"].append((frame, piece))
                    break
            else:
                lines.append({"top": frame[1], "bottom": frame[3], "pieces": [(frame, piece)]})
        whole = []
        for line in lines:
            members = [p for _, p in sorted(line["pieces"], key=lambda fp: fp[0][0])]
            piece = members[0]
            if len(members) > 1:
                piece = Piece(" ".join(p.text for p in members), min(p.confidence for p in members),
                              union(p.box for p in members), turn, members[0].reader)
            whole.append((turn, union(frame for frame, _ in line["pieces"]), piece, len(members)))
        out += sorted(whole, key=lambda line: line[1][1])
    return out


def join_lines(pieces):
    """Pieces that sit side by side on one line, joined into one piece: a date is
    often read as separate bits, "Dec." "25," "1990". Also each line with the one
    written just below it ("Easter Sunday" over "1966"), trusted less, since the
    two lines may have nothing to do with each other."""
    lines = group_lines(pieces)
    out = [piece for _, _, piece, count in lines if count > 1]
    by_turn = {}
    for turn, frame, piece, _ in lines:
        by_turn.setdefault(turn, []).append((frame, piece))
    for turn, whole in by_turn.items():
        for (upper, a), (lower, b) in zip(whole, whole[1:]):
            height = upper[3] - upper[1]
            below = -0.2 * height <= lower[1] - upper[3] <= 1.5 * height
            if below and lower[0] < upper[2] and upper[0] < lower[2]:
                out.append(Piece(a.text + " " + b.text, min(a.confidence, b.confidence) * STACKED,
                                 union((a.box, b.box)), turn, a.reader))
    return out


def camera_imprint(picture):
    """The orange or red glowing digits a film camera prints in a corner, as black
    on white, or None if the photo has none (or is mostly orange, like a sunset)."""
    small = np.asarray(picture.convert("RGB"), dtype=np.int16)
    r, g, b = small[..., 0], small[..., 1], small[..., 2]
    glow = (r > 170) & (r - b > 90) & (g > 50) & (g < r - 20)
    share = glow.mean()
    if not 0.00003 < share < 0.03:
        return None
    mask = Image.fromarray(np.where(glow, 0, 255).astype(np.uint8), "L")
    return mask.filter(ImageFilter.MinFilter(3)).convert("RGB")


# ------------------------------------------------------------------ RapidOCR

def quiet_rapidocr():
    for name in ("RapidOCR", "rapidocr"):
        logging.getLogger(name).setLevel(logging.ERROR)


class Reader:
    """Reads pictures. Makes the readers the first time they are needed (a second or two)."""

    def __init__(self, use_tesseract=True):
        self._finder = None          # finds patches of writing (PP-OCRv5)
        self._english = None         # reads English, print and handwriting (PP-OCRv5)
        self._second = None          # a second opinion that reads some hands better (PP-OCRv6)
        self.use_tesseract = use_tesseract and find_tesseract()

    def _engines(self):
        if self._finder is None:
            from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR
            quiet_rapidocr()
            v5 = {"Det.ocr_version": OCRVersion.PPOCRV5, "Det.model_type": ModelType.MOBILE,
                  "Rec.ocr_version": OCRVersion.PPOCRV5, "Rec.lang_type": LangRec.EN,
                  "Rec.model_type": ModelType.MOBILE, "Global.use_cls": False,
                  "Global.text_score": 0.0}
            self._finder = self._english = RapidOCR(params=v5)
            # only its reading part is used (patches are found once, by the first), so it
            # is set to the same finder to need no data of its own for finding
            self._second = RapidOCR(params={"Det.ocr_version": OCRVersion.PPOCRV5,
                                            "Det.model_type": ModelType.MOBILE,
                                            "Global.use_cls": False, "Global.text_score": 0.0})
            quiet_rapidocr()
        return self._finder, self._english, self._second

    def warm_up(self):
        """Load the readers now (and fetch their data, the first time, when not packed in)."""
        finder, english, second = self._engines()
        blank = np.full((64, 256, 3), 255, np.uint8)
        blank[20:44, 30:220] = 0
        finder(blank, use_det=True, use_cls=False, use_rec=False)
        self._recognize(english, [blank])
        self._recognize(second, [blank])

    @staticmethod
    def _recognize(engine, crops):
        """[(text, confidence)] for each crop."""
        if not crops:
            return []
        try:
            result = engine.recognize_txt(crops)
            return list(zip(result.txts, result.scores))
        except Exception:
            out = []
            for crop in crops:
                try:
                    result = engine(crop, use_det=False, use_cls=False, use_rec=True)
                    out.append((result.txts[0], result.scores[0]) if result.txts else ("", 0.0))
                except Exception:
                    out.append(("", 0.0))
            return out

    def find_patches(self, picture):
        """[(four-cornered box, level strip)] for each patch of writing."""
        finder, _, _ = self._engines()
        bgr = np.ascontiguousarray(np.asarray(picture)[:, :, ::-1])
        found = finder(bgr, use_det=True, use_cls=False, use_rec=False)
        boxes = getattr(found, "boxes", None)
        patches = []
        for quad in ([] if boxes is None else boxes):
            strip = crop_quad(bgr, quad)
            if strip is not None:
                patches.append((quad, strip))
        return patches

    def read_patches(self, patches, engine, scale, reader_name, by_patch=False):
        """Read each patch both ways round (four ways for upright strips) and keep the better.
        by_patch=True gives {patch number: Piece} instead of a list."""
        tries, owners = [], []
        for index, (quad, strip) in enumerate(patches):
            height, width = strip.shape[:2]
            if height > 1.5 * width:          # standing on end: the words run up or down
                turns = ((90, np.rot90(strip, -1)), (270, np.rot90(strip, 1)))
            else:
                turns = ((0, strip), (180, np.rot90(strip, 2)))
            for turn, image in turns:
                tries.append(np.ascontiguousarray(image))
                owners.append((index, turn))
        best = {}
        for (index, turn), (text, confidence) in zip(owners, self._recognize(engine, tries)):
            text = (text or "").strip()
            if not text:
                continue
            kept = best.get(index)
            # Most writing is the right way up, and a camera's digits read almost the same
            # upside down ('98 10 12 as 21 01 86'), so the level reading wins unless the
            # upside-down one is clearly better.
            margin = UPSIDE_DOWN_MARGIN if kept and kept[2] == 0 and turn == 180 else 0.0
            if kept is None or confidence > kept[1] + margin:
                best[index] = (text, float(confidence), turn)
        pieces = {}
        for index, (text, confidence, turn) in sorted(best.items()):
            if confidence < MIN_CONFIDENCE:
                continue
            quad = np.asarray(patches[index][0], dtype=float) / scale
            box = (float(quad[:, 0].min()), float(quad[:, 1].min()),
                   float(quad[:, 0].max()), float(quad[:, 1].max()))
            strip_height, strip_width = patches[index][1].shape[:2]
            across = (strip_width if turn in (90, 270) else strip_height) / scale
            pieces[index] = Piece(text, confidence, box, turn, reader_name,
                                  tuple(tuple(float(v) for v in point) for point in quad), float(across))
        return pieces if by_patch else list(pieces.values())

    def passes(self, picture):
        """Read a picture in passes, each trying harder than the last. Each pass yields
        its new pieces plus lines joined from them; stop when the words are enough."""
        picture = upright(picture)
        small, scale = shrunk(picture)
        patches = self.find_patches(small)
        _, english, second = self._engines()

        # Both readers read the same patches (finding them is the slow part), so
        # each can outvote the other's slip: 3/15/91 read as 3/25/91 by one of them.
        first = self.read_patches(patches, english, scale, "rapid")
        other = self.read_patches(patches, second, scale, "rapid6")
        yield first + join_lines(first) + other + join_lines(other)

        if self.use_tesseract:
            gray, tscale = shrunk(ImageOps.autocontrast(picture.convert("L")), 3000)
            try:
                yield tesseract_pieces(gray, tscale)
            except Exception:       # a picture Tesseract can't take shouldn't stop the rest
                yield []

        imprint = camera_imprint(small)
        if imprint is not None:
            found = self.read_patches(self.find_patches(imprint), english, scale, "imprint")
            yield found + join_lines(found)

    def read_best(self, picture, longest=LONGEST):
        """Every patch of writing read by both readers, keeping the more confident
        reading of each: the English one is best on handwriting, the second on
        small print, and each is sure of itself where it is right."""
        picture = upright(picture)
        small, scale = shrunk(picture, longest)
        patches = self.find_patches(small)
        _, english, second = self._engines()
        first = self.read_patches(patches, english, scale, "rapid", by_patch=True)
        other = self.read_patches(patches, second, scale, "rapid6", by_patch=True)
        best = []
        for index in sorted(set(first) | set(other)):
            choices = [p for p in (first.get(index), other.get(index)) if p is not None]
            best.append(max(choices, key=lambda p: p.confidence))
        return best

    def read_lines(self, picture, longest=LONGEST):
        """The lines of writing in a picture, each one piece (for Photos to Word)."""
        return [piece for _, _, piece, _ in group_lines(self.read_best(picture, longest))]

    def read_all(self, picture):
        """Every pass at once: all the pieces a picture has."""
        return [piece for pieces in self.passes(picture) for piece in pieces]
