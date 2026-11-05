"""
clean_scans.py

Pick a folder of pictures. Each one gets a cleaned copy in a new "cleaned"
subfolder. The originals are never changed, and a "cleaned" folder from an
earlier run is left alone (the next run uses "cleaned (2)"):

  - A page photographed or scanned at an angle is found, its perspective is
    corrected and it is cropped to the page, so its edges come out square.
  - Lines of text that still sit crooked are turned level.
  - Screenshots are already square, so they are left uncropped and unturned,
    unless one shows a page in a viewer with dark controls around it.
  - An obvious document (marks on a plain light background) gets a pure white
    background, with shadows evened out. Pictures and colored areas on the page
    keep their look, and ordinary photos are not whitened at all.
  - Every copy is resized to fit the frame (1200 x 900 unless changed in the
    window), keeping its shape with white filling what is left, and marked
    600 DPI.

Only the top level of the folder is read unless the subfolders box is ticked.
The copies then keep the same folder layout, and "cleaned" folders from earlier
runs are skipped. "what was done.txt" in the new folder lists what happened to
each picture.

Requirements (one-time): numpy, opencv-python-headless and pillow, plus
pillow-heif for iPhone .heic photos. If the Python running this lacks them, it
offers to install them.
"""

import importlib.util
import os
import queue
import re
import site
import subprocess
import sys
import threading
import traceback
from math import hypot

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

TITLE = "Clean Up Scans"
REQUIRED = {"PIL": "pillow", "numpy": "numpy", "cv2": "opencv-python-headless"}


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
    messagebox.showerror(TITLE, "Something went wrong:\n\n" + text[-1500:])


def install_missing():
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add whatever this one lacks."""
    missing = [pip for module, pip in REQUIRED.items()
               if importlib.util.find_spec(module) is None]
    if not missing:
        return
    root = dialog_root()
    if not messagebox.askyesno(
            TITLE,
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

import cv2
import numpy as np
from PIL import Image, ImageOps

try:  # optional HEIC/HEIF support
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

# ---------------------------------------------------------------- settings
WIDTH, HEIGHT, DPI = 1200, 900, 600
PICTURE_TYPES = (".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".tif", ".tiff", ".bmp",
                 ".gif", ".webp", ".heic", ".heif")
KEEPS_DPI = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")   # others are saved as .jpg
FULL_SIZE = "full size"    # subfolder of the results that holds the full-size copies
WORK_SIDE = 1000           # long side of the small copy the analysis runs on
FOCAL = 0.9                # typical phone lens, as a share of the picture's long side
MIN_PAGE = 0.10            # a page must cover at least this share of the picture
MIN_MARKS = 25             # small dark marks needed before they count as text


def odd(n):
    return max(3, int(n)) | 1


def shrink(img, side):
    h, w = img.shape[:2]
    scale = side / max(h, w)
    if scale >= 1:
        return img, 1.0
    size = (max(1, round(w * scale)), max(1, round(h * scale)))
    return cv2.resize(img, size, interpolation=cv2.INTER_AREA), scale


def box(n):
    return cv2.getStructuringElement(cv2.MORPH_RECT, (n, n))


# ---------------------------------------------------------- finding the page
def order_corners(quad):
    """Top-left, top-right, bottom-right, bottom-left, starting at whichever
    corner needs the page turned least."""
    middle = quad.mean(axis=0)
    around = quad[np.argsort(np.arctan2(*(quad - middle).T[::-1]))]
    low, high = quad.min(axis=0), quad.max(axis=0)
    frame = np.array([low, [high[0], low[1]], high, [low[0], high[1]]])
    return min((np.roll(around, -i, axis=0) for i in range(4)),
               key=lambda q: np.linalg.norm(q - frame, axis=1).sum()).astype(np.float32)


def cross_point(a, b):
    (ax, ay, apx, apy), (bx, by, bpx, bpy) = a, b
    d = ax * by - ay * bx
    if abs(d) < 1e-6:
        return None
    t = ((bpx - apx) * by - (bpy - apy) * bx) / d
    return [apx + t * ax, apy + t * ay]


def fit_quad(contour):
    """Four corners for an outline. A straight line is fitted to the points along
    each side, so a thumb over one corner or a torn edge doesn't drag it."""
    hull = cv2.convexHull(contour)
    around = cv2.arcLength(hull, True)
    quad = None
    for share in np.arange(0.01, 0.12, 0.005):
        rough = cv2.approxPolyDP(hull, share * around, True)
        if len(rough) == 4:
            quad = rough.reshape(4, 2).astype(np.float32)
            break
    if quad is None:
        return None
    quad = order_corners(quad)

    points = contour.reshape(-1, 2).astype(np.float32)
    lines = []
    for i in range(4):
        a, b = quad[i], quad[(i + 1) % 4]
        along = b - a
        length = max(np.linalg.norm(along), 1)
        along = along / length
        rel = points - a
        off = np.abs(rel[:, 0] * along[1] - rel[:, 1] * along[0])
        pos = rel @ along
        near = points[(off < 0.01 * around) & (pos > 0.1 * length) & (pos < 0.9 * length)]
        if len(near) >= 10:
            lines.append(cv2.fitLine(near, cv2.DIST_HUBER, 0, 0.01, 0.01).ravel())
        else:
            lines.append(np.array([along[0], along[1], a[0], a[1]]))
    fitted = [cross_point(lines[i - 1], lines[i]) for i in range(4)]
    if any(p is None for p in fitted):
        return quad
    fitted = np.array(fitted, dtype=np.float32)
    if np.linalg.norm(fitted - quad, axis=1).max() > 0.04 * around:
        return quad
    return fitted


def page_outlines(small, lab):
    """Possible page outlines from several clues: what differs from the colors
    around the rim of the picture, the same ignoring brightness (a shadow or a
    black banner doesn't split the page from a colored table), what is bright,
    and sharp edges."""
    h, w = lab.shape[:2]
    rim = max(2, int(0.03 * min(h, w)))
    around = np.vstack([lab[:rim].reshape(-1, 3), lab[-rim:].reshape(-1, 3),
                        lab[:, :rim].reshape(-1, 3), lab[:, -rim:].reshape(-1, 3)])
    offset = lab.astype(np.float32) - np.median(around, axis=0)
    away = np.clip(np.linalg.norm(offset, axis=2), 0, 255).astype(np.uint8)
    hue_away = np.clip(np.linalg.norm(offset[:, :, 1:], axis=2) * 4, 0, 255).astype(np.uint8)
    masks = [cv2.threshold(away, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
             cv2.threshold(hue_away, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
             cv2.threshold(lab[:, :, 0], 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
             cv2.Canny(lab[:, :, 0], 30, 90)]
    closer = box(odd(max(h, w) / 66))
    for mask in masks:
        closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, closer)
        contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
        for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:2]:
            if cv2.contourArea(contour) >= MIN_PAGE * h * w:
                quad = fit_quad(contour)
                if quad is not None:
                    yield quad


def side_angles(quad):
    """How far each side is from level or plumb, and each corner's angle."""
    tilts, corners = [], []
    for i in range(4):
        a, b, c = quad[i], quad[(i + 1) % 4], quad[(i + 2) % 4]
        x, y = b - a
        tilt = abs(np.degrees(np.arctan2(y, x))) % 90
        tilts.append(min(tilt, 90 - tilt))
        u, v = a - b, c - b
        cosine = (u @ v) / max(np.linalg.norm(u) * np.linalg.norm(v), 1e-6)
        corners.append(np.degrees(np.arccos(np.clip(cosine, -1, 1))))
    return tilts, corners


def edge_support(lab, quad):
    """For each side, the share of it where the color really changes from just
    inside to just outside. A side lying along the picture's own edge counts as
    supported, since a page can run off the frame."""
    h, w = lab.shape[:2]
    gap = max(3, max(h, w) // 150)
    edge = 0.015 * max(h, w)
    middle = quad.mean(axis=0)
    sides = []
    for i in range(4):
        a, b = quad[i], quad[(i + 1) % 4]
        on_frame = any((abs(a[n]) < edge and abs(b[n]) < edge) or
                       (abs(a[n] - size) < edge and abs(b[n] - size) < edge)
                       for n, size in ((0, w - 1), (1, h - 1)))
        if on_frame:
            sides.append(1.0)
            continue
        along = (b - a) / max(np.linalg.norm(b - a), 1)
        out = np.array([along[1], -along[0]])
        if out @ ((a + b) / 2 - middle) < 0:
            out = -out
        spots = a + (b - a) * np.linspace(0.1, 0.9, 40)[:, None]

        def read(points):
            x = np.clip(points[:, 0].round().astype(int), 0, w - 1)
            y = np.clip(points[:, 1].round().astype(int), 0, h - 1)
            return lab[y, x].astype(np.float32)

        change = np.linalg.norm(read(spots - out * gap) - read(spots + out * gap), axis=1)
        sides.append(float((change > 12).mean()))
    return sides


def true_shape(quad, shape):
    """Width over height of the real page. Side lengths alone understate whichever
    direction the camera was tilted along, so this works it back through a typical
    phone lens (the rectangle method of Zhang and He)."""
    h, w = shape[:2]
    focal = FOCAL * max(h, w)
    tl, tr, br, bl = [np.array([x - w / 2, y - h / 2, 1.0]) for x, y in quad]
    k2 = np.dot(np.cross(tl, br), bl) / np.dot(np.cross(tr, br), bl)
    k3 = np.dot(np.cross(tl, br), tr) / np.dot(np.cross(bl, br), tr)
    across, down = k2 * tr - tl, k3 * bl - tl
    top = (across[0] ** 2 + across[1] ** 2) / focal ** 2 + across[2] ** 2
    bottom = (down[0] ** 2 + down[1] ** 2) / focal ** 2 + down[2] ** 2
    return (top / bottom) ** 0.5


def flatten(rgb, quad):
    """Warp the outlined page to a rectangle, then trim a sliver so no background
    is left along its edges."""
    tl, tr, br, bl = quad
    wide = (np.linalg.norm(tr - tl) + np.linalg.norm(br - bl)) / 2
    tall = (np.linalg.norm(bl - tl) + np.linalg.norm(br - tr)) / 2
    plain = wide / max(tall, 1)
    shape = true_shape(quad, rgb.shape)
    if not (np.isfinite(shape) and 0.75 < shape / plain < 1.33):
        shape = plain
    out_h = max(np.linalg.norm(bl - tl), np.linalg.norm(br - tr), 2)
    out_w = max(out_h * shape, 2)
    if max(out_w, out_h) > max(rgb.shape[:2]):       # no more pixels than the picture had
        fit = max(rgb.shape[:2]) / max(out_w, out_h)
        out_w, out_h = out_w * fit, out_h * fit
    size = (int(round(out_w)), int(round(out_h)))
    target = np.array([[0, 0], [size[0] - 1, 0], [size[0] - 1, size[1] - 1], [0, size[1] - 1]],
                      dtype=np.float32)
    matrix = cv2.getPerspectiveTransform(quad.astype(np.float32), target)
    page = cv2.warpPerspective(rgb, matrix, size, flags=cv2.INTER_CUBIC,
                               borderMode=cv2.BORDER_REPLICATE)
    trim = max(1, int(0.004 * min(size)))
    return page[trim:size[1] - trim, trim:size[0] - trim]

# ------------------------------------------------------ reading what's on it
class Look:
    """What a page is made of, measured on a small copy."""


def study_page(rgb, on_sheet=False):
    """on_sheet: this was already found to be a separate sheet lying on something,
    so big lettering with few small marks can still count as a document."""
    small, _ = shrink(rgb, WORK_SIDE)
    long_side = max(small.shape[:2])
    k = odd(long_side / 66)
    # Closing paints over anything thinner than k, which leaves the paper. A wider
    # window does worse along the hard edge of a shadow.
    paper = cv2.medianBlur(cv2.morphologyEx(small, cv2.MORPH_CLOSE, box(k)), odd(k * 1.5))
    gray = cv2.cvtColor(small, cv2.COLOR_RGB2GRAY).astype(np.float32)
    paper_gray = cv2.cvtColor(paper, cv2.COLOR_RGB2GRAY).astype(np.float32)
    lab = cv2.cvtColor(paper, cv2.COLOR_RGB2LAB).astype(np.float32)
    level = np.percentile(paper_gray, 75)
    lit = paper_gray > 0.8 * level
    tint = np.median(lab[lit][:, 1:], axis=0) - 128 if lit.any() else np.array([0.0, 0.0])

    # Solid areas are pictures, dark bands and colored panels: paper estimates
    # must not come from them, and they keep their look.
    colored = np.hypot(lab[:, :, 1] - 128 - tint[0], lab[:, :, 2] - 128 - tint[1]) > 22
    candidates = (gray < 0.88 * paper_gray) | (paper_gray < 0.55 * level) | colored
    s = odd(long_side / 110)
    solid = cv2.morphologyEx(candidates.astype(np.uint8), cv2.MORPH_OPEN, box(s))
    solid = cv2.morphologyEx(solid, cv2.MORPH_CLOSE, box(3 * s))
    count, labels, stats, _ = cv2.connectedComponentsWithStats(solid, connectivity=8)
    big = np.flatnonzero(stats[:, cv2.CC_STAT_AREA] >= 0.005 * solid.size)
    big = big[big != 0]
    solid = np.isin(labels, big)

    # Small dark marks, the bottoms of which line up if they are text.
    ink = (gray < 0.72 * paper_gray) & ~solid
    count, _, stats, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
    stats = stats[1:]
    tall, wide = stats[:, cv2.CC_STAT_HEIGHT], stats[:, cv2.CC_STAT_WIDTH]
    keep = ((tall >= 3) & (tall <= 0.06 * long_side) & (wide <= 0.12 * long_side)
            & (stats[:, cv2.CC_STAT_AREA] >= 4))
    marks = stats[keep]
    points = np.column_stack((marks[:, cv2.CC_STAT_LEFT] + marks[:, cv2.CC_STAT_WIDTH] / 2,
                              marks[:, cv2.CC_STAT_TOP] + marks[:, cv2.CC_STAT_HEIGHT]))
    mark_size = float(np.median(marks[:, cv2.CC_STAT_HEIGHT])) if len(marks) else 0.0

    # A clean graphic or scan has one exact background color. Paper lit by a lamp
    # spreads over a range of values instead.
    counts = np.bincount(paper_gray.astype(np.uint8).ravel(), minlength=256)
    mode = int(counts.argmax())
    near_mode = np.abs(paper_gray - mode) <= 1
    flat_share = near_mode.mean()
    spread = max((np.abs(paper_gray - mode) <= 8).mean(), 1e-6)

    look = Look()
    look.small, look.paper, look.solid, look.points = small, paper, solid, points
    look.mark_size = mark_size
    look.paper_color = (np.median(paper[lit], axis=0) if lit.any()
                        else np.median(paper.reshape(-1, 3), axis=0))
    look.flat_color = np.median(paper[near_mode], axis=0)
    look.flat = bool(mode >= 170 and flat_share >= 0.1 and flat_share / spread >= 0.6)
    look.digital_white = bool(look.flat and mode >= 250)   # no photo of paper is this clean

    open_part = ~solid
    open_count = max(int(open_part.sum()), 1)
    relative = gray / np.maximum(paper_gray, 1)
    plain = (relative >= 0.8) & open_part
    grain = float(relative[plain].std()) if plain.any() else 1.0
    ink_share = (ink & open_part).sum() / open_count
    _, line_gain = best_skew(points, 15, mark_size)
    look.stats = dict(level=float(level), tint=[float(t) for t in tint], flat_share=float(flat_share),
                      mode=mode, solid=float(solid.mean()), plain=plain.sum() / open_count,
                      grain=grain, ink=float(ink_share), marks=len(points), gain=float(line_gain))

    if look.flat:
        pale_paper = abs(tint[0]) <= 8 and -5 <= tint[1] <= 18
        look.is_document = bool(pale_paper and solid.mean() <= 0.6 and
                                (mode >= 225 or (flat_share >= 0.35 and len(points) >= MIN_MARKS)))
    else:
        paper_like = level >= 90 and -8 <= tint[0] <= 12 and -6 <= tint[1] <= 36
        text = len(points) >= MIN_MARKS and line_gain >= 1.5
        big_lettering = on_sheet and solid.mean() <= 0.3
        look.is_document = bool(paper_like and solid.mean() <= 0.5 and plain.sum() / open_count >= 0.6
                                and grain <= 0.07 and ink_share <= 0.35
                                and ((text and ink_share >= 0.002) or big_lettering))
    return look


def best_skew(points, limit, mark_size):
    """The turn in degrees that lines the marks' bottoms up best, and how much
    better that is than a typical angle (1 means no sign of lines)."""
    if len(points) < MIN_MARKS:
        return 0.0, 1.0
    x, y = points[:, 0], points[:, 1]
    step = max(mark_size / 2, 1.5)

    def score(degrees):
        t = np.radians(degrees)
        rows = ((y * np.cos(t) - x * np.sin(t)) / step)
        rows = (rows - rows.min()).astype(np.int64)
        return float((np.bincount(rows) ** 2).sum())

    coarse = np.arange(-limit, limit + 0.25, 0.5)
    scores = np.array([score(a) for a in coarse])
    best = float(coarse[scores.argmax()])
    fine = np.arange(best - 0.5, best + 0.5, 0.05)
    fine_scores = np.array([score(a) for a in fine])
    best = float(fine[fine_scores >= 0.995 * fine_scores.max()].mean())   # middle of any tie
    return best, score(best) / max(float(np.median(scores)), 1.0)


def level_text(page, look, limit):
    """Turn the page if its text is clearly crooked. A clean graphic is usually
    straight already, and sloped labels on a drawing can look like crooked text,
    so it must show stronger evidence."""
    angle, gain = best_skew(look.points, limit, look.mark_size)
    if gain < (2.5 if look.flat else 1.6) or abs(angle) < 0.3:
        return page, 0.0
    if not lines_run_across(look.points, angle, look.mark_size, limit) or already_square(page):
        return page, 0.0
    h, w = page.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    cos, sin = abs(matrix[0, 0]), abs(matrix[0, 1])
    new_w, new_h = int(h * sin + w * cos + 0.5), int(h * cos + w * sin + 0.5)
    matrix[0, 2] += (new_w - w) / 2
    matrix[1, 2] += (new_h - h) / 2
    fill = [float(c) for c in look.paper_color]
    return (cv2.warpAffine(page, matrix, (new_w, new_h), flags=cv2.INTER_CUBIC,
                           borderMode=cv2.BORDER_CONSTANT, borderValue=fill), angle)


def already_square(page):
    """Long straight lines settle it: if most near-level ones are exactly level or
    plumb, the picture isn't crooked, whatever its labels suggest."""
    small, _ = shrink(page, 1600)
    edges = cv2.Canny(cv2.cvtColor(small, cv2.COLOR_RGB2GRAY), 50, 150)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 720, threshold=80,
                            minLineLength=int(0.15 * max(small.shape[:2])), maxLineGap=4)
    if lines is None:
        return False
    exact = near = 0.0
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        dx, dy = abs(int(x2) - int(x1)), abs(int(y2) - int(y1))
        tilt = np.degrees(np.arctan2(min(dx, dy), max(dx, dy)))
        if tilt <= 5:
            near += hypot(dx, dy)
            exact += hypot(dx, dy) if tilt <= 0.12 else 0
    return near > 0 and exact / near >= 0.6


def lines_run_across(points, angle, mark_size, limit):
    """A few labels on a drawing can agree on any angle. Real text has several
    separate lines, and the left and right of the page lean the same way."""
    x, y = points[:, 0], points[:, 1]
    t = np.radians(angle)
    rows = (y * np.cos(t) - x * np.sin(t)) / max(mark_size / 2, 1.5)
    filled = np.bincount((rows - rows.min()).astype(np.int64)) >= 3
    lines = int((filled[1:] & ~filled[:-1]).sum()) + int(filled[0])
    middle = np.median(x)
    left, left_gain = best_skew(points[x < middle], limit, mark_size)
    right, right_gain = best_skew(points[x >= middle], limit, mark_size)
    return lines >= 5 and min(left_gain, right_gain) >= 1.4 and abs(left - right) <= 0.4


def surroundings_differ(small, quad, page_look):
    """A real page stands out from what it lies on. A ruled box printed on a white
    page doesn't, and cropping to it would throw away the rest of the page."""
    mask = np.ones(small.shape[:2], np.uint8)
    cv2.fillConvexPoly(mask, quad.astype(np.int32), 0)
    if mask.sum() < 0.005 * mask.size:
        return True
    outside = cv2.cvtColor(small, cv2.COLOR_RGB2LAB)[mask > 0].astype(np.float32)
    swatch = np.uint8([[page_look.paper_color]])
    inside = cv2.cvtColor(swatch, cv2.COLOR_RGB2LAB)[0, 0].astype(np.float32)
    return np.linalg.norm(np.median(outside, axis=0) - inside) > 15


def surroundings(small, quad):
    """How busy the area outside an outline is (share of it that is edges), and how
    bright. Whatever a page lies on is plain next to the rest of a composite graphic."""
    mask = np.full(small.shape[:2], 255, np.uint8)
    cv2.fillConvexPoly(mask, quad.astype(np.int32), 0)
    mask = cv2.erode(mask, box(odd(max(small.shape[:2]) / 80)))
    if cv2.countNonZero(mask) < 0.005 * mask.size:
        return 0.0, 255.0
    gray = cv2.GaussianBlur(cv2.cvtColor(small, cv2.COLOR_RGB2GRAY), (5, 5), 0)
    edges = cv2.Canny(gray, 30, 90)
    return float((edges[mask > 0] > 0).mean()), float(np.median(gray[mask > 0]))


def square_up(rgb):
    """Crop to the page if there is one. Returns the page, its Look and whether a
    page was found. Better to miss a page than to crop away half a picture: one
    panel of a composite graphic or a window in a photo must not become the whole
    thing. So the outline must be tilted, as a real sheet nearly always is (or sit
    on black), with the color changing along most of its outline, on a plain background."""
    whole = study_page(rgb)
    small, scale = shrink(rgb, WORK_SIDE)
    lab = cv2.GaussianBlur(cv2.cvtColor(small, cv2.COLOR_RGB2LAB), (5, 5), 0)
    area = small.shape[0] * small.shape[1]
    ranked = []
    for quad in page_outlines(small, lab):
        share = cv2.contourArea(quad) / area
        if not MIN_PAGE <= share <= 0.98:
            continue
        tilts, corners = side_angles(quad)
        if max(abs(c - 90) for c in corners) > 25:
            continue
        sides = edge_support(lab, quad)
        if min(sides) < 0.25 or sum(sides) / 4 < 0.8:     # one edge can be lost in dim light
            continue
        ranked.append((share * min(sides) * sum(sides) / 4, quad, share, tilts, corners))
    ranked.sort(key=lambda item: item[0], reverse=True)

    tried = []
    for _, quad, share, tilts, corners in ranked:
        if any(np.abs(quad - other).max() < 8 for other in tried):
            continue
        tried.append(quad)
        page = flatten(rgb, quad / scale)
        look = study_page(page, on_sheet=True)
        busy, brightness = surroundings(small, quad)
        tilted = max(tilts) > 0.7
        if whole.digital_white and brightness >= 40:     # a panel in a graphic
            continue
        if look.is_document:
            has_content = len(look.points) >= MIN_MARKS or look.solid.mean() > 0.02
            good = (has_content and surroundings_differ(small, quad, look) and busy <= 0.03
                    and (share >= 0.25 or busy <= 0.02)
                    and (tilted or (brightness < 40 and len(look.points) >= MIN_MARKS)))
        else:
            good = (tilted and share >= 0.2 and busy <= 0.01
                    and max(abs(c - 90) for c in corners) <= 20)
        if good:
            return page, look, True
    return rgb, whole, False


# -------------------------------------------------------------- whitening
def whiten_flat(page, look):
    """One exact background color: send it to white and lift everything else by
    the same proportion, so pale fills in a drawing stay visible."""
    white = np.maximum(look.flat_color, 1) * (1 - 6 / 255)
    table = np.clip(np.arange(256, dtype=np.float32)[:, None] * 255 / white[None, :], 0, 255)
    table = table.astype(np.uint8)
    return np.dstack([cv2.LUT(page[:, :, c], table[:, c]) for c in range(3)])


def whiten_photo(page, look):
    """Paper photographed or scanned: divide out the lighting, measured from the
    paper itself and bridged across solid areas, then push the paper to white."""
    h, w = page.shape[:2]
    paper = look.paper.astype(np.float32)
    usable = (~look.solid).astype(np.float32)
    reach = max(paper.shape[:2]) / 25
    weight = cv2.GaussianBlur(usable, (0, 0), reach)
    bridged = cv2.GaussianBlur(paper * usable[:, :, None], (0, 0), reach)
    bridged = bridged / np.maximum(weight, 1e-3)[:, :, None]
    far = weight < 0.02
    if far.any():
        bridged[far] = look.paper_color
    light = np.where(look.solid[:, :, None], bridged, paper)
    light = cv2.GaussianBlur(light, (0, 0), max(paper.shape[:2]) / 130)
    light = cv2.resize(light, (w, h), interpolation=cv2.INTER_LINEAR)

    ratio = page.astype(np.float32) / np.maximum(light, 1)
    keep = cv2.resize(look.solid.astype(np.float32), (w, h), interpolation=cv2.INTER_LINEAR)
    sample = ratio.mean(axis=2)[::4, ::4]
    sample = sample[keep[::4, ::4] < 0.5]
    if not sample.size:
        sample = ratio.mean(axis=2)[::4, ::4]
    paper_part = sample[sample > 0.75]
    white = float(np.clip(np.percentile(paper_part, 8), 0.80, 0.97)) if paper_part.size else 0.9
    black = float(np.clip(np.percentile(sample, 0.5), 0, 0.25))
    stretched = (ratio - black) / (white - black)
    out = stretched * (1 - keep[:, :, None]) + ratio * keep[:, :, None]
    return clear_slivers((np.clip(out, 0, 1) * 255).astype(np.uint8))


def clear_slivers(page):
    """Paint out a thin dark strip along an edge, left over from the background.
    A band wider than the limit is part of the page and stays."""
    h, w = page.shape[:2]
    dark = cv2.cvtColor(page, cv2.COLOR_RGB2GRAY) < 200
    across, down = dark.mean(axis=1), dark.mean(axis=0)
    limit_h, limit_w = max(2, int(0.02 * h)), max(2, int(0.02 * w))

    def depth(values, limit):
        n = 0
        while n < limit and values[n] > 0.5:
            n += 1
        return n if n < limit else 0

    top, bottom = depth(across, limit_h), depth(across[::-1], limit_h)
    left, right = depth(down, limit_w), depth(down[::-1], limit_w)
    page[:top], page[h - bottom:], page[:, :left], page[:, w - right:] = 255, 255, 255, 255
    return page


# -------------------------------------------------------------- one picture
def open_picture(path):
    with Image.open(path) as img:
        pages = getattr(img, "n_frames", 1)
        icc = img.info.get("icc_profile") if img.mode == "RGB" else None
        img = ImageOps.exif_transpose(img)
        if img.mode.startswith("I"):         # 16-bit scans
            values = np.asarray(img, dtype=np.float32)
            img = Image.fromarray((values / 257 if values.max() > 255 else values).astype(np.uint8))
        if img.mode in ("RGBA", "LA", "PA") or "transparency" in img.info:
            img = img.convert("RGBA")
            back = Image.new("RGBA", img.size, "white")
            back.alpha_composite(img)
            img = back
        return np.array(img.convert("RGB")), pages, icc


def is_screen_capture(name, rgb):
    """A screenshot is square already, and its panels and buttons look like pages.
    Phone screens are also taller than any camera's 16:9."""
    h, w = rgb.shape[:2]
    return bool(re.search(r"screen[ _-]?shot", name, re.I)) or max(h, w) / min(h, w) >= 1.9


def viewer_page(rgb):
    """The page inside a screenshot of a photo viewer: a bright band running edge
    to edge, with dark app controls above and below it. None if there isn't one."""
    small, scale = shrink(rgb, WORK_SIDE)
    rows = np.median(cv2.cvtColor(small, cv2.COLOR_RGB2GRAY), axis=1)
    h = len(rows)
    bright = (rows >= 170).astype(np.uint8)[:, None]
    bright = cv2.morphologyEx(bright, cv2.MORPH_CLOSE, np.ones((odd(0.03 * h), 1), np.uint8))
    count, _, stats, _ = cv2.connectedComponentsWithStats(bright, connectivity=8)
    if count < 2:
        return None
    biggest = 1 + stats[1:, cv2.CC_STAT_AREA].argmax()
    top = stats[biggest, cv2.CC_STAT_TOP]
    bottom = top + stats[biggest, cv2.CC_STAT_HEIGHT]
    if bottom - top < 0.3 * h or top + h - bottom < 0.1 * h:
        return None
    if (np.concatenate((rows[:top], rows[bottom:])) < 60).mean() < 0.9:    # content would be lost
        return None
    return rgb[int(top / scale):int(bottom / scale)]


def clean_screen(rgb):
    """A screenshot: only a page shown in a viewer is cropped, and only a clean
    background is whitened. Evening out lighting has no place on a screen."""
    done = []
    page, look = rgb, study_page(rgb)
    inner = viewer_page(rgb)
    if inner is not None:
        inner_look = study_page(inner)
        if inner_look.is_document and inner_look.flat:
            page, look = inner, inner_look
            done.append("page squared up")
    if look.is_document and look.flat:
        page = whiten_flat(page, look)
        done.append("background whitened")
    return page, done


def clean_page(rgb):
    """Square up, level and whiten. Returns the page and a list of what was done."""
    done = []
    page, look, squared = square_up(rgb)
    if squared:
        done.append("page squared up")
    if look.is_document:
        page, angle = level_text(page, look, 4 if squared else 15)
        if angle:
            done.append(f"leveled {abs(angle):.1f} degrees")
            look = study_page(page)
        if look.is_document:
            page = whiten_flat(page, look) if look.flat else whiten_photo(page, look)
            done.append("background whitened")
    return page, done


def fit_frame(page, width, height):
    """Resize to fit, keeping the page's shape, with white filling the rest."""
    h, w = page.shape[:2]
    scale = min(width / w, height / h)
    size = (max(1, min(width, round(w * scale))), max(1, min(height, round(h * scale))))
    method = cv2.INTER_AREA if scale < 1 else cv2.INTER_CUBIC
    frame = np.full((height, width, 3), 255, np.uint8)
    x, y = (width - size[0]) // 2, (height - size[1]) // 2
    frame[y:y + size[1], x:x + size[0]] = cv2.resize(page, size, interpolation=method)
    return frame


def save_picture(rgb, path, dpi, icc):
    options = {"dpi": (dpi, dpi)}
    if icc:
        options["icc_profile"] = icc
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        options.update(quality=95, subsampling=0)
    elif ext in (".tif", ".tiff"):
        options["compression"] = "tiff_lzw"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    Image.fromarray(rgb).save(path, **options)


def clean_file(source, target, full_target, width, height, dpi, turn_frame):
    rgb, pages, icc = open_picture(source)
    if not full_target:      # a sharper working copy than this would never reach the file
        rgb, _ = shrink(rgb, 2 * max(width, height))
    screen = is_screen_capture(os.path.basename(source), rgb)
    page, done = clean_screen(rgb) if screen else clean_page(rgb)
    if full_target:
        save_picture(page, full_target, dpi, icc)
    upright = turn_frame and page.shape[0] > page.shape[1]
    frame_w, frame_h = (min(width, height), max(width, height)) if upright else (width, height)
    save_picture(fit_frame(page, frame_w, frame_h), target, dpi, icc)
    if pages > 1:
        done.append(f"only page 1 of {pages} was used")
    return done or ["resized only"]

# ------------------------------------------------------------------ the batch
def natural_key(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def skip_folder(name):
    """Results of an earlier run would be cleaned a second time, and hidden or
    system folders (.git, $RECYCLE.BIN) don't hold the user's pictures."""
    return name.startswith((".", "$")) or bool(re.fullmatch(r"cleaned( \(\d+\))?", name, re.I))


def list_pictures(folder, subfolders=False):
    """Picture paths relative to the folder, in natural order."""
    found = []
    for here, folders, files in os.walk(folder):
        folders[:] = [d for d in folders if subfolders and not skip_folder(d)]
        found += [os.path.relpath(os.path.join(here, f), folder)
                  for f in files if f.lower().endswith(PICTURE_TYPES)]
    return sorted(found, key=natural_key)


def new_folder(folder):
    target = os.path.join(folder, "cleaned")
    n = 1
    while os.path.exists(target):
        n += 1
        target = os.path.join(folder, f"cleaned ({n})")
    return target


def output_name(relative, taken):
    """Keep the path, name and type where the type can hold a DPI. Two originals
    can land on one name (a.heic and a.jpg), so the later one gets a number. Its
    full-size name is reserved too, in case the folder has a "full size" subfolder."""
    stem, ext = os.path.splitext(relative)
    ext = ext if ext.lower() in KEEPS_DPI else ".jpg"
    result, n = stem + ext, 1
    while {result.lower(), os.path.join(FULL_SIZE, result).lower()} & taken:
        n += 1
        result = f"{stem} ({n}){ext}"
    taken.update((result.lower(), os.path.join(FULL_SIZE, result).lower()))
    return result


def run_batch(folder, target, settings, report, stop):
    """Clean every picture, calling report(kind, ...) as it goes."""
    width, height, dpi, turn_frame, keep_full, subfolders = settings
    pictures = list_pictures(folder, subfolders)
    full_folder = os.path.join(target, FULL_SIZE)
    os.makedirs(target, exist_ok=True)
    taken = set()
    log, whitened, squared, leveled, failed = [], 0, 0, 0, 0
    finished = 0
    for n, name in enumerate(pictures, 1):
        if stop.is_set():
            break
        report("progress", n, len(pictures), name)
        out_name = output_name(name, taken)
        try:
            done = clean_file(os.path.join(folder, name), os.path.join(target, out_name),
                              os.path.join(full_folder, out_name) if keep_full else None,
                              width, height, dpi, turn_frame)
            text = ", ".join(done)
            whitened += "background whitened" in done
            squared += "page squared up" in done
            leveled += any(d.startswith("leveled") for d in done)
        except Exception as e:
            traceback.print_exc()
            text = f"PROBLEM: {e}"
            failed += 1
        finished += 1
        line = f"{name}: {text}"
        log.append(line)
        report("line", line)
    with open(os.path.join(target, "what was done.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n")
    summary = dict(total=len(pictures), finished=finished, squared=squared, leveled=leveled,
                   whitened=whitened, failed=failed, stopped=stop.is_set(), target=target)
    report("finished", summary)
    return summary


# ----------------------------------------------------------------- the window
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(TITLE)
        self.root.report_callback_exception = show_crash
        self.messages = queue.Queue()
        self.stop = threading.Event()
        self.target = None
        self.folder = tk.StringVar()
        self.width = tk.StringVar(value=str(WIDTH))
        self.height = tk.StringVar(value=str(HEIGHT))
        self.dpi = tk.StringVar(value=str(DPI))
        self.turn_frame = tk.BooleanVar(value=False)
        self.keep_full = tk.BooleanVar(value=False)
        self.subfolders = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Choose a folder of pictures to begin.")
        self.build()

    def build(self):
        frame = ttk.Frame(self.root, padding=12)
        frame.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(7, weight=1)

        ttk.Label(frame, text="Folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.folder, width=60).grid(row=0, column=1, sticky="ew", padx=6)
        self.browse_button = ttk.Button(frame, text="Choose folder...", command=self.browse)
        self.browse_button.grid(row=0, column=2)
        self.subfolders_box = ttk.Checkbutton(frame, variable=self.subfolders,
                                              command=self.count_pictures,
                                              text="Include subfolders (the copies keep the same "
                                                   "folder layout)")
        self.subfolders_box.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        size = ttk.Frame(frame)
        size.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 2))
        ttk.Label(size, text="Size:").pack(side="left")
        ttk.Entry(size, textvariable=self.width, width=6).pack(side="left", padx=(6, 2))
        ttk.Label(size, text="x").pack(side="left")
        ttk.Entry(size, textvariable=self.height, width=6).pack(side="left", padx=2)
        ttk.Label(size, text="pixels at").pack(side="left", padx=(2, 4))
        ttk.Entry(size, textvariable=self.dpi, width=6).pack(side="left")
        ttk.Label(size, text="DPI").pack(side="left", padx=4)

        ttk.Checkbutton(frame, variable=self.turn_frame,
                        text="Turn the frame for upright pictures, so they come out larger "
                             "(900 x 1200 instead of 1200 x 900)"
                        ).grid(row=3, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(frame, variable=self.keep_full,
                        text='Also keep a full-size copy of each, in "full size" '
                             "(sharper for printing or reading the text)"
                        ).grid(row=4, column=0, columnspan=3, sticky="w")

        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=3, sticky="w", pady=8)
        self.start_button = ttk.Button(buttons, text="Start", command=self.start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(buttons, text="Stop", command=self.stop.set, state="disabled")
        self.stop_button.pack(side="left", padx=6)
        self.open_button = ttk.Button(buttons, text="Open the new folder", command=self.open_results,
                                      state="disabled")
        self.open_button.pack(side="left")

        self.bar = ttk.Progressbar(frame, mode="determinate")
        self.bar.grid(row=6, column=0, columnspan=3, sticky="ew")

        lines = ttk.Frame(frame)
        lines.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=6)
        lines.columnconfigure(0, weight=1)
        lines.rowconfigure(0, weight=1)
        self.list = tk.Text(lines, height=14, width=90, wrap="none", state="disabled")
        self.list.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(lines, command=self.list.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.list.configure(yscrollcommand=scroll.set)

        ttk.Label(frame, textvariable=self.status, wraplength=700, justify="left"
                  ).grid(row=8, column=0, columnspan=3, sticky="w")

    def browse(self):
        chosen = filedialog.askdirectory(title="Pick the folder of pictures", parent=self.root)
        if chosen:
            self.folder.set(os.path.normpath(chosen))
            self.count_pictures()

    def count_pictures(self):
        folder = self.folder.get().strip().strip('"')
        if os.path.isdir(folder):
            count = len(list_pictures(folder, self.subfolders.get()))
            self.status.set(f"{count} picture(s) found. Copies will go in a new subfolder; "
                            "the originals are not changed.")

    def settings(self):
        """The numbers in the boxes, or None after telling the user what is wrong."""
        try:
            width, height, dpi = int(self.width.get()), int(self.height.get()), int(self.dpi.get())
            if min(width, height, dpi) < 1 or max(width, height) > 30000:
                raise ValueError
        except ValueError:
            messagebox.showerror(TITLE, "Size and DPI must be whole numbers above zero.",
                                 parent=self.root)
            return None
        return (width, height, dpi, self.turn_frame.get(), self.keep_full.get(),
                self.subfolders.get())

    def start(self):
        folder = self.folder.get().strip().strip('"')
        if not os.path.isdir(folder):
            messagebox.showerror(TITLE, "Please choose a folder first.", parent=self.root)
            return
        settings = self.settings()
        if settings is None:
            return
        if not list_pictures(folder, self.subfolders.get()):
            where = "in that folder or its subfolders" if self.subfolders.get() else "directly in that folder"
            messagebox.showinfo(TITLE, f"There are no pictures {where}.", parent=self.root)
            return
        self.target = new_folder(folder)
        self.stop.clear()
        self.list.configure(state="normal")
        self.list.delete("1.0", "end")
        self.list.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        self.subfolders_box.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.open_button.configure(state="disabled")

        def work():
            try:
                run_batch(folder, self.target, settings, lambda *m: self.messages.put(m), self.stop)
            except Exception as e:
                traceback.print_exc()
                self.messages.put(("crashed", str(e)))

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, self.check_messages)

    def check_messages(self):
        carry_on = True
        while not self.messages.empty():
            message = self.messages.get()
            if message[0] == "progress":
                _, n, total, name = message
                self.bar.configure(maximum=total, value=n)
                self.status.set(f"{n} of {total}: {name}")
            elif message[0] == "line":
                self.list.configure(state="normal")
                self.list.insert("end", message[1] + "\n")
                self.list.see("end")
                self.list.configure(state="disabled")
            elif message[0] == "finished":
                self.show_summary(message[1])
                carry_on = False
            elif message[0] == "crashed":
                self.work_ended()
                messagebox.showerror(TITLE, "Something went wrong:\n\n" + message[1], parent=self.root)
                carry_on = False
        if carry_on:
            self.root.after(100, self.check_messages)

    def work_ended(self):
        self.start_button.configure(state="normal")
        self.browse_button.configure(state="normal")
        self.subfolders_box.configure(state="normal")
        self.stop_button.configure(state="disabled")
        if self.target and os.path.isdir(self.target):
            self.open_button.configure(state="normal")

    def show_summary(self, s):
        self.work_ended()
        self.bar.configure(value=s["finished"], maximum=max(s["total"], 1))
        text = (f"{s['finished']} of {s['total']} done. Pages squared up: {s['squared']}. "
                f"Leveled: {s['leveled']}. Whitened: {s['whitened']}.")
        if s["stopped"]:
            text = "Stopped. " + text
        if s["failed"]:
            text += f" Problems: {s['failed']} (see the list)."
        self.status.set(text + f"\nSaved in: {s['target']}")

    def open_results(self):
        if self.target:
            os.startfile(self.target)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()