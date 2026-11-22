"""Photo Stitcher's photo work: opening photos, joining them into one picture,
finding the ragged black edges to trim, and saving. No window code is here, so
it can be tested on its own (test_photo_stitcher.py).

The joining itself is OpenCV's ready-made stitcher: it finds the same details
in overlapping photos, works out how they fit, evens out the brightness and
hides the seams. The photos can be in any order, side by side, top to bottom,
or in rows and columns.
"""

from __future__ import annotations

import ctypes
import datetime
import os
import re
from dataclasses import dataclass, field

import cv2
import numpy as np

# OpenCV's graphics-card part (OpenCL) is left off. With some fits its warping
# failed ("CL_OUT_OF_RESOURCES", seen 2026-11-22 with OpenCV 5.0 on an NVIDIA
# card) and took the whole stitching process down; graphics drivers differ from
# PC to PC. The ordinary processor was as fast or faster on every test.
cv2.ocl.setUseOpenCL(False)

PHOTO_TYPES = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp", ".heic", ".heif")
SAVE_TYPES = {".jpg": "JPG", ".jpeg": "JPG", ".png": "PNG", ".tif": "TIFF", ".tiff": "TIFF"}
JPEG_MAX = 65500                  # libjpeg can't write a JPG wider or taller than this
TURNING, FLAT = "turning", "flat"
KIND_NAMES = {TURNING: "turned-camera photos", FLAT: "flat things"}
TRIM_WORK_SIZE = 1000             # trimming looks at the picture shrunk to about this size

WHY_NOT_JOINED = {
    cv2.Stitcher_ERR_NEED_MORE_IMGS:
        "The photos didn't overlap enough to be joined. Photos that sit next to each other "
        "need to share about a third of their picture, with clear details in that part "
        "(not just sky, water or a plain wall).",
    cv2.Stitcher_ERR_HOMOGRAPHY_EST_FAIL:
        "The matching details didn't fit together. Try photos taken from one spot, "
        "with about a third of each photo overlapping the next.",
    cv2.Stitcher_ERR_CAMERA_PARAMS_ADJUST_FAIL:
        "The photos matched, but how they fit together couldn't be worked out. "
        "Leave out blurred photos and try again.",
}


class PhotoError(Exception):
    """A photo couldn't be opened. The message finishes "<file name> ..."."""


class StitchError(Exception):
    """The photos couldn't be joined. The message says why in plain words."""


class SaveError(Exception):
    """The panorama couldn't be saved. The message says why in plain words."""


class _OutOfMemory(Exception):
    pass


# ------------------------------------------------------------------ opening photos

def _open_heic_with_pillow() -> None:
    """Teach Pillow to open iPhone HEIC photos (needs pillow-heif)."""
    import pillow_heif
    pillow_heif.register_heif_opener()


def load_photo(path: str) -> np.ndarray:
    """The photo as a picture OpenCV can use (blue-green-red, 8 bits), turned
    the right way up. Works with any letters in the path."""
    if os.path.splitext(path)[1].lower() in (".heic", ".heif"):
        return _load_heic(path)
    try:
        data = np.fromfile(path, dtype=np.uint8)
    except OSError as error:
        raise PhotoError(f"can't be read ({error.strerror or error})") from error
    image = cv2.imdecode(data, cv2.IMREAD_COLOR) if data.size else None
    if image is None:
        raise PhotoError("can't be opened as a photo")
    return image


def _load_heic(path: str) -> np.ndarray:
    try:
        from PIL import Image, ImageOps
        _open_heic_with_pillow()
    except ImportError as error:
        raise PhotoError("is an iPhone (HEIC) photo, and the part that opens those "
                         "is missing") from error
    try:
        with Image.open(path) as picture:
            upright = ImageOps.exif_transpose(picture).convert("RGB")
    except OSError as error:
        raise PhotoError("can't be opened as a photo") from error
    except Exception as error:            # pillow-heif raises its own kinds for damaged files
        raise PhotoError("can't be opened as a photo") from error
    return cv2.cvtColor(np.asarray(upright), cv2.COLOR_RGB2BGR)


def photo_size(path: str) -> tuple[int, int]:
    """(width, height) of the photo the right way up, read quickly from the
    start of the file."""
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None         # big scans are the user's own files
    try:
        _open_heic_with_pillow()
    except ImportError:
        pass
    try:
        with Image.open(path) as picture:
            width, height = picture.size
            turn = picture.getexif().get(0x0112, 1)      # EXIF orientation
    except Exception:                     # Pillow raises many kinds for files it can't read
        height, width = load_photo(path).shape[:2]     # OpenCV may still read it
        return width, height
    return (height, width) if turn in (5, 6, 7, 8) else (width, height)


def _natural(name: str) -> list:
    """Sorts IMG_9 before IMG_10."""
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", name)]


def photos_in_folder(folder: str) -> list[str]:
    """The photos directly in the folder (not in folders inside it), in name order."""
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    paths = [os.path.join(folder, name) for name in names
             if os.path.splitext(name)[1].lower() in PHOTO_TYPES]
    return sorted((p for p in paths if os.path.isfile(p)),
                  key=lambda p: _natural(os.path.basename(p)))


# ------------------------------------------------------------------ joining

@dataclass
class StitchResult:
    image: np.ndarray             # the joined picture, untrimmed
    used: list[int]               # which of the given photos are in it
    kind: str                     # TURNING or FLAT, whichever made it
    switched: bool = False        # True if the other kind than the one asked for made it
    scale: float = 1.0            # below 1 if memory was short and it was made smaller
    coverage: np.ndarray | None = None   # True where the photos are, if it was worked out
    stitcher: object = field(default=None, repr=False)   # OpenCV's stitcher that made it


@dataclass
class _Plan:
    """How the photos fit for one kind, found but not yet made into a picture."""
    kind: str
    status: int
    stitcher: object = None
    used: list[int] = field(default_factory=list)
    believable: bool = True       # False for a turned-camera fit with an impossible zoom
    straighten: bool = True       # the turned-camera kind's horizon straightening was on


def _out_of_memory(error: Exception) -> bool:
    if isinstance(error, MemoryError):
        return True
    text = str(error).lower()
    return isinstance(error, cv2.error) and (
        getattr(error, "code", None) == cv2.Error.StsNoMem
        or "memory" in text or "alloc" in text)      # std::bad_alloc reads "bad allocation"


def _check_opencv_error(error: Exception) -> None:
    """Out of memory and OpenCV's "Unknown C++ exception" are raised on (the
    second can leave its stitcher frozen for the next try, seen with 4.14, 5.0
    and 5.1). Any other OpenCV error is OpenCV refusing these photos this way,
    and the caller goes on with the next way."""
    if _out_of_memory(error):
        raise _OutOfMemory() from error
    if "unknown c++ exception" in str(error).lower():
        raise StitchError("The joining part of the program hit an internal error with these "
                          "photos. Try again with one photo fewer, or the other kind in Step 2."
                          ) from error


def _mode(kind: str) -> int:
    return cv2.Stitcher_PANORAMA if kind == TURNING else cv2.Stitcher_SCANS


_CLAHE = None


def _stretched(image: np.ndarray) -> np.ndarray:
    """A copy brightened evenly: its darkest-to-brightest range (ignoring the
    darkest and brightest half percent, so one lamp can't spoil it) stretched
    to the full 0-255. Used only to find the matching details."""
    low, high = np.percentile(image, (0.5, 99.5))
    if high - low < 1:
        return image
    scaled = (image.astype(np.float32) - low) * (255.0 / (high - low))
    return np.clip(scaled, 0, 255).astype(np.uint8)


def _for_matching(image: np.ndarray) -> np.ndarray:
    """A copy with local contrast raised strongly (CLAHE on lightness), used
    only to find the matching details."""
    global _CLAHE
    if _CLAHE is None:
        _CLAHE = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    lab[:, :, 0] = _CLAHE.apply(lab[:, :, 0])
    return cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)


def _as_is(image: np.ndarray) -> np.ndarray:
    return image


MATCHING_STAGES = (_as_is, _stretched, _for_matching)   # as is, evenly brightened, strongly
LONGEST_ZOOM = 2.5    # a turned-camera fit whose zoom (focal length) is over 2.5 photo widths
                      # means the photos are really flat things (or a strong telephoto)


def _components(stitcher) -> list[int]:
    return sorted(int(i) for i in np.asarray(stitcher.component()).ravel())


def _match(images: list[np.ndarray], kind: str, prepare,
           straighten: bool = True) -> tuple[int, object, list[int]]:
    """Finds how the photos fit, on copies made by `prepare`. Returns (status,
    stitcher, used photos). If some photos don't fit, the matching is done again
    with only the ones that do: OpenCV 5.0 can't make the picture (or the white
    coverage picture) while left-out photos are still in its list."""
    used = list(range(len(images)))
    try:
        for _ in range(3):
            # OpenCV's matching draws random samples; the same start every time
            # makes the same photos give the same panorama every time.
            cv2.setRNGSeed(1953)
            stitcher = cv2.Stitcher.create(_mode(kind))
            if kind == TURNING:
                stitcher.setWaveCorrection(straighten)
            status = stitcher.estimateTransform([prepare(images[i]) for i in used])
            if status != cv2.Stitcher_OK:
                return status, None, []
            fits = _components(stitcher)
            if len(fits) == len(used):
                return status, stitcher, used
            used = [used[i] for i in fits]
            if len(used) < 2:
                break
    except cv2.error as error:
        _check_opencv_error(error)        # e.g. "knn" from a photo with almost no details
    return cv2.Stitcher_ERR_NEED_MORE_IMGS, None, []


def _zoom(stitcher, images: list[np.ndarray], used: list[int]) -> float:
    """The turned-camera fit's focal length in photo widths (about 1 for a
    phone's main camera, 2 to 3 for its telephoto)."""
    focal = float(np.median([camera.focal for camera in stitcher.cameras()]))
    width = float(np.median([images[i].shape[1] for i in used])) * stitcher.workScale()
    return focal / width if width > 0 else 0.0


def _plan(images: list[np.ndarray], kind: str, straighten: bool = True) -> _Plan:
    """Matching for one kind, in stages: the photos as they are, then evenly
    brightened copies, then strongly brightened ones (dark photos have too
    little contrast to match as they are, but brightening can hurt bright
    photos). It stops when every photo fits; otherwise the stage that fits the
    most photos wins. The picture is made later, from the real photos."""
    plan = _Plan(kind, cv2.Stitcher_ERR_NEED_MORE_IMGS, straighten=straighten)
    for prepare in MATCHING_STAGES:
        status, stitcher, used = _match(images, kind, prepare, straighten)
        if stitcher is None:
            if plan.stitcher is None:
                plan.status = status
        elif len(used) > len(plan.used):
            plan.status, plan.stitcher, plan.used = status, stitcher, used
        if len(plan.used) == len(images):
            break
    if plan.stitcher is not None and kind == TURNING:
        try:
            plan.believable = _zoom(plan.stitcher, images, plan.used) <= LONGEST_ZOOM
        except cv2.error:
            plan.believable = True
    return plan


def _memory_cap() -> int:
    """The most pixels a panorama may have with this computer's free memory
    (joining needs about 40 bytes per pixel), between 50 and 400 megapixels."""
    class Status(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
    try:
        status = Status()
        status.dwLength = ctypes.sizeof(Status)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            return int(min(400e6, max(50e6, status.ullAvailPhys / 40)))
    except (AttributeError, OSError):
        pass
    return int(150e6)


def believable_size(width: float, height: float, photos: list[np.ndarray]) -> bool:
    """A joined picture can't be much wider or taller than the photos laid end
    to end, nor much bigger than all of them together; one that is was twisted
    by a wrong fit."""
    widths = sum(photo.shape[1] for photo in photos)
    heights = sum(photo.shape[0] for photo in photos)
    area = sum(photo.shape[0] * photo.shape[1] for photo in photos)
    return width <= 1.5 * widths and height <= 1.5 * heights and width * height <= 2.0 * area


TRIAL_MEGAPIXELS = 0.05   # the trial picture is made from photos shrunk to this size


def _compose(plan: _Plan, images: list[np.ndarray],
             max_pixels: int) -> tuple[np.ndarray, float] | None:
    """Makes the picture for a plan from the real photos, or None if it can't.
    A tiny trial picture comes first: it shows the full size, so a twisted fit is
    refused and a picture too big for the computer's memory is made smaller."""
    photos = [images[i] for i in plan.used]
    stitcher = plan.stitcher
    first = photos[0].shape[0] * photos[0].shape[1]
    try:
        stitcher.setCompositingResol(TRIAL_MEGAPIXELS)
        status, trial = stitcher.composePanorama(photos)
        if status != cv2.Stitcher_OK or trial is None or trial.size == 0:
            return None
        shrink = min(1.0, (TRIAL_MEGAPIXELS * 1e6 / first) ** 0.5)
        width, height = trial.shape[1] / shrink, trial.shape[0] / shrink
        if not believable_size(width, height, photos):
            return None
        scale = 1.0
        if width * height > max_pixels:
            scale = (max_pixels / (width * height)) ** 0.5
            stitcher.setCompositingResol(first * scale * scale / 1e6)
        else:
            stitcher.setCompositingResol(-1.0)        # OpenCV's "full size"
        status, picture = stitcher.composePanorama(photos)
    except cv2.error as error:
        _check_opencv_error(error)
        return None
    if status != cv2.Stitcher_OK or picture is None or picture.size == 0:
        return None
    return picture, scale


def _best_kind(images: list[np.ndarray], kind: str, max_pixels: int) -> StitchResult:
    """Works out the fit for the chosen kind, and for the other kind when the
    chosen one leaves photos out or isn't believable. The fits are tried in
    this order: believable ones first, then the one with more photos, then the
    chosen kind. The first that makes a sensible picture wins."""
    other = FLAT if kind == TURNING else TURNING
    plans = [_plan(images, kind)]
    first = plans[0]
    if not (first.stitcher is not None and len(first.used) == len(images) and first.believable):
        try:
            plans.append(_plan(images, other))
        except _OutOfMemory:
            if first.stitcher is None:
                raise                         # nothing to fall back on
    found = [plan for plan in plans if plan.stitcher is not None]
    found.sort(key=lambda plan: (plan.believable, len(plan.used), plan.kind == kind), reverse=True)
    for plan in found:
        made = _compose(plan, images, max_pixels)
        if made is None and plan.kind == TURNING and plan.straighten:
            # The turned-camera kind straightens the horizon assuming the camera
            # turned sideways. Photos taken by tilting it up and down come out
            # twisted that way, so they get one more try without straightening.
            plan = _plan(images, TURNING, straighten=False)
            made = _compose(plan, images, max_pixels) if plan.stitcher is not None else None
        if made is not None:
            picture, scale = made
            return StitchResult(picture, plan.used, plan.kind, switched=plan.kind != kind,
                                scale=scale, stitcher=plan.stitcher)
    status = first.status if first.status != cv2.Stitcher_OK else cv2.Stitcher_ERR_NEED_MORE_IMGS
    if found:
        raise StitchError("The photos matched, but only in a twisted way that can't be right. "
                          "Try the other kind in Step 2, or leave out photos that don't belong.")
    raise StitchError(WHY_NOT_JOINED.get(status, "The photos couldn't be joined. Check that "
                                         "neighboring photos overlap by about a third."))


DARK = 8    # photos with pixels this dark need the true coverage for trimming


def _coverage(result: StitchResult, images: list[np.ndarray]) -> np.ndarray | None:
    """Where the photos really are in the joined picture (True), or None if
    it isn't needed or can't be worked out.

    Trimming normally takes the stitcher's pure black fill as the edge. Very
    dark parts of photos (a night sky) can also come out pure black, so for
    dark photos the stitcher's last step is run again with plain white
    stand-ins: white is picture, black is fill. It is made at the same size as
    the picture, because the stitcher keeps the size it was last set to."""
    if result.stitcher is None:
        return None
    used = [images[i] for i in result.used]
    if not any((photo.max(axis=2) <= DARK).any() for photo in used):
        return None
    try:
        status, cover = result.stitcher.composePanorama(
            [np.full_like(photo, 255) for photo in used])
        if status != cv2.Stitcher_OK or cover is None or cover.shape[:2] != result.image.shape[:2]:
            return None
        return cover.max(axis=2) > 127
    except (cv2.error, MemoryError):
        return None


def _even_sizes(images: list[np.ndarray]) -> list[np.ndarray]:
    """Photos much smaller than the biggest (another camera, or zoomed in) are
    enlarged to about its size. OpenCV sizes every photo for matching by the
    first one, so a small photo has too few details to match, or gets placed
    wrongly; even at half size it was put in the wrong spot."""
    biggest = max(image.shape[0] * image.shape[1] for image in images)
    evened = []
    for image in images:
        grow = (biggest / (image.shape[0] * image.shape[1])) ** 0.5
        if grow > 1.25:
            image = cv2.resize(image, None, fx=grow, fy=grow, interpolation=cv2.INTER_CUBIC)
        evened.append(image)
    return evened


def stitch(images: list[np.ndarray], kind: str, max_pixels: int | None = None) -> StitchResult:
    """Joins the photos into one picture (see _best_kind). If the computer runs
    out of memory anyway, it halves the photos and tries again (twice at most)."""
    if len(images) < 2:
        raise StitchError("Add at least 2 overlapping photos first.")
    max_pixels = max_pixels or _memory_cap()
    # OpenCV takes the first photo as its reference; a small photo first left
    # doubled shapes where the photos meet. So the biggest go first (photos of
    # the same size keep their order), and `order` maps back to the caller's.
    order = sorted(range(len(images)), key=lambda i: -images[i].shape[0] * images[i].shape[1])
    # OpenCV's stitcher fails with "Unknown C++ exception" on a cut-out of a
    # bigger picture that isn't stored in one block; give it its own copy.
    images = _even_sizes([np.ascontiguousarray(images[i]) for i in order])
    halved = 1.0
    for attempt in range(3):
        try:
            result = _best_kind(images, kind, max_pixels)
        except _OutOfMemory:
            if attempt == 2:
                break
            halved /= 2
            images = [cv2.resize(image, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
                      for image in images]
            continue
        result.scale *= halved
        result.coverage = _coverage(result, images)
        result.stitcher = None            # done with it; frees its memory
        result.used = sorted(order[i] for i in result.used)
        return result
    raise StitchError("This computer ran out of memory, even with the photos at a quarter "
                      "of their size. Try fewer photos, or close other programs and try again.")


# ------------------------------------------------------------------ trimming the edges

def _largest_rectangle(inside: np.ndarray) -> tuple[int, int, int, int]:
    """(x, y, w, h) of the largest rectangle that is all 1 in a 0/1 grid
    (the row-by-row histogram method)."""
    rows, cols = inside.shape
    heights = np.zeros(cols, dtype=np.int32)
    best, best_area = (0, 0, 0, 0), 0
    for y in range(rows):
        heights = np.where(inside[y] > 0, heights + 1, 0)
        column_heights = heights.tolist() + [0]
        stack: list[int] = []
        for x, height in enumerate(column_heights):
            while stack and column_heights[stack[-1]] >= height:
                top_height = column_heights[stack.pop()]
                left = stack[-1] + 1 if stack else 0
                area = top_height * (x - left)
                if area > best_area:
                    best_area = area
                    best = (left, y - top_height + 1, x - left, top_height)
            stack.append(x)
    return best


def trim_rectangle(image: np.ndarray,
                   coverage: np.ndarray | None = None) -> tuple[int, int, int, int]:
    """(x, y, width, height) of the largest rectangle with picture everywhere
    inside it. `coverage` (True where the photos are) comes from stitch() for
    dark photos; without it, pure black counts as the stitcher's fill. Only
    fill that reaches the outside counts as edge: black inside the picture is kept."""
    height, width = image.shape[:2]
    if coverage is not None:
        covered = coverage
    else:
        covered = image.max(axis=2) > 0 if image.ndim == 3 else image > 0
    if covered.all():
        return 0, 0, width, height
    # Work on blocks of `step` x `step` pixels: a block counts as picture only
    # if every pixel in it is picture. Then the rectangle found is safe to cut
    # at full size too.
    step = max(1, -(-max(height, width) // TRIM_WORK_SIZE))
    tall, wide = -(-height // step), -(-width // step)
    # Partial blocks at the right and bottom repeat the last real pixels.
    padded = np.pad(covered, ((0, tall * step - height), (0, wide * step - width)), mode="edge")
    blocks = padded.reshape(tall, step, wide, step).all(axis=(1, 3))
    # Black reachable from the border is edge; fill it from a one-block frame.
    edge = np.pad((~blocks).astype(np.uint8), 1, constant_values=1)
    mask = np.zeros((edge.shape[0] + 2, edge.shape[1] + 2), dtype=np.uint8)
    cv2.floodFill(edge, mask, (0, 0), 2)
    inside = (edge[1:-1, 1:-1] != 2).astype(np.uint8)
    x, y, w, h = _largest_rectangle(inside)
    if w == 0 or h == 0:
        return 0, 0, width, height
    left, top = x * step, y * step
    right, bottom = min(width, (x + w) * step), min(height, (y + h) * step)
    return left, top, right - left, bottom - top


def crop(image: np.ndarray, rectangle: tuple[int, int, int, int]) -> np.ndarray:
    x, y, w, h = rectangle
    return image[y:y + h, x:x + w]


# ------------------------------------------------------------------ showing and saving

def fit_for_preview(image: np.ndarray, width: int, height: int) -> np.ndarray:
    """The picture shrunk to fit inside width x height (never enlarged)."""
    tall, wide = image.shape[:2]
    scale = min(width / wide, height / tall, 1.0)
    if scale >= 1.0:
        return image
    size = (max(1, round(wide * scale)), max(1, round(tall * scale)))
    return cv2.resize(image, size, interpolation=cv2.INTER_AREA)


def preview_png(image: np.ndarray, width: int, height: int) -> bytes:
    """The picture shrunk to fit width x height, as PNG bytes for the window."""
    ok, data = cv2.imencode(".png", fit_for_preview(image, width, height),
                            [cv2.IMWRITE_PNG_COMPRESSION, 1])
    if not ok:
        raise SaveError("The preview couldn't be made.")
    return data.tobytes()


def save_image(image: np.ndarray, path: str) -> None:
    """Saves as JPG (quality 95), PNG or TIFF, picked by the file's ending.
    Writes to a temporary file first, so a half-written file is never left."""
    ending = os.path.splitext(path)[1].lower()
    kind = SAVE_TYPES.get(ending)
    if kind is None:
        raise SaveError("A panorama can be saved as .jpg, .png or .tif. "
                        "Pick one of those in the Save box.")
    height, width = image.shape[:2]
    if kind == "JPG" and max(height, width) > JPEG_MAX:
        raise SaveError(f"This panorama is {width:,} x {height:,} pixels, which is too big "
                        f"for JPG ({JPEG_MAX:,} pixels on a side at most). "
                        "Save it as PNG or TIFF instead.")
    settings = {
        "JPG": [cv2.IMWRITE_JPEG_QUALITY, 95],
        "PNG": [cv2.IMWRITE_PNG_COMPRESSION, 3],
        "TIFF": [cv2.IMWRITE_TIFF_COMPRESSION, getattr(cv2, "IMWRITE_TIFF_COMPRESSION_LZW", 5)],
    }[kind]
    try:
        ok, data = cv2.imencode(ending, np.ascontiguousarray(image), settings)
    except (cv2.error, MemoryError) as error:
        raise SaveError("There wasn't enough memory to save it. Close other programs "
                        "and try again, or save as JPG.") from error
    if not ok:
        raise SaveError(f"The panorama couldn't be made into a {kind} file.")
    temporary = path + ".saving"
    try:
        data.tofile(temporary)
        os.replace(temporary, path)
    except OSError as error:
        try:
            os.remove(temporary)
        except OSError:
            pass
        raise SaveError(f"It couldn't be saved there ({error.strerror or error}). "
                        "Try another folder.") from error


def _registry_pictures() -> str | None:
    """The Pictures folder as set in the registry. Windows writes it here when
    the folder is moved (also when OneDrive takes it over or gives it back)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows"
                            r"\CurrentVersion\Explorer\User Shell Folders") as key:
            value, _ = winreg.QueryValueEx(key, "My Pictures")
        return os.path.expandvars(value) if value else None
    except (ImportError, OSError):
        return None


def _in_onedrive(path: str) -> bool:
    return any(part.lower().startswith("onedrive") for part in os.path.normpath(path).split(os.sep))


def _onedrive_installed() -> bool:
    places = (r"%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe",
              r"%ProgramFiles%\Microsoft OneDrive\OneDrive.exe",
              r"%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe")
    return any(os.path.isfile(os.path.expandvars(place)) for place in places)


def pictures_folder() -> str:
    """Windows' own Pictures folder (wherever it has been moved to).

    If that is a OneDrive folder but OneDrive isn't installed, the plain
    Pictures folder is used. On a PC seen 2026-11-20, OneDrive had been removed
    and had given Pictures back, but Python started through py/pyw still read
    the old OneDrive\\Pictures setting, while reg.exe and PowerShell read the
    plain one."""
    plain = os.path.join(os.path.expanduser("~"), "Pictures")
    for candidate in (_registry_pictures(), _known_folder_pictures(), plain):
        if not candidate or not os.path.isdir(candidate):
            continue
        if _in_onedrive(candidate) and os.path.isdir(plain) and not _onedrive_installed():
            return plain
        return candidate
    return os.path.expanduser("~")


def _known_folder_pictures() -> str | None:
    """The Pictures folder as Windows' shell reports it."""
    try:
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                        ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

        pictures = GUID(0x33E28130, 0x4E1E, 0x4676,
                        (ctypes.c_ubyte * 8)(0x83, 0x5A, 0x98, 0x39, 0x5C, 0x3B, 0xC3, 0xBB))
        found = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(pictures), 0, None,
                                                      ctypes.byref(found)) == 0:
            path = found.value
            ctypes.windll.ole32.CoTaskMemFree(found)
            return path or None
    except (AttributeError, OSError, ValueError):
        pass
    return None


def free_name(path: str) -> str:
    """`path`, or "name (2).png", "name (3).png" ... if that file already exists,
    so a file the person didn't pick in the Save box is never replaced."""
    base, ending = os.path.splitext(path)
    number = 2
    while os.path.exists(path):
        path = f"{base} ({number}){ending}"
        number += 1
    return path


def default_save_folder() -> str:
    return os.path.join(pictures_folder(), "Panoramas")


def default_file_name(folder: str, ending: str = ".jpg",
                      when: datetime.datetime | None = None) -> str:
    """'Panorama 2026-11-20 14-05.jpg', with (2), (3) ... if that name is taken."""
    when = when or datetime.datetime.now()
    base = f"Panorama {when:%Y-%m-%d %H-%M}"
    name, number = f"{base}{ending}", 2
    while os.path.exists(os.path.join(folder, name)):
        name, number = f"{base} ({number}){ending}", number + 1
    return name
