"""Dating Photographs: prints a date on a photo in glowing orange digits, the way
a film camera did ('98 10 12 in the bottom-right corner).

The digits are drawn as seven-segment figures, like a camera's, sized to the
photo (about 4% of its shorter side) and set inside any white border. Only
what is known is printed: a date found as "Oct-98" prints as '98 10, not as
a day nobody wrote.
"""

from __future__ import annotations

import datetime as dt
import io

from PIL import Image, ImageDraw, ImageFilter, ImageOps, PngImagePlugin

import dating_photos_dates as dates
import dating_photos_files as files
from photo_text_reader import eight_bit

ORANGE = (255, 138, 28)
DIGIT_HEIGHT = 0.04        # of the photo's shorter side
MARGIN = 0.05              # from the photo's edges, of its shorter side
SMALLEST = 14              # pixels: never smaller than this
STYLES = {                 # name shown in the window -> how a full date looks
    "'98 10 12": "year first",
    "10 12 '98": "month first",
    "1998-10-12": "full year",
}
DEFAULT_STYLE = "'98 10 12"
CAN_PRINT = (".jpg", ".jpeg", ".jfif", ".png", ".tif", ".tiff", ".bmp", ".webp")
SEGMENTS = {"0": "abcdef", "1": "bc", "2": "abged", "3": "abgcd", "4": "fgbc", "5": "afgcd",
            "6": "afgedc", "7": "abc", "8": "abcdefg", "9": "abcfgd", "-": "g"}


def stamp_text(date, kind, style=DEFAULT_STYLE):
    """What gets printed: only the parts of the date that are known."""
    year = date.year
    short = f"'{year % 100:02}"
    if kind in (dates.YEAR, dates.DECADE):
        return str(year) if style == "1998-10-12" else short
    with_day = kind in (dates.DAY, dates.HOLIDAY)
    if style == "1998-10-12":
        return date.isoformat() if with_day else f"{year}-{date.month:02}"
    if style == "10 12 '98":
        return f"{date.month} {date.day} {short}" if with_day else f"{date.month} {short}"
    return f"{short} {date.month} {date.day}" if with_day else f"{short} {date.month}"


def content_box(picture, threshold=232):
    """The photo inside any white (or nearly white) border a scan has around it."""
    gray = eight_bit(picture).convert("L")
    small = gray.resize((max(1, gray.width // 8), max(1, gray.height // 8)))
    width, height = small.size
    pixels = small.load()

    def plain(values):
        values = list(values)
        return sum(v >= threshold for v in values) >= 0.97 * len(values)

    left, top, right, bottom = 0, 0, width, height
    while left < right - 1 and plain(pixels[left, y] for y in range(height)):
        left += 1
    while right - 1 > left and plain(pixels[right - 1, y] for y in range(height)):
        right -= 1
    while top < bottom - 1 and plain(pixels[x, top] for x in range(width)):
        top += 1
    while bottom - 1 > top and plain(pixels[x, bottom - 1] for x in range(width)):
        bottom -= 1
    box = (left * 8, top * 8, min(picture.width, right * 8), min(picture.height, bottom * 8))
    # a border only counts if a real photo is left inside it
    if (box[2] - box[0]) < 0.4 * picture.width or (box[3] - box[1]) < 0.4 * picture.height:
        return (0, 0, picture.width, picture.height)
    return box


def segment(draw, start, end, thickness, fill):
    """One bar of a seven-segment figure, pointed at both ends."""
    (x1, y1), (x2, y2) = start, end
    half = thickness / 2
    if y1 == y2:     # across
        points = [(x1, y1), (x1 + half, y1 - half), (x2 - half, y1 - half), (x2, y1),
                  (x2 - half, y1 + half), (x1 + half, y1 + half)]
    else:            # up and down
        points = [(x1, y1), (x1 + half, y1 + half), (x1 + half, y2 - half), (x1, y2),
                  (x1 - half, y2 - half), (x1 - half, y1 + half)]
    draw.polygon(points, fill=fill)


def render(text, height, color=ORANGE):
    """The digits on a see-through layer, slanted slightly like a camera's."""
    height = max(SMALLEST, int(height))
    width = height * 0.55
    thick = max(2.0, height / 8)
    gap = thick * 0.18
    advance = {" ": height * 0.45, "'": height * 0.3}
    total = sum(advance.get(ch, width + height * 0.3) for ch in text) + thick * 2
    pad = int(height * 0.4)
    layer = Image.new("RGBA", (int(total) + 2 * pad, height + 2 * pad), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    x = pad + thick
    top = pad
    for ch in text:
        if ch == " ":
            x += advance[" "]
            continue
        if ch == "'":
            draw.line((x + thick, top, x, top + height * 0.28), fill=color, width=int(thick))
            x += advance["'"]
            continue
        left, right = x + thick / 2, x + width - thick / 2
        upper, middle, lower = top + thick / 2, top + height / 2, top + height - thick / 2
        bars = {"a": ((left + gap, upper), (right - gap, upper)), "g": ((left + gap, middle), (right - gap, middle)),
                "d": ((left + gap, lower), (right - gap, lower)),
                "f": ((left, upper + gap), (left, middle - gap)), "b": ((right, upper + gap), (right, middle - gap)),
                "e": ((left, middle + gap), (left, lower - gap)), "c": ((right, middle + gap), (right, lower - gap))}
        for name in SEGMENTS.get(ch, ""):
            segment(draw, *bars[name], thick, color)
        x += width + height * 0.3
    slant = 0.12
    return layer.transform(layer.size, Image.AFFINE, (1, slant, -slant * layer.height / 2, 0, 1, 0),
                           resample=Image.BICUBIC)


def draw_stamp(picture, text):
    """A copy of the picture with the date printed in its bottom-right corner."""
    out = eight_bit(picture).convert("RGB") if picture.mode != "RGB" else picture.copy()
    left, top, right, bottom = content_box(out)
    short = min(right - left, bottom - top)
    digits = render(text, short * DIGIT_HEIGHT)
    margin = int(short * MARGIN)
    x = max(left, right - margin - digits.width)
    y = max(top, bottom - margin - digits.height)
    halo = Image.new("RGBA", digits.size, (60, 20, 0, 0))
    halo.putalpha(digits.getchannel("A").filter(ImageFilter.GaussianBlur(max(1, digits.height / 25)))
                  .point(lambda a: a * 45 // 100))
    glow = digits.filter(ImageFilter.GaussianBlur(max(1, digits.height / 18)))
    for layer in (halo, glow, digits):
        out.paste(layer, (x, y), layer)
    return out


def why_not(path):
    """Why the date can't be printed on this file, or '' if it can."""
    if not path.lower().endswith(CAN_PRINT):
        return "the date can't be printed on this kind of picture"
    try:
        with Image.open(path) as picture:
            if getattr(picture, "n_frames", 1) > 1:
                return "it has several pages, so the date wasn't printed"
    except Exception as error:
        return f"it couldn't be opened ({error})"
    return ""


# TIFF packings that work for a color picture; any other (a fax-style one for black
# and white scans) can't hold the orange, so the photo is saved packed the usual way
COLOR_TIFF = ("tiff_lzw", "tiff_adobe_deflate", "tiff_deflate", "packbits", "jpeg")
# A TIFF's own tags describe how its picture is laid out (size, bits, packing, strips,
# its color profile). They describe the old picture, not the stamped RGB one, and
# copied over they make the saved file unreadable, so only the rest is kept.
TIFF_LAYOUT_TAGS = {254, 255, 256, 257, 258, 259, 262, 266, 273, 277, 278, 279, 280, 281, 284, 292,
                    293, 317, 320, 322, 323, 324, 325, 330, 338, 339, 347, 530, 531, 532, 34675}


def print_date(path, text):
    """Print the date onto the picture file itself, keeping its type, resolution,
    color profile and photo information (EXIF, and XMP where the type has it).
    Works on a read-only photo too. Returns '' or why it couldn't."""
    problem = why_not(path)
    if problem:
        return problem
    temp = path + files.TEMP_ENDING
    with open(path, "rb") as f:
        # opened from memory: Pillow maps an uncompressed gray TIFF straight from the
        # file, and Windows won't let a mapped file be replaced
        data = io.BytesIO(f.read())
    try:
        with Image.open(data) as source:
            source.load()
            save_stamped(source, draw_date_on(source, text), temp)
        files.replace_with(temp, path)
    finally:
        files.discard(temp)
    return ""


def draw_date_on(source, text):
    """(the picture turned upright with the date printed on it, whether it was turned, its EXIF)."""
    exif = source.getexif()
    turned = exif.get(0x0112, 1) not in (None, 1)
    picture = ImageOps.exif_transpose(source) if turned else source
    return draw_stamp(picture, text), turned, picture.getexif() if turned else exif


def save_stamped(source, drawn, temp):
    """Save the stamped picture to temp in the source's own type and settings."""
    stamped, turned, exif = drawn
    kind, info = source.format, source.info
    options = {}
    if info.get("dpi"):
        options["dpi"] = info["dpi"]
    if info.get("icc_profile") and source.mode == "RGB":
        options["icc_profile"] = info["icc_profile"]
    if len(exif):
        options["exif"] = exif.tobytes()
    if info.get("xmp") and kind in ("JPEG", "WEBP"):
        options["xmp"] = info["xmp"]
    if kind == "JPEG":
        if source.mode == "RGB" and not turned and getattr(source, "quantization", None):
            # re-use the photo's own quality settings, so the rest of it looks the same
            source.paste(stamped)
            source.save(temp, "JPEG", quality="keep", subsampling="keep", **options)
        else:
            stamped.save(temp, "JPEG", quality=95, **options)
    elif kind == "TIFF":
        if len(exif):
            for tag in TIFF_LAYOUT_TAGS & set(exif):
                del exif[tag]
            options["exif"] = exif.tobytes()
        compression = info.get("compression")
        if compression in (None, "raw"):
            compression = None
        elif compression not in COLOR_TIFF:
            compression = "tiff_lzw"
        if compression == "jpeg":
            options["quality"] = 95
        stamped.save(temp, "TIFF", compression=compression, **options)
    elif kind == "PNG":
        words = PngImagePlugin.PngInfo()          # its text, XMP included
        for key, value in getattr(source, "text", {}).items():
            words.add_itxt(key, value)
        stamped.save(temp, "PNG", pnginfo=words, **options)
    elif kind == "WEBP":
        stamped.save(temp, "WEBP", quality=95, **options)
    else:
        options.pop("exif", None)
        stamped.save(temp, kind, **options)


def today_text(style=DEFAULT_STYLE):
    """How today's date would look, for the window's example."""
    return stamp_text(dt.date.today(), dates.DAY, style)
