"""Dating Photographs: the photo files themselves.

Finds the pictures in a folder, pairs each front with its back (IMG_0001.jpg
and IMG_0001_b.jpg), works out new names (1998-10-01.jpg, 1998-10-01 (2).jpg,
1998-10-01_b.jpg), moves or copies them into "Dated Photographs", keeps a list
of every one in an Excel-readable file there, and can undo the last batch
(a moved photo goes back; a copy is taken away, as its original never changed).

It can also set a photo's "Date taken" (the EXIF DateTimeOriginal that Windows
and photo programs sort by). In a JPG or TIFF that already has one, only those
19 characters change; a JPG without one gets it added. The picture itself is
never re-saved. Every file also gets its created and modified dates set.
"""

from __future__ import annotations

import csv
import ctypes
import datetime as dt
import functools
import os
import re
import shutil
import stat
import struct
import time
from ctypes import wintypes
from dataclasses import dataclass, field

DEST = "Dated Photographs"
LOG_NAME = "Dating Photographs log.csv"
ORIGINALS = "Originals (no printed date)"   # untouched copies of photos the date was printed on
PICTURE_ENDINGS = (".jpg", ".jpeg", ".jfif", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp",
                   ".heic", ".heif")
# iPhone photos: listed, so they aren't missed, but Pillow can't open them without
# pillow-heif, which is left out of the Setup because of its license
CANT_OPEN = (".heic", ".heif")
TEMP_ENDING = ".dating-tmp"
LOG_BUSY = (f"Close “{LOG_NAME}” first: it is open in another program, like Excel (it is in "
            f"the “{DEST}” folder). Then try again.")
SAVE_EVERY = 20           # files moved between saves of the list, so a crash loses little
# The back of a photo, scanned as its own file: IMG_0001_b, IMG_0001 back, IMG_0001 (back)
BACK = re.compile(r"(?i)^(?P<base>.+?)(?P<suffix>(?:[ _-]+|\s*\()(?:back|reverse|verso|rev|b)\)?)$")
# A scanner's touched-up copy of the front (Epson FastFoto saves IMG_0001_a)
COPY = re.compile(r"(?i)^(?P<base>.+?)(?P<suffix>[ _-]a)$")
COPIED, MOVED = "copied", "moved"
# "Copied or moved" came in 1.1: an older list has none, and its rows were all moved
LOG_COLUMNS = ["Batch", "When", "Copied or moved", "Photo before", "Photo now", "Date", "Found from",
               "Found in", "Date printed", "Original kept", "Old date taken", "Old file dates", "Undone"]
NOON = dt.time(12, 0)     # midday, so no time-zone shift can move the date


def natural_key(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def is_picture_name(name):
    # "._IMG_1.jpg" is a Mac's hidden note about IMG_1.jpg, not a picture
    return name.lower().endswith(PICTURE_ENDINGS) and not name.startswith(("~$", "._"))


def list_pictures(folder, subfolders=False):
    """Every picture in a folder (and its sub-folders, if asked), skipping
    "Dated Photographs", in the order Explorer shows them."""
    found = []
    if subfolders:
        for here, dirs, files in os.walk(folder):
            dirs[:] = sorted((d for d in dirs if d.lower() != DEST.lower() and not d.startswith(".")),
                             key=natural_key)
            found += [os.path.join(here, f) for f in files if is_picture_name(f)]
    else:
        found = [os.path.join(folder, f) for f in os.listdir(folder)
                 if is_picture_name(f) and os.path.isfile(os.path.join(folder, f))]
    return sorted(found, key=lambda p: natural_key(os.path.relpath(p, folder)))


@dataclass
class Photo:
    """One photograph: its front, plus any other files that belong with it."""
    front: str
    back: str | None = None
    extras: list = field(default_factory=list)     # [(path, suffix kept in the new name)]

    def files(self):
        """[(path, suffix)] for every file, the front first with no suffix."""
        return [(self.front, "")] + self.extras

    def name(self):
        return os.path.basename(self.front)


def group_photos(paths, every_second_is_back=False):
    """Pair fronts with their backs. A back named after its front (IMG_1_b.jpg,
    IMG_1 back.jpg) always goes with it; with every_second_is_back, the rest are
    taken in pairs, the second of each pair being the back of the first."""
    by_folder = {}
    for path in paths:
        by_folder.setdefault(os.path.dirname(path), []).append(path)
    photos = []
    for folder_paths in by_folder.values():
        stems = {}
        for path in folder_paths:
            stems.setdefault(os.path.splitext(os.path.basename(path))[0].lower(), []).append(path)
        belongs = {}       # path of a back or copy -> (front path, suffix, is it the back)
        for path in folder_paths:
            stem = os.path.splitext(os.path.basename(path))[0]
            for pattern, is_back in ((BACK, True), (COPY, False)):
                match = pattern.match(stem)
                fronts = stems.get(match.group("base").lower(), []) if match else []
                fronts = [f for f in fronts if f != path]
                if fronts:
                    belongs[path] = (fronts[0], match.group("suffix"), is_back)
                    break
        made = {}
        loose = []
        for path in folder_paths:
            if path in belongs:
                continue
            made[path] = Photo(path)
        for path, (front, suffix, is_back) in belongs.items():
            photo = made.get(front)
            if photo is None:            # its "front" is itself a back: keep it on its own
                made[path] = Photo(path)
                continue
            photo.extras.append((path, suffix))
            if is_back and photo.back is None:
                photo.back = path
        ordered = [made[p] for p in folder_paths if p in made]
        if every_second_is_back:
            paired = []
            for photo in ordered:
                if photo.back is None and not photo.extras:
                    loose.append(photo)
                else:
                    paired.append(photo)
            for front, back in zip(loose[::2], loose[1::2]):
                front.back = back.front
                front.extras.append((back.front, " back"))
            if len(loose) % 2:
                paired.append(loose[-1])
            keep = set(id(p) for p in paired) | {id(p) for p in loose[::2]}
            ordered = [p for p in ordered if id(p) in keep]
        photos += ordered
    order = {path: i for i, path in enumerate(paths)}
    return sorted(photos, key=lambda p: order.get(p.front, 0))


# ------------------------------------------------------------------ new names

@dataclass
class Move:
    photo: Photo
    date: dt.date
    pairs: list               # [(from path, to path)]


def new_names(dated, root, keep_old_name=False, year_folders=False, taken=None, rename=True):
    """Where each (photo, date) goes: [Move]. Names never clash with each other or
    with files already there: 1998-10-01.jpg, then 1998-10-01 (2).jpg, ...
    With rename=False the photos keep their names (numbered only if they'd clash)."""
    taken = {p.lower() for p in (taken or ())}
    moves = []
    for photo, date in dated:
        relative = os.path.relpath(os.path.dirname(photo.front), root)
        folder = os.path.join(root, DEST, "" if relative == "." else relative,
                              str(date.year) if year_folders else "")
        folder = os.path.normpath(folder)
        old = os.path.splitext(photo.name())[0]
        if not rename:
            base = old
        else:
            base = date.isoformat() + (" " + old if keep_old_name else "")
        number = 1
        while True:
            stem = base if number == 1 else f"{base} ({number})"
            pairs = [(path, os.path.join(folder, stem + suffix + os.path.splitext(path)[1]))
                     for path, suffix in photo.files()]
            if not any(to.lower() in taken or os.path.exists(to) for _, to in pairs):
                break
            number += 1
        taken.update(to.lower() for _, to in pairs)
        moves.append(Move(photo, date, pairs))
    return moves


# ------------------------------------------------------------------ "Date taken" inside the file

DATE_TAKEN_TAG = 0x9003        # EXIF DateTimeOriginal
EXIF_POINTER = 0x8769


def _tiff_tag(tiff, tag):
    """Where the 19 date characters of `tag` are in TIFF-structured bytes, or None."""
    if tiff[:2] not in (b"II", b"MM") or len(tiff) < 8:
        return None
    order = "<" if tiff[:2] == b"II" else ">"

    def entries(offset):
        if offset + 2 > len(tiff):
            return
        count = struct.unpack(order + "H", tiff[offset:offset + 2])[0]
        for i in range(count):
            at = offset + 2 + 12 * i
            if at + 12 > len(tiff):
                return
            number, kind, length = struct.unpack(order + "HHI", tiff[at:at + 8])
            yield number, kind, length, at + 8

    first = struct.unpack(order + "I", tiff[4:8])[0]
    exif = next((struct.unpack(order + "I", tiff[at:at + 4])[0]
                 for number, _, _, at in entries(first) if number == EXIF_POINTER), None)
    if exif is None:
        return None
    for number, kind, length, at in entries(exif):
        if number == tag and kind == 2 and length >= 20:
            place = struct.unpack(order + "I", tiff[at:at + 4])[0]
            if place + 19 <= len(tiff):
                return place
    return None


def _jpeg_exif(data):
    """(start, end) of the JPG's EXIF block contents (after its 4-byte header), or None,
    and where a new one would go."""
    if data[:2] != b"\xff\xd8":
        raise ValueError("not a JPG")
    at, insert = 2, 2
    while at + 4 <= len(data) and data[at] == 0xFF:
        marker = data[at + 1]
        if marker in (0xD9, 0xDA):       # end, or the picture data begins
            break
        length = struct.unpack(">H", data[at + 2:at + 4])[0]
        if marker == 0xE1 and data[at + 4:at + 10] == b"Exif\x00\x00":
            return (at + 4, at + 2 + length), insert
        if marker == 0xE0:               # a new block goes after the JFIF one
            insert = at + 2 + length
        at += 2 + length
    return None, insert


def exif_text(date):
    return None if date is None else dt.datetime.combine(date, NOON).strftime("%Y:%m:%d %H:%M:%S")


def read_date_taken(path):
    """The photo's "Date taken" as stored ("1998:10:01 12:00:00"), or None."""
    try:
        with open(path, "rb") as f:
            data = f.read()
        if path.lower().endswith((".tif", ".tiff")):
            place, tiff = _tiff_tag(data, DATE_TAKEN_TAG), data
        else:
            block, _ = _jpeg_exif(data)
            if block is None:
                return None
            tiff = data[block[0] + 6:block[1]]
            place = _tiff_tag(tiff, DATE_TAKEN_TAG)
        return tiff[place:place + 19].decode("ascii", "replace") if place is not None else None
    except (OSError, ValueError, struct.error):
        return None


def can_set_date_taken(path):
    return path.lower().endswith((".jpg", ".jpeg", ".jfif", ".tif", ".tiff"))


def write_date_taken(path, text):
    """Put a "Date taken" (as "YYYY:MM:DD HH:MM:SS") into a JPG or TIFF, changing
    nothing else. Returns True if it was written. A TIFF without one is left alone."""
    with open(path, "rb") as f:
        data = bytearray(f.read())
    if path.lower().endswith((".tif", ".tiff")):
        place = _tiff_tag(data, DATE_TAKEN_TAG)
        if place is None:
            return False
        data[place:place + 19] = text.encode("ascii")
    else:
        block, insert = _jpeg_exif(bytes(data))
        if block is not None:
            tiff_start = block[0] + 6
            place = _tiff_tag(bytes(data[tiff_start:block[1]]), DATE_TAKEN_TAG)
            if place is not None:                  # change just those characters
                data[tiff_start + place:tiff_start + place + 19] = text.encode("ascii")
            else:                                  # add it, keeping the rest of the block
                from PIL import Image
                exif = Image.Exif()
                exif.load(bytes(data[block[0]:block[1]]))
                exif_ifd = exif.get_ifd(EXIF_POINTER)
                exif_ifd[DATE_TAKEN_TAG] = text
                exif[EXIF_POINTER] = exif_ifd
                data[block[0] - 4:block[1]] = _segment(exif.tobytes())
        else:
            from PIL import Image
            exif = Image.Exif()
            exif[EXIF_POINTER] = {DATE_TAKEN_TAG: text}
            data[insert:insert] = _segment(exif.tobytes())
    write_file(path, data)
    return True


def _segment(payload):
    if len(payload) + 2 > 0xFFFF:
        raise ValueError("the photo's information block is too big to change")
    return b"\xff\xe1" + struct.pack(">H", len(payload) + 2) + payload


# ------------------------------------------------------------------ changing a file safely

def discard(path):
    """Delete a leftover temporary file, if it is there."""
    try:
        os.remove(path)
    except OSError:
        pass


def replace_with(temp, path):
    """os.replace(temp, path), which also works on a read-only photo (as copied from
    a CD); the photo stays read-only afterwards."""
    read_only = not os.stat(path).st_mode & stat.S_IWRITE
    if read_only:
        os.chmod(path, stat.S_IREAD | stat.S_IWRITE)
    try:
        os.replace(temp, path)
    finally:
        if read_only:
            os.chmod(path, stat.S_IREAD)


def write_file(path, data):
    """Replace a file's contents in one step (a half-written photo is never left),
    read-only or not. The temporary file never stays behind."""
    temp = path + TEMP_ENDING
    try:
        with open(temp, "wb") as f:
            f.write(data)
        replace_with(temp, path)
    finally:
        discard(temp)


def remove_file(path):
    """Delete a file, read-only or not."""
    try:
        os.remove(path)
    except PermissionError:
        os.chmod(path, stat.S_IREAD | stat.S_IWRITE)
        os.remove(path)


# ------------------------------------------------------------------ file dates

class FILETIME(ctypes.Structure):
    _fields_ = [("low", wintypes.DWORD), ("high", wintypes.DWORD)]


def _filetime(seconds):
    ticks = int(round((seconds + 11644473600) * 10_000_000))
    return FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)


def file_dates(path):
    """(modified, created) as seconds since 1970 (negative before it)."""
    info = os.stat(path)
    return info.st_mtime, getattr(info, "st_birthtime", info.st_ctime)


@functools.cache
def _kernel():
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                                   wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    return kernel


def _open_handle(path, access):
    """A Windows handle to an existing file that shares with everyone, or None."""
    handle = _kernel().CreateFileW(path, access, 0x7, None, 3, 0x80, None)   # share all, open existing
    return None if handle in (None, wintypes.HANDLE(-1).value) else handle


def set_file_dates(path, modified, created):
    """Set a file's modified and created dates; works for years before 1970,
    which os.utime refuses on Windows."""
    kernel = _kernel()
    handle = _open_handle(path, 0x100)        # FILE_WRITE_ATTRIBUTES
    if handle is None:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        made, changed = _filetime(created), _filetime(modified)
        if not kernel.SetFileTime(handle, ctypes.byref(made), None, ctypes.byref(changed)):
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        kernel.CloseHandle(handle)


EPOCH = dt.datetime(1970, 1, 1)


def local_offset():
    """How far this PC's clock is from UTC now. Windows can't convert times before
    1970 to local time (Errno 22), so old dates use today's offset; at noon, an
    hour of summer time either way can't change the day."""
    return dt.datetime.now().astimezone().utcoffset() or dt.timedelta(0)


def seconds_at_noon(date):
    """Noon on a date, this PC's time, as seconds since 1970 (negative before it)."""
    return (dt.datetime.combine(date, NOON) - local_offset() - EPOCH).total_seconds()


def day_of(seconds):
    """The local date of a file time, before 1970 too."""
    return (EPOCH + dt.timedelta(seconds=seconds) + local_offset()).date()


# ------------------------------------------------------------------ moving, the list, undo

@dataclass
class Done:
    moved: int = 0                       # photos (fronts) moved or copied
    files: int = 0                       # files moved or copied, backs included
    problems: list = field(default_factory=list)   # (photo name, why)
    batch: int = 0


def log_path(root):
    return os.path.join(root, DEST, LOG_NAME)


def log_in_use(root):
    """True if the list of moves can't be changed now, most often because it is open
    in Excel: a move or Undo then couldn't record what it did."""
    path = log_path(root)
    if not os.path.exists(path):
        return False
    handle = _open_handle(path, 0x40000000 | 0x10000)       # GENERIC_WRITE and DELETE, as a save needs
    if handle is None:
        return True
    _kernel().CloseHandle(handle)
    return False


def batch_of(row):
    try:
        return int(row.get("Batch") or 0)
    except ValueError:                   # changed by hand in Excel
        return 0


def read_log(root):
    # Excel may have saved it the old Windows way; as a last resort odd letters are
    # read as "?", never the list dropped (a move would then write over it)
    for encoding, errors in (("utf-8-sig", "strict"), ("cp1252", "strict"), ("cp1252", "replace")):
        try:
            with open(log_path(root), encoding=encoding, errors=errors, newline="") as f:
                return [dict(row) for row in csv.DictReader(f)]
        except FileNotFoundError:
            return []
        except UnicodeDecodeError:
            continue
    return []


def write_log(root, rows):
    path = log_path(root)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temp = path + ".tmp"
    try:
        with open(temp, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, LOG_COLUMNS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temp, path)
    finally:
        discard(temp)


def save_log(root, rows, tries=4):
    """write_log, trying again for a moment if something holds the list open.
    Returns '' or why it couldn't be saved."""
    for attempt in range(tries):
        try:
            write_log(root, rows)
            return ""
        except OSError as error:
            if attempt == tries - 1:
                return (f"the list of moves couldn't be saved ({why(error)}), so Undo won't know about "
                        f"these. {LOG_BUSY}")
            time.sleep(0.5)
    return ""


SPARE_ROOM = 200 * 1024 * 1024      # left free after copying, so Windows isn't left with a full disk


def no_room_to_copy(root, moves):
    """'' if the disk has room for copies of these photos, or else why not, in words."""
    try:
        needed = sum(os.path.getsize(source) for move in moves for source, _ in move.pairs)
        free = shutil.disk_usage(root).free
    except OSError:
        return ""                       # can't tell: let the copying itself say
    needed = int(needed * 1.1)          # a printed date can make a photo a little bigger
    if needed + SPARE_ROOM <= free:
        return ""
    return (f"There isn't room on this disk to copy them: they need about {megabytes(needed)} and "
            f"only {megabytes(free)} is free. Free up some room, copy fewer at a time, or move them instead.")


def megabytes(size):
    return f"{size / 1024 ** 3:.1f} GB" if size >= 1024 ** 3 else f"{max(1, round(size / 1024 ** 2))} MB"


def apply_moves(root, moves, clues, set_date_taken=True, now=None, printed=None, copy=False):
    """Move every photo (and its back) to its new name, or with copy=True make a
    copy there and leave the photo untouched where it is; then set the dates on
    it (or the copy) if asked, and add each one to the list in "Dated Photographs".
    clues maps a photo's front path to the Clue its date came from (or None if
    it was typed). printed maps a front path to the date to print on it
    ('98 10 12); when moving, an untouched copy of each such photo is kept in
    "Originals (no printed date)" (a copy's original is untouched already).
    The list is saved as it goes and at the end, whatever happens on the way,
    so Undo always knows what was done. If the list is open in Excel, nothing is."""
    printed = printed or {}
    now = now or dt.datetime.now()
    done = Done()
    if log_in_use(root):
        done.problems.append((LOG_NAME, LOG_BUSY))
        return done
    rows = read_log(root)
    done.batch = max((batch_of(r) for r in rows), default=0) + 1
    saved = len(rows)
    try:
        for move in moves:
            finished = (copy_group if copy else move_group)(move, done)
            if not finished:
                continue
            for source, target in finished:
                rows.append(finish_move(root, move, source, target, clues.get(move.photo.front), printed,
                                        set_date_taken, now, done, copy))
                done.files += 1
            done.moved += 1
            if len(rows) - saved >= SAVE_EVERY:
                problem = save_log(root, rows)
                if problem:
                    done.problems.append((LOG_NAME, problem))
                    break               # stop: further moves couldn't be undone
                saved = len(rows)
    finally:
        if len(rows) > saved:
            problem = save_log(root, rows)       # once more: Excel may have been closed by now
            if problem and (LOG_NAME, problem) not in done.problems:
                done.problems.append((LOG_NAME, problem))
    return done


def move_group(move, done):
    """Move a photo's files to their new names, all or none: [(from, to)], or []
    (with the reason added to done.problems) if any couldn't be moved."""
    finished = []
    try:
        for source, target in move.pairs:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            os.rename(source, target)          # never replaces a file that's there
            finished.append((source, target))
    except Exception as error:
        for source, target in reversed(finished):      # put the group back as it was
            try:
                os.rename(target, source)
            except OSError:
                pass
        done.problems.append((move.photo.name(), why(error)))
        return []
    return finished


def copy_file(source, target):
    """Copy a file to a new name, its dates and all. It never replaces a file that's
    there, and a half-made copy never stays behind (it is made under another name
    first, and only named when whole)."""
    temp = target + TEMP_ENDING
    if os.path.exists(temp):             # left by a copy that was cut off
        remove_file(temp)
    try:
        shutil.copy2(source, temp)
        os.rename(temp, target)          # never replaces a file that's there
    except BaseException:
        try:
            if os.path.exists(temp):
                remove_file(temp)
        except OSError:
            pass                         # the real reason is the one worth reporting
        raise


def copy_group(move, done):
    """Copy a photo's files to their new names, all or none: [(from, to)], or []
    (with the reason added to done.problems) if any couldn't be copied. The
    photo itself is never changed."""
    finished = []
    try:
        for source, target in move.pairs:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            copy_file(source, target)
            finished.append((source, target))
    except Exception as error:
        for _, target in reversed(finished):            # take away the copies already made
            try:
                remove_file(target)
            except OSError:
                pass
        done.problems.append((move.photo.name(), why(error)))
        return []
    return finished


def finish_move(root, move, source, target, clue, printed, set_date_taken, now, done, copy=False):
    """Print the date and set the dates on one moved (or copied) file, as asked.
    Returns its row for the list; a problem on the way is noted in done.problems,
    but the file is still listed, so Undo can put it back (or take the copy away)."""
    row = {"Batch": done.batch, "When": now.strftime("%Y-%m-%d %H:%M"), "Copied or moved": COPIED if copy else MOVED,
           "Photo before": os.path.relpath(source, root), "Photo now": os.path.relpath(target, root),
           "Date": move.date.isoformat(), "Found from": clue.words if clue else "typed in",
           "Found in": clue.where if clue else "", "Date printed": "", "Original kept": "",
           "Old date taken": "-", "Old file dates": "", "Undone": ""}
    name = os.path.basename(target)
    done_word = row["Copied or moved"]
    try:
        row["Old file dates"] = "{:.3f};{:.3f}".format(*file_dates(target))
        if source == move.photo.front and printed.get(source):
            kept, problem = print_on(target, printed[source], keep_original=not copy)
            if problem:
                done.problems.append((name, problem))
            else:
                row["Date printed"] = printed[source]
                row["Original kept"] = os.path.relpath(kept, root) if kept else ""
        if set_date_taken:
            row["Old date taken"], problem = stamp_date(target, move.date, done_word)
            if problem:
                done.problems.append((name, problem))
    except Exception as error:
        done.problems.append((name, f"{done_word}, but not finished ({why(error)})"))
    return row


def print_on(path, text, keep_original=True):
    """Keep an untouched copy of the photo in "Originals (no printed date)" (unless
    keep_original is False: a copy's original is still where it was), then print
    the date on it. Returns (path of the untouched copy or '', problem or '')."""
    import dating_photos_stamp as stamp
    problem = stamp.why_not(path)
    if problem:
        return "", problem
    if not keep_original:
        try:
            return "", stamp.print_date(path, text)
        except Exception as error:
            return "", f"the date couldn't be printed ({why(error)})"
    folder = os.path.join(os.path.dirname(path), ORIGINALS)
    copy = ""
    try:
        os.makedirs(folder, exist_ok=True)
        stem, ending = os.path.splitext(os.path.basename(path))
        copy, number = os.path.join(folder, stem + ending), 2
        while os.path.exists(copy):
            copy = os.path.join(folder, f"{stem} ({number}){ending}")
            number += 1
        shutil.copy2(path, copy)
        problem = stamp.print_date(path, text)
    except Exception as error:
        problem = f"the date couldn't be printed ({why(error)})"
    if problem:
        if copy:
            try:
                remove_file(copy)
            except OSError:
                pass
        return "", problem
    return copy, ""


def stamp_date(path, date, done_word=MOVED):
    """Set a moved (or copied) file's Date taken and file dates. Returns (old date
    taken as stored, problem or ''). "-" means the date taken wasn't changed."""
    old_taken = "-"
    problem = ""
    if can_set_date_taken(path):
        try:
            before = read_date_taken(path)
            if write_date_taken(path, exif_text(date)):
                old_taken = before or ""
        except Exception as error:
            problem = f"{done_word}, but its Date taken couldn't be set ({why(error)})"
    try:
        noon = seconds_at_noon(date)
        set_file_dates(path, noon, noon)
    except OSError as error:
        problem = problem or f"{done_word}, but its file dates couldn't be set ({why(error)})"
    return old_taken, problem


def why(error):
    if isinstance(error, PermissionError):
        return "it is open in another program, or Windows won't allow it"
    if isinstance(error, FileExistsError):
        return "a file with the new name is already there"
    if isinstance(error, FileNotFoundError):
        return "it is no longer there"
    return str(error) or type(error).__name__


def last_batch(root, rows=None):
    """The number of the last batch that hasn't been undone, or None."""
    batches = [batch_of(r) for r in (read_log(root) if rows is None else rows) if not r.get("Undone")]
    batches = [b for b in batches if b]
    return max(batches) if batches else None


def was_copied(row):
    return (row.get("Copied or moved") or "").strip().lower() == COPIED


def last_batch_copied(root, rows=None):
    """True if the last batch not undone was copied, False if it was moved, None if
    there is nothing to undo."""
    rows = read_log(root) if rows is None else rows
    batch = last_batch(root, rows)
    if batch is None:
        return None
    return any(was_copied(r) for r in rows if batch_of(r) == batch and not r.get("Undone"))


def undo_last(root, now=None):
    """Move the last batch back to where and what it was, with its old dates; if
    it was copied, take the copies away (their originals never changed). Returns
    (files put back or copies taken away, [(name, problem or note)]). If the list
    is open in Excel, nothing is done: the list couldn't record it."""
    now = now or dt.datetime.now()
    if log_in_use(root):
        return 0, [(LOG_NAME, LOG_BUSY)]
    rows = read_log(root)
    batch = last_batch(root, rows)
    if batch is None:
        return 0, []
    back, problems = 0, []
    folders = set()
    try:
        for row in reversed(rows):
            if row.get("Undone") or batch_of(row) != batch:
                continue
            before = os.path.join(root, row["Photo before"])
            now_at = os.path.join(root, row["Photo now"])
            if was_copied(row):
                there = os.path.exists(now_at)
                if take_copy_away(row, before, now_at, problems):
                    row["Undone"] = now.strftime("%Y-%m-%d %H:%M")
                    folders.add(os.path.dirname(now_at))
                    back += there              # a copy already gone wasn't taken away now
                continue
            if not os.path.exists(now_at):
                problems.append((row["Photo now"], f"couldn't be put back: it is no longer in {DEST}"))
                continue
            if os.path.exists(before):
                problems.append((row["Photo before"], "couldn't be put back: a file with its old name is "
                                                      "back where it was"))
                continue
            try:
                note = put_back(root, row, before, now_at, folders)
            except Exception as error:
                problems.append((row["Photo now"], f"couldn't be put back ({why(error)})"))
                continue
            if note:
                problems.append((row["Photo before"], note))
            row["Undone"] = now.strftime("%Y-%m-%d %H:%M")
            folders.add(os.path.dirname(now_at))
            back += 1
    finally:
        problem = save_log(root, rows)
        if problem:
            problems.append((LOG_NAME, problem.replace("Undo won't know about these",
                                                       "it still shows these as moved")))
    # tidy away year and sub-folders the batch made, if they are now empty
    dest = os.path.normcase(os.path.join(root, DEST))
    for folder in sorted(folders, key=len, reverse=True):
        while os.path.normcase(folder).startswith(dest + os.sep) and os.path.normcase(folder) != dest:
            try:
                os.rmdir(folder)
            except OSError:
                break
            folder = os.path.dirname(folder)
    return back, problems


def take_copy_away(row, before, now_at, problems):
    """Delete one listed copy, but only while the photo it was copied from is still
    where it was: otherwise the copy may be the only one left, so it stays.
    Returns True if the row is finished with (taken away, or already gone)."""
    if not os.path.exists(now_at):
        problems.append((row["Photo now"], f"was already gone from {DEST}"))
        return True
    if not os.path.exists(before):
        problems.append((row["Photo now"], f"was kept: the photo it was copied from, {row['Photo before']}, "
                                           "is no longer there, so this copy may be the only one"))
        return False
    try:
        remove_file(now_at)
    except OSError as error:
        problems.append((row["Photo now"], f"couldn't be taken away ({why(error)})"))
        return False
    return True


def copied_already(root):
    """{photo (its path, as os.path.normcase gives it): (where its copy is, its date)}
    for every photo copied into "Dated Photographs" and not undone, whose copy is
    still there: so it isn't copied a second time by mistake."""
    found = {}
    for row in read_log(root):
        # a list changed by hand in Excel may have empty cells, or columns gone
        if not was_copied(row) or row.get("Undone") or not row.get("Photo now") or not row.get("Photo before"):
            continue
        if os.path.exists(os.path.join(root, row["Photo now"])):
            try:
                date = dt.date.fromisoformat(row.get("Date") or "")
            except ValueError:
                date = None
            found[os.path.normcase(os.path.normpath(os.path.join(root, row["Photo before"])))] = \
                (row["Photo now"], date)
    return found


def put_back(root, row, before, now_at, folders):
    """Put one listed file back as it was. Returns '' or a note about how it went."""
    note = ""
    kept = os.path.join(root, row["Original kept"]) if row.get("Original kept") else ""
    os.makedirs(os.path.dirname(before), exist_ok=True)
    if kept and os.path.exists(kept):
        # the date was printed on it: the untouched copy goes back instead
        os.rename(kept, before)
        folders.add(os.path.dirname(kept))
        try:
            remove_file(now_at)
        except OSError as error:
            note = (f"put back, but the copy with the printed date couldn't be deleted from "
                    f"{DEST} ({why(error)})")
    else:
        if kept:
            note = ("put back, but its untouched copy was gone from “" + ORIGINALS
                    + "”, so the copy with the printed date went back instead")
        old = row.get("Old date taken", "-")
        if old != "-":
            if old:
                write_date_taken(now_at, old)
            else:
                remove_date_taken(now_at)
        os.rename(now_at, before)
    if row.get("Old file dates"):
        # the file is back by now, so a problem here is only a note: its dates stay as they are
        try:
            modified, created = (float(x) for x in row["Old file dates"].split(";"))
            set_file_dates(before, modified, created)
        except (OSError, ValueError) as error:
            note = note or f"put back, but its old file dates couldn't be set ({why(error)})"
    return note


def remove_date_taken(path):
    """Take out a Date taken this program added to a JPG that had none."""
    if path.lower().endswith((".tif", ".tiff")):
        return
    with open(path, "rb") as f:
        data = bytearray(f.read())
    block, _ = _jpeg_exif(bytes(data))
    if block is None:
        return
    from PIL import Image
    exif = Image.Exif()
    exif.load(bytes(data[block[0]:block[1]]))
    exif_ifd = exif.get_ifd(EXIF_POINTER)
    exif_ifd.pop(DATE_TAKEN_TAG, None)
    if exif_ifd:
        exif[EXIF_POINTER] = exif_ifd
    else:
        exif.pop(EXIF_POINTER, None)
    if len(exif) == 0:
        data[block[0] - 4:block[1]] = b""
    else:
        data[block[0] - 4:block[1]] = _segment(exif.tobytes())
    write_file(path, data)
