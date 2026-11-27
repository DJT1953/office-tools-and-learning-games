"""Dating Photographs: works out the best date for one photograph.

It reads the back first, if the back was scanned as its own file, then the
front (unless the back already gave an exact date), and adds dates written
in the file's and folder's names. A date in a file name that matches the day
the file was made is the scan date, not the photo's, so it is left out (only
from 1990 on: no scan is older, and a photo already dated by this program
carries its date in both). When two different readers agree on a date, it
counts for more.
"""

from __future__ import annotations

import datetime as dt
import os
from dataclasses import dataclass, field, replace

from PIL import Image, UnidentifiedImageError

import dating_photos_dates as dates
import dating_photos_files as files

Image.MAX_IMAGE_PIXELS = None          # big scans are fine
GOOD_ENOUGH = 0.6                      # stop reading a side once a date scores this
SURE = 0.8                             # an exact date this sure on the back: skip the front
TICK_FROM = 0.45                       # dates scoring less start unticked, marked "check this"
AGREE_BONUS = 0.1
FIRST_SCAN_YEAR = 1990                 # a file-name date before this can't be the day it was scanned


@dataclass
class Finding:
    """What was found for one photo."""
    photo: files.Photo
    clues: list = field(default_factory=list)       # every date found, best first
    words: list = field(default_factory=list)       # (where, text) of everything read
    problem: str = ""                                # why it couldn't be read, if it couldn't

    @property
    def best(self):
        return self.clues[0] if self.clues else None


def name_clues(path, root, today=None, month_first=True):
    """Dates in the file's name (unless it's the scan date) and its folders' names."""
    today = today or dt.date.today()
    stem = os.path.splitext(os.path.basename(path))[0]
    clues = dates.from_name(stem, "file name", today, month_first)
    try:
        own = {files.day_of(t) for t in files.file_dates(path)}
    except OSError:
        own = set()
    clues = [c for c in clues if c.date.year < FIRST_SCAN_YEAR
             or not any(abs((c.date - day).days) <= 2 for day in own)]
    folders = [os.path.basename(os.path.normpath(root))]
    relative = os.path.relpath(os.path.dirname(path), root)
    if relative != ".":
        folders += relative.split(os.sep)
    for name in folders:
        clues += dates.from_name(name, "folder name", today, month_first)
    return clues


def read_side(reader, path, where, today, month_first, stop=lambda: False):
    """Read one side: ([(reader name, Clue)], [(where, text)])."""
    found, words = [], []
    with Image.open(path) as picture:
        picture.load()
        for pieces in reader.passes(picture):
            if stop():
                break
            for piece in pieces:
                words.append((where, piece.text))
                found += [(piece.reader, clue) for clue in dates.find_dates(
                    piece.text, where, piece.confidence, piece.box, today, month_first)]
            best = dates.best([clue for _, clue in found])
            if best and best.score >= GOOD_ENOUGH:
                break
    return found, words


def agree(found):
    """Clues, one per date, best first; a date that two different readers found gets a bonus."""
    readers = {}
    for reader, clue in found:
        readers.setdefault(clue.date, set()).add(reader)
    merged = dates.merge([clue for _, clue in found])
    out = []
    for clue in merged:
        if len(readers.get(clue.date, ())) > 1:
            clue = replace(clue, score=round(min(1.0, clue.score + AGREE_BONUS), 4))
        out.append(clue)
    return sorted(out, key=lambda c: (-c.score, -dates.EXACTNESS[c.kind]))


def unreadable(path, error):
    """Plain words for why a picture couldn't be read."""
    if path.lower().endswith(files.CANT_OPEN):
        return ("couldn't be read: it is an iPhone HEIC photo, which this program can't open "
                "(save it as a JPG first)")
    try:
        if os.path.getsize(path) == 0:
            return "couldn't be read: the file is empty"
    except OSError:
        return "couldn't be read: it is no longer there"
    if isinstance(error, (UnidentifiedImageError, SyntaxError)) or "truncated" in str(error) \
            or "broken data" in str(error):
        return "couldn't be read: the file is damaged, or isn't really a picture"
    return f"couldn't be read ({error})"


def find(reader, photo, root, today=None, month_first=True, stop=lambda: False):
    """The Finding for one photo. A side that can't be read (a damaged back) is
    reported and the other side is still read."""
    today = today or dt.date.today()
    finding = Finding(photo)
    found, problems = [], []

    def read(path, where):
        try:
            clues, words = read_side(reader, path, where, today, month_first, stop)
        except Exception as error:      # a damaged picture is reported, not fatal
            words = unreadable(path, error)
            problems.append(words if where == "front" else f"its back {words}")
            return
        found.extend(clues)
        finding.words.extend(words)

    if photo.back:
        read(photo.back, "back")
    best = dates.best([clue for _, clue in found])
    exact_on_back = best and best.score >= SURE and best.kind in (dates.DAY, dates.HOLIDAY)
    if not exact_on_back and not stop():
        read(photo.front, "front")
    finding.problem = "; ".join(problems)
    found += [("name", clue) for clue in name_clues(photo.front, root, today, month_first)]
    finding.clues = agree(found)
    return finding
