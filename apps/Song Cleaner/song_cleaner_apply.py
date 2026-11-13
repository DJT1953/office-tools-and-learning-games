"""Song Cleaner: carrying out the changes, and undoing them.

Nothing is ever deleted. A copy that is not wanted is moved to a folder beside
the music folder; a song that is trimmed has its untouched original moved to
another. Every step is written to a log before it is done, so the last Apply
can be put back exactly, even if the computer stopped half way through it.
"""

import base64
import dataclasses
import json
import os
import shutil
import time

import song_cleaner_trim as trim
from song_cleaner_engine import TEMP_PREFIX, SongError, app_folder

REMOVED = "Song Cleaner - removed copies"
ORIGINALS = "Song Cleaner - originals before trimming"
SPARE_BYTES = 200 * 1024 * 1024


@dataclasses.dataclass
class Change:
    kind: str                 # "remove", "trim" or "rename"
    path: str
    size: int = 0             # as the file was when scanned; a file that differs is left alone
    mtime: float = 0.0
    start: float = 0.0        # trim: keep from ... to
    end: float = 0.0
    artist: str = ""          # rename
    title: str = ""
    new_name: str = ""


@dataclasses.dataclass
class Report:
    done: int = 0
    skipped: list = dataclasses.field(default_factory=list)     # (file name, why)
    stopped: bool = False


def default_log():
    return os.path.join(app_folder(), "last apply.log")


def build_changes(songs, renames=(), removals=(), trims=()):
    """The changes in the order to carry them out.

    songs: {path: Song}. renames: NameChange list. removals: paths.
    trims: Trim list. A file being removed gets nothing else done to it; a
    file both trimmed and renamed is trimmed first.
    """
    gone = set(removals)
    changes = []

    def facts(path):
        song = songs.get(path)
        return (song.size, song.mtime) if song else (0, 0.0)

    for path in removals:
        changes.append(Change("remove", path, *facts(path)))
    for cut in trims:
        if cut.path not in gone:
            changes.append(Change("trim", cut.path, *facts(cut.path), start=cut.start, end=cut.end))
    for name in renames:
        if name.path not in gone and (name.new_name != name.old_name or name.artist or name.title):
            changes.append(Change("rename", name.path, *facts(name.path), artist=name.artist,
                                  title=name.title, new_name=name.new_name))
    return changes


def space_needed(changes):
    """Bytes of free disk the changes need: a trimmed song exists twice for a moment."""
    return sum(c.size for c in changes if c.kind == "trim") + SPARE_BYTES


def free_space(folder):
    return shutil.disk_usage(folder).free


def review_folder(folder, name):
    parent = os.path.dirname(os.path.abspath(folder).rstrip("\\/"))
    if not parent or os.path.normcase(parent) == os.path.normcase(os.path.abspath(folder).rstrip("\\/")):
        parent = folder                              # the music is a whole drive: keep it inside
    return os.path.join(parent, name)


def free_name(path, taken=()):
    """This path, or the same with (2), (3)... if something is already there."""
    stem, ending = os.path.splitext(path)
    candidate, number = path, 1
    while os.path.exists(candidate) or os.path.normcase(candidate) in taken:
        number += 1
        candidate = f"{stem} ({number}){ending}"
    return candidate


def _plain(error):
    if isinstance(error, PermissionError):
        return "another program is using it, or Windows would not allow it"
    if isinstance(error, FileNotFoundError):
        return "it is no longer there"
    if isinstance(error, OSError) and getattr(error, "winerror", None) == 112 or \
            getattr(error, "errno", None) == 28:
        return "the disk is full"
    return str(getattr(error, "strerror", None) or error)


class _Log:
    def __init__(self, path, folder):
        self.file = open(path, "w", encoding="utf-8")
        self.count = 0
        self.write({"folder": folder, "when": time.strftime("%Y-%m-%d %H:%M:%S")})

    def write(self, entry):
        self.file.write(json.dumps(entry) + "\n")
        self.file.flush()
        os.fsync(self.file.fileno())

    def step(self, **entry):
        self.count += 1
        self.write({"n": self.count, **entry})

    def close(self):
        self.file.close()


def _tag_edges(file, size):
    """(where the sound starts, where it ends) in an MP3: between the tags at each end."""
    file.seek(0)
    head = file.read(10)
    start = 0
    if len(head) == 10 and head[:3] == b"ID3":
        start = 10 + ((head[6] & 127) << 21 | (head[7] & 127) << 14 | (head[8] & 127) << 7 | (head[9] & 127))
        if head[5] & 0x10:
            start += 10                                  # a footer
    end = size
    if size >= start + 128:
        file.seek(size - 128)
        if file.read(3) == b"TAG":
            end = size - 128
    return min(start, size), end


def _tag_bytes(path):
    """The tags at both ends of an MP3, exactly as they are, as text for the log."""
    size = os.path.getsize(path)
    with open(path, "rb") as file:
        start, end = _tag_edges(file, size)
        file.seek(0)
        head = file.read(start)
        file.seek(end)
        tail = file.read()
    return {"head": base64.b64encode(head).decode("ascii"), "tail": base64.b64encode(tail).decode("ascii")}


def _restore_tag_bytes(path, saved):
    """Give an MP3 back the exact tags it had; the sound between them is not touched."""
    size = os.path.getsize(path)
    temp = os.path.join(os.path.dirname(path), TEMP_PREFIX + " tags.mp3")
    with open(path, "rb") as file, open(temp, "wb") as out:
        start, end = _tag_edges(file, size)
        out.write(base64.b64decode(saved["head"]))
        file.seek(start)
        left = end - start
        while left > 0:
            block = file.read(min(left, 1 << 20))
            if not block:
                break
            out.write(block)
            left -= len(block)
        out.write(base64.b64decode(saved["tail"]))
    os.replace(temp, path)


def _read_tags(path):
    import mutagen.id3
    try:
        tags = mutagen.id3.ID3(path)
    except mutagen.id3.ID3NoHeaderError:
        return None
    text = lambda key: str(tags[key].text[0]) if key in tags and tags[key].text else None
    return {"artist": text("TPE1"), "title": text("TIT2")}


def _write_tags(path, artist, title, had_tags=True):
    """Set (or with None, take away) the artist and title inside the file."""
    import mutagen.id3
    try:
        tags = mutagen.id3.ID3(path)
    except mutagen.id3.ID3NoHeaderError:
        if not had_tags or (artist is None and title is None):
            return
        tags = mutagen.id3.ID3()
    for key, frame, value in (("TPE1", mutagen.id3.TPE1, artist), ("TIT2", mutagen.id3.TIT2, title)):
        tags.delall(key)
        if value is not None:
            tags.add(frame(encoding=3, text=[value]))
    if not had_tags:
        tags.delete(path)
        return
    tags.save(path, v2_version=3 if tags.version[1] == 3 else 4)


def apply(changes, folder, log_path=None, progress=None, should_stop=None, free=free_space):
    """Carry out the changes. Returns a Report; raises SongError if it cannot start."""
    log_path = log_path or default_log()
    if space_needed(changes) > free(folder) and any(c.kind == "trim" for c in changes):
        raise SongError("There is not enough free disk space to do this safely. Nothing was changed.")
    removed, originals = review_folder(folder, REMOVED), review_folder(folder, ORIGINALS)
    try:
        if any(c.kind == "remove" for c in changes):
            os.makedirs(removed, exist_ok=True)
        if any(c.kind == "trim" for c in changes):
            os.makedirs(originals, exist_ok=True)
        log = _Log(log_path, folder)
    except OSError as error:
        raise SongError(f"The review folders or the log could not be made ({_plain(error)}). "
                        "Nothing was changed.") from error
    report = Report()
    ours = {}                              # path -> its new path, for files this run changed
    try:
        for number, change in enumerate(changes, 1):
            if should_stop and should_stop():
                report.stopped = True
                break
            path = ours.get(os.path.normcase(change.path), change.path)
            name = os.path.basename(path)
            if progress:
                progress(number, len(changes), name)
            try:
                stat = os.stat(path)
                touched = os.path.normcase(change.path) in ours
                if not touched and change.size and (stat.st_size != change.size
                                                    or abs(stat.st_mtime - change.mtime) > 2):
                    report.skipped.append((name, "it has changed since the scan; scan again"))
                    continue
                if not os.access(path, os.W_OK):
                    report.skipped.append((name, "it is marked read-only"))
                    continue
                if change.kind == "remove":
                    target = free_name(os.path.join(removed, name))
                    log.step(op="move", src=path, dst=target)
                    os.rename(path, target)
                elif change.kind == "trim":
                    temp = os.path.join(os.path.dirname(path), TEMP_PREFIX + " cutting.mp3")
                    kept = free_name(os.path.join(originals, name))
                    log.step(op="trim", src=path, kept=kept, temp=temp)
                    trim.cut_mp3(path, temp, change.start, change.end)
                    try:
                        os.rename(path, kept)
                    except OSError:
                        os.remove(temp)
                        raise
                    os.rename(temp, path)
                    ours[os.path.normcase(change.path)] = path
                else:
                    target = os.path.join(os.path.dirname(path), change.new_name or name)
                    same_file = os.path.normcase(target) == os.path.normcase(path)
                    if not same_file:
                        target = free_name(target)
                    old = _read_tags(path)
                    new_artist = change.artist or (old or {}).get("artist")
                    new_title = change.title or (old or {}).get("title")
                    retag = old is None or old["artist"] != new_artist or old["title"] != new_title
                    if same_file and target == path and not retag:
                        continue
                    saved = _tag_bytes(path) if retag else None
                    log.step(op="rename", src=path, dst=target, old=old, retag=retag, saved=saved)
                    if retag:
                        _write_tags(path, new_artist, new_title)
                    if target != path:
                        os.rename(path, target)
                    ours[os.path.normcase(change.path)] = target
                log.write({"n": log.count, "done": True})
                report.done += 1
            except SongError as error:
                report.skipped.append((name, str(error)))
            except OSError as error:
                report.skipped.append((name, _plain(error)))
            except Exception as error:  # noqa: BLE001 - one odd file must not stop the rest
                report.skipped.append((name, f"it could not be changed ({type(error).__name__})"))
    finally:
        log.close()
    return report


def has_undo(log_path=None):
    return os.path.exists(log_path or default_log())


def undo(log_path=None, progress=None):
    """Put back everything the last Apply did, newest first."""
    log_path = log_path or default_log()
    steps = []
    try:
        with open(log_path, encoding="utf-8") as file:
            for line in file:
                try:
                    entry = json.loads(line)
                except ValueError:
                    continue                             # a line cut short by a crash
                if "op" in entry:
                    steps.append(entry)
    except OSError as error:
        raise SongError("There is nothing to undo.") from error
    report = Report()
    for number, step in enumerate(reversed(steps), 1):
        name = os.path.basename(step["src"])
        if progress:
            progress(number, len(steps), name)
        try:
            if step["op"] == "move":
                if os.path.exists(step["dst"]):
                    if os.path.exists(step["src"]):
                        report.skipped.append((name, "another file now has its old name; it is still in "
                                               + os.path.basename(os.path.dirname(step["dst"]))))
                        continue
                    os.rename(step["dst"], step["src"])
                    report.done += 1
            elif step["op"] == "trim":
                if os.path.exists(step["temp"]):
                    os.remove(step["temp"])              # a cut that was never put in place
                if os.path.exists(step["kept"]):
                    if os.path.exists(step["src"]):
                        os.remove(step["src"])           # the trimmed version this app made
                    os.rename(step["kept"], step["src"])
                    report.done += 1
            elif step["op"] == "rename":
                now = step["dst"] if os.path.exists(step["dst"]) else step["src"]
                if not os.path.exists(now):
                    report.skipped.append((name, "it is no longer there"))
                    continue
                if step.get("retag"):
                    old = step.get("old")
                    if step.get("saved"):
                        _restore_tag_bytes(now, step["saved"])
                    else:
                        _write_tags(now, old and old["artist"], old and old["title"],
                                    had_tags=old is not None)
                if now != step["src"]:
                    if os.path.exists(step["src"]) and os.path.normcase(now) != os.path.normcase(step["src"]):
                        report.skipped.append((name, "another file now has its old name"))
                        continue
                    os.rename(now, step["src"])
                report.done += 1
        except OSError as error:
            report.skipped.append((name, _plain(error)))
    try:
        os.replace(log_path, log_path + ".undone")
    except OSError:
        pass
    for folder in {os.path.dirname(s.get("dst") or s.get("kept") or "") for s in steps
                   if s["op"] in ("move", "trim")}:
        try:
            os.rmdir(folder)                             # only goes if it is empty
        except OSError:
            pass
    return report
