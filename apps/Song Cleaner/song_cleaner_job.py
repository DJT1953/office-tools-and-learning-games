"""Song Cleaner: putting the proposals together.

Takes the scanned songs and works out everything the window shows: names to
change, copies to remove, trims that are sure, trims to listen to first, and
problems. Also keeps which of them are ticked. No window needed.
"""

import dataclasses
import os

import song_cleaner_apply as applying
import song_cleaner_copies as copies
import song_cleaner_names as names
import song_cleaner_trim as trim


@dataclasses.dataclass
class Proposals:
    songs: dict = dataclasses.field(default_factory=dict)        # path -> Song (readable ones)
    names: list = dataclasses.field(default_factory=list)        # NameChange
    groups: list = dataclasses.field(default_factory=list)       # CopyGroup
    different: list = dataclasses.field(default_factory=list)    # [Song, ...] same name, other recording
    trims: dict = dataclasses.field(default_factory=dict)        # path -> Trim (sure and not)
    problems: list = dataclasses.field(default_factory=list)     # (file name, what is wrong)
    name_ticks: dict = dataclasses.field(default_factory=dict)   # path -> bool
    remove_ticks: dict = dataclasses.field(default_factory=dict)
    trim_ticks: dict = dataclasses.field(default_factory=dict)
    copy_ends: dict = dataclasses.field(default_factory=dict)    # path -> (start, end) from copies
    known_length: dict = dataclasses.field(default_factory=dict)  # path -> seconds, from the lookup

    # ---- what is on each page -------------------------------------------------
    def removed(self):
        return {path for path, on in self.remove_ticks.items() if on}

    def sure_trims(self):
        gone = self.removed()
        return [t for t in self.trims.values() if t.sure and t.path not in gone]

    def listen_trims(self):
        gone = self.removed()
        return [t for t in self.trims.values() if not t.sure and t.path not in gone]

    def doubtful(self):
        return [c for c in self.names if c.doubtful]

    # ---- changing one's mind --------------------------------------------------
    def keep_instead(self, group, index):
        """Keep another copy of a song; the rest of its copies become the ones to remove."""
        group.keep = index
        for n, member in enumerate(group.members):
            self.remove_ticks[member.song.path] = n != index

    def rename(self, change, artist, title):
        """Use a name typed by hand (or found online) for a song."""
        change.artist, change.title = names.squeeze(artist), names.squeeze(title)
        change.new_name = names.safe_file_name(change.artist, change.title)
        change.doubtful = False
        self.name_ticks[change.path] = True

    def nudge(self, path, side, seconds):
        """Move where a trim cuts: side is "start" or "end", seconds may be negative."""
        cut = self.trims[path]
        if side == "start":
            cut.start = min(max(0.0, cut.start + seconds), cut.end - 30.0)
        else:
            cut.end = max(min(cut.length, cut.end + seconds), cut.start + 30.0)

    def found_online(self, change, match):
        """Take a name (and the song's real length) from the online lookup."""
        self.rename(change, names.tidy(match.artist), names.tidy(match.title))
        change.reason = f"Found online ({match.score:.0%} match)"
        if match.length:
            self.known_length[change.path] = match.length
            self._propose_trim(self.songs[change.path])

    def _propose_trim(self, song):
        found = trim.propose_trim(song, self.copy_ends.get(song.path), self.known_length.get(song.path))
        if found is None:
            self.trims.pop(song.path, None)
            self.trim_ticks.pop(song.path, None)
        else:
            self.trims[song.path] = found
            self.trim_ticks[song.path] = found.sure

    # ---- what Apply will do ---------------------------------------------------
    def changes(self):
        renames = [c for c in self.names if self.name_ticks.get(c.path) and c.changes]
        removals = sorted(self.removed())
        trims = [t for t in self.trims.values() if self.trim_ticks.get(t.path)]
        return applying.build_changes(self.songs, renames, removals, trims)

    def counts(self):
        changes = self.changes()
        return {kind: sum(c.kind == kind for c in changes) for kind in ("rename", "remove", "trim")}


def analyse(songs, progress=None, should_stop=None):
    """Proposals for a list of scanned songs. progress(done, total, text)."""
    found = Proposals()
    good = []
    for song in songs:
        if song.error:
            found.problems.append((song.name, song.error))
        else:
            good.append(song)
            found.songs[song.path] = song

    found.names = names.propose_names(good)
    by_path = {c.path: c for c in found.names}
    for change in found.names:
        found.name_ticks[change.path] = change.changes

    def name_of(song):
        change = by_path.get(song.path)
        return (change.artist, change.title) if change else copies.default_name(song)

    def copy_progress(done, total, text):
        if progress:
            progress(done, total, "Comparing songs with the same name: " + text)

    found.groups, found.different = copies.find_copies(
        good, progress=copy_progress, name_of=name_of, should_stop=should_stop, is_break=trim.is_break)
    _untick_clashing_names(found, good)
    for group in found.groups:
        unsure = any(m.unsure for m in group.members)
        for n, member in enumerate(group.members):
            # where it is not known which copy is fuller, nothing is ticked: the user chooses
            found.remove_ticks[member.song.path] = n != group.keep and not unsure
            if not member.unsure:
                found.copy_ends[member.song.path] = (member.start, member.end)

    for done, song in enumerate(good, 1):
        if should_stop and should_stop():
            break
        found._propose_trim(song)
        if progress and (done % 25 == 0 or done == len(good)):
            progress(done, len(good), "Looking for talk at the ends: " + song.name)
    return found


def _untick_clashing_names(found, songs):
    """A rename to a name another song has (and is keeping) would only swap one number for another."""
    unsure_groups = [g for g in found.groups if any(m.unsure for m in g.members)]
    going = {m.song.path for g in found.groups if g not in unsure_groups
             for n, m in enumerate(g.members) if n != g.keep}
    taken = {song.name.lower() for song in songs if song.path not in going}
    for change in found.names:
        if not change.changes or change.path in going:
            continue
        old, new = change.old_name.lower(), change.new_name.lower()
        if new != old and new in taken:
            found.name_ticks[change.path] = False
            change.reason = "Another song already has that name, so it is left as it is"
        else:
            taken.discard(old)
            taken.add(new)


def describe_changes(counts, folder):
    """The words of the "are you sure" box before Apply."""
    def many(number, one, more):
        return f"{number:,} {one if number == 1 else more}"

    lines = []
    if counts["rename"]:
        lines.append(f"Rename {many(counts['rename'], 'song', 'songs')}.")
    if counts["remove"]:
        lines.append(f"Move {many(counts['remove'], 'copy', 'copies')} out of the music folder, to:\n"
                     f"    {applying.review_folder(folder, applying.REMOVED)}")
    if counts["trim"]:
        lines.append(f"Trim {many(counts['trim'], 'song', 'songs')}. The untouched originals go to:\n"
                     f"    {applying.review_folder(folder, applying.ORIGINALS)}")
    return "\n\n".join(lines)


def clock(seconds):
    seconds = int(round(seconds))
    return f"{seconds // 60}:{seconds % 60:02d}"


def file_label(path):
    return os.path.basename(path)
