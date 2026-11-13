"""Song Cleaner: finding copies of the same recording, and choosing the one to keep.

Two files are copies when their sound lines up: take twelve seconds of one,
slide it along the fingerprint of the other, and at the right place nearly
every bit agrees. A live version, a cover or a re-recording never lines up,
whatever it is called, so those are left alone. It is done twelve seconds at
a time because copies do not always run side by side all the way: a radio
edit drops a verse, and one station plays a touch faster than another.
Where copies line up also shows what each has extra at its ends (talk, the
song before) and which copy has more of the song.
"""

import collections
import dataclasses

import numpy as np

import song_cleaner_names as names
from song_cleaner_engine import FRAME_S

SEGMENT = 128                 # frames compared at a time, about 12 s
STEP = 64
MATCH_ERROR = 0.30            # share of fingerprint bits that may differ in a matching segment
MIN_SHARED_S = 60.0           # copies share at least this much (or 60% of the shorter one)
PATCH = 21                    # frames, about 2 s: the stretch judged "same sound or not"
FINE = 7                      # frames, about 0.65 s: for finding the edge more exactly
ALIKE = 0.5                   # how alike a stretch must be to count as the same sound
LONG_S = 20.0                 # more than this at one end is more likely song than talk
SAME_S = 1.5                  # ends closer than this count as the same place


@dataclasses.dataclass
class LineUp:
    start_a: int              # the part of a that b also has: first and one-past-last frame
    end_a: int
    start_b: int
    end_b: int
    shared: float             # seconds of sound that line up
    error: float

    @property
    def offset(self):
        """Frames by which a's song starts later than b's."""
        return self.start_a - self.start_b


def _bits(fp):
    """Each fingerprint word as 32 values of +1 or -1."""
    table = ((fp[:, None] >> np.arange(32, dtype=np.uint32)) & 1).astype(np.float32)
    return table * 2 - 1


def _segment_offsets(fp_a, fp_b):
    """For each 12 s of a: (frame, offset where it sits best in b, share of bits that differ)."""
    size = 1 << int(np.ceil(np.log2(len(fp_b) + SEGMENT)))
    whole_b = np.conj(np.fft.rfft(_bits(fp_b), size, axis=0))
    bits_a = _bits(fp_a)
    found = []
    for at in range(0, len(fp_a) - SEGMENT + 1, STEP):
        agree = np.fft.irfft(np.fft.rfft(bits_a[at:at + SEGMENT], size, axis=0) * whole_b,
                             size, axis=0).sum(axis=1)
        lag = int(np.argmax(agree))
        in_b = (size - lag) % size
        if in_b + SEGMENT > len(fp_b) + SEGMENT // 2:
            continue                                # wrapped round: not a real place in b
        found.append((at, at - in_b, (1 - float(agree[lag]) / (SEGMENT * 32)) / 2))
    return found


def _chains(segments):
    """Matching segments whose offsets agree with a neighbour's (allowing a little drift)."""
    good = [seg for seg in segments if seg[2] < MATCH_ERROR]
    kept = []
    for n, (at, off, err) in enumerate(good):
        for other in (good[n - 1] if n else None, good[n + 1] if n + 1 < len(good) else None):
            if other and abs(other[0] - at) <= STEP * 4 \
                    and abs(other[1] - off) <= 1 + 0.06 * abs(other[0] - at):
                kept.append((at, off, err))
                break
    return kept


def _patch_alike(a, b, patch=PATCH):
    """How alike two equally long runs of band levels are, per 2 s stretch (-1 to 1).

    Each band's own average over the stretch is taken away first, so what is
    compared is how the sound moves, not how loud or bright the station made it.
    """
    a = a.astype(np.float64)
    b = b.astype(np.float64)

    def run(x):
        total = np.concatenate([np.zeros((1,) + x.shape[1:]), np.cumsum(x, axis=0)])
        return total[patch:] - total[:-patch]

    sum_a, sum_b = run(a), run(b)
    top = (run(a * b) - sum_a * sum_b / patch).sum(axis=1)
    spread = (run(a * a) - sum_a ** 2 / patch).sum(axis=1) * (run(b * b) - sum_b ** 2 / patch).sum(axis=1)
    return np.where(spread > 1e-6, top / np.sqrt(np.maximum(spread, 1e-6)), 0.0)


def _runs(mask, shortest):
    """[(start, end)] of runs of True at least `shortest` long."""
    edges = np.flatnonzero(np.diff(np.concatenate([[0], mask.astype(np.int8), [0]])))
    return [(int(s), int(e)) for s, e in zip(edges[::2], edges[1::2]) if e - s >= shortest]


def line_up(song_a, song_b):
    """How two songs line up, or None when they are not the same recording."""
    fp_a, fp_b = song_a.fp, song_b.fp
    if fp_a is None or fp_b is None or min(len(fp_a), len(fp_b)) < SEGMENT * 2:
        return None
    chain = _chains(_segment_offsets(fp_a, fp_b))
    need = min(MIN_SHARED_S, 0.6 * min(len(fp_a), len(fp_b)) * FRAME_S)
    if len(chain) * STEP * FRAME_S < need * 0.8:
        return None
    # where each frame of a sits in b: follow the nearest matching segment
    centres = np.array([at + SEGMENT // 2 for at, _, _ in chain])
    offsets = np.array([off for _, off, _ in chain])
    frames_a = np.arange(len(fp_a))
    nearest = np.abs(frames_a[:, None] - centres[None, :]).argmin(axis=1)
    frames_b = frames_a - offsets[nearest]
    inside = (frames_b >= 0) & (frames_b < len(fp_b))
    frames_a, frames_b = frames_a[inside], frames_b[inside]
    if len(frames_a) < PATCH * 2:
        return None
    alike = _patch_alike(song_a.bands[frames_a], song_b.bands[frames_b]) > ALIKE
    runs = _runs(alike, shortest=PATCH // 2)
    if not runs:
        return None
    shared = sum(e - s for s, e in runs) * FRAME_S
    if shared < need:
        return None
    # a stretch is judged whole; one that straddles the edge of the shared part
    # still passes, so the real edge is about half a stretch further in
    first = runs[0][0]
    last = runs[-1][1] - 1 + PATCH - 1
    if first > 0:
        first += PATCH // 2
    if last < len(frames_a) - 1:
        last -= PATCH // 2
    # Talk giving way to music is a step that looks alike in any two files, and
    # can pass as "the same sound" for a second or two. Short stretches are not
    # fooled: move each edge in (2 s at most) to where three in a row agree.
    fine = _patch_alike(song_a.bands[frames_a], song_b.bands[frames_b], FINE) > ALIKE
    steady = fine[:-2] & fine[1:-1] & fine[2:] if len(fine) > 2 else fine
    for move in range(PATCH + 1):
        if first + move < len(steady) and steady[first + move]:
            first += move
            break
    for move in range(PATCH + 1):
        at = last - move - (FINE + 2 - 1)           # the three stretches ending at this frame
        if 0 <= at < len(steady) and steady[at]:
            last -= move
            break
    return LineUp(int(frames_a[first]), int(frames_a[last]) + 1, int(frames_b[first]),
                  int(frames_b[last]) + 1, shared, float(np.mean([err for _, _, err in chain])))


@dataclasses.dataclass
class Member:
    song: object
    start: float = 0.0        # where the song itself starts and ends in this file, in seconds
    end: float = 0.0
    lead_extra: float = 0.0   # sound before the song that no other copy has
    tail_extra: float = 0.0
    missing: float = 0.0      # seconds of the song this copy lacks, compared with the fullest
    note: str = ""
    unsure: bool = False      # a long stretch only this copy has: song or talk, not known


@dataclasses.dataclass
class CopyGroup:
    members: list
    keep: int = 0

    @property
    def keeper(self):
        return self.members[self.keep]


def default_name(song):
    artist, title = names.split_name(song.stem)
    return names.tidy_artist(artist), names.tidy(names.clean_title(title))


def candidate_pairs(songs, name_of=default_name):
    """Pairs of songs (by position) worth lining up.

    Same artist and title once tidied; same title under artists spelt nearly
    alike or one containing the other; same artist with one title starting
    with the other (a cut-off name) or a letter or two different.
    """
    keyed = [(names.letters_only(a), names.letters_only(t)) for a, t in map(name_of, songs)]
    pairs = set()
    by_title = collections.defaultdict(list)
    by_artist = collections.defaultdict(list)
    for index, (artist, title) in enumerate(keyed):
        if songs[index].fp is None:
            continue
        by_title[title].append(index)
        by_artist[artist].append(index)
    for title, group in by_title.items():
        if not title or len(group) > 40:
            continue
        for x, i in enumerate(group):
            for j in group[x + 1:]:
                a, b = keyed[i][0], keyed[j][0]
                if a == b or not a or not b or a in b or b in a \
                        or names.edit_distance(a, b, 2) <= 2:
                    pairs.add((i, j))
    for artist, group in by_artist.items():
        if not artist:
            continue
        for x, i in enumerate(group):
            for j in group[x + 1:]:
                a, b = keyed[i][1], keyed[j][1]
                if not a or not b or a == b:
                    continue
                short, long = sorted((a, b), key=len)
                if (len(short) >= 8 and long.startswith(short)) \
                        or (len(short) >= 6 and names.edit_distance(a, b, 2) <= 2):
                    pairs.add((i, j))
    return pairs


def _plainness(song):
    """Lower is plainer: no copy number, no junk, proper capitals."""
    stem = song.stem
    return (names.clean_title(stem) != stem, " - " not in stem,
            stem.islower() or stem.isupper(), len(stem))


def choose_keeper(members):
    """Position of the copy to keep: whole song, least extra, best sound, plainest name."""
    def rank(item):
        index, m = item
        return (round(m.missing / SAME_S),                      # whole song first
                round((m.lead_extra + m.tail_extra) / SAME_S),  # then least extra at the ends
                -m.song.bitrate, -m.song.channels, _plainness(m.song), index)
    return min(enumerate(members), key=rank)[0]


def _describe(m):
    parts = [f"{m.missing:.0f} s shorter than the fullest copy" if m.missing >= SAME_S
             else "whole song"]
    if m.lead_extra >= 0.5:
        parts.append(f"{m.lead_extra:.0f} s extra at the start")
    if m.tail_extra >= 0.5:
        parts.append(f"{m.tail_extra:.0f} s extra at the end")
    if m.unsure:
        parts.append("has a long stretch the others lack (a longer version, or talk)")
    return ", ".join(parts)


def measure_group(songs, lineups, is_break=None):
    """Members for songs known to be copies of each other.

    lineups: {(i, j): LineUp} between positions in songs. The song in each
    file runs from the first to the last sound that any other copy also has.
    What only this file has before or after that is extra (talk, a jingle,
    the song before), unless it is longer than LONG_S: then it is more likely
    a longer version of the song and counts as song, unless
    is_break(song, seconds, "start" or "end") says the sound changes there.
    """
    spans = {}
    for (a, b), lu in lineups.items():
        for index, s, e in ((a, lu.start_a, lu.end_a), (b, lu.start_b, lu.end_b)):
            old = spans.get(index)
            spans[index] = (min(s, old[0]), max(e, old[1])) if old else (s, e)
    members, content = [], []
    for i, song in enumerate(songs):
        frames = len(song.fp)
        s, e = spans.get(i, (0, frames))
        m = Member(song, s * FRAME_S, e * FRAME_S)
        lead, tail = s * FRAME_S, (frames - e) * FRAME_S
        has = (e - s) * FRAME_S
        for side, amount in (("start", lead), ("end", tail)):
            at = m.start if side == "start" else m.end
            if amount > LONG_S and not (is_break and is_break(song, at, side)):
                has += amount                      # a longer version, as far as can be told
                m.unsure = True
                if side == "start":
                    m.start = 0.0
                else:
                    m.end = frames * FRAME_S
            elif side == "start":
                m.lead_extra = amount
            else:
                m.tail_extra = amount
        members.append(m)
        content.append(has)
    for m, has in zip(members, content):
        m.missing = max(content) - has
        m.note = _describe(m)
    return members


def find_copies(songs, progress=None, name_of=default_name, should_stop=None, is_break=None):
    """(copy groups, groups with the same name but different recordings)."""
    pairs = sorted(candidate_pairs(songs, name_of))
    lineups = {}
    for done, (i, j) in enumerate(pairs, 1):
        if should_stop and should_stop():
            break
        found = line_up(songs[i], songs[j])
        if found:
            lineups[(i, j)] = found
        if progress and (done % 5 == 0 or done == len(pairs)):
            progress(done, len(pairs), songs[i].name)
    parent = list(range(len(songs)))

    def root(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i, j in lineups:
        parent[root(i)] = root(j)
    families = collections.defaultdict(list)
    for index in range(len(songs)):
        families[root(index)].append(index)
    groups = []
    for family in families.values():
        if len(family) < 2:
            continue
        place = {index: n for n, index in enumerate(family)}
        inner = {(place[i], place[j]): lu for (i, j), lu in lineups.items() if i in place}
        members = measure_group([songs[i] for i in family], inner, is_break)
        groups.append(CopyGroup(members, choose_keeper(members)))
    groups.sort(key=lambda g: g.keeper.song.name.lower())

    # same name, sound that does not line up
    by_name = collections.defaultdict(set)
    for index, song in enumerate(songs):
        if song.fp is not None:
            by_name[names.loose_key(*name_of(song))].add(root(index))
    different = []
    for key, roots in by_name.items():
        if len(roots) > 1 and key != "|":
            different.append(sorted((songs[i] for r in roots for i in families[r]),
                                    key=lambda s: s.name.lower()))
    different.sort(key=lambda group: group[0].name.lower())
    return groups, different
