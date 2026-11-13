"""Song Cleaner: finding where a song really starts and ends, and cutting there.

A radio recording often has something before or after the song: a DJ, a
jingle, the tail of the song before. Three kinds of clue say where:

  a gap      the sound dips nearly to nothing for a moment
  a change   the sound on one side of a moment is unlike the sound on the other
  copies     another recording of the same song has nothing there

No one clue is trusted alone. Music dips and changes by itself, and what one
copy has "extra" is often just more of the song, which the other station
faded early. A cut is called sure only where two clues point at the same
second. Anything less is a guess, offered for listening.

Cutting copies the MP3's own frames into a new file, so the sound that stays
is exactly what it was: nothing is decoded and encoded again.
"""

import dataclasses
import os
import subprocess

import numpy as np

from song_cleaner_engine import (FRAME_S, HOP, NO_WINDOW, RATE, WINDOW, SongError, decode_mono,
                                 ffmpeg_exe)

MAX_TRIM_S = 45.0             # never cut more than this from one end without being asked
AGREE_S = 1.0                 # two clues this close are pointing at the same place
SEARCH_S = 30.0               # how far in from each end to look
MIN_TRIM_S = 0.5              # less than this is not worth a cut
SIDE = 32                     # frames each side of a moment compared for "change" (3 s)
CHANGE = 2.2                  # how unlike the two sides must be to count as a change
LISTEN_MIN_S = 2.0            # a guess smaller than this is not worth listening to
GAP_DB = 18.0                 # a gap is at least this far below the sound either side
SILENT_DB = -78.0             # quieter than this is silence (a quiet intro is not)
MP3_FRAME_S = 1152 / 44100    # an MP3 frame: the finest a cut can be placed


def frame_time(frame):
    """The moment at the middle of a frame, in seconds."""
    return (frame * HOP + WINDOW / 2) / RATE


def frame_at(seconds):
    return int(round((seconds * RATE - WINDOW / 2) / HOP))


@dataclasses.dataclass
class Trim:
    path: str
    start: float              # keep from here ...
    end: float                # ... to here
    length: float
    sure: bool
    reason: str = ""

    @property
    def cut_start(self):
        return self.start

    @property
    def cut_end(self):
        return max(0.0, self.length - self.end)


def change_curve(song):
    """For each frame, how unlike the 3 s after it are from the 3 s before it.

    Compared: the shape of the sound (which bands are strong) and its level.
    Scaled by how much the sound varies anyway, so steady music scores low
    and a real change of what is playing scores high.
    """
    shape = song.bands.astype(np.float64) / 2.0                      # dB
    # a band 40 dB under the loudest cannot be heard; without this floor an
    # empty band jumping about would count as the sound "varying anyway"
    shape = np.maximum(shape, shape.max(axis=1, keepdims=True) - 40.0)
    level = shape.mean(axis=1, keepdims=True)
    feats = np.concatenate([shape - level, level * 0.5], axis=1)
    n = len(feats)
    total = np.concatenate([np.zeros((1, feats.shape[1])), np.cumsum(feats, axis=0)])
    squares = np.concatenate([np.zeros((1, feats.shape[1])), np.cumsum(feats ** 2, axis=0)])
    frames = np.arange(n + 1)
    lo = np.maximum(frames - SIDE, 0)
    hi = np.minimum(frames + SIDE, n)
    n_before = np.maximum(frames - lo, 1)[:, None]
    n_after = np.maximum(hi - frames, 1)[:, None]
    mean_before = (total[frames] - total[lo]) / n_before
    mean_after = (total[hi] - total[frames]) / n_after
    var_before = (squares[frames] - squares[lo]) / n_before - mean_before ** 2
    var_after = (squares[hi] - squares[frames]) / n_after - mean_after ** 2
    spread = np.sqrt(np.maximum((var_before + var_after) / 2, 0).mean(axis=1) + 1.0)
    curve = np.sqrt(((mean_after - mean_before) ** 2).mean(axis=1)) / spread
    short = np.minimum(frames - lo, hi - frames) < 8                  # too little on one side
    curve[short] = 0.0
    return curve[:n]


def _strongest_change(curve, first, last):
    """(frame, strength) of the biggest change between two frames, or (None, 0)."""
    first, last = max(first, 0), min(last, len(curve))
    if last - first < 3:
        return None, 0.0
    at = first + int(np.argmax(curve[first:last]))
    strength = float(curve[at])
    return (at, strength) if strength >= CHANGE else (None, strength)


def _gaps(song, first, last):
    """Frames where the sound dips at least GAP_DB below what is within 1.5 s on both sides."""
    level = song.level
    first, last = max(first, 1), min(last, len(level) - 1)
    reach = int(1.5 / FRAME_S)
    found = []
    for at in range(first, last):
        here = level[at]
        if here > level[at - 1] or here > level[at + 1]:
            continue                                               # not the bottom of a dip
        before = level[max(0, at - reach):at]
        after = level[at + 1:at + 1 + reach]
        if len(before) and len(after) and before.max() - here >= GAP_DB \
                and after.max() - here >= GAP_DB:
            found.append(at)
    return found


def silent_ends(song):
    """(seconds of silence at the start, at the end)."""
    quiet = song.level < SILENT_DB
    loud = np.flatnonzero(~quiet)
    if not len(loud):
        return 0.0, 0.0
    # a silent frame is silent from its first sample to its last
    lead = ((loud[0] - 1) * HOP + WINDOW) / RATE if loud[0] > 0 else 0.0
    tail = song.length - (loud[-1] + 1) * HOP / RATE if loud[-1] < len(quiet) - 1 else 0.0
    return max(0.0, lead), max(0.0, tail)


def ends_from_sound(song):
    """(start by gap, start by change, end by gap, end by change), in seconds or None.

    Looked for within SEARCH_S of each end. A gap is reported at its quietest
    moment; where there are several, the one with the biggest change of sound
    across it. A change is the strongest one.
    """
    curve = change_curve(song)
    n = len(song.level)
    reach = min(int(SEARCH_S / FRAME_S), n // 3)
    results = []
    for first, last in ((2, reach), (n - reach, n - 2)):
        gaps = _gaps(song, first, last)
        gap = max(gaps, key=lambda g: curve[max(0, g - 3):g + 4].max()) if gaps else None
        change, _ = _strongest_change(curve, first, last)
        results += [frame_time(gap) if gap is not None else None,
                    frame_time(change) if change is not None else None]
    return tuple(results)


def is_break(song, seconds, side=None):
    """Whether the sound has a gap or a change within a second of this moment."""
    at = frame_at(seconds)
    reach = int(AGREE_S / FRAME_S)
    if _gaps(song, at - reach, at + reach + 1):
        return True
    return _strongest_change(change_curve(song), at - reach, at + reach + 1)[0] is not None


def quiet_point(song, seconds, within=0.3):
    """The quietest moment within a little of the one given: the kindest place to cut."""
    at = frame_at(seconds)
    reach = max(1, int(round(within / FRAME_S)))
    first, last = max(0, at - reach), min(len(song.level), at + reach + 1)
    if last <= first:
        return seconds
    return frame_time(first + int(np.argmin(song.level[first:last])))


def _past_silence(song, seconds, step):
    """A cut placed in a gap is moved on to where the silence ends (3 s at most).

    step is +1 for the start of the song (move later), -1 for its end (move earlier).
    """
    at = frame_at(seconds)
    moved = 0
    while 0 <= at + step < len(song.level) and song.level[at] < SILENT_DB and moved < int(3 / FRAME_S):
        at += step
        moved += 1
    if not moved:
        return seconds
    # every frame passed over was silent from its first sample to its last, so the
    # cut can go right up to the edge of the last of them
    if step > 0:
        return ((at - 1) * HOP + WINDOW) / RATE
    return (at + 1) * HOP / RATE


def _one_end(song, side, gap, change, copy_edge):
    """(seconds or None, sure, reason) for one end. side is "start" or "end"."""
    def close(a, b):
        return a is not None and b is not None and abs(a - b) <= AGREE_S

    if close(gap, change):
        return gap, True, "a gap, and the sound changes there"
    if copy_edge is not None and close(copy_edge, gap):
        return gap, True, "a gap, where another copy shows the song " + ("starts" if side == "start" else "ends")
    if copy_edge is not None and close(copy_edge, change):
        return change, True, "the sound changes where another copy shows the song " + (
            "starts" if side == "start" else "ends")
    if copy_edge is not None:
        extra = copy_edge if side == "start" else song.length - copy_edge
        if extra >= LISTEN_MIN_S:
            return copy_edge, False, "another copy has nothing here (it may be talk, or more of the song)"
    # One clue alone is not offered even as a guess. Measured on 552 real songs
    # whose ends were known from copies: a gap alone cut into the song 68% of
    # the time and a change alone 30-50%, however strong; two together, never.
    return None, False, ""


def propose_trim(song, copy_ends=None, known_length=None):
    """Where to cut this song, or None when there is nothing worth cutting.

    copy_ends: (start, end) in seconds as shown by other copies, or None.
    known_length: the real recording's length in seconds when looked up online.
    """
    if song.fp is None:
        return None
    length = song.length
    start_gap, start_change, end_gap, end_change = ends_from_sound(song)
    copy_start = copy_end = None
    if copy_ends:
        if copy_ends[0] >= MIN_TRIM_S:
            copy_start = copy_ends[0] + WINDOW / 2 / RATE
        if length - copy_ends[1] >= MIN_TRIM_S:
            copy_end = copy_ends[1]
    start, start_sure, start_why = _one_end(song, "start", start_gap, start_change, copy_start)
    end, end_sure, end_why = _one_end(song, "end", end_gap, end_change, copy_end)

    # the real length, when known, can confirm a guess or expose a bad one
    if known_length and known_length > 30:
        extra = length - known_length
        guess = (start or 0.0) + (length - end if end is not None else 0.0)
        if extra < MIN_TRIM_S and guess > 2.0:
            start_sure = end_sure = False                 # nothing to cut, the lookup says
            start_why = end_why = "the song's known length says there is nothing extra"
        elif abs(extra - guess) <= 1.5 and guess >= MIN_TRIM_S:
            if start is not None and not start_sure:
                start_sure, start_why = True, start_why + ", and the song's known length agrees"
            if end is not None and not end_sure:
                end_sure, end_why = True, end_why + ", and the song's known length agrees"

    silent_lead, silent_tail = silent_ends(song)
    # silence is cut whatever else is guessed: a guess must not hide it
    if not start_sure and silent_lead >= MIN_TRIM_S:
        start, start_sure, start_why = silent_lead, True, "silence"
    if not end_sure and silent_tail >= MIN_TRIM_S:
        end, end_sure, end_why = length - silent_tail, True, "silence"

    if start is not None:
        start = _past_silence(song, quiet_point(song, start), +1) if start_why != "silence" else start
        if start < MIN_TRIM_S:
            start = None
    if end is not None:
        end = _past_silence(song, quiet_point(song, end), -1) if end_why != "silence" else end
        if length - end < MIN_TRIM_S:
            end = None
    if start is None and end is None:
        return None
    if start is not None:
        start_sure = start_sure and start <= MAX_TRIM_S
    if end is not None:
        end_sure = end_sure and length - end <= MAX_TRIM_S
    # One end sure and the other a guess: propose the sure end alone. Otherwise
    # the sure cut would wait on the guess, and show up only on a later scan.
    if start is not None and end is not None and start_sure != end_sure:
        if start_sure:
            end = None
        else:
            start = None
    sure = True
    reasons = []
    if start is not None:
        sure = sure and start_sure
        reasons.append("Start: " + start_why)
    if end is not None:
        sure = sure and end_sure
        reasons.append("End: " + end_why)
    start = start or 0.0
    end = end if end is not None else length
    if end - start < 30.0 or end - start < length * 0.5:
        return None                                       # that would not leave a song
    return Trim(song.path, float(start), float(end), float(length), bool(sure), ". ".join(reasons))


def cut_mp3(src, dst, start, end):
    """Write the part of an MP3 between two times to a new file, tags and all.

    The MP3's frames are copied as they are, so the cut lands on a frame edge
    (within 26 ms) and the sound is not changed. Raises SongError, and leaves
    no file behind, if the result is not right.
    """
    import mutagen.id3
    import mutagen.mp3
    wanted = end - start
    if wanted <= 0:
        raise SongError("There would be nothing left of this song.")
    try:
        done = subprocess.run(
            [ffmpeg_exe(), "-v", "error", "-nostdin", "-y", "-i", src, "-ss", f"{start:.3f}",
             "-t", f"{wanted:.3f}", "-map", "0:a:0", "-c", "copy", "-map_metadata", "-1",
             "-id3v2_version", "0", "-f", "mp3", dst],
            capture_output=True, creationflags=NO_WINDOW)
        if done.returncode != 0 or not os.path.exists(dst):
            raise SongError("FFmpeg could not cut this song.")
        try:
            tags = mutagen.id3.ID3(src)
        except mutagen.id3.ID3NoHeaderError:
            tags = None
        if tags is not None:
            tags.save(dst, v2_version=3 if tags.version[1] == 3 else 4)
        length = mutagen.mp3.MP3(dst).info.length
        if abs(length - wanted) > 0.2:
            raise SongError("The cut song came out the wrong length, so it was not used.")
        heard = len(decode_mono(dst)) / RATE
        if abs(heard - wanted) > 0.5:
            raise SongError("The cut song did not play right, so it was not used.")
    except (OSError, SongError) as error:
        try:
            os.remove(dst)
        except OSError:
            pass
        if isinstance(error, SongError):
            raise
        raise SongError(f"Windows would not let this song be cut ({error.strerror or error}).") from error
    except Exception:
        try:
            os.remove(dst)
        except OSError:
            pass
        raise
