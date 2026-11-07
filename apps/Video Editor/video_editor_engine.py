"""
video_editor_engine.py

The editing rules behind video_editor.py, kept free of windows so they can be
tested on their own:

  Piece ........ a stretch of one video or sound file, with its own sound
                 settings (mute, volume, fade in, fade out)
  AddedSound ... music or narration laid over the edit at a chosen time,
                 mixed with the pieces' own sound or replacing it
  Edit ......... the pieces in playing order and the added sounds, with
                 split, insert, delete, keep, copy, trim and move
  probe() ...... reads a file's length, picture size, frame rate and sound
  plan_export()  the ffmpeg command that saves an edit as a new file

Nothing here changes the files that were opened; saving always writes a new file.
"""

from __future__ import annotations

import copy
import os
import re
import subprocess
from dataclasses import dataclass, field, replace

MIN_PIECE = 0.05        # seconds; a split this close to a piece's edge snaps to the edge
SAMPLE_RATE = 48000

VIDEO_FORMATS = {       # extension: (name shown in the Save window, ffmpeg settings)
    ".mp4": ("MP4 video, plays almost anywhere",
             ["-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
              "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart"]),
    ".mov": ("MOV video, for Apple devices",
             ["-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
              "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart"]),
    ".m4v": ("M4V video",
             ["-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
              "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart"]),
    ".mkv": ("MKV video",
             ["-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
              "-c:a", "aac", "-b:a", "192k"]),
    ".webm": ("WebM video, for web pages",
              ["-c:v", "libvpx-vp9", "-crf", "32", "-b:v", "0", "-row-mt", "1",
               "-deadline", "good", "-cpu-used", "4", "-c:a", "libopus", "-b:a", "160k"]),
    ".avi": ("AVI video, for older players",
             ["-c:v", "mpeg4", "-q:v", "3", "-c:a", "libmp3lame", "-q:a", "2"]),
    ".wmv": ("WMV video, Windows Media",
             ["-c:v", "wmv2", "-b:v", "6M", "-c:a", "wmav2", "-b:a", "192k"]),
    ".mpg": ("MPG video, DVD style",
             ["-c:v", "mpeg2video", "-q:v", "3", "-c:a", "mp2", "-b:a", "224k"]),
}
SOUND_FORMATS = {
    ".mp3": ("MP3 sound", ["-c:a", "libmp3lame", "-q:a", "2"]),
    ".m4a": ("M4A sound (AAC)", ["-c:a", "aac", "-b:a", "192k"]),
    ".wav": ("WAV sound, uncompressed", ["-c:a", "pcm_s16le"]),
    ".flac": ("FLAC sound, full quality", ["-c:a", "flac"]),
    ".ogg": ("OGG sound (Vorbis)", ["-c:a", "libvorbis", "-q:a", "5"]),
    ".opus": ("Opus sound", ["-c:a", "libopus", "-b:a", "128k"]),
    ".wma": ("WMA sound, Windows Media", ["-c:a", "wmav2", "-b:a", "192k"]),
}
# MPEG-2 (.mpg) only allows these frame rates.
MPEG2_RATES = ["24000/1001", "24", "25", "30000/1001", "30", "50", "60000/1001", "60"]


# ------------------------------------------------------------------ pieces

@dataclass
class Piece:
    """A stretch of one file, from start to end (seconds into the file)."""
    source: str
    start: float
    end: float
    source_length: float
    has_video: bool = True
    has_audio: bool = True
    width: int = 0
    height: int = 0
    fps: float = 0.0
    video_stream: int = -1       # which stream of the file holds the picture / the sound
    audio_stream: int = -1
    volume: float = 1.0          # 1.0 is as recorded; up to 2.0
    muted: bool = False
    fade_in: float = 0.0
    fade_out: float = 0.0

    @property
    def length(self):
        return self.end - self.start

    @property
    def silent(self):
        return self.muted or not self.has_audio or self.volume <= 0


@dataclass
class AddedSound:
    """Music or narration: start to end of a file, playing from `at` seconds into the edit."""
    source: str
    start: float
    end: float
    at: float
    source_length: float
    audio_stream: int = -1
    volume: float = 1.0
    fade_in: float = 0.0
    fade_out: float = 0.0
    replace: bool = False        # silence the pieces' own sound while this plays

    @property
    def length(self):
        return self.end - self.start

    @property
    def finish(self):
        return self.at + self.length


def piece_from(info, start=0.0, end=None):
    end = info.duration if end is None else end
    return Piece(info.path, start, end, info.duration, info.has_video, info.has_audio,
                 info.width, info.height, info.fps, info.video_stream, info.audio_stream)


def sound_from(info, at):
    if not info.has_audio:
        raise ProbeError(f"{os.path.basename(info.path)} has no sound to add.")
    return AddedSound(info.path, 0.0, info.duration, at, info.duration, info.audio_stream)


def fade_gain(offset, length, fade_in, fade_out):
    """How loud (0 to 1) a fading stretch is at `offset` seconds into it."""
    gain = 1.0
    if fade_in > 0:
        gain = min(gain, max(0.0, offset / fade_in))
    if fade_out > 0:
        gain = min(gain, max(0.0, (length - offset) / fade_out))
    return gain


# -------------------------------------------------------------------- edit

class Edit:
    """The pieces in playing order, and the sounds added over them."""

    def __init__(self, pieces=None, sounds=None):
        self.pieces = list(pieces or [])
        self.sounds = list(sounds or [])

    # ------------------------------------------------------------ reading
    @property
    def duration(self):
        return sum(p.length for p in self.pieces)

    def starts(self):
        """Where each piece starts in the edit."""
        out, total = [], 0.0
        for p in self.pieces:
            out.append(total)
            total += p.length
        return out

    def has_video(self):
        return any(p.has_video for p in self.pieces)

    def locate(self, t):
        """(piece index, seconds into it) for a time in the edit. A time on the
        line between two pieces belongs to the later one; the very end belongs
        to the last piece. (-1, 0) when there are no pieces."""
        if not self.pieces:
            return -1, 0.0
        t = min(max(0.0, t), self.duration)
        for i, start in enumerate(self.starts()):
            if t < start + self.pieces[i].length:
                return i, t - start
        return len(self.pieces) - 1, self.pieces[-1].length

    def first_video(self):
        return next((p for p in self.pieces if p.has_video), None)

    def snapshot(self):
        return Edit(copy.deepcopy(self.pieces), copy.deepcopy(self.sounds))

    # ------------------------------------------------------------ changing
    def split(self, t):
        """Cut the piece under time t in two. Returns the index of the piece
        that now starts at t (len(pieces) at the end). Within MIN_PIECE of an
        edge, the existing edge is used instead of making a sliver."""
        if t <= MIN_PIECE / 2:
            return 0
        start = 0.0
        for i, p in enumerate(self.pieces):
            inside = t - start
            if inside < MIN_PIECE:
                return i
            if inside <= p.length - MIN_PIECE:
                cut = p.start + inside
                left = replace(p, end=cut, fade_out=0.0)
                right = replace(p, start=cut, fade_in=0.0)
                self.pieces[i:i + 1] = [left, right]
                return i + 1
            start += p.length
        return len(self.pieces)

    def insert(self, t, pieces):
        """Put copies of pieces in at time t. Returns the index of the first one."""
        i = self.split(t)
        self.pieces[i:i] = copy.deepcopy(list(pieces))
        return i

    def append(self, pieces):
        self.pieces.extend(copy.deepcopy(list(pieces)))

    def delete_range(self, a, b):
        """Remove everything between times a and b and join what is left."""
        a, b = sorted((max(0.0, a), min(self.duration, b)))
        if b - a <= 0:
            return
        i = self.split(a)
        j = self.split(b)
        del self.pieces[i:j]

    def keep_range(self, a, b):
        """Keep only what is between times a and b."""
        a, b = sorted((a, b))
        self.delete_range(b, self.duration)
        self.delete_range(0, a)

    def copy_range(self, a, b):
        """Copies of the pieces between times a and b, cut to fit."""
        part = self.snapshot()
        part.keep_range(a, b)
        return part.pieces

    def remove(self, i):
        del self.pieces[i]

    def move(self, i, step):
        """Swap piece i with its neighbour. Returns its new index."""
        j = i + step
        if not (0 <= i < len(self.pieces) and 0 <= j < len(self.pieces)):
            return i
        self.pieces[i], self.pieces[j] = self.pieces[j], self.pieces[i]
        return j

    def set_trim(self, i, start, end):
        """Set where piece i starts and ends in its file, kept inside the file
        and at least MIN_PIECE long. Fades shrink to fit."""
        p = self.pieces[i]
        start = min(max(0.0, start), p.source_length - MIN_PIECE)
        end = min(max(start + MIN_PIECE, end), p.source_length)
        p.start, p.end = start, end
        fit_fades(p)

    def trim_to_playhead(self, t, keep_after):
        """Trim the piece under t so it starts (keep_after) or ends at t.
        Returns the piece index, or -1 when t is on an edge already."""
        i, inside = self.locate(t)
        if i < 0:
            return -1
        p = self.pieces[i]
        if keep_after:
            if inside < MIN_PIECE:
                return -1
            self.set_trim(i, p.start + inside, p.end)
        else:
            if p.length - inside < MIN_PIECE:
                return -1
            self.set_trim(i, p.start, p.start + inside)
        return i


def fit_fades(item):
    """Keep fades inside the stretch they belong to."""
    item.fade_in = min(max(0.0, item.fade_in), item.length)
    item.fade_out = min(max(0.0, item.fade_out), item.length - item.fade_in)


# ------------------------------------------------------------------- probe

class ProbeError(ValueError):
    pass


@dataclass
class MediaInfo:
    path: str
    duration: float
    has_video: bool
    has_audio: bool
    width: int = 0
    height: int = 0
    fps: float = 0.0
    video_stream: int = -1
    audio_stream: int = -1


def parse_probe(text, path):
    """Read what `ffmpeg -i file` printed about a file."""
    found = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", text)
    if not found:
        raise ProbeError(f"Can't tell how long {os.path.basename(path)} is.")
    h, m, s = found.groups()
    duration = int(h) * 3600 + int(m) * 60 + float(s)
    if duration <= 0:
        raise ProbeError(f"{os.path.basename(path)} is empty.")
    video = audio = None
    video_index = audio_index = -1
    for line in text.splitlines():
        line_match = re.search(r"Stream #\d+:(\d+)[^:]*: (Video|Audio): (.*)", line)
        if not line_match:
            continue
        index, kind, rest = int(line_match.group(1)), line_match.group(2), line_match.group(3)
        if kind == "Video" and video is None and "(attached pic)" not in rest:
            video, video_index = rest, index  # cover pictures don't count as video
        elif kind == "Audio" and audio is None:
            audio, audio_index = rest, index
    if video is None and audio is None:
        raise ProbeError(f"{os.path.basename(path)} has no picture or sound that can be used.")
    info = MediaInfo(path, duration, video is not None, audio is not None,
                     video_stream=video_index, audio_stream=audio_index)
    if video is not None:
        size = re.search(r"(?<![\dx])(\d{2,5})x(\d{2,5})(?![\dx])", video)
        if size:
            info.width, info.height = int(size.group(1)), int(size.group(2))
        shape = re.search(r"SAR (\d+):(\d+) DAR (\d+):(\d+)", video)
        if shape and info.height and shape.group(1) != shape.group(2):
            dar = int(shape.group(3)) / int(shape.group(4))       # stored squeezed:
            info.width = int(round(info.height * dar))           # use the shape it is shown at
        rate = re.search(r"(\d+(?:\.\d+)?) fps", video) or re.search(r"(\d+(?:\.\d+)?) tbr", video)
        if rate:
            info.fps = float(rate.group(1))
        turned = re.search(r"rotation of (-?\d+(?:\.\d+)?) degrees", text)
        if turned and round(abs(float(turned.group(1)))) % 180 == 90:
            info.width, info.height = info.height, info.width   # phone videos filmed upright
    return info


def probe(path, ffmpeg):
    """Length, picture size, frame rate and sound of a file, via ffmpeg."""
    result = subprocess.run([ffmpeg, "-hide_banner", "-i", path], capture_output=True,
                            text=True, encoding="utf-8", errors="replace",
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return parse_probe(result.stderr, os.path.abspath(path))


# ------------------------------------------------------------------ export

def rate_text(fps, for_mpeg2=False):
    """A frame rate as ffmpeg wants it: 29.97 becomes 30000/1001."""
    fps = fps if 1 <= fps <= 240 else 30.0
    fps = min(fps, 60.0)
    known = {"24000/1001": 24000 / 1001, "30000/1001": 30000 / 1001, "60000/1001": 60000 / 1001}
    if for_mpeg2:
        return min(MPEG2_RATES, key=lambda r: abs(eval_rate(r) - fps))
    for text, value in known.items():
        if abs(fps - value) < 0.01:
            return text
    return f"{fps:g}" if fps == round(fps, 3) else f"{fps:.3f}"


def eval_rate(text):
    top, _, bottom = text.partition("/")
    return float(top) / float(bottom or 1)


def even(n):
    return max(2, int(n) - int(n) % 2)


def stream(k, index, kind):
    """The filter graph's name for a stream of input k."""
    return f"[{k}:{index}]" if index >= 0 else f"[{k}:{kind}:0]"


def seconds(x):
    return f"{max(0.0, x):.3f}"


def fades(item, length):
    parts = []
    if item.fade_in > 0:
        parts.append(f"afade=t=in:st=0:d={seconds(min(item.fade_in, length))}")
    if item.fade_out > 0:
        d = min(item.fade_out, length)
        parts.append(f"afade=t=out:st={seconds(length - d)}:d={seconds(d)}")
    return parts


SOUND_SHAPE = (f"aresample={SAMPLE_RATE},aformat=sample_fmts=fltp:"
               f"channel_layouts=stereo,asetpts=PTS-STARTPTS")


@dataclass
class ExportPlan:
    inputs: list
    graph: str
    output: list
    duration: float
    sources: list = field(default_factory=list)

    def command(self, ffmpeg, graph_file=None):
        """The ffmpeg command. A long filter graph is passed in a file instead,
        because Windows limits how long a command can be."""
        graph = ["-/filter_complex", graph_file] if graph_file else ["-filter_complex", self.graph]
        return [ffmpeg, "-hide_banner", "-y", *self.inputs, *graph, *self.output]


def same_path(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def plan_export(edit, out_path):
    """Work out how ffmpeg should save the edit to out_path. The extension
    picks the format; a sound extension saves the sound alone."""
    ext = os.path.splitext(out_path)[1].lower()
    if ext in VIDEO_FORMATS:
        with_picture, settings = True, VIDEO_FORMATS[ext][1]
    elif ext in SOUND_FORMATS:
        with_picture, settings = False, SOUND_FORMATS[ext][1]
    else:
        raise ValueError(f"Can't save as {ext or 'a file without an extension'}.")
    if not edit.pieces:
        raise ValueError("There is nothing to save yet.")
    if with_picture and not edit.has_video():
        raise ValueError("This edit is sound only. Save it as a sound file, such as MP3.")
    used = [p.source for p in edit.pieces] + [s.source for s in edit.sounds]
    if any(same_path(out_path, s) for s in used):
        raise ValueError("That would write over a file the edit uses. Pick another name.")

    first = edit.first_video()
    width, height = (even(first.width or 1280), even(first.height or 720)) if first else (0, 0)
    rate = rate_text(first.fps if first else 30.0, for_mpeg2=(ext == ".mpg"))
    total = edit.duration

    inputs, chains, joins = [], [], []
    for k, p in enumerate(edit.pieces):
        length = seconds(p.length)
        inputs += ["-ss", seconds(p.start), "-t", length, "-i", p.source]
        if with_picture:
            if p.has_video:
                chains.append(
                    f"{stream(k, p.video_stream, 'v')}scale=trunc(iw*sar/2)*2:ih,setsar=1,"
                    f"scale={width}:{height}:force_original_aspect_ratio=decrease,"
                    f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2:color=black,setsar=1,"
                    f"fps={rate},format=yuv420p,setpts=PTS-STARTPTS,"
                    f"tpad=stop_mode=clone:stop_duration={length},trim=duration={length}[v{k}]")
            else:                                   # a sound file in a video: black picture
                chains.append(f"color=c=black:s={width}x{height}:r={rate}:d={length},"
                              f"format=yuv420p,setsar=1[v{k}]")
        if p.silent:
            chains.append(f"anullsrc=r={SAMPLE_RATE}:cl=stereo,atrim=0:{length},"
                          f"aformat=sample_fmts=fltp:channel_layouts=stereo[a{k}]")
        else:
            steps = [f"{stream(k, p.audio_stream, 'a')}{SOUND_SHAPE}", f"volume={p.volume:g}",
                     *fades(p, p.length),
                     f"apad=whole_dur={length}", f"atrim=0:{length}"]
            chains.append(",".join(steps) + f"[a{k}]")
        joins.append(f"[v{k}][a{k}]" if with_picture else f"[a{k}]")
    n = len(edit.pieces)
    if with_picture:
        chains.append("".join(joins) + f"concat=n={n}:v=1:a=1[vout][amain]")
    else:
        chains.append("".join(joins) + f"concat=n={n}:v=0:a=1[amain]")

    # Added sounds: cut to the end of the edit, placed, and mixed in.
    main, added = "amain", []
    playing = [s for s in edit.sounds if s.at < total - 0.01]
    quiet = [s for s in playing if s.replace]
    if quiet:
        spans = ",".join(f"volume=0:enable='between(t,{seconds(s.at)},"
                         f"{seconds(min(s.finish, total))})'" for s in quiet)
        chains.append(f"[amain]{spans}[aquiet]")
        main = "aquiet"
    for j, s in enumerate(playing):
        k = n + j
        length = min(s.length, total - s.at)
        inputs += ["-ss", seconds(s.start), "-t", seconds(length), "-i", s.source]
        delay = int(round(s.at * 1000))
        steps = [f"{stream(k, s.audio_stream, 'a')}{SOUND_SHAPE}", f"volume={s.volume:g}",
                 *fades(s, length),
                 f"adelay=delays={delay}:all=1"]
        chains.append(",".join(steps) + f"[s{j}]")
        added.append(f"[s{j}]")
    if added:
        chains.append(f"[{main}]{''.join(added)}amix=inputs={1 + len(added)}:duration=first:"
                      f"normalize=0:dropout_transition=0,"
                      f"alimiter=limit=0.97:level=disabled[aout]")
    else:
        chains.append(f"[{main}]anull[aout]")

    output = (["-map", "[vout]"] if with_picture else []) + ["-map", "[aout]"]
    output += ["-ar", str(SAMPLE_RATE), *settings, "-progress", "pipe:1", "-nostats", out_path]
    return ExportPlan(inputs, ";".join(chains), output, total, used)


def progress_seconds(line):
    """Seconds written so far, from one line of ffmpeg's -progress output."""
    if line.startswith("out_time_us="):
        value = line.split("=", 1)[1].strip()
        if value.isdigit():
            return int(value) / 1_000_000
    return None
