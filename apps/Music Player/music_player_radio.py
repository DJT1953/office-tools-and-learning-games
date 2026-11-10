"""
music_player_radio.py

Internet radio for the Music Player, kept free of Qt so it can be tested on
its own:

  - RadioDirectory looks stations up in Radio Browser (radio-browser.info), a
    free, open list of about 50,000 stations run by volunteers.
  - RadioRecorder brings a station in: it reads the stream, takes out the song
    names many stations send along with the sound, and has FFmpeg turn the
    sound into plain 16-bit stereo.
  - TimeshiftBuffer keeps the last hour of that sound in a temporary file,
    reused round and round, which is what lets live radio be paused, rewound,
    replayed and repeated while the station carries on arriving.
"""

import _socket
import http.client
import json
import random
import re
import socket
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass, fields

USER_AGENT = "MusicPlayer/1.0"        # Radio Browser asks programs to name themselves
DIRECTORY_FALLBACK = ("de1.api.radio-browser.info", "de2.api.radio-browser.info",
                      "fi1.api.radio-browser.info")
TIMEOUT = 12
BUFFER_MINUTES = 60
CHUNK = 8192
RECONNECT_WAITS = (1, 3, 6, 10, 15)   # seconds between tries after a station drops out


# ------------------------------------------------------------------ stations

@dataclass
class Station:
    name: str
    url: str
    homepage: str = ""
    favicon: str = ""
    tags: str = ""
    country: str = ""
    countrycode: str = ""
    codec: str = ""
    bitrate: int = 0
    uuid: str = ""
    votes: int = 0
    clicks: int = 0

    @classmethod
    def from_directory(cls, data):
        return cls(
            name=(data.get("name") or "").strip() or "Unnamed station",
            url=data.get("url_resolved") or data.get("url") or "",
            homepage=data.get("homepage") or "",
            favicon=data.get("favicon") or "",
            tags=", ".join(t.strip() for t in (data.get("tags") or "").split(",") if t.strip()),
            country=data.get("country") or "",
            countrycode=data.get("countrycode") or "",
            codec=data.get("codec") or "",
            bitrate=int(data.get("bitrate") or 0),
            uuid=data.get("stationuuid") or "",
            votes=int(data.get("votes") or 0),
            clicks=int(data.get("clickcount") or 0),
        )

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, data):
        """A saved station. Damaged entries are mended where they can be;
        one with no web address raises ValueError."""
        url = data.get("url") if isinstance(data, dict) else None
        if not isinstance(url, str) or not url.strip():
            raise ValueError("a saved station without a web address")
        values = {}
        for f in fields(cls):
            value = data.get(f.name)
            if f.type is int:
                try:
                    values[f.name] = int(value or 0)
                except (TypeError, ValueError, OverflowError):
                    values[f.name] = 0
            else:
                values[f.name] = value.strip() if isinstance(value, str) else ""
        values["name"] = values["name"] or values["url"]
        return cls(**values)

    @property
    def quality(self):
        parts = [self.codec.upper()] if self.codec and self.codec.upper() != "UNKNOWN" else []
        if self.bitrate:
            parts.append(f"{self.bitrate} kbps")
        return " ".join(parts)


class RadioDirectory:
    """Searches Radio Browser. Its servers come and go, so the list of them is
    asked for first and the next one is tried if one fails."""

    def __init__(self, opener=None):
        self.opener = opener or urllib.request.urlopen
        self.servers = None

    def _servers(self):
        if self.servers is None:
            try:
                found = self._fetch("https://all.api.radio-browser.info/json/servers")
                names = sorted({s["name"] for s in found if s.get("name")})
                random.shuffle(names)      # share the load, as Radio Browser asks
                self.servers = names or list(DIRECTORY_FALLBACK)
            except Exception:
                self.servers = list(DIRECTORY_FALLBACK)
        return self.servers

    def _fetch(self, url):
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with self.opener(request, timeout=TIMEOUT) as response:
            return json.loads(response.read().decode("utf-8"))

    def _get(self, path, **params):
        query = urllib.parse.urlencode({k: v for k, v in params.items() if v not in ("", None)})
        problem = None
        for server in list(self._servers()):
            try:
                return self._fetch(f"https://{server}{path}" + (f"?{query}" if query else ""))
            except Exception as error:
                problem = error
                try:                            # try it last next time
                    self.servers.remove(server)
                    self.servers.append(server)
                except ValueError:              # another search moved it already
                    pass
        raise ConnectionError("The radio station list couldn't be reached. "
                              "Check the internet connection and try again.") from problem

    def _stations(self, found):
        stations, seen = [], set()
        for data in found:
            station = Station.from_directory(data)
            key = station.uuid or station.url
            if station.url and key not in seen:
                seen.add(key)
                stations.append(station)
        return stations

    def search(self, text, countrycode="", limit=100):
        """Stations whose name or type of music matches, most listened first."""
        text = text.strip()
        common = dict(countrycode=countrycode, order="clickcount", reverse="true",
                      limit=limit, hidebroken="true")
        by_name = self._get("/json/stations/search", name=text, **common)
        by_type = self._get("/json/stations/search", tag=text, **common) if text else []
        stations = self._stations(by_name + by_type)
        stations.sort(key=lambda s: s.clicks, reverse=True)
        return stations[:limit]

    def popular(self, countrycode="", limit=100):
        return self.search("", countrycode, limit)

    def countries(self):
        """(name, two-letter code) for every country with stations, by name."""
        found = self._get("/json/countries", order="name", hidebroken="true")
        result = [(c.get("name", ""), c.get("iso_3166_1", "")) for c in found
                  if c.get("iso_3166_1") and int(c.get("stationcount") or 0) > 0]
        return sorted(result, key=lambda c: c[0].casefold())

    def count_click(self, station):
        """Tell Radio Browser the station was played, which keeps its
        "most listened" order honest. Failures don't matter."""
        if station.uuid:
            try:
                self._get(f"/json/url/{station.uuid}")
            except Exception:
                pass


# -------------------------------------------------------------- song names

def parse_stream_title(meta):
    """The song name from a station's StreamTitle='Artist - Song'; block."""
    raw = meta.rstrip(b"\x00")
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    match = re.search(r"StreamTitle='(.*?)';", text, re.DOTALL)
    if not match:
        match = re.search(r"StreamTitle='(.*)'", text, re.DOTALL)
    return match.group(1).strip() if match else None


class IcySplitter:
    """Stations that send song names put a small block of text into the sound
    every metaint bytes: one length byte (times 16), then the text."""

    def __init__(self, metaint):
        self.metaint = metaint
        self.audio_left = metaint
        self.meta_left = None          # None while reading sound
        self.meta = b""

    def feed(self, data):
        """Returns (sound bytes, [(sound bytes before it, song name), ...])."""
        sound, titles = [], []
        sent = 0
        i = 0
        while i < len(data):
            if self.meta_left is None:
                take = min(self.audio_left, len(data) - i)
                sound.append(data[i:i + take])
                sent += take
                i += take
                self.audio_left -= take
                if self.audio_left == 0:
                    self.meta_left = -1          # the length byte comes next
            elif self.meta_left == -1:
                self.meta_left = data[i] * 16
                i += 1
                self.meta = b""
                if self.meta_left == 0:
                    self.meta_left = None
                    self.audio_left = self.metaint
            else:
                take = min(self.meta_left, len(data) - i)
                self.meta += data[i:i + take]
                i += take
                self.meta_left -= take
                if self.meta_left == 0:
                    title = parse_stream_title(self.meta)
                    if title is not None:
                        titles.append((sent, title))
                    self.meta_left = None
                    self.audio_left = self.metaint
        return b"".join(sound), titles


# ---------------------------------------------------------- the time shift

class TimeshiftBuffer:
    """The last hour (or so) of a station as 16-bit stereo sound. Frames are
    counted from when the station started; once the file is full, the oldest
    sound is written over. Song changes are kept as marks."""

    def __init__(self, rate, channels=2, minutes=BUFFER_MINUTES, folder=None):
        self.rate = rate
        self.channels = channels
        self.frame_bytes = 2 * channels
        self.capacity = max(rate, int(minutes * 60 * rate))
        self.file = tempfile.TemporaryFile(prefix="music-player-radio-", dir=folder)
        self.lock = threading.Lock()
        self.written = 0                 # frames received so far: the live edge
        self.marks = []                  # (frame, song name, clock time)
        self.closed = False

    @property
    def oldest(self):
        return max(0, self.written - self.capacity)

    @property
    def seconds_kept(self):
        return (self.written - self.oldest) / self.rate

    def write(self, data):
        """Add sound; data must be whole frames."""
        with self.lock:
            if self.closed:
                return
            frames = len(data) // self.frame_bytes
            done = 0
            while done < frames:
                slot = self.written % self.capacity
                count = min(frames - done, self.capacity - slot)
                self.file.seek(slot * self.frame_bytes)
                self.file.write(data[done * self.frame_bytes:(done + count) * self.frame_bytes])
                done += count
                self.written += count
            oldest = self.oldest
            # Keep the last mark before the oldest sound: that song is still partly here.
            while len(self.marks) > 1 and self.marks[1][0] <= oldest:
                self.marks.pop(0)

    def read(self, frame, count):
        """Up to count frames from frame on (fewer near the live edge)."""
        with self.lock:
            if self.closed:
                return b""
            frame = max(frame, self.oldest)
            count = min(count, self.written - frame)
            parts = []
            while count > 0:
                slot = frame % self.capacity
                take = min(count, self.capacity - slot)
                self.file.seek(slot * self.frame_bytes)
                parts.append(self.file.read(take * self.frame_bytes))
                frame += take
                count -= take
            return b"".join(parts)

    def mark_song(self, title, frame=None, when=None):
        """A new song starts at frame (default: the live edge)."""
        with self.lock:
            if self.marks and self.marks[-1][1] == title:
                return
            frame = self.written if frame is None else max(int(frame), 0)
            if self.marks:
                frame = max(frame, self.marks[-1][0])          # marks stay in order
            self.marks.append((frame, title, time.time() if when is None else when))

    def song_at(self, frame):
        """(start frame, name) of the song playing at frame, or None."""
        with self.lock:
            found = None
            for mark in self.marks:
                if mark[0] <= frame:
                    found = mark
            if found is None:
                return None
            return max(found[0], self.oldest), found[1]

    def songs(self):
        with self.lock:
            oldest = self.oldest
            return [(max(frame, oldest), title, when) for frame, title, when in self.marks]

    def close(self):
        with self.lock:
            if not self.closed:
                self.closed = True
                self.file.close()            # a temporary file: Windows deletes it


# ------------------------------------------------------------ the recorder

def ffmpeg_path():
    import imageio_ffmpeg
    return imageio_ffmpeg.get_ffmpeg_exe()


class NotAStation(Exception):
    pass


def _close_quietly(response):
    try:
        response.close()
    except Exception:
        pass


def hang_up(response):
    """Close a station's connection without waiting. Closing it the usual way
    first waits for the read in progress to finish, which takes up to TIMEOUT
    seconds when the station has gone quiet. Closing the connection itself
    makes that read end now (on Windows only a real close does; shutdown()
    alone doesn't), and the rest of the closing happens on a thread of its own."""
    raw = getattr(getattr(response, "fp", None), "raw", None)
    connection = getattr(raw, "_sock", None)
    if isinstance(connection, socket.socket):
        try:
            connection.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            # urllib has already "closed" this socket object, which leaves the
            # real connection open for the response; this closes that.
            _socket.socket.close(connection)
        except OSError:
            pass
    threading.Thread(target=_close_quietly, args=(response,), name="radio-hang-up",
                     daemon=True).start()


class RadioRecorder:
    """Brings one station into a TimeshiftBuffer until stopped.

    on_event(kind, value) is called from the recorder's own threads with:
    "name" (what the station calls itself), "started" (sound is arriving),
    "song" (a new song name), "reconnecting", "error" (a message) and
    "ended"."""

    def __init__(self, url, buffer, on_event, ffmpeg=None, opener=None):
        self.url = url.strip()
        self.buffer = buffer
        self.on_event = on_event
        self.ffmpeg = ffmpeg or ffmpeg_path()
        self.opener = opener or urllib.request.urlopen
        self.stopping = threading.Event()
        self.process = None
        self.response = None
        self.lock = threading.Lock()
        self.started = False
        self.errors = []                 # FFmpeg's last few complaints
        self.thread = None
        self.readers = []
        self.last_song = None
        self.reconnect_waits = RECONNECT_WAITS
        # Song names arrive with the station's sound, before FFmpeg has turned
        # that sound into frames, so each is placed by how far into the
        # stream it came: bytes sent so far, at the station's bytes per second.
        self.bytes_fed = 0
        self.byte_rate = None
        self.pending_songs = []          # (bytes into the stream, song name)
        self.songs_lock = threading.Lock()

    # --- control
    def start(self):
        self.thread = threading.Thread(target=self._run, name="radio-in", daemon=True)
        self.thread.start()

    def stop(self):
        """Stop at once. Never waits on the station, so the window doesn't
        freeze when a station has gone quiet."""
        self.stopping.set()
        with self.lock:
            response, process = self.response, self.process
        if response is not None:
            hang_up(response)
        if process is not None:
            try:
                process.kill()
            except OSError:
                pass

    def join(self, timeout=None):
        if self.thread:
            self.thread.join(timeout)

    def _event(self, kind, value=None):
        if not self.stopping.is_set():
            try:
                self.on_event(kind, value)
            except Exception:
                pass

    # --- connecting
    def _open(self, url):
        request = urllib.request.Request(url, headers={
            "User-Agent": USER_AGENT, "Icy-MetaData": "1", "Accept": "*/*"})
        return self.opener(request, timeout=TIMEOUT)

    def _connect(self, url, hops=0):
        """Returns ("pipe", response) for a plain stream, or ("direct", url)
        for one FFmpeg must fetch itself (HLS, very old servers, and the
        older kinds of address: rtmp://, rtsp://, mms://)."""
        scheme = urllib.parse.urlparse(url).scheme.lower()
        if scheme == "icy":                                # Shoutcast's name for plain http
            url, scheme = "http" + url[3:], "http"
        if scheme == "mms":                                # FFmpeg reads these as mmsh
            return "direct", "mmsh" + url[3:]
        if scheme not in ("http", "https"):
            return "direct", url
        if urllib.parse.urlparse(url).path.lower().endswith(".m3u8"):
            return "direct", url
        try:
            response = self._open(url)
        except http.client.BadStatusLine:                  # "ICY 200 OK" servers
            return "direct", url
        except urllib.error.HTTPError as error:
            error.close()
            raise ConnectionError(f"The station's server said \"{error.code} {error.reason}\".")
        kind = (response.headers.get("Content-Type") or "").split(";")[0].strip().lower()
        path = urllib.parse.urlparse(response.geturl() or url).path.lower()
        if kind in ("application/vnd.apple.mpegurl", "application/x-mpegurl") or path.endswith(".m3u8"):
            response.close()
            return "direct", url
        playlist_kinds = ("audio/x-scpls", "audio/scpls", "audio/x-mpegurl", "audio/mpegurl",
                          "application/pls+xml", "application/x-scpls")
        if kind in playlist_kinds or path.endswith((".pls", ".m3u")) or kind.startswith("text/"):
            body = response.read(65536)
            response.close()
            text = body.decode("utf-8", "replace")
            if "#EXT-X-" in text:
                return "direct", url
            if kind == "text/html" or "<html" in text[:500].lower():
                raise NotAStation("That web address is a web page, not a radio stream. "
                                  "Look on the station's site for a 'listen' link.")
            entries = re.findall(r"(?im)^(?:file\d+\s*=\s*)?\s*((?:https?|icy)://\S+)", text)
            if not entries or hops >= 3:
                raise NotAStation("No radio stream was found at that web address.")
            return self._connect(entries[0].replace("icy://", "http://", 1), hops + 1)
        return "pipe", response

    # --- running
    def _start_ffmpeg(self, source):
        command = [self.ffmpeg, "-hide_banner", "-nostats", "-loglevel", "info"]
        if source == "pipe:0":
            command += ["-probesize", "131072", "-analyzeduration", "2000000"]
        else:
            command += ["-reconnect", "1", "-reconnect_streamed", "1",
                        "-reconnect_delay_max", "10", "-user_agent", USER_AGENT]
        command += ["-i", source, "-vn", "-ac", str(self.buffer.channels),
                    "-ar", str(self.buffer.rate), "-f", "s16le", "pipe:1"]
        process = subprocess.Popen(
            command, stdin=subprocess.PIPE if source == "pipe:0" else subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        with self.lock:
            self.process = process
        self.readers = [
            threading.Thread(target=self._store_sound, args=(process,), name="radio-sound",
                             daemon=True),
            threading.Thread(target=self._read_errors, args=(process,), name="radio-errors",
                             daemon=True)]
        for reader in self.readers:
            reader.start()
        if self.stopping.is_set():
            process.kill()
        return process

    def _end_process(self, process):
        """Stop FFmpeg and close its pipes, so nothing is left half-open."""
        try:
            process.kill()
        except OSError:
            pass
        try:
            process.wait(5)
        except subprocess.TimeoutExpired:
            pass
        for reader in self.readers:
            reader.join(2)
        for pipe in (process.stdin, process.stdout, process.stderr):
            if pipe is not None:
                try:
                    pipe.close()
                except (OSError, ValueError):
                    pass

    def _store_sound(self, process):
        leftover = b""
        frame_bytes = self.buffer.frame_bytes
        while True:
            try:
                data = process.stdout.read1(32768)
            except (OSError, ValueError):
                break
            if not data:
                break
            data = leftover + data
            whole = len(data) - len(data) % frame_bytes
            leftover = data[whole:]
            if whole:
                try:
                    self.buffer.write(data[:whole])
                except OSError as error:        # the disk is full, say
                    self._event("error", "The station's sound couldn't be kept on this "
                                         f"computer's disk ({error.strerror or error}).")
                    self.stop()
                    break
                self._place_songs()
                if not self.started:
                    self.started = True
                    self._event("started")

    def _read_errors(self, process):
        reading_input = True             # FFmpeg describes the station, then its own output
        for raw in iter(process.stderr.readline, b""):
            line = raw.decode("utf-8", "replace").strip()
            if line.startswith("Output #"):
                reading_input = False
            rate = re.search(r"Audio:.*?(\d+) kb/s", line)
            if rate and reading_input and self.byte_rate is None and line.startswith("Stream #0"):
                self.byte_rate = int(rate.group(1)) * 125
            if re.search(r"(?i)error|invalid|failed|refused|not found|denied", line):
                self.errors = (self.errors + [line])[-5:]

    def _place_songs(self):
        if not self.pending_songs:
            return
        rate = self.byte_rate
        if rate is None and self.buffer.written > self.buffer.rate * 10:
            # No bitrate given: estimate it from what has come through so far.
            rate = self.bytes_fed / (self.buffer.written / self.buffer.rate)
        if not rate:
            return
        with self.songs_lock:             # names arrive on one thread, sound on another
            songs, self.pending_songs = self.pending_songs, []
        for position, title in songs:
            self.buffer.mark_song(title, frame=position / rate * self.buffer.rate)

    def _run(self):
        try:
            mode, source = self._connect(self.url)
            if mode == "direct":
                self._run_direct(source)
            else:
                self._run_pipe(source)
        except NotAStation as error:
            self._event("error", str(error))
        except Exception as error:
            if not self.stopping.is_set():
                self._event("error", "The station couldn't be reached. It may be off the air "
                                     f"or the internet may be down.\n\n({error})")
        finally:
            with self.lock:
                process, self.process = self.process, None
                response, self.response = self.response, None
            if process is not None:
                self._end_process(process)
            if response is not None:
                try:
                    response.close()
                except Exception:
                    pass
            self._event("ended")

    def _run_direct(self, url):
        process = self._start_ffmpeg(url)
        process.wait()
        if not self.stopping.is_set():
            detail = f"\n\n({self.errors[-1]})" if self.errors else ""
            self._event("error", "The station stopped sending." + detail)

    def _run_pipe(self, response):
        with self.lock:
            self.response = response
        name = response.headers.get("icy-name")
        if name:
            self._event("name", name.strip())
        bitrate = (response.headers.get("icy-br") or "").split(",")[0].strip()
        if bitrate.isdigit() and 8 <= int(bitrate) <= 1500:
            self.byte_rate = int(bitrate) * 125
        process = self._start_ffmpeg("pipe:0")
        tries = 0
        while not self.stopping.is_set():
            metaint = int(response.headers.get("icy-metaint") or 0)
            splitter = IcySplitter(metaint) if metaint > 0 else None
            while not self.stopping.is_set():
                try:
                    data = response.read1(CHUNK) if hasattr(response, "read1") else response.read(CHUNK)
                except Exception:
                    data = b""
                if not data:
                    break
                tries = 0
                if splitter:
                    data, titles = splitter.feed(data)
                    for offset, title in titles:
                        if title and title != self.last_song:
                            self.last_song = title
                            with self.songs_lock:
                                self.pending_songs.append((self.bytes_fed + offset, title))
                            self._place_songs()
                            self._event("song", title)
                if process.poll() is not None:
                    raise RuntimeError(self.errors[-1] if self.errors else "FFmpeg stopped")
                try:
                    process.stdin.write(data)
                    process.stdin.flush()             # small pieces go straight on, too
                    self.bytes_fed += len(data)
                except (OSError, ValueError):
                    if self.stopping.is_set():
                        return
                    raise RuntimeError(self.errors[-1] if self.errors else "FFmpeg stopped")
            if self.stopping.is_set():
                return
            # Dropped out: try again a few times, carrying on into the same FFmpeg.
            try:
                response.close()
            except Exception:
                pass
            if tries >= len(self.reconnect_waits):
                # Let FFmpeg finish the sound it already has before giving up.
                try:
                    process.stdin.close()
                    process.wait(10)
                except (OSError, ValueError, subprocess.TimeoutExpired):
                    pass
                self._event("error", "The station stopped sending and didn't come back.")
                return
            self._event("reconnecting")
            if self.stopping.wait(self.reconnect_waits[tries]):
                return
            tries += 1
            try:
                mode, response = self._connect(self.url)
            except Exception:
                continue
            if mode != "pipe":
                with self.lock:
                    self.process = None
                self._end_process(process)
                return self._run_direct(response)
            with self.lock:
                self.response = response
