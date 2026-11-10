"""
music_player_spotify.py

Spotify for the Music Player, kept free of Qt so it can be tested on its own.

Spotify doesn't let other programs play its music, so the player works the
Spotify app instead:
  - saved playlist links open in the Spotify app on this PC;
  - Play, Pause, Next, Previous, the song name and picture go through Windows'
    own media controls (the same ones the volume pop-up uses);
  - volume and mute set Spotify's own slider in the Windows volume mixer.
Nothing here needs a Spotify account setting or developer sign-up.
"""

import asyncio
import datetime
import json
import os
import re
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import dataclass

USER_AGENT = "MusicPlayer/1.0"
TIMEOUT = 10
KINDS = ("playlist", "album", "track", "artist", "show", "episode")
KIND_WORDS = {"playlist": "Playlist", "album": "Album", "track": "Song", "artist": "Artist",
              "show": "Podcast", "episode": "Podcast episode"}
LINK = re.compile(
    r"(?:open\.spotify\.com/(?:intl-[\w-]+/)?(?:embed/)?(?:user/[^/\s]+/)?"
    r"|spotify:(?:user:[^:\s]+:)?)"
    r"(" + "|".join(KINDS) + r")[/:]([A-Za-z0-9]{22})", re.IGNORECASE)
SHORT_LINK = re.compile(r"https?://(?:spotify\.link|spotify\.app\.link)/\S+", re.IGNORECASE)


@dataclass
class SpotifyLink:
    kind: str
    id: str
    name: str = ""

    @property
    def uri(self):
        return f"spotify:{self.kind}:{self.id}"

    @property
    def web_url(self):
        return f"https://open.spotify.com/{self.kind}/{self.id}"

    @property
    def label(self):
        return self.name or f"Spotify {KIND_WORDS.get(self.kind, self.kind).lower()}"

    def to_dict(self):
        return {"kind": self.kind, "id": self.id, "name": self.name}

    @classmethod
    def from_dict(cls, data):
        """A saved link; ValueError if it's too damaged to open."""
        data = data if isinstance(data, dict) else {}
        kind, link_id, name = data.get("kind"), data.get("id"), data.get("name")
        if not isinstance(link_id, str) or not re.fullmatch(r"[A-Za-z0-9]{22}", link_id):
            raise ValueError("a saved Spotify link without a proper id")
        kind = kind.lower() if isinstance(kind, str) and kind.lower() in KINDS else "playlist"
        return cls(kind, link_id, name.strip() if isinstance(name, str) else "")


def parse_link(text):
    """A Spotify link or URI, as Spotify's Share > Copy link gives it."""
    match = LINK.search(text or "")
    if not match:
        return None
    return SpotifyLink(match.group(1).lower(), match.group(2))


def _open(url, opener):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    return opener(request, timeout=TIMEOUT)


def resolve_link(text, opener=None):
    """parse_link, also following the short spotify.link links the phone app
    shares. Returns None if no Spotify link is there."""
    found = parse_link(text)
    if found or not text:
        return found
    short = SHORT_LINK.search(text)
    if not short:
        return None
    try:
        with _open(short.group(0), opener or urllib.request.urlopen) as response:
            found = parse_link(response.geturl())
            if not found:
                found = parse_link(response.read(200000).decode("utf-8", "replace"))
    except Exception:
        return None
    return found


def look_up_name(link, opener=None):
    """The playlist's (or album's...) name from Spotify's public oEmbed page;
    None if it can't be reached."""
    url = "https://open.spotify.com/oembed?url=" + urllib.parse.quote(link.web_url, safe="")
    try:
        with _open(url, opener or urllib.request.urlopen) as response:
            data = json.loads(response.read().decode("utf-8"))
    except Exception:
        return None
    title = (data.get("title") or "").strip()
    return title or None


def app_installed(scheme="spotify"):
    """True if a program on this PC opens spotify: links. Asking Windows to
    open one with nothing installed shows its own "get an app" box instead
    of failing, so ask first."""
    try:
        import ctypes
        from ctypes import wintypes
        size = wintypes.DWORD(260)
        name = ctypes.create_unicode_buffer(260)
        ASSOCF_IS_PROTOCOL, ASSOCSTR_FRIENDLYAPPNAME = 0x1000, 4
        answer = ctypes.windll.shlwapi.AssocQueryStringW(
            ASSOCF_IS_PROTOCOL, ASSOCSTR_FRIENDLYAPPNAME, scheme, None, name, ctypes.byref(size))
        return answer == 0
    except (AttributeError, OSError):       # not Windows: just try it
        return True


def is_running(names=("spotify.exe",)):
    """True if the Spotify app is open on this PC (or if that can't be told,
    so no box offers to open it for nothing)."""
    try:
        import psutil                      # comes with pycaw
    except ImportError:
        return True
    for process in psutil.process_iter(["name"]):
        if (process.info.get("name") or "").lower() in names:
            return True
    return False


def open_in_spotify(link):
    """Open in the Spotify app; falls back to the web player in the browser
    when the app isn't installed. Returns "app" or "web"."""
    if app_installed():
        try:
            os.startfile(link.uri)
            return "app"
        except OSError:
            pass
    webbrowser.open(link.web_url)
    return "web"


def open_spotify_app():
    if not app_installed():
        return False
    try:
        os.startfile("spotify:")
        return True
    except OSError:
        return False


# ------------------------------------------------------------- now playing

@dataclass
class SpotifyNow:
    title: str = ""
    artist: str = ""
    album: str = ""
    playing: bool = False
    position: float = 0.0          # seconds, at the time this was read
    length: float = 0.0
    shuffle: object = None         # True, False, or None when Spotify doesn't say
    repeat: object = None          # "off", "all", "one" or None
    can_next: bool = True
    can_previous: bool = True
    picture: object = None         # bytes of the song's picture, when Spotify gives one
    read_at: float = 0.0           # time.monotonic() when read

    def position_now(self):
        if not self.playing:
            return self.position
        moved = self.position + time.monotonic() - self.read_at
        return min(moved, self.length) if self.length else moved

    @property
    def song_key(self):
        return (self.title, self.artist, self.album)


REPEAT_FROM_WINDOWS = {0: "off", 1: "one", 2: "all"}
REPEAT_TO_WINDOWS = {"off": 0, "one": 1, "all": 2}


def is_spotify_session(app_id):
    return "spotify" in (app_id or "").lower()


class SpotifyRemote:
    """Works Spotify through Windows' media controls on a thread of its own,
    so a slow answer never holds up the window. Callbacks run on that thread."""

    def __init__(self):
        self.loop = None
        self.thread = None
        self.manager = None
        self.picture_key = None
        self.picture = None
        self.ready = threading.Event()
        self.problem = None
        self.timeout = 6                 # seconds before giving up on an answer
        self.closed = False

    # --- the worker thread
    def _ensure(self):
        if self.closed:
            return False
        if self.thread is None:
            self.thread = threading.Thread(target=self._run, name="spotify", daemon=True)
            self.thread.start()
            self.ready.wait(5)
        return self.loop is not None

    def _run(self):
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.loop = loop
        finally:
            self.ready.set()
        try:
            loop.run_forever()
        finally:
            loop.close()

    def _submit(self, make, callback=None):
        if not self._ensure():
            if callback:
                callback(None)
            return

        async def guarded():
            try:
                return await asyncio.wait_for(make(), self.timeout)   # never wait for ever
            except Exception as error:      # Spotify closed mid-call, Windows said no...
                self.problem = error
                self.manager = None
                return None

        coroutine = guarded()
        try:
            future = asyncio.run_coroutine_threadsafe(coroutine, self.loop)
        except RuntimeError:                # closed meanwhile (the window closing)
            coroutine.close()
            if callback:
                callback(None)
            return
        if callback:
            future.add_done_callback(lambda f: callback(None if f.cancelled() else f.result()))

    async def _session(self):
        if self.manager is None:
            from winrt.windows.media.control import (
                GlobalSystemMediaTransportControlsSessionManager as Manager)
            self.manager = await Manager.request_async()
        for session in self.manager.get_sessions():
            if is_spotify_session(session.source_app_user_model_id):
                return session
        return None

    # --- reading
    def request_now(self, callback):
        """callback(SpotifyNow) or callback(None) when Spotify isn't running
        or isn't playing anything Windows knows of."""
        self._submit(self._now, callback)

    async def _now(self):
        session = await self._session()
        if session is None:
            return None
        properties = await session.try_get_media_properties_async()
        info = session.get_playback_info()
        timeline = session.get_timeline_properties()
        now = SpotifyNow(read_at=time.monotonic())
        now.title = properties.title or ""
        now.artist = properties.artist or properties.album_artist or ""
        now.album = properties.album_title or ""
        now.playing = int(info.playback_status) == 4          # 4 = playing
        try:
            controls = info.controls
            now.can_next = bool(controls.is_next_enabled)
            now.can_previous = bool(controls.is_previous_enabled)
        except Exception:
            pass
        try:
            now.shuffle = None if info.is_shuffle_active is None else bool(info.is_shuffle_active)
        except Exception:
            now.shuffle = None
        try:
            mode = info.auto_repeat_mode
            now.repeat = None if mode is None else REPEAT_FROM_WINDOWS.get(int(mode))
        except Exception:
            now.repeat = None
        try:
            position = timeline.position.total_seconds()
            now.length = max(0.0, (timeline.end_time - timeline.start_time).total_seconds())
            updated = timeline.last_updated_time
            if updated is not None and updated.tzinfo is None:
                updated = updated.replace(tzinfo=datetime.timezone.utc)
            if now.playing and updated is not None and updated.year > 1601:
                age = (datetime.datetime.now(datetime.timezone.utc) - updated).total_seconds()
                if 0 <= age < 3600:
                    position += age
            now.position = position
        except Exception:
            pass
        if now.song_key != self.picture_key:
            self.picture_key = now.song_key
            self.picture = await self._picture(properties)
        now.picture = self.picture
        return now

    async def _picture(self, properties):
        try:
            reference = properties.thumbnail
            if reference is None:
                return None
            from winrt.windows.storage.streams import Buffer, InputStreamOptions
            stream = await reference.open_read_async()
            size = int(stream.size)
            if not size or size > 20 * 1024 * 1024:
                return None
            buffer = Buffer(size)
            result = await stream.read_async(buffer, size, InputStreamOptions.READ_AHEAD)
            return bytes(memoryview(result))[:result.length]
        except Exception:
            return None

    # --- buttons
    def command(self, name, value=None, callback=None):
        """name: play, pause, toggle, next, previous, seek (value = seconds),
        shuffle (value = True/False) or repeat (value = off/all/one)."""
        async def run():
            session = await self._session()
            if session is None:
                return False
            if name == "play":
                return await session.try_play_async()
            if name == "pause":
                return await session.try_pause_async()
            if name == "toggle":
                return await session.try_toggle_play_pause_async()
            if name == "next":
                return await session.try_skip_next_async()
            if name == "previous":
                return await session.try_skip_previous_async()
            if name == "seek":
                return await session.try_change_playback_position_async(int(max(0, value) * 1e7))
            if name == "shuffle":
                return await session.try_change_shuffle_active_async(bool(value))
            if name == "repeat":
                from winrt.windows.media import MediaPlaybackAutoRepeatMode
                return await session.try_change_auto_repeat_mode_async(
                    MediaPlaybackAutoRepeatMode(REPEAT_TO_WINDOWS[value]))
            raise ValueError(name)
        self._submit(run, callback)

    def close(self):
        """Stop for good. Asking anything after this answers None rather
        than failing ("Event loop is closed")."""
        self.closed = True
        if self.loop is not None:
            try:
                self.loop.call_soon_threadsafe(self.loop.stop)
            except RuntimeError:            # already closed
                pass


# ----------------------------------------------------------------- volume

class SpotifyVolume:
    """Spotify's slider in the Windows volume mixer. Spotify only has one once
    it has played something since it started. Call from the window's thread."""

    REFRESH_SECONDS = 4

    def __init__(self):
        self.sessions = []
        self.found_at = -1e9

    def _volumes(self):
        if time.monotonic() - self.found_at > self.REFRESH_SECONDS or not self.sessions:
            self.found_at = time.monotonic()
            self.sessions = []
            try:
                from pycaw.pycaw import AudioUtilities
                found = AudioUtilities.GetAllSessions()
            except Exception:
                found = []
            for session in found:
                try:        # a program that has just closed mustn't hide Spotify's slider
                    if session.Process and session.Process.name().lower() == "spotify.exe":
                        self.sessions.append(session.SimpleAudioVolume)
                except Exception:
                    continue
        return self.sessions

    def get(self):
        """(volume 0 to 1, muted) or None when Spotify has no slider yet."""
        for volume in self._volumes():
            try:
                return float(volume.GetMasterVolume()), bool(volume.GetMute())
            except Exception:
                self.sessions = []
        return None

    def set(self, level=None, muted=None):
        done = False
        for volume in self._volumes():
            try:
                if level is not None:
                    volume.SetMasterVolume(max(0.0, min(1.0, float(level))), None)
                if muted is not None:
                    volume.SetMute(bool(muted), None)
                done = True
            except Exception:
                self.sessions = []
        return done
