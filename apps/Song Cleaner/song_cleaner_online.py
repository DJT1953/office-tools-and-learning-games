"""Song Cleaner: asking the internet what a song is.

Only used for names the rules could not settle, and only if a key has been
typed into Settings. The song's sound is turned into a Chromaprint
fingerprint by a small free program (fpcalc), and the fingerprint is sent to
AcoustID, a free public database, which answers with the artist, title and
length it knows for that sound. Nothing but the fingerprint, the song's
length and the key is sent.
"""

import collections
import dataclasses
import io
import json
import os
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

from song_cleaner_engine import NO_WINDOW, SongError, app_folder

FPCALC_VERSION = "1.6.2"                    # checked on acoustid.org/chromaprint, 2026-11-12
FPCALC_URL = ("https://github.com/acoustid/chromaprint/releases/download/"
              f"v{FPCALC_VERSION}/chromaprint-fpcalc-{FPCALC_VERSION}-windows-x86_64.zip")
LOOKUP_URL = "https://api.acoustid.org/v2/lookup"
KEY_PAGE = "https://acoustid.org/new-application"
MIN_SCORE = 0.8
PER_SECOND = 3                              # AcoustID's stated limit
WAITS = (2, 4, 8)                           # seconds to wait before trying again


class OnlineError(SongError):
    """The lookup could not be done; said in plain words, and never with the key in it."""


@dataclasses.dataclass
class Match:
    artist: str
    title: str
    length: float
    score: float


def fpcalc_path():
    path = os.path.join(app_folder(), "fpcalc.exe")
    return path if os.path.isfile(path) else None


def fetch_fpcalc(opener=urllib.request.urlopen):
    """Download fpcalc from the Chromaprint project's own release page. Returns its path."""
    try:
        with opener(FPCALC_URL, timeout=60) as reply:
            data = reply.read()
    except (OSError, urllib.error.URLError) as error:
        raise OnlineError("The fingerprint program could not be downloaded. "
                          "Check the internet connection and try again.") from error
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            inside = [n for n in archive.namelist() if n.lower().endswith("fpcalc.exe")]
            if not inside:
                raise OnlineError("The download did not contain the fingerprint program.")
            target = os.path.join(app_folder(), "fpcalc.exe")
            with open(target, "wb") as out:
                out.write(archive.read(inside[0]))
    except zipfile.BadZipFile as error:
        raise OnlineError("The download was damaged. Try again.") from error
    except OSError as error:
        raise OnlineError("The fingerprint program could not be saved.") from error
    return target


def chromaprint(path, program=None):
    """(length in whole seconds, fingerprint text) for a song."""
    program = program or fpcalc_path()
    if not program:
        raise OnlineError("The fingerprint program is not installed yet.")
    try:
        done = subprocess.run([program, "-json", path], capture_output=True, timeout=120,
                              creationflags=NO_WINDOW)
        data = json.loads(done.stdout.decode("utf-8", "replace"))
        return int(round(float(data["duration"]))), str(data["fingerprint"])
    except (OSError, ValueError, KeyError, subprocess.TimeoutExpired) as error:
        raise OnlineError("This song could not be fingerprinted.") from error


def parse_reply(data):
    """The best Match in AcoustID's answer, or None when it does not know the song."""
    if not isinstance(data, dict) or data.get("status") != "ok":
        error = data.get("error") if isinstance(data, dict) else None
        message = str((error or {}).get("message", "")).lower() if isinstance(error, dict) else ""
        if "key" in message or "client" in message:
            raise OnlineError("AcoustID did not accept the key. Check it in Settings.")
        if "rate" in message or "limit" in message:
            raise OnlineError("AcoustID says too many questions were asked at once.")
        raise OnlineError("AcoustID could not answer.")
    best = None
    for result in data.get("results") or []:
        try:
            score = float(result.get("score") or 0)
        except (TypeError, ValueError):
            continue
        if score < MIN_SCORE:
            continue
        named = collections.Counter()
        lengths = {}
        for recording in result.get("recordings") or []:
            title = str(recording.get("title") or "").strip()
            artists = recording.get("artists") or []
            artist = "".join(str(a.get("name") or "") + str(a.get("joinphrase") or "")
                             for a in artists if isinstance(a, dict)).strip()
            if title and artist:
                named[artist, title] += 1
                lengths.setdefault((artist, title), float(recording.get("duration") or 0))
        if named and (best is None or score > best.score):
            artist, title = named.most_common(1)[0][0]
            best = Match(artist, title, lengths[artist, title], score)
    return best


class Lookup:
    """Asks AcoustID about songs, never faster than it allows."""

    def __init__(self, key, opener=urllib.request.urlopen, sleep=time.sleep, clock=time.monotonic,
                 fingerprinter=chromaprint):
        self.key = (key or "").strip()
        self.opener, self.sleep, self.clock = opener, sleep, clock
        self.fingerprinter = fingerprinter
        self.last = None
        if not self.key:
            raise OnlineError("No AcoustID key has been entered in Settings.")

    def _wait_turn(self):
        gap = 1.0 / PER_SECOND + 0.02
        now = self.clock()
        if self.last is not None and now - self.last < gap:
            self.sleep(gap - (now - self.last))
        self.last = self.clock()

    def _ask(self, length, fingerprint):
        body = urllib.parse.urlencode({"client": self.key, "duration": length, "meta": "recordings",
                                       "fingerprint": fingerprint, "format": "json"}).encode("ascii")
        request = urllib.request.Request(LOOKUP_URL, data=body, method="POST", headers={
            "Content-Type": "application/x-www-form-urlencoded", "User-Agent": "SongCleaner/1.0"})
        for wait in (*WAITS, None):
            self._wait_turn()
            try:
                with self.opener(request, timeout=30) as reply:
                    return json.loads(reply.read().decode("utf-8", "replace"))
            except urllib.error.HTTPError as error:
                self._close(error)
                if error.code == 429 or error.code >= 500:
                    if wait is None:
                        raise OnlineError("AcoustID is too busy just now. Try again later.") from None
                    self.sleep(wait)
                    continue
                try:
                    return json.loads(self._body)                 # says what was wrong
                except ValueError:
                    raise OnlineError("AcoustID could not answer.") from None
            except (OSError, urllib.error.URLError, ValueError):
                if wait is None:
                    raise OnlineError("AcoustID could not be reached. "
                                      "Check the internet connection.") from None
                self.sleep(wait)
        return None

    def _close(self, error):
        """Read what an error reply says, then let go of it."""
        try:
            self._body = error.read().decode("utf-8", "replace")
        except (OSError, ValueError):
            self._body = ""
        try:
            error.close()
        except OSError:
            pass

    def find(self, path):
        """What AcoustID knows this song as, or None."""
        length, fingerprint = self.fingerprinter(path)
        return parse_reply(self._ask(length, fingerprint))
