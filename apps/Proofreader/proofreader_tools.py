"""
proofreader_tools.py

The free checker behind the Proofreader: LanguageTool, a program that runs on
this computer only. It needs Java 17, so the first run downloads LanguageTool
and its own private copy of Java into

    %LOCALAPPDATA%\\Proofreader\\tools      (about 700 MB once unpacked)

Nothing is installed into Windows and no document text leaves the computer:
the checker listens on 127.0.0.1 only and is stopped when the app closes.
"""

import atexit
import concurrent.futures
import ctypes
import glob
import hashlib
import json
import os
import shutil
import socket
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

LANGUAGETOOL_URL = "https://languagetool.org/download/LanguageTool-stable.zip"
JAVA_INDEX_URL = ("https://api.adoptium.net/v3/assets/latest/17/hotspot"
                  "?architecture=x64&image_type=jre&os=windows")
DOWNLOAD_MB = 300
AGENT = {"User-Agent": "Proofreader"}


def computer_memory_gb():
    """The memory fitted in this computer, in GB (0 when it can't be told)."""
    try:
        kilobytes = ctypes.c_ulonglong(0)
        if ctypes.windll.kernel32.GetPhysicallyInstalledSystemMemory(ctypes.byref(kilobytes)):
            return kilobytes.value // (1024 * 1024)
    except (AttributeError, OSError):
        pass
    return 0


def checker_memory(gb=None):
    """How much memory the checker may use. More lets it work on long batches
    without stopping to tidy up; a small computer keeps enough for itself."""
    gb = computer_memory_gb() if gb is None else gb
    if gb >= 16:
        return "4g"
    if gb >= 8:
        return "2g"
    return "1g"


def tools_folder():
    override = os.environ.get("PROOFREADER_TOOLS")
    if override:
        return override
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "Proofreader", "tools")


def find_java(folder=None):
    found = glob.glob(os.path.join(folder or tools_folder(), "jre", "*", "bin", "javaw.exe"))
    return found[0] if found else None


def find_languagetool(folder=None):
    found = glob.glob(os.path.join(folder or tools_folder(), "lt", "*", "languagetool-server.jar"))
    return found[0] if found else None


def installed(folder=None):
    return bool(find_java(folder) and find_languagetool(folder))


# ---------------------------------------------------------------- one-time download
def _download(url, target, progress, label, done_before, total):
    request = urllib.request.Request(url, headers=AGENT)
    digest = hashlib.sha256()
    with urllib.request.urlopen(request, timeout=60) as response, open(target, "wb") as out:
        done = 0
        while True:
            block = response.read(1 << 18)
            if not block:
                break
            out.write(block)
            digest.update(block)
            done += len(block)
            if progress:
                progress(label, min(0.95, (done_before + done) / total))
    return digest.hexdigest(), done


def _unpack(archive, folder):
    """Unpack beside the final place, then swap in, so a stopped download
    never leaves a half-made folder that looks finished."""
    fresh = folder + ".new"
    shutil.rmtree(fresh, ignore_errors=True)
    with zipfile.ZipFile(archive) as z:
        for name in z.namelist():                     # nothing may land outside the folder
            path = os.path.realpath(os.path.join(fresh, name))
            if not path.startswith(os.path.realpath(fresh) + os.sep):
                raise RuntimeError("The download contains an unsafe file name: " + name)
        z.extractall(fresh)
    shutil.rmtree(folder, ignore_errors=True)
    os.replace(fresh, folder)
    os.remove(archive)


def install(progress=None, folder=None):
    """Download whatever is missing. progress(label, fraction 0..1)."""
    try:
        _install(progress, folder or tools_folder())
    except (urllib.error.URLError, TimeoutError, ConnectionError) as error:
        raise RuntimeError("The free checker could not be downloaded. Check the internet "
                           f"connection and press Start again.\n\n({error})") from error


def _install(progress, folder):
    os.makedirs(folder, exist_ok=True)
    total = DOWNLOAD_MB * 1024 * 1024
    done = 0
    if not find_java(folder):
        with urllib.request.urlopen(urllib.request.Request(JAVA_INDEX_URL, headers=AGENT),
                                    timeout=60) as response:
            package = json.load(response)[0]["binary"]["package"]
        archive = os.path.join(folder, "jre.zip")
        digest, size = _download(package["link"], archive, progress, "Downloading Java", done, total)
        if digest != package["checksum"]:
            os.remove(archive)
            raise RuntimeError("The Java download arrived damaged. Please try again.")
        done += size
        _unpack(archive, os.path.join(folder, "jre"))
    if not find_languagetool(folder):
        archive = os.path.join(folder, "lt.zip")
        _download(LANGUAGETOOL_URL, archive, progress, "Downloading LanguageTool", done, total)
        if progress:
            progress("Unpacking LanguageTool", 0.97)
        _unpack(archive, os.path.join(folder, "lt"))
    if not installed(folder):
        raise RuntimeError("The checker did not unpack properly in " + folder)
    if progress:
        progress("Ready", 1.0)


# ---------------------------------------------------------------- the checker itself
class LanguageTool:
    """LanguageTool running hidden on this computer."""

    CHUNK = 2000            # characters in one piece; several pieces are checked at once
    AT_ONCE = 6             # measured: 6 pieces at once took 0.48 s where one took 1.14 s
    BREAK = "\n\n"

    def __init__(self, folder=None, language="en-US"):
        self.java = find_java(folder)
        self.jar = find_languagetool(folder)
        self.language = language
        self.process = None
        self.port = None
        self.off_rules = ()
        self.off_categories = ()

    def start(self, wait=90):
        if not (self.java and self.jar):
            raise RuntimeError("The free checker is not downloaded yet.")
        with socket.socket() as probe:
            probe.bind(("127.0.0.1", 0))
            self.port = probe.getsockname()[1]
        self.process = subprocess.Popen(
            [self.java, "-Xmx" + checker_memory(), "-cp", self.jar, "org.languagetool.server.HTTPServer",
             "--port", str(self.port)],
            cwd=os.path.dirname(self.jar), stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        atexit.register(self.stop)          # never leave it running behind a closed app
        deadline = time.time() + wait
        while time.time() < deadline:
            if self.process.poll() is not None:
                raise RuntimeError("The free checker stopped as soon as it started.")
            try:
                with urllib.request.urlopen(self._url("languages"), timeout=3):
                    return self
            except OSError:
                time.sleep(0.5)
        self.stop()
        raise RuntimeError("The free checker took too long to start.")

    def stop(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(10)
            except subprocess.TimeoutExpired:
                self.process.kill()
        self.process = None

    def __enter__(self):
        return self.start()

    def __exit__(self, *_):
        self.stop()

    def _url(self, what):
        return f"http://127.0.0.1:{self.port}/v2/{what}"

    def check(self, text):
        fields = {"text": text, "language": self.language}
        if self.off_rules:
            fields["disabledRules"] = ",".join(self.off_rules)
        if self.off_categories:
            fields["disabledCategories"] = ",".join(self.off_categories)
        body = urllib.parse.urlencode(fields).encode()
        with urllib.request.urlopen(urllib.request.Request(self._url("check"), body),
                                    timeout=300) as response:
            return json.load(response)["matches"]

    def check_many(self, texts):
        """Matches for each text, with offsets counted inside that text. Texts
        go in pieces of a few paragraphs each, and several pieces are checked
        at the same time, which the checker handles on separate processor
        cores. The answers are the same as checking the text in one go."""
        answers = [[] for _ in texts]
        groups, group, size = [], [], 0
        for index, text in enumerate(texts):
            if not text.strip():
                continue
            if group and size + len(text) > self.CHUNK:
                groups.append(group)
                group, size = [], 0
            group.append(index)
            size += len(text) + len(self.BREAK)
        if group:
            groups.append(group)

        def send(group):
            starts, joined, at = [], [], 0
            for index in group:
                starts.append(at)
                joined.append(texts[index])
                at += len(texts[index]) + len(self.BREAK)
            for match in self.check(self.BREAK.join(joined)):
                offset, length = match["offset"], match["length"]
                for place, index in enumerate(group):
                    inside = offset - starts[place]
                    if 0 <= inside and inside + length <= len(texts[index]):
                        answers[index].append(dict(match, offset=inside))   # each piece its own
                        break

        if len(groups) <= 1:
            for group in groups:
                send(group)
        else:
            with concurrent.futures.ThreadPoolExecutor(min(self.AT_ONCE, len(groups))) as pool:
                list(pool.map(send, groups))        # list() brings back any error
        return answers
