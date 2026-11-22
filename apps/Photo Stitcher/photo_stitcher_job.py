"""Runs one stitch for the Photo Stitcher in a second process (a second copy of
Python working in the background). The window never freezes while big photos
are joined, and Cancel can stop the work at once, even if OpenCV gets stuck.

The window starts a StitchJob, calls poll() a few times a second, and gets:
    ("step", "Opening photo 2 of 7: IMG_0041.jpg")
    ("done", info)        info: see worker()
    ("failed", "why, in plain words")
"""

from __future__ import annotations

import multiprocessing
import os
import queue
import shutil
import tempfile
import threading
import time
import traceback

FOLDER_PREFIX = "Photo Stitcher "


def _end_with_the_window() -> None:
    """If the window's process ends without stopping this one (it was killed,
    or Windows closed it), stop too instead of joining photos for nobody."""
    parent = multiprocessing.parent_process()
    if parent is None:
        return

    def watch():
        parent.join()
        os._exit(1)

    threading.Thread(target=watch, daemon=True).start()


def clean_old_folders(days: float = 1.0) -> int:
    """Removes work folders left in the temporary folder by a stitch that was
    cut off (the window killed, the PC switched off). Only folders older than
    `days` go, so a second open Photo Stitcher keeps its own. Returns how many."""
    removed = 0
    top = tempfile.gettempdir()
    try:
        names = os.listdir(top)
    except OSError:
        return 0
    for name in names:
        path = os.path.join(top, name)
        if not (name.startswith(FOLDER_PREFIX) and os.path.isdir(path)):
            continue
        try:
            if time.time() - os.path.getmtime(path) > days * 86400:
                shutil.rmtree(path)
                removed += 1
        except OSError:
            pass
    return removed


def worker(paths: list[str], kind: str, folder: str, messages) -> None:
    """Opens the photos, joins them, finds the trim, and leaves the joined
    picture in folder/panorama.npy. The last message is "done" or "failed".

    info = {"image_file": str, "used": [indexes into paths], "unreadable":
            [(path, reason)], "kind": str, "switched": bool, "scale": float,
            "trim": (x, y, width, height)}
    """
    _end_with_the_window()
    try:
        import numpy as np
        import photo_stitcher_engine as engine
    except Exception:
        messages.put(("failed", "A part the Photo Stitcher needs is missing:\n"
                      + traceback.format_exc(limit=1)))
        return
    try:
        images, kept, unreadable = [], [], []
        for number, path in enumerate(paths, 1):
            messages.put(("step", f"Opening photo {number} of {len(paths)}: "
                                  f"{os.path.basename(path)}"))
            try:
                images.append(engine.load_photo(path))
                kept.append(number - 1)
            except engine.PhotoError as error:
                unreadable.append((path, str(error)))
            except MemoryError:
                unreadable.append((path, "is too big to open on this computer"))
        if len(images) < 2:
            names = "\n".join(f"  {os.path.basename(p)} {why}" for p, why in unreadable)
            messages.put(("failed", "There aren't 2 photos that can be opened, so nothing could "
                                    "be joined." + (f"\n\n{names}" if names else "")))
            return
        messages.put(("step", f"Finding the matching details in {len(images)} photos "
                              "and joining them"))
        result = engine.stitch(images, kind)
        del images
        messages.put(("step", "Finding the ragged edges to trim"))
        trim = engine.trim_rectangle(result.image, result.coverage)
        messages.put(("step", "Handing the panorama to the window"))
        image_file = os.path.join(folder, "panorama.npy")
        np.save(image_file, result.image)
        messages.put(("done", {"image_file": image_file,
                               "used": [kept[i] for i in result.used],
                               "unreadable": unreadable, "kind": result.kind,
                               "switched": result.switched, "scale": result.scale,
                               "trim": tuple(int(v) for v in trim)}))
    except engine.StitchError as error:
        messages.put(("failed", str(error)))
    except MemoryError:
        messages.put(("failed", "This computer ran out of memory joining these photos. "
                                "Try fewer photos, or close other programs and try again."))
    except Exception:
        messages.put(("failed", "Something went wrong while joining the photos:\n\n"
                                + traceback.format_exc()[-1500:]))


def read_picture(info: dict):
    """The joined picture the worker left behind (a numpy array)."""
    import numpy as np
    return np.load(info["image_file"])


class StitchJob:
    """One stitch in a second process. Make it, start() it, then poll()."""

    def __init__(self, paths: list[str], kind: str):
        self.paths, self.kind = list(paths), kind
        self.folder = tempfile.mkdtemp(prefix=FOLDER_PREFIX)
        context = multiprocessing.get_context("spawn")
        self.messages = context.Queue()
        self.process = context.Process(target=worker, daemon=True,
                                       args=(self.paths, kind, self.folder, self.messages))
        self.finished = False

    def start(self) -> None:
        self.process.start()

    @property
    def running(self) -> bool:
        return not self.finished and self.process.is_alive()

    def poll(self) -> tuple | None:
        """The next message, or None if there is none yet. When the process has
        ended without a last message (it crashed), this says so."""
        try:
            message = self.messages.get_nowait()
        except queue.Empty:
            if self.finished or self.process.is_alive():
                return None
            try:                       # the very last message may still be on its way
                message = self.messages.get(timeout=0.5)
            except queue.Empty:
                self.finished = True
                return ("failed", "The joining stopped by itself before it finished. The "
                                  "photos may be too big for this computer's memory: try "
                                  "fewer photos, or close other programs.")
        if message[0] in ("done", "failed"):
            self.finished = True
            self._stop_process()       # OpenCV can hang while that process closes
        return message

    def _stop_process(self) -> None:
        """The last message is in, so the worker has nothing left to do: stop it
        now rather than wait, because OpenCV can hang while a process closes."""
        if self.process.is_alive():
            self.process.terminate()
            self.process.join(2)

    def cancel(self) -> None:
        """Stops the work at once and removes its temporary files."""
        self.finished = True
        if self.process.is_alive():
            self.process.terminate()
            self.process.join(5)
        self.clean_up()

    def clean_up(self) -> None:
        shutil.rmtree(self.folder, ignore_errors=True)
