"""Photo Stitcher: joins overlapping photos into one big picture (a panorama),
side by side, top to bottom, or in rows and columns, in any order.

Step 1 pick the photos, Step 2 say what kind they are, Step 3 stitch,
Step 4 look at the result and save it as JPG, PNG or TIFF.
The original photos are never changed, moved or deleted.

Start it by double-clicking Run_PhotoStitcher.bat (or its RUN shortcut).
Requirements (one-time, offered on first start): opencv-python-headless,
numpy, Pillow and pillow-heif.
"""

from __future__ import annotations

import base64
import ctypes
import importlib
import importlib.util
import json
import multiprocessing
import os
import site
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

TITLE = "Photo Stitcher"
VERSION = "1.0.2"    # the Setup's AppVersion must match (test_the_version_matches_the_setup)
# The "Check for updates" button opens this page, which has the newest version of every program.
UPDATES_PAGE = "https://djt1953.github.io/office-tools-and-learning-games/"
NEEDS = {   # module: (pip name, what it's for, size in MB)
    "cv2": ("opencv-python-headless", "OpenCV (finds the matching details and joins the photos)", 40),
    "numpy": ("numpy", "numpy (holds the pictures while they are joined)", 15),
    "PIL": ("pillow", "Pillow (reads photo sizes)", 7),
    "pillow_heif": ("pillow-heif", "pillow-heif (opens iPhone HEIC photos)", 10),
}
SETTINGS = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), TITLE, "settings.json")
PHOTO_PATTERN = "*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp *.heic *.heif"
SAVE_CHOICES = (   # ending, words beside the button, words in the Save box
    (".jpg", "JPG (smaller file)", "JPG picture"),
    (".png", "PNG (full quality, bigger file)", "PNG picture"),
    (".tif", "TIFF", "TIFF picture"),
)


# ------------------------------------------------------------------ start-up checks

def show_crash(kind, value, tb):
    text = "".join(traceback.format_exception(kind, value, tb))
    try:
        messagebox.showerror(TITLE, "Something went wrong and Photo Stitcher has to close.\n\n"
                             + text[-1500:])
    finally:
        sys.exit(1)


def has_module(name):
    return importlib.util.find_spec(name) is not None


def install_missing():
    """Double-clicking may start a Python that lacks the add-ons, so offer to
    add them. A built copy (from Setup) already has everything inside it."""
    if getattr(sys, "frozen", False):
        return
    missing = [need for module, need in NEEDS.items() if not has_module(module)]
    if not missing:
        return
    root = tk.Tk()
    root.withdraw()
    names = " and ".join(what for _, what, _ in missing)
    size = sum(mb for *_, mb in missing)
    if not messagebox.askyesno(TITLE, f"Photo Stitcher still needs {names}.\n\nDownload and "
                               f"install now? This only happens once. It is about {size} MB; "
                               "a window shows the progress."):
        sys.exit()
    root.destroy()
    command = [console_python(), "-m", "pip", "install", *(pip for pip, _, _ in missing)]
    if subprocess.run(command, creationflags=subprocess.CREATE_NEW_CONSOLE).returncode:
        raise RuntimeError(f"Could not install {names}. Check the internet connection and try again.")
    site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()


def console_python(executable=None):
    """The Python that shows a window. Run_PhotoStitcher.bat starts pythonw.exe,
    which never opens one, so pip's progress would be invisible under it."""
    executable = executable or sys.executable
    folder, name = os.path.split(executable)
    if name.lower() == "pythonw.exe" and os.path.isfile(os.path.join(folder, "python.exe")):
        return os.path.join(folder, "python.exe")
    return executable


def sharp_on_big_screens():
    """Tell Windows this program handles screen scaling itself, so the window
    and the preview are sharp."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        pass


def set_icon(root):
    """The joined-photos icon in the title bar: packed beside the program in a
    built copy, or in the installer folder when run from the source files."""
    here = getattr(sys, "_MEIPASS", None) or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "Installers", TITLE)
    icon = os.path.join(here, f"{TITLE}.ico")
    if os.path.exists(icon):
        try:
            root.iconbitmap(default=icon)
        except tk.TclError:
            pass


def load_settings():
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            values = json.load(f)
            return values if isinstance(values, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(values):
    try:
        os.makedirs(os.path.dirname(SETTINGS), exist_ok=True)
        with open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(values, f)
    except OSError:
        pass


def clock(seconds):
    seconds = max(0, int(seconds))
    hours, rest = divmod(seconds, 3600)
    return f"{hours}:{rest // 60:02}:{rest % 60:02}" if hours else f"{rest // 60}:{rest % 60:02}"


def describe_result(info, paths, width, height, kind_names):
    """The words under the preview: how many photos joined, which were left
    out and why, and how big the panorama is."""
    unreadable = [path for path, _ in info["unreadable"]]
    used = set(info["used"])
    left_out = [i for i, path in enumerate(paths) if i not in used and path not in unreadable]
    if len(used) == len(paths):
        lines = [f"Joined all {len(paths)} photos."]
    else:
        lines = [f"Joined {len(used)} of {len(paths)} photos."]
    if left_out:
        names = ", ".join(os.path.basename(paths[i]) for i in left_out)
        lines.append(f"Left out: {names} (didn't overlap enough with the others).")
    for path, why in info["unreadable"]:
        lines.append(f"Left out: {os.path.basename(path)} {why}.")
    if info["switched"]:
        lines.append(f"They fit better as {kind_names[info['kind']]} than the kind picked in "
                     "Step 2, so they were joined that way.")
    if info["scale"] < 1:
        lines.append(f"This computer ran short of memory, so the photos were joined at "
                     f"{round(info['scale'] * 100)}% of their size.")
    lines.append(f"Size: {width:,} x {height:,} pixels.")
    return "\n".join(lines)


# ------------------------------------------------------------------ the window

class App:
    def __init__(self, root, engine, jobs):
        self.root, self.engine, self.jobs = root, engine, jobs
        self.paths = []                  # the photos in Step 1, in list order
        self.job = None                  # the stitch running now
        self.started = 0.0
        self.step_text = ""
        self.stitched_paths = []         # the photos the result was made from
        self.result = None               # the joined picture, untrimmed
        self.info = None                 # what the job said about it
        self.preview_source = None       # a smaller copy for drawing the preview
        self.preview_photo = None
        self.resize_waiting = None
        self.saved_path = None
        self.unsaved = False
        self.saving = False
        self.results = []                # things the saving thread asks the window to do
        self.settings = load_settings()
        kind = self.settings.get("kind")
        save_type = self.settings.get("save_type")
        self.kind = tk.StringVar(value=kind if kind in engine.KIND_NAMES else engine.TURNING)
        self.trim_on = tk.BooleanVar(value=self.settings.get("trim", True) is not False)
        endings = [ending for ending, _, _ in SAVE_CHOICES]
        self.save_type = tk.StringVar(value=save_type if save_type in endings else ".jpg")

        root.title(TITLE)
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        style.configure("Big.TButton", font=("Segoe UI", 13, "bold"), padding=(14, 9),
                        foreground="black")
        style.configure("TButton", foreground="black")
        style.configure("TCheckbutton", font=("Segoe UI", 12))
        style.configure("TRadiobutton", font=("Segoe UI", 12))
        style.configure("TLabelframe.Label", font=("Segoe UI", 12, "bold"))
        style.configure("Note.TLabel", font=("Segoe UI", 11))
        list_font = tkfont.Font(root, family="Segoe UI", size=11)
        style.configure("Treeview", font=list_font, rowheight=int(list_font.metrics("linespace") * 1.35))
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))

        page = ttk.Frame(root, padding=16)
        page.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(2, weight=1)

        heading = ttk.Frame(page)            # the name, with "Check for updates" in the right corner
        heading.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(heading, text=TITLE, font=("Segoe UI", 20, "bold")).pack(side="left")
        self.updates_button = ttk.Button(heading, text="Check for updates",
                                         command=self.check_for_updates)
        self.updates_button.pack(side="right", anchor="n")
        ttk.Label(page, style="Note.TLabel",
                  text="Joins overlapping photos into one big picture (a panorama): side by side, "
                       "top to bottom, or in rows and columns. Your photos are never changed."
                  ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        left = ttk.Frame(page)
        left.grid(row=2, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        # Step 1
        pick = ttk.LabelFrame(left, text="Step 1 - Pick the photos", padding=12)
        pick.grid(row=0, sticky="nsew", pady=(0, 6))
        pick.columnconfigure((0, 1), weight=1)
        pick.rowconfigure(2, weight=1)
        self.add_button = self.big(pick, "Add photos", self.add_photos, 0, 0)
        self.folder_button = self.big(pick, "Add a folder", self.add_folder, 0, 1)
        self.remove_button = self.big(pick, "Remove picked", self.remove_picked, 1, 0)
        self.clear_button = self.big(pick, "Clear all", self.clear_all, 1, 1)
        holder = ttk.Frame(pick)
        holder.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(holder, columns=("name", "size"), show="headings", height=7,
                                 selectmode="extended")
        self.tree.heading("name", text="Photo", anchor="w")
        self.tree.heading("size", text="Size", anchor="w")
        self.tree.column("name", width=290, anchor="w")
        self.tree.column("size", width=150, anchor="w", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.bind("<<TreeviewSelect>>", lambda event: self.update_buttons())
        self.tree.bind("<Delete>", lambda event: self.remove_picked())
        self.count = ttk.Label(pick, style="Note.TLabel", wraplength=440, justify="left")
        self.count.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Step 2
        kinds = ttk.LabelFrame(left, text="Step 2 - What kind of photos?", padding=12)
        kinds.grid(row=1, sticky="ew", pady=6)
        ttk.Radiobutton(kinds, text="Taken by turning the camera (landscapes, rooms, buildings)",
                        value=engine.TURNING, variable=self.kind,
                        command=self.kind_changed).grid(sticky="w")
        ttk.Radiobutton(kinds, text="Flat things shot or scanned in parts (documents, maps, walls)",
                        value=engine.FLAT, variable=self.kind,
                        command=self.kind_changed).grid(sticky="w", pady=(4, 0))

        # Step 3
        run = ttk.LabelFrame(left, text="Step 3 - Stitch", padding=12)
        run.grid(row=2, sticky="ew", pady=(6, 0))
        run.columnconfigure(0, weight=1)
        self.stitch_button = self.big(run, "Stitch the photos", self.start_stitch, 0, 0)
        self.cancel_button = self.big(run, "Cancel", self.cancel, 0, 1)
        self.progress = ttk.Label(run, style="Note.TLabel", wraplength=440, justify="left",
                                  text="Joining takes from a few seconds to a few minutes, "
                                       "depending on how many photos and how big they are.")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # Step 4
        right = ttk.Frame(page)
        right.grid(row=2, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        self.look = ttk.LabelFrame(right, text="Step 4 - Look and save", padding=12)
        self.look.grid(sticky="nsew")
        self.look.columnconfigure(0, weight=1)
        self.look.rowconfigure(0, weight=1, minsize=300)
        self.preview = tk.Label(self.look, text="The joined picture shows here.", bg="#e4e4e4",
                                fg="black", font=("Segoe UI", 12), width=1, height=1)
        self.preview.grid(row=0, sticky="nsew")
        self.preview.bind("<Configure>", self.preview_resized)
        options = ttk.Frame(self.look)
        options.grid(row=1, sticky="w", pady=(10, 0))
        ttk.Checkbutton(options, text="Trim the ragged black edges", variable=self.trim_on,
                        command=self.trim_changed).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(options, text="Save as:", font=("Segoe UI", 12)).grid(row=1, column=0,
                                                                        sticky="w", pady=(6, 0))
        for column, (ending, words, _) in enumerate(SAVE_CHOICES, 1):
            ttk.Radiobutton(options, text=words, value=ending, variable=self.save_type,
                            command=self.save_type_changed).grid(row=1, column=column, sticky="w",
                                                                 padx=(10, 0), pady=(6, 0))
        buttons = ttk.Frame(self.look)
        buttons.grid(row=2, sticky="ew", pady=(10, 0))
        buttons.columnconfigure((0, 1, 2), weight=1)
        self.save_button = self.big(buttons, "Save the panorama", self.save, 0, 0)
        self.open_button = self.big(buttons, "Open it", self.open_saved, 0, 1)
        self.open_folder_button = self.big(buttons, "Open the folder", self.open_folder, 0, 2)
        self.note = ttk.Label(self.look, style="Note.TLabel", justify="left", wraplength=640)
        self.note.grid(row=3, sticky="w", pady=(8, 0))
        step5 = ttk.LabelFrame(right, text="Step 5 - Adjust (optional)", padding=12)
        step5.grid(row=1, sticky="ew", pady=(8, 0))
        step5.columnconfigure(0, weight=1)
        self.adjust_button = self.big(step5, "Turn, square up, crop or resize", self.open_adjust,
                                      0, 0)
        self.adjust = None                   # the Adjust window, once opened
        self.adjust_after_save = False       # open it as soon as the plain panorama is saved
        self.saving_picture = self.saved_picture = None

        self.status = ttk.Label(page, style="Note.TLabel", wraplength=1100, justify="left",
                                text="Ready. Start with Step 1: click Add photos or Add a folder.")
        self.status.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self.show_count()
        self.update_buttons()
        self.poll_waiting = root.after(200, self.poll)

    @staticmethod
    def big(parent, text, command, row, column):
        button = ttk.Button(parent, text=text, style="Big.TButton", command=command)
        button.grid(row=row, column=column, sticky="ew", padx=4, pady=3)
        return button

    def fit_to_screen(self):
        """Size the window to the screen: on a short screen the photo list and
        the preview get shorter, so every button stays in view."""
        root = self.root
        root.update_idletasks()
        room = root.winfo_screenheight() - int(110 * root.winfo_fpixels("1i") / 96)
        rows = int(self.tree.cget("height"))
        while root.winfo_reqheight() > room and rows > 3:
            rows -= 1
            self.tree.configure(height=rows)
            root.update_idletasks()
        if root.winfo_reqheight() > room:
            self.look.rowconfigure(0, minsize=160)
            root.update_idletasks()
        scale = root.winfo_fpixels("1i") / 96
        width = min(max(root.winfo_reqwidth(), int(1180 * scale)), root.winfo_screenwidth() - 20)
        height = min(max(root.winfo_reqheight(), int(720 * scale)), room)
        left = max(0, (root.winfo_screenwidth() - width) // 2)
        top = max(0, (room - height) // 2)
        root.geometry(f"{width}x{height}+{left}+{top}")
        root.minsize(min(width, int(900 * scale)), min(height, root.winfo_reqheight()))

    def bring_to_front(self):
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(400, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    def say(self, text):
        self.status.configure(text=text)

    def busy(self):
        """Why the window can't take a new request right now, or ''."""
        if self.job:
            return "The photos are being joined. Wait for it to finish, or click Cancel."
        if self.saving:
            return "Wait a moment: the panorama is being saved."
        return ""

    def remember(self, **values):
        self.settings.update(values)
        save_settings(self.settings)

    def update_buttons(self):
        def ready(button, ok):
            button.state(["!disabled"] if ok else ["disabled"])
        ready(self.remove_button, bool(self.tree.selection()))
        ready(self.clear_button, bool(self.paths))
        ready(self.cancel_button, self.job is not None)
        ready(self.save_button, self.result is not None)
        saved = bool(self.saved_path and os.path.exists(self.saved_path))
        ready(self.open_button, saved)
        ready(self.open_folder_button, saved)
        ready(self.adjust_button, self.result is not None)

    def show_count(self):
        count = len(self.paths)
        if count == 0:
            text = "No photos yet. Click Add photos, or Add a folder to take every photo in a folder."
        elif count == 1:
            text = "1 photo. Add at least one more that overlaps it."
        else:
            text = f"{count} photos. The order doesn't matter: the program works out where each one goes."
        self.count.configure(text=text)

    def check_for_updates(self):
        """Opens UPDATES_PAGE in the web browser, or says where to go when that can't be done."""
        try:
            opened = webbrowser.open(UPDATES_PAGE)
        except Exception:
            opened = False
        if opened:
            self.say("The download page is opening in your web browser.")
        else:
            messagebox.showinfo(TITLE, "The web browser could not be opened. To check for a newer "
                                f"version, go to:\n\n{UPDATES_PAGE}", parent=self.root)

    # -- Step 1

    def add_photos(self):
        if why := self.busy():
            self.say(why)
            return
        start = self.settings.get("photo_folder") or self.engine.pictures_folder()
        picked = filedialog.askopenfilenames(
            parent=self.root, title="Pick the photos to join", initialdir=start,
            filetypes=[("Photos", PHOTO_PATTERN), ("All files", "*.*")])
        if not picked:
            self.say("No photos were added.")
            return
        self.add(list(self.root.tk.splitlist(picked)) if isinstance(picked, str) else list(picked))

    def add_folder(self):
        if why := self.busy():
            self.say(why)
            return
        start = self.settings.get("photo_folder") or self.engine.pictures_folder()
        folder = filedialog.askdirectory(parent=self.root, title="Pick a folder of photos to join",
                                         initialdir=start, mustexist=True)
        if not folder:
            self.say("No folder was picked.")
            return
        photos = self.engine.photos_in_folder(folder)
        if not photos:
            self.say(f"There are no photos in {os.path.normpath(folder)}. Photos are .jpg, .png, "
                     ".tif, .bmp, .webp or .heic files (folders inside it aren't looked in).")
            return
        self.add(photos)

    def add(self, paths):
        known = {os.path.normcase(p) for p in self.paths}
        added = 0
        for path in paths:
            path = os.path.normpath(path)
            if os.path.normcase(path) in known:
                continue
            known.add(os.path.normcase(path))
            try:
                width, height = self.engine.photo_size(path)
                size = f"{width:,} x {height:,}"
            except self.engine.PhotoError:
                size = "can't be opened"
            self.paths.append(path)
            self.tree.insert("", "end", iid=path, values=(os.path.basename(path), size))
            added += 1
        if paths:
            self.remember(photo_folder=os.path.dirname(os.path.normpath(paths[0])))
        skipped = len(paths) - added
        text = f"Added {added} photo{'s' if added != 1 else ''}."
        if skipped:
            text += f" {skipped} {'was' if skipped == 1 else 'were'} already in the list."
        if self.result is not None and added:
            text += " Click Stitch the photos again to include the new ones."
        self.say(text)
        self.show_count()
        self.update_buttons()

    def remove_picked(self):
        if why := self.busy():
            self.say(why)
            return
        picked = self.tree.selection()
        if not picked:
            self.say("Click a photo in the list first (hold Ctrl to pick several), "
                     "then click Remove picked.")
            return
        for path in picked:
            self.tree.delete(path)
        gone = set(picked)
        self.paths = [p for p in self.paths if p not in gone]
        self.say(f"Took {len(picked)} photo{'s' if len(picked) != 1 else ''} off the list. "
                 "The files themselves are not touched.")
        self.show_count()
        self.update_buttons()

    def clear_all(self):
        if why := self.busy():
            self.say(why)
            return
        if not self.paths:
            self.say("The list is already empty.")
            return
        self.tree.delete(*self.tree.get_children())
        self.paths = []
        self.say("The list is empty. Your photos themselves are not touched.")
        self.show_count()
        self.update_buttons()

    # -- Step 2

    def kind_changed(self):
        self.remember(kind=self.kind.get())
        self.say(f"The photos will be joined as {self.engine.KIND_NAMES[self.kind.get()]}.")

    # -- Step 3

    def start_stitch(self):
        if self.job:
            self.say("The photos are already being joined. Wait for it to finish, or click Cancel.")
            return
        if self.saving:
            self.say(self.busy())
            return
        if len(self.paths) < 2:
            self.say("Add at least 2 overlapping photos first (Step 1).")
            return
        if self.unsaved and not messagebox.askyesno(
                TITLE, "The panorama you made isn't saved yet. Join the photos again anyway?\n\n"
                       "Click No, then Save the panorama, to keep it.", parent=self.root):
            self.say("Nothing was started. Save the panorama first if you want to keep it.")
            return
        # The words under the picture keep naming the photos of the picture shown
        # until this stitch has finished (stitch_done takes job.paths).
        job = self.jobs.StitchJob(list(self.paths), self.kind.get())
        try:
            job.start()
        except Exception as error:
            job.clean_up()
            self.say("The joining couldn't be started.")
            messagebox.showerror(TITLE, "The joining couldn't be started:\n\n" + str(error)
                                 + "\n\nClose other programs and try again.", parent=self.root)
            return
        self.job = job
        self.started = time.time()
        self.step_text = "Starting"
        self.progress.configure(text="Starting ...")
        self.say(f"Joining {len(self.paths)} photos. You can keep using the computer meanwhile.")
        self.update_buttons()

    def cancel(self):
        if not self.job:
            self.say("Nothing is being joined right now.")
            return
        self.job.cancel()
        self.job = None
        self.progress.configure(text="Stopped.")
        self.say("Stopped. Nothing was changed.")
        self.update_buttons()

    def poll(self):
        try:
            if self.job:
                for _ in range(100):
                    message = self.job.poll() if self.job else None
                    if message is None:
                        break
                    what, detail = message
                    if what == "step":
                        self.step_text = detail
                    elif what == "done":
                        self.stitch_done(detail)
                    else:
                        self.stitch_failed(detail)
                if self.job:
                    self.progress.configure(
                        text=f"{self.step_text} ...   {clock(time.time() - self.started)}")
            while self.results:
                self.results.pop(0)()
        finally:
            if self.poll_waiting:          # one timer only, even if poll() is called directly
                try:
                    self.root.after_cancel(self.poll_waiting)
                except tk.TclError:
                    pass
            self.poll_waiting = self.root.after(200, self.poll)

    def stop_polling(self):
        for waiting in (self.poll_waiting, self.resize_waiting):
            if waiting:
                try:
                    self.root.after_cancel(waiting)
                except tk.TclError:
                    pass

    def stitch_done(self, info):
        job, self.job = self.job, None
        took = clock(time.time() - self.started)
        try:
            picture = self.jobs.read_picture(info)
        except (OSError, ValueError, MemoryError) as error:
            job.clean_up()
            self.stitch_failed(f"The joined picture couldn't be handed to the window ({error}). "
                               "Try again, or try fewer photos.")
            return
        job.clean_up()
        self.result, self.info, self.stitched_paths = picture, info, job.paths
        self.saved_path, self.unsaved = None, True
        self.show_result()
        self.progress.configure(text=f"Finished in {took}.")
        self.say("Done. Look at the picture in Step 4, then click Save the panorama.")
        self.update_buttons()

    def stitch_failed(self, text):
        if self.job:
            self.job.clean_up()
            self.job = None
        self.progress.configure(text="The photos weren't joined.")
        self.say("The photos weren't joined. The message says why and what to try.")
        self.update_buttons()
        messagebox.showwarning(TITLE, text, parent=self.root)

    # -- Step 4

    def shown_picture(self):
        if self.trim_on.get() and self.info:
            return self.engine.crop(self.result, self.info["trim"])
        return self.result

    def show_result(self):
        picture = self.shown_picture()
        self.preview_source = self.engine.fit_for_preview(picture, 2400, 1400)
        self.draw_preview()
        height, width = picture.shape[:2]
        self.note.configure(text=describe_result(self.info, self.stitched_paths, width, height,
                                                 self.engine.KIND_NAMES))

    def draw_preview(self):
        self.resize_waiting = None
        if self.preview_source is None:
            return
        width = max(40, self.preview.winfo_width() - 6)
        height = max(40, self.preview.winfo_height() - 6)
        data = self.engine.preview_png(self.preview_source, width, height)
        self.preview_photo = tk.PhotoImage(master=self.root,
                                           data=base64.b64encode(data).decode("ascii"))
        self.preview.configure(image=self.preview_photo, text="")

    def preview_resized(self, event=None):
        if self.preview_source is None:
            return
        if self.resize_waiting:
            self.root.after_cancel(self.resize_waiting)
        self.resize_waiting = self.root.after(150, self.draw_preview)

    def trim_changed(self):
        self.remember(trim=self.trim_on.get())
        if self.result is None:
            self.say("The ragged edges will be trimmed." if self.trim_on.get()
                     else "The ragged edges will be left on.")
            return
        self.show_result()
        self.say("Showing the panorama with the ragged edges trimmed." if self.trim_on.get()
                 else "Showing the whole panorama, ragged edges and all.")

    def save_type_changed(self):
        self.remember(save_type=self.save_type.get())
        words = {ending: name for ending, _, name in SAVE_CHOICES}[self.save_type.get()]
        self.say(f"Save the panorama will save a {words}.")

    def save(self):
        if why := self.busy():
            self.say(why)
            return
        if self.result is None:
            self.say("There's nothing to save yet. Click Stitch the photos first (Step 3).")
            return
        ending = self.save_type.get()
        folder = self.settings.get("save_folder") or self.engine.default_save_folder()
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            folder = self.engine.pictures_folder()
        kinds = [(name, f"*{e}") for e, _, name in SAVE_CHOICES]
        chosen = [k for k in kinds if k[1] == f"*{ending}"]
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Save the panorama", initialdir=folder,
            initialfile=self.engine.default_file_name(folder, ending), defaultextension=ending,
            filetypes=chosen + [k for k in kinds if k not in chosen])
        if not path:
            self.adjust_after_save = False
            self.say("Not saved. Click Save the panorama when you're ready.")
            return
        path = os.path.normpath(path)
        picture = self.shown_picture()
        height, width = picture.shape[:2]
        if (os.path.splitext(path)[1].lower() in (".jpg", ".jpeg")
                and max(height, width) > self.engine.JPEG_MAX):
            if not messagebox.askyesno(
                    TITLE, f"This panorama is {width:,} x {height:,} pixels, which is too big for "
                           f"JPG ({self.engine.JPEG_MAX:,} pixels on a side at most).\n\n"
                           "Save it as PNG instead? PNG keeps full quality but makes a bigger file.",
                    parent=self.root):
                self.adjust_after_save = False
                self.say("Not saved. Pick PNG or TIFF under Save as, then click Save the panorama.")
                return
            path = self.engine.free_name(os.path.splitext(path)[0] + ".png")
        self.saving = True
        self.saving_picture = picture
        self.say(f"Saving {os.path.basename(path)} ...")
        self.update_buttons()
        threading.Thread(target=self.save_in_background, args=(picture, path), daemon=True).start()

    def save_in_background(self, picture, path):
        try:
            self.engine.save_image(picture, path)
        except self.engine.SaveError as error:
            text = str(error)
            self.results.append(lambda: self.save_failed(text))
        except Exception:
            text = "Something went wrong while saving:\n\n" + traceback.format_exc()[-1200:]
            self.results.append(lambda: self.save_failed(text))
        else:
            self.results.append(lambda: self.save_done(path))

    def save_done(self, path):
        self.saving = False
        self.saved_path, self.unsaved = path, False
        self.saved_picture = self.saving_picture
        self.remember(save_folder=os.path.dirname(path))
        self.say(f"Saved {os.path.basename(path)} in {os.path.dirname(path)}. "
                 "Click Open it to look at it.")
        self.update_buttons()
        if self.adjust_after_save:
            self.adjust_after_save = False
            self.open_adjust()

    def save_failed(self, text):
        self.saving = False
        self.adjust_after_save = False
        self.say("The panorama wasn't saved. The message says why.")
        self.update_buttons()
        messagebox.showerror(TITLE, text, parent=self.root)

    def open_saved(self):
        if not (self.saved_path and os.path.exists(self.saved_path)):
            self.say("Save the panorama first. Then Open it shows it.")
            self.update_buttons()
            return
        try:
            os.startfile(self.saved_path)
            self.say(f"Opening {os.path.basename(self.saved_path)} ...")
        except OSError as error:
            self.say(f"Windows couldn't open it ({error.strerror or error}).")

    def open_folder(self):
        if not (self.saved_path and os.path.exists(self.saved_path)):
            self.say("Save the panorama first. Then Open the folder shows where it is.")
            self.update_buttons()
            return
        subprocess.Popen(["explorer", "/select,", self.saved_path])
        self.say(f"Opening {os.path.dirname(self.saved_path)} ...")

    # -- Step 5

    def open_adjust(self):
        if why := self.busy():
            self.say(why)
            return
        if self.result is None:
            self.say("Make a panorama first (Step 3), then adjust it.")
            return
        adjusting = self.adjust is not None and self.adjust.exists()
        if adjusting and self.adjust_shows_saved():
            self.adjust.bring_to_front()
            self.say("The Adjust window is already open.")
            return
        if self.unsaved or not (self.saved_path and os.path.exists(self.saved_path)):
            if not messagebox.askyesno(
                    TITLE, "Save the plain panorama first? Then you keep it as well as the "
                           "adjusted one.", parent=self.root):
                self.say("The Adjust window opens once the plain panorama is saved.")
                return
            self.adjust_after_save = True
            self.save()
            return
        if adjusting:
            # It still shows an earlier panorama. Close it first (it asks if its
            # adjustments aren't saved), then open this one.
            self.adjust.bring_to_front()
            if not self.adjust.close():
                self.say("The Adjust window still has the earlier panorama. Save its "
                         "adjustments there (Save as) or close it, then click this again.")
                return
        import photo_stitcher_adjust
        picture = self.saved_picture if self.saved_picture is not None else self.shown_picture()
        self.adjust = photo_stitcher_adjust.AdjustWindow(self.root, picture, self.saved_path,
                                                         SAVE_CHOICES, self.save_type.get())
        self.say("The Adjust window is open. Your plain panorama stays as it is.")

    def adjust_shows_saved(self):
        """True if the open Adjust window holds the panorama saved now (not one
        from before the last stitch)."""
        if self.unsaved or self.adjust.plain_path != self.saved_path:
            return False
        return self.saved_picture is None or self.adjust.picture is self.saved_picture

    # -- closing

    def close(self):
        if self.adjust is not None and self.adjust.exists() and not self.adjust.close():
            self.say("Not closed: the Adjust window is still open.")
            return
        if self.job:
            if not messagebox.askyesno(TITLE, "The photos are still being joined. Stop and close?",
                                       parent=self.root):
                self.say("Still joining. Not closed.")
                return
            self.job.cancel()
            self.job = None
        elif self.saving:
            self.say("Wait a moment: the panorama is being saved. Then close.")
            return
        elif self.unsaved and not messagebox.askyesno(
                TITLE, "The panorama isn't saved yet. Close without saving it?", parent=self.root):
            self.say("Not closed. Click Save the panorama to keep it.")
            return
        self.stop_polling()
        self.root.destroy()


def self_test(folder):
    """For the Setup's test on a clean PC ("Test in Sandbox"): joins made-up
    photos and opens an iPhone photo using only the parts packed into the
    program, then writes what happened to self-test.txt in the folder."""
    lines = []
    try:
        import numpy as np
        import cv2
        import photo_stitcher_engine as engine
        import photo_stitcher_job as jobs
        os.makedirs(folder, exist_ok=True)
        rng = np.random.default_rng(7)
        scene = np.full((400, 1100, 3), 150, np.uint8)
        scene[:, :, 0] = np.linspace(60, 200, 1100, dtype=np.uint8)
        for _ in range(170):
            x, y = int(rng.integers(0, 1100)), int(rng.integers(0, 400))
            color = tuple(int(c) for c in rng.integers(20, 255, 3))
            cv2.circle(scene, (x, y), int(rng.integers(5, 28)), color, -1)
            cv2.putText(scene, "AB7KX"[int(rng.integers(0, 5))], (x + 9, y - 9),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.1, color[::-1], 2)
        paths = []
        for number, left in enumerate((600, 0, 300)):
            path = os.path.join(folder, f"self-test piece {number}.jpg")
            engine.save_image(scene[:, left:left + 500], path)
            paths.append(path)
        from PIL import Image
        engine._open_heic_with_pillow()
        heic = os.path.join(folder, "self-test iphone.heic")
        Image.fromarray(np.ascontiguousarray(scene[:, :300, ::-1])).save(heic, quality=80)
        lines.append(f"HEIC photo opens: {engine.load_photo(heic).shape[:2] == (400, 300)}")
        job = jobs.StitchJob(paths, engine.TURNING)
        job.start()
        started, message = time.time(), None
        while time.time() - started < 180:
            message = job.poll()
            if message and message[0] != "step":
                break
            time.sleep(0.1)
        if message and message[0] == "done":
            info = message[1]
            picture = engine.crop(jobs.read_picture(info), info["trim"])
            job.clean_up()
            engine.save_image(picture, os.path.join(folder, "self-test panorama.png"))
            height, width = picture.shape[:2]
            lines.append(f"Joined {len(info['used'])} of 3 photos in the background, "
                         f"{width} x {height} pixels, in {time.time() - started:.1f} s")
            lines.append(f"Join worked: {len(info['used']) == 3 and abs(width - 1100) < 60}")
        else:
            job.cancel()
            lines.append(f"Join worked: False ({message})")
        if message and message[0] == "done":
            import photo_stitcher_adjust as adjust
            steps = adjust.Adjustments()
            steps.turn(clockwise=True)
            steps.set_straighten(3)
            steps.square([(0.05, 0.05), (0.95, 0.03), (0.97, 0.95), (0.03, 0.97)])
            steps.crop((0.1, 0.1, 0.9, 0.9))
            steps.size = 0.5
            adjusted = steps.apply(picture)
            engine.save_image(adjusted, os.path.join(folder, "self-test adjusted.jpg"))
            height, width = adjusted.shape[:2]
            lines.append(f"Adjust worked: {height > width > 50}  ({width} x {height} pixels)")
        lines.append(f"Pictures folder: {engine.pictures_folder()}")
    except Exception:
        lines.append("Self-test crashed:\n" + traceback.format_exc())
    with open(os.path.join(folder, "self-test.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def main():
    multiprocessing.freeze_support()        # the background stitching starts this program again
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        self_test(sys.argv[2])
        return
    sys.excepthook = show_crash
    sharp_on_big_screens()
    install_missing()
    import photo_stitcher_engine as engine
    import photo_stitcher_job as jobs

    jobs.clean_old_folders()
    root = tk.Tk()
    root.report_callback_exception = show_crash
    set_icon(root)
    app = App(root, engine, jobs)
    app.fit_to_screen()
    app.bring_to_front()
    root.mainloop()


if __name__ == "__main__":
    main()
