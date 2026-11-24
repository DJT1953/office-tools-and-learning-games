"""Dating Photographs: reads the dates written on old photographs (on the front,
the white border or the back) and renames each photo by its date, like
1998-10-01.jpg, in a "Dated Photographs" folder.

Step 1 pick the folder, Step 2 find the dates, Step 3 check them (change any,
untick any to leave alone), Step 4 rename and move. "Oct-98" and "98-Oct" both
become 1998-10-01: a date with no day gets the 1st. Photos with no date stay
where they are. Undo puts the last batch back, names and dates and all.

Start it by double-clicking Run_DatingPhotos.bat (or its RUN shortcut).
Requirements (offered on first start): Pillow, rapidocr and onnxruntime (the
handwriting reader, which does all the reading) and pytesseract, plus the
reader's data (about 35 MB, fetched once). Tesseract OCR, if it is on the PC,
adds a pass for printed words but isn't needed, and the Setup leaves it out:
https://github.com/UB-Mannheim/tesseract/wiki
"""

from __future__ import annotations

import ctypes
import importlib
import importlib.util
import json
import os
import queue
import site
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

TITLE = "Dating Photographs"
VERSION = "1.0"
# The "Check for updates" button opens this page, which has the newest version of every program.
UPDATES_PAGE = "https://djt1953.github.io/office-tools-and-learning-games/"
NEEDS = {   # module: (pip name, what it's for, size in MB)
    "PIL": ("pillow", "Pillow (opens the photos)", 7),
    "rapidocr": ("rapidocr>=3.10,<4", "RapidOCR (reads handwriting)", 60),
    "onnxruntime": ("onnxruntime", "onnxruntime (runs the handwriting reader)", 15),
    "pytesseract": ("pytesseract", "pytesseract (reads printed words with Tesseract)", 1),
}
READER_DATA = ("ch_PP-OCRv5_det_mobile.onnx", "en_PP-OCRv5_rec_mobile.onnx", "PP-OCRv6_rec_small.onnx")
SETTINGS = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), TITLE, "settings.json")
TICKED, UNTICKED = "\u2611", "\u2610"
LIST_ROOM = "1.2i"      # room the Step 3 list keeps even on a short screen (about 3 photos)
PHOTO_ROOM = "0.75i"    # room the photo keeps on a short screen
HERE = os.path.dirname(os.path.abspath(__file__))


# ------------------------------------------------------------------ start-up checks

def show_crash(kind, value, tb):
    text = "".join(traceback.format_exception(kind, value, tb))
    try:
        messagebox.showerror(TITLE, "Something went wrong and Dating Photographs has to close.\n\n"
                             + text[-1500:])
    finally:
        sys.exit(1)


def has_module(name):
    return importlib.util.find_spec(name) is not None


def console_python(executable=None):
    """The Python that shows a window. Run_DatingPhotos.bat starts pythonw.exe,
    which never opens one, so pip's progress would be invisible under it."""
    executable = executable or sys.executable
    folder, name = os.path.split(executable)
    if name.lower() == "pythonw.exe" and os.path.isfile(os.path.join(folder, "python.exe")):
        return os.path.join(folder, "python.exe")
    return executable


def reader_data_missing():
    """The handwriting reader's data files not on this PC yet ([] in a built copy)."""
    spec = importlib.util.find_spec("rapidocr")
    if spec is None or not spec.submodule_search_locations:
        return list(READER_DATA)
    models = os.path.join(list(spec.submodule_search_locations)[0], "models")
    return [name for name in READER_DATA if not os.path.isfile(os.path.join(models, name))]


def install_missing():
    """Double-clicking may start a Python that lacks the add-ons, so offer to
    add them, and fetch the handwriting reader's data the first time. A built
    copy (from Setup) already has everything inside it."""
    if getattr(sys, "frozen", False):
        return
    missing = [need for module, need in NEEDS.items() if not has_module(module)]
    data = bool(missing) or bool(reader_data_missing())
    if not missing and not data:
        return
    root = tk.Tk()
    root.withdraw()
    names = " and ".join(what for _, what, _ in missing)
    size = sum(mb for *_, mb in missing) + 35
    words = (f"Dating Photographs still needs {names}, and the handwriting reader's data"
             if missing else "Dating Photographs still needs the handwriting reader's data")
    if not messagebox.askyesno(TITLE, f"{words}.\n\nDownload and install now? This only happens once. "
                               f"It is about {size} MB; a window shows the progress."):
        sys.exit()
    root.destroy()
    python = console_python()
    if missing:
        command = [python, "-m", "pip", "install", *(pip for pip, _, _ in missing)]
        if subprocess.run(command, creationflags=subprocess.CREATE_NEW_CONSOLE).returncode:
            raise RuntimeError(f"Could not install {names}. Check the internet connection and try again.")
        site.addsitedir(site.getusersitepackages())
        importlib.invalidate_caches()
    fetch = (f"import sys; sys.path.insert(0, {HERE!r}); import photo_text_reader as r; "
             "print('Fetching the handwriting reader data...'); r.Reader(use_tesseract=False).warm_up(); "
             "print('Done.')")
    if subprocess.run([python, "-c", fetch], creationflags=subprocess.CREATE_NEW_CONSOLE).returncode:
        raise RuntimeError("Could not fetch the handwriting reader's data. Check the internet "
                           "connection and try again.")


def no_lost_output():
    """pythonw.exe has nowhere to print; give it somewhere so nothing that prints can fail."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def sharp_on_big_screens():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        pass


def set_icon(root):
    here = getattr(sys, "_MEIPASS", None) or os.path.join(HERE, "Installers", TITLE)
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


def pictures_folder():
    """Windows' Pictures folder; the plain one if Windows still points at a OneDrive
    folder that OneDrive isn't there to keep."""
    plain = os.path.join(os.path.expanduser("~"), "Pictures")
    path = None
    try:
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                        ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

        pictures = GUID(0x33E28130, 0x4E1E, 0x4676,
                        (ctypes.c_ubyte * 8)(0x83, 0x5A, 0x98, 0x39, 0x5C, 0x3B, 0xC3, 0xBB))
        found = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(pictures), 0, None,
                                                      ctypes.byref(found)) == 0:
            path = found.value
            ctypes.windll.ole32.CoTaskMemFree(found)
    except (AttributeError, OSError):
        path = None
    onedrive_gone = not any(os.path.isfile(os.path.expandvars(p)) for p in (
        r"%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe", r"%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"))
    if path and os.path.isdir(path) and not ("onedrive" in path.lower() and onedrive_gone
                                             and os.path.isdir(plain)):
        return path
    return plain if os.path.isdir(plain) else os.path.expanduser("~")


def clock(seconds):
    seconds = max(0, int(seconds))
    if seconds < 60:
        return "under a minute"
    minutes = round(seconds / 60)
    return f"about {minutes} minute{'s' if minutes != 1 else ''}"


# ------------------------------------------------------------------ reading in the background

class Reading(threading.Thread):
    """Finds the dates for a list of photos, one at a time, in the background."""

    def __init__(self, photos, root, month_first, make_reader=None):
        super().__init__(daemon=True)
        self.photos, self.root, self.month_first = photos, root, month_first
        self.make_reader = make_reader       # tests hand in a stand-in reader
        self.messages = queue.Queue()
        self.stopping = threading.Event()

    def run(self):
        try:
            import dating_photos_find as find
            import photo_text_reader
            self.messages.put(("step", "Getting the handwriting reader ready..."))
            reader = (self.make_reader or photo_text_reader.Reader)()
            for number, photo in enumerate(self.photos, 1):
                if self.stopping.is_set():
                    break
                self.messages.put(("reading", number, photo))
                finding = find.find(reader, photo, self.root, month_first=self.month_first,
                                    stop=self.stopping.is_set)
                if self.stopping.is_set() and not finding.clues and not finding.problem:
                    break               # stopped half way: leave it as not read yet
                self.messages.put(("found", finding))
            self.messages.put(("done", self.stopping.is_set()))
        except Exception:
            self.messages.put(("failed", traceback.format_exc()))


class Moving(threading.Thread):
    """Renames and moves the ticked photos in the background."""

    def __init__(self, root, moves, clues, set_date_taken, printed=None):
        super().__init__(daemon=True)
        self.args = root, moves, clues, set_date_taken
        self.printed = printed or {}
        self.result = None
        self.error = None

    def run(self):
        try:
            import dating_photos_files as files
            self.result = files.apply_moves(*self.args[:3], set_date_taken=self.args[3], printed=self.printed)
        except Exception:
            self.error = traceback.format_exc()


# ------------------------------------------------------------------ the window

class Row:
    """One photo in the Step 3 list."""

    def __init__(self, photo):
        self.photo = photo
        self.finding = None          # what Step 2 found, once read
        self.date = None             # the date it will get
        self.kind = None             # how exact that date is (day, month, year...)
        self.clue = None             # where that date came from (None if typed)
        self.ticked = False
        self.typed = False           # the date was typed in by hand
        self.state = "waiting"       # waiting, read, moved
        self.moved_to = ""


class App:
    def __init__(self, root, dates, files):
        self.root, self.dates, self.files = root, dates, files
        self.settings = load_settings()
        self.folder = None
        self.rows = {}                   # list item id -> Row
        self.order = []                  # item ids in list order
        self.findings = {}               # front path -> Finding, kept across Undo and re-listing
        self.reading = None
        self.moving = None
        self.started = 0.0
        self.read_count = 0
        self.preview_images = []
        self.preview_cache = {}
        self.current = None              # item id shown on the right
        self.choices = {}                # front path -> (date, kind) typed in, or None if skipped
        self.make_reader = None          # None: the real handwriting reader
        self.stamp = importlib.import_module("dating_photos_stamp")

        get = self.settings.get
        self.subfolders = tk.BooleanVar(value=get("subfolders") is True)
        self.every_second = tk.BooleanVar(value=get("every_second") is True)
        self.day_first = tk.BooleanVar(value=get("day_first") is True)
        self.set_taken = tk.BooleanVar(value=get("set_taken", True) is not False)
        self.rename = tk.BooleanVar(value=get("rename", True) is not False)
        self.keep_name = tk.BooleanVar(value=get("keep_name") is True)
        self.print_date = tk.BooleanVar(value=get("print_date") is True)
        style = get("print_style")
        self.print_style = tk.StringVar(value=style if style in self.stamp.STYLES else self.stamp.DEFAULT_STYLE)
        self.year_folders = tk.BooleanVar(value=get("year_folders") is True)
        self.date_text = tk.StringVar()
        self.other = tk.StringVar()

        root.title(TITLE)
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        style.configure("Big.TButton", font=("Segoe UI", 13, "bold"), padding=(12, 8), foreground="black")
        style.configure("TButton", foreground="black", font=("Segoe UI", 11))
        style.configure("TCheckbutton", font=("Segoe UI", 11))
        style.configure("TLabelframe.Label", font=("Segoe UI", 12, "bold"))
        style.configure("Note.TLabel", font=("Segoe UI", 11))
        style.configure("Found.TLabel", font=("Segoe UI", 12))
        list_font = tkfont.Font(root, family="Segoe UI", size=11)
        style.configure("Treeview", font=list_font, rowheight=int(list_font.metrics("linespace") * 1.4))
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))

        page = ttk.Frame(root, padding=14)
        page.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(2, weight=1)

        heading = ttk.Frame(page)
        heading.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(heading, text=TITLE, font=("Segoe UI", 20, "bold")).pack(side="left")
        self.updates_button = ttk.Button(heading, text="Check for updates", command=self.check_for_updates)
        self.updates_button.pack(side="right", anchor="n")
        ttk.Label(page, style="Note.TLabel",
                  text="Reads the dates written on your photos (front, border or back) and renames each "
                       "one by its date, like 1998-10-01.jpg, in a \u201cDated Photographs\u201d folder."
                  ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        left = ttk.Frame(page)
        left.grid(row=2, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        # Step 1
        pick = ttk.LabelFrame(left, text="Step 1 - Pick the photos", padding=10)
        pick.grid(row=0, sticky="ew", pady=(0, 6))
        pick.columnconfigure(0, weight=1)
        self.pick_button = self.big(pick, "Pick a folder", self.pick_folder, 0, 0)
        self.folder_label = ttk.Label(pick, style="Note.TLabel", wraplength=470, justify="left",
                                      text="No folder yet.")
        self.folder_label.grid(row=1, sticky="w", pady=(4, 4))
        ttk.Checkbutton(pick, text="Include photos in sub-folders", variable=self.subfolders,
                        command=self.options_changed).grid(row=2, sticky="w")
        ttk.Checkbutton(pick, text="Every second photo is the back of the one before it",
                        variable=self.every_second, command=self.options_changed).grid(row=3, sticky="w")
        ttk.Label(pick, style="Note.TLabel", wraplength=470, justify="left",
                  text="(A back named like IMG_0001_b.jpg is found anyway.)").grid(
            row=4, sticky="w", padx=(24, 0))

        # Step 2
        run = ttk.LabelFrame(left, text="Step 2 - Find the dates", padding=10)
        run.grid(row=1, sticky="ew", pady=6)
        run.columnconfigure((0, 1), weight=1)
        self.find_button = self.big(run, "Find the dates", self.start_reading, 0, 0)
        self.stop_button = self.big(run, "Stop", self.stop_reading, 0, 1)
        self.progress = ttk.Progressbar(run, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=(6, 0))
        self.progress_label = ttk.Label(run, style="Note.TLabel", wraplength=470, justify="left",
                                        text="Each photo takes a few seconds.")
        self.progress_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Checkbutton(run, text="Day before month in dates like 3/4/91 (not the US way)",
                        variable=self.day_first, command=self.remember_options).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Step 3: the list
        check = ttk.LabelFrame(left, text="Step 3 - Check the dates", padding=10)
        check.grid(row=2, sticky="nsew", pady=(6, 0))
        check.columnconfigure(0, weight=1)
        check.rowconfigure(0, weight=1, minsize=LIST_ROOM)    # the list keeps room for a few photos
        holder = ttk.Frame(check)
        holder.grid(row=0, column=0, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(holder, columns=("use", "photo", "date", "how"), show="headings",
                                 height=10, selectmode="browse")
        for column, words, width, stretch in (("use", "Move", 56, False), ("photo", "Photo", 190, True),
                                              ("date", "Date", 110, False), ("how", "Found", 120, True)):
            self.tree.heading(column, text=words, anchor="w")
            self.tree.column(column, width=width, anchor="w", stretch=stretch)
        self.tree.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.tag_configure("moved", foreground="#707070")
        self.tree.tag_configure("unsure", foreground="#9a5a00")
        self.tree.tag_configure("none", foreground="#b00020")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_selected())
        self.tree.bind("<Button-1>", self.list_clicked)
        self.tree.bind("<space>", lambda e: self.toggle(self.tree.focus()))
        self.tree.bind("<Delete>", lambda e: self.skip_photo())
        self.tree.bind("<Return>", lambda e: self.focus_date_box())
        self.count_label = ttk.Label(check, style="Note.TLabel", wraplength=470, justify="left",
                                     text="Click a photo to see it on the right. Untick any you "
                                          "don't want moved.")
        self.count_label.grid(row=1, sticky="w", pady=(6, 0))

        # Step 3: the photo
        right = ttk.Frame(page)
        right.grid(row=2, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        look = ttk.LabelFrame(right, text="Step 3 - The photo and its date", padding=10)
        look.grid(row=0, sticky="nsew")
        look.columnconfigure(0, weight=1)
        # on a short screen (a laptop's 768) the photo gets smaller, so the date box,
        # Other dates found and the hint under them all stay in sight
        look.rowconfigure(0, weight=1, minsize=PHOTO_ROOM)
        self.canvas = tk.Canvas(look, background="#e4e4e4", highlightthickness=0, width=600, height=360)
        self.canvas.grid(row=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda e: self.draw_preview())
        self.found_label = ttk.Label(look, style="Found.TLabel", wraplength=640, justify="left",
                                     text="Pick a folder and find the dates; then click a photo in the list.")
        self.found_label.grid(row=1, sticky="w", pady=(8, 4))
        entry_row = ttk.Frame(look)
        entry_row.grid(row=2, sticky="ew")
        ttk.Label(entry_row, text="Date:", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.date_box = ttk.Entry(entry_row, textvariable=self.date_text, width=14, font=("Segoe UI", 13))
        self.date_box.pack(side="left", padx=(6, 8))
        self.date_box.bind("<Return>", lambda e: self.use_date())
        self.use_button = ttk.Button(entry_row, text="Use this date", style="Big.TButton",
                                     command=self.use_date)
        self.use_button.pack(side="left", padx=(0, 6))
        self.skip_button = ttk.Button(entry_row, text="Skip this photo", style="Big.TButton",
                                      command=self.skip_photo)
        self.skip_button.pack(side="left")
        other_row = ttk.Frame(look)
        other_row.grid(row=3, sticky="ew", pady=(6, 0))
        ttk.Label(other_row, text="Other dates found:", style="Note.TLabel").pack(side="left")
        self.other_box = ttk.Combobox(other_row, textvariable=self.other, state="readonly", width=48,
                                      font=("Segoe UI", 11))
        self.other_box.pack(side="left", padx=(6, 0), fill="x", expand=True)
        self.other_box.bind("<<ComboboxSelected>>", lambda e: self.other_picked())
        self.words_label = ttk.Label(look, style="Note.TLabel", wraplength=640, justify="left",
                                     foreground="#404040",
                                     text="Type a date like 1998-10-01, Oct 98 or 1998, then click Use this date "
                                          "(or press Enter).")
        self.words_label.grid(row=4, sticky="w", pady=(6, 0))

        # Step 4
        go = ttk.LabelFrame(right, text="Step 4 - Date the photos and move them", padding=10)
        go.grid(row=1, sticky="ew", pady=(8, 0))
        go.columnconfigure((0, 1, 2), weight=1)
        options = ttk.Frame(go)
        options.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(options, text="Rename each photo by its date (1998-10-01.jpg)",
                        variable=self.rename, command=self.go_options_changed).grid(row=0, sticky="w")
        self.keep_name_box = ttk.Checkbutton(options, text="Keep the old name after the date "
                                                           "(1998-10-01 IMG_0001.jpg)",
                                             variable=self.keep_name, command=self.go_options_changed)
        self.keep_name_box.grid(row=1, sticky="w", padx=(24, 0))
        printing = ttk.Frame(options)
        printing.grid(row=2, sticky="w")
        ttk.Checkbutton(printing, text="Print the date on each photo in orange, like a film camera:",
                        variable=self.print_date, command=self.go_options_changed).pack(side="left")
        self.style_box = ttk.Combobox(printing, textvariable=self.print_style, state="readonly", width=11,
                                      values=list(self.stamp.STYLES), font=("Segoe UI", 11))
        self.style_box.pack(side="left", padx=(6, 0))
        self.style_box.bind("<<ComboboxSelected>>", lambda e: self.go_options_changed())
        ttk.Label(options, style="Note.TLabel",
                  text=f"An untouched copy of each is kept in \u201c{self.files.ORIGINALS}\u201d.").grid(
            row=3, sticky="w", padx=(24, 0))
        ttk.Checkbutton(options, text="Set each photo's \u201cDate taken\u201d to its date (photo programs "
                                      "sort by it; scans otherwise carry the scan date)",
                        variable=self.set_taken, command=self.go_options_changed).grid(row=4, sticky="w")
        ttk.Checkbutton(options, text="Put them in a folder for each year (Dated Photographs\\1998)",
                        variable=self.year_folders, command=self.go_options_changed).grid(row=5, sticky="w")
        self.move_button = self.big(go, "Rename and move", self.rename_and_move, 1, 0)
        self.undo_button = self.big(go, "Undo the last move", self.undo, 1, 1)
        self.open_button = self.big(go, "Open Dated Photographs", self.open_dated, 1, 2)

        self.status = ttk.Label(page, style="Note.TLabel", wraplength=1200, justify="left",
                                text="Ready. Start with Step 1: click Pick a folder.")
        self.status.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.update_buttons()
        self.poll_waiting = root.after(150, self.poll)

    @staticmethod
    def big(parent, text, command, row, column):
        button = ttk.Button(parent, text=text, style="Big.TButton", command=command)
        button.grid(row=row, column=column, sticky="ew", padx=4, pady=3)
        return button

    def fit_to_screen(self):
        root = self.root
        root.update_idletasks()
        room = root.winfo_screenheight() - int(110 * root.winfo_fpixels("1i") / 96)
        rows = int(self.tree.cget("height"))
        while root.winfo_reqheight() > room and rows > 4:
            rows -= 1
            self.tree.configure(height=rows)
            root.update_idletasks()
        scale = root.winfo_fpixels("1i") / 96
        width = min(max(root.winfo_reqwidth(), int(1240 * scale)), root.winfo_screenwidth() - 20)
        height = min(max(root.winfo_reqheight(), int(760 * scale)), room)
        root.geometry(f"{width}x{height}+{max(0, (root.winfo_screenwidth() - width) // 2)}"
                      f"+{max(0, (room - height) // 2)}")
        root.minsize(min(width, int(980 * scale)), min(height, root.winfo_reqheight()))

    def bring_to_front(self):
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(400, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    def say(self, text):
        self.status.configure(text=text)

    def remember_options(self):
        self.settings.update(subfolders=self.subfolders.get(), every_second=self.every_second.get(),
                             day_first=self.day_first.get(), set_taken=self.set_taken.get(),
                             rename=self.rename.get(), keep_name=self.keep_name.get(),
                             print_date=self.print_date.get(), print_style=self.print_style.get(),
                             year_folders=self.year_folders.get())
        save_settings(self.settings)

    def go_options_changed(self):
        self.remember_options()
        self.update_buttons()
        self.draw_preview()
        if self.print_date.get():
            self.say(f"The date will be printed like {self.print_style.get()} (only what's known: a date "
                     "with no day prints without one). The photo on the right shows how it will look.")
        elif not self.rename.get():
            self.say("The photos will keep their names. Tick Rename or Print the date, or both.")
        else:
            self.say("Options saved.")

    def go_words(self):
        """The Step 4 button's words, for the choices ticked."""
        rename, printing = self.rename.get(), self.print_date.get()
        if rename and printing:
            return "Rename, print dates, move"
        if printing:
            return "Print dates and move"
        return "Rename and move"

    def busy(self):
        if self.reading:
            return "The photos are being read. Wait for it to finish, or click Stop."
        if self.moving:
            return "Wait a moment: the photos are being moved."
        return ""

    def update_buttons(self):
        self.stop_button.state(["!disabled"] if self.reading else ["disabled"])
        # gray only while they can't apply; ticking Rename or Print turns them black again
        self.keep_name_box.state(["!disabled"] if self.rename.get() else ["disabled"])
        self.style_box.state(["readonly", "!disabled"] if self.print_date.get() else ["disabled"])
        self.move_button.configure(text=self.go_words())

    def check_for_updates(self):
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

    def pick_folder(self):
        if why := self.busy():
            self.say(why)
            return
        start = self.folder or self.settings.get("folder") or pictures_folder()
        if not os.path.isdir(start):
            start = pictures_folder()
        folder = filedialog.askdirectory(parent=self.root, title="Pick the folder of photographs",
                                         initialdir=start, mustexist=True)
        if not folder:
            self.say("No folder picked." if not self.folder else f"Still using {self.folder}.")
            return
        folder = os.path.normpath(folder)
        if self.inside_dated(folder):
            self.say(f"That is the “{self.files.DEST}” folder this program made (or a folder in it). "
                     "Click Pick a folder again and pick the folder above it, the one your photos were in.")
            return
        self.folder = folder
        self.settings["folder"] = self.folder
        save_settings(self.settings)
        self.list_photos()

    def above_dated(self, folder):
        """The folder that holds the "Dated Photographs" this program made (the one
        with its list of moves), if the folder is that one or inside it; else None."""
        here = os.path.normpath(folder)
        while True:
            if os.path.basename(here).lower() == self.files.DEST.lower() \
                    and os.path.isfile(os.path.join(here, self.files.LOG_NAME)):
                return os.path.dirname(here)
            if os.path.dirname(here) == here:
                return None
            here = os.path.dirname(here)

    def inside_dated(self, folder):
        return self.above_dated(folder) is not None

    def options_changed(self):
        self.remember_options()
        if self.reading:
            self.say("That takes effect after the reading stops. Click Stop, then Find the dates again.")
            return
        if self.folder:
            self.list_photos()

    def list_photos(self):
        """Fill the list from the folder, keeping anything already found."""
        try:
            paths = self.files.list_pictures(self.folder, self.subfolders.get())
        except OSError as error:
            self.say(f"That folder can't be opened: {error}")
            return
        photos = self.files.group_photos(paths, self.every_second.get())
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.order = []
        for photo in photos:
            row = Row(photo)
            finding = self.findings.get(photo.front)
            if finding is not None and self.same_sides(finding.photo, photo):
                self.take_finding(row, finding)
            self.apply_choice(row)
            item = self.tree.insert("", "end")
            self.rows[item] = row
            self.order.append(item)
            self.refresh_row(item)
        backs = sum(1 for p in photos if p.back)
        words = f"{self.folder}\n{len(photos)} photo{'s' if len(photos) != 1 else ''}"
        if backs:
            words += f", {backs} with {'its' if backs == 1 else 'their'} back scanned too"
        self.folder_label.configure(text=words + ".")
        self.current = None
        self.draw_preview()
        self.show_counts()
        if not photos:
            self.say("There are no photos in that folder" + (" or its sub-folders." if self.subfolders.get()
                                                              else ". Tick Include photos in sub-folders if "
                                                                   "they are in folders inside it."))
        else:
            self.say("Now Step 2: click Find the dates." if self.unread() else
                     "These were read before. Check them in Step 3, then click Rename and move.")
        if self.order:
            self.tree.selection_set(self.order[0])
            self.tree.focus(self.order[0])

    @staticmethod
    def same_sides(a, b):
        return a.front == b.front and a.back == b.back

    # -- Step 2

    def unread(self):
        return [item for item in self.order if self.rows[item].state == "waiting"]

    def start_reading(self):
        if why := self.busy():
            self.say(why)
            return
        if not self.folder:
            self.say("First pick a folder in Step 1.")
            return
        items = self.unread()
        if not items:
            self.say("Every photo in the list has been read. Check the dates in Step 3, or pick another folder.")
            return
        self.reading_items = items
        self.reading = Reading([self.rows[i].photo for i in items], self.folder, not self.day_first.get(),
                               self.make_reader)
        self.reading.start()
        self.started = time.time()
        self.read_count = 0
        self.progress.configure(maximum=len(items), value=0)
        self.progress_label.configure(text="Getting the handwriting reader ready...")
        self.say(f"Reading {len(items)} photo{'s' if len(items) != 1 else ''}. You can check the ones "
                 "already read while you wait.")
        self.update_buttons()

    def stop_reading(self):
        if not self.reading:
            self.say("Nothing is being read now.")
            return
        self.reading.stopping.set()
        self.progress_label.configure(text="Stopping after this photo...")
        self.say("Stopping. Find the dates carries on from where it stopped.")

    def poll(self):
        try:
            if self.reading:
                self.read_messages()
            if self.moving and not self.moving.is_alive():
                self.moving_done()
        finally:
            self.poll_waiting = self.root.after(150, self.poll)

    def read_messages(self):
        reading = self.reading
        while True:
            try:
                message = reading.messages.get_nowait()
            except queue.Empty:
                return
            kind = message[0]
            if kind == "step":
                self.progress_label.configure(text=message[1])
            elif kind == "reading":
                number, photo = message[1], message[2]
                left = ""
                if self.read_count:
                    each = (time.time() - self.started) / self.read_count
                    left = f" - {clock(each * (len(reading.photos) - number + 1))} left"
                self.progress_label.configure(text=f"Reading {number} of {len(reading.photos)}: "
                                                   f"{photo.name()}{left}")
            elif kind == "found":
                finding = message[1]
                self.read_count += 1
                self.progress.configure(value=self.read_count)
                self.findings[finding.photo.front] = finding
                item = self.item_for(finding.photo.front)
                if item:
                    self.take_finding(self.rows[item], finding)
                    self.apply_choice(self.rows[item])
                    self.refresh_row(item)
                    if item == self.current:
                        self.show_selected(force=True)
                self.show_counts()
            elif kind == "done":
                self.reading = None
                self.update_buttons()
                found = sum(1 for i in self.order if self.rows[i].date)
                read = sum(1 for i in self.order if self.rows[i].state != "waiting")
                if message[1]:
                    self.progress_label.configure(text=f"Stopped. {read} of {len(self.order)} read.")
                    self.say("Stopped. Click Find the dates to read the rest.")
                else:
                    self.progress_label.configure(text=f"Done: {read} photo{'s' if read != 1 else ''} read in "
                                                       f"{clock(time.time() - self.started).replace('about ', '')}.")
                    self.say(f"Found dates for {found} of {len(self.order)} photos. Check them in Step 3 "
                             "(click each one), then click Rename and move.")
                return
            elif kind == "failed":
                self.reading = None
                self.update_buttons()
                self.progress_label.configure(text="The reading stopped with a problem.")
                messagebox.showerror(TITLE, "The photos couldn't be read:\n\n" + message[1][-1200:],
                                     parent=self.root)
                return

    def item_for(self, front):
        return next((i for i in self.order if self.rows[i].photo.front == front), None)

    def take_finding(self, row, finding):
        row.finding = finding
        row.state = "read"
        best = finding.best
        row.date = best.date if best else None
        row.kind = best.kind if best else None
        row.clue = best
        row.ticked = bool(best and best.score >= self.tick_from())

    def apply_choice(self, row):
        """Keep a date typed in by hand, or a skip, whatever the reading finds."""
        if row.photo.front not in self.choices:
            return
        choice = self.choices[row.photo.front]
        if choice is None:
            row.ticked = False
            return
        date, kind = choice
        row.state = "read" if row.state == "waiting" else row.state
        row.date, row.kind, row.typed, row.ticked = date, kind, True, True
        row.clue = next((c for c in (row.finding.clues if row.finding else []) if c.date == date), None)

    @staticmethod
    def tick_from():
        import dating_photos_find
        return dating_photos_find.TICK_FROM

    # -- Step 3

    def refresh_row(self, item):
        row = self.rows[item]
        if row.state == "moved":
            values = ("", row.photo.name(), row.date.isoformat(), "moved to " + row.moved_to)
            tags = ("moved",)
        elif row.state == "waiting":
            values = ("", row.photo.name(), "", "not read yet")
            tags = ()
        else:
            if row.date:
                how = self.how(row)
                unsure = row.clue is not None and row.clue.score < self.tick_from()
                tags = ("unsure",) if unsure else ()
                values = (TICKED if row.ticked else UNTICKED, row.photo.name(), row.date.isoformat(),
                          ("check this: " if unsure else "") + how)
            else:
                problem = row.finding.problem if row.finding else ""
                values = (UNTICKED, row.photo.name(), "none", problem or "no date found")
                tags = ("none",)
        if row.photo.back and row.state != "moved":
            values = (values[0], values[1] + "  (+ back)") + tuple(values[2:])
        self.tree.item(item, values=values, tags=tags)

    @staticmethod
    def how(row):
        if row.clue is None:
            return "typed in"
        where = {"back": "on the back", "front": "on the front", "file name": "in the file name",
                 "folder name": "in the folder name"}.get(row.clue.where, row.clue.where)
        return where

    def show_counts(self):
        total = len(self.order)
        ticked = sum(1 for i in self.order if self.rows[i].ticked and self.rows[i].state == "read")
        moved = sum(1 for i in self.order if self.rows[i].state == "moved")
        waiting = len(self.unread())
        parts = [f"{ticked} of {total} ticked to move"]
        if waiting:
            parts.append(f"{waiting} not read yet")
        if moved:
            parts.append(f"{moved} moved")
        self.count_label.configure(text=", ".join(parts) + ". Click a photo to see it; untick any you "
                                                           "don't want moved.")

    def list_clicked(self, event):
        if self.tree.identify_region(event.x, event.y) == "cell" and self.tree.identify_column(event.x) == "#1":
            item = self.tree.identify_row(event.y)
            if item:
                self.toggle(item)

    def toggle(self, item):
        if not item or item not in self.rows:
            return
        row = self.rows[item]
        if row.state == "moved":
            self.say(f"{row.photo.name()} has already been moved.")
            return
        if row.state == "waiting":
            self.say("That photo hasn't been read yet. Click Find the dates, or type its date on the right.")
            return
        if not row.date:
            self.say("That photo has no date yet. Type one in the Date box on the right, then click Use this date.")
            return
        row.ticked = not row.ticked
        if row.ticked and row.typed:
            self.choices[row.photo.front] = (row.date, row.kind)
        elif row.ticked:
            self.choices.pop(row.photo.front, None)
        else:
            self.choices[row.photo.front] = None
        self.refresh_row(item)
        self.show_counts()
        self.say(f"{row.photo.name()} will {'be moved' if row.ticked else 'stay where it is'}.")

    def selected(self):
        item = self.tree.focus() or (self.tree.selection() or [None])[0]
        return item if item in self.rows else None

    def show_selected(self, force=False):
        item = self.selected()
        if item is None:
            return
        if item == self.current and not force:
            return
        self.current = item
        row = self.rows[item]
        self.date_text.set(row.date.isoformat() if row.date else "")
        clues = row.finding.clues if row.finding else []
        self.other_choices = [c for c in clues if not (row.date and c.date == row.date)][:12]
        self.other_box.configure(values=[f"{c.date.isoformat()}  from {c.describe()} "
                                         f"{self.how_clue(c)}" for c in self.other_choices])
        self.other.set("" if self.other_choices else "(none)")
        self.found_label.configure(text=self.found_words(row))
        words = [text for _, text in (row.finding.words if row.finding else [])]
        unique = list(dict.fromkeys(w.strip() for w in words if w.strip()))
        self.words_label.configure(text=("Words read: " + "  |  ".join(unique)[:400]) if unique else
                                   "Type a date like 1998-10-01, Oct 98 or 1998, then click Use this date "
                                   "(or press Enter).")
        self.draw_preview()

    @staticmethod
    def how_clue(clue):
        return {"back": "on the back", "front": "on the front", "file name": "in the file name",
                "folder name": "in the folder name"}.get(clue.where, "")

    def found_words(self, row):
        name = row.photo.name()
        if row.state == "moved":
            return f"{name} was moved to {row.moved_to}."
        if row.state == "waiting":
            return f"{name} hasn't been read yet."
        if not row.date:
            problem = row.finding.problem if row.finding else ""
            return f"{name}: {problem or 'no date was found'}. Type its date below if you know it."
        if row.clue is None:
            return f"{name}: {row.date.isoformat()}, typed in."
        return (f"{name}: {row.date.isoformat()}, from {row.clue.describe()} {self.how(row)}."
                + ("  Check this one: the reading wasn't sure." if row.clue.score < self.tick_from() else ""))

    def focus_date_box(self):
        self.date_box.focus_set()
        self.date_box.select_range(0, "end")

    def use_date(self):
        item = self.selected()
        if item is None:
            self.say("First click a photo in the list.")
            return
        row = self.rows[item]
        if row.state == "moved":
            self.say(f"{row.photo.name()} has already been moved. Undo the last move to change it.")
            return
        typed = self.dates.typed_clue(self.date_text.get(), month_first=not self.day_first.get())
        if typed is None:
            self.say("That isn't a date I can read. Type it like 1998-10-01, Oct 98 or 1998 "
                     "(it can't be after today or before 1839).")
            self.focus_date_box()
            return
        date, kind = typed.date, typed.kind
        found = next((c for c in (row.finding.clues if row.finding else []) if c.date == date), None)
        if found:                # e.g. picked from Other dates: as exact as the writing was
            kind = found.kind
        self.choices[row.photo.front] = (date, kind)
        self.apply_choice(row)
        self.refresh_row(item)
        self.show_counts()
        self.draw_preview()
        self.say(f"{row.photo.name()} will be {date.isoformat()}.")
        self.next_photo(item)

    def other_picked(self):
        index = self.other_box.current()
        if 0 <= index < len(getattr(self, "other_choices", [])):
            self.date_text.set(self.other_choices[index].date.isoformat())
            self.say("Click Use this date to give the photo that date.")

    def skip_photo(self):
        item = self.selected()
        if item is None:
            self.say("First click a photo in the list.")
            return
        row = self.rows[item]
        if row.state == "moved":
            self.say(f"{row.photo.name()} has already been moved.")
            return
        row.ticked = False
        if row.date:
            self.choices[row.photo.front] = None
        self.refresh_row(item)
        self.show_counts()
        self.say(f"{row.photo.name()} will stay where it is.")
        self.next_photo(item)

    def next_photo(self, item):
        index = self.order.index(item)
        if index + 1 < len(self.order):
            following = self.order[index + 1]
            self.tree.selection_set(following)
            self.tree.focus(following)
            self.tree.see(following)

    # -- the preview

    def load_preview(self, path):
        """(small upright picture, full upright size) for the preview, cached."""
        if path in self.preview_cache:
            return self.preview_cache[path]
        from PIL import Image, ImageOps
        from photo_text_reader import eight_bit
        try:
            with Image.open(path) as picture:
                full = picture.size
                orientation = picture.getexif().get(0x0112, 1)
                if orientation in (5, 6, 7, 8):
                    full = (full[1], full[0])
                picture.draft("RGB", (1400, 1400))
                small = eight_bit(ImageOps.exif_transpose(picture)).convert("RGB")
                small.thumbnail((1400, 1400))
        except Exception:
            small, full = None, (1, 1)
        if len(self.preview_cache) > 40:
            self.preview_cache.clear()
        self.preview_cache[path] = (small, full)
        return small, full

    def draw_preview(self):
        from PIL import ImageTk
        canvas = self.canvas
        canvas.delete("all")
        self.preview_images = []
        width, height = max(50, canvas.winfo_width()), max(50, canvas.winfo_height())
        item = self.current
        if item is None or item not in self.rows:
            canvas.create_text(width // 2, height // 2, text="The photo shows here.", fill="#505050",
                               font=("Segoe UI", 13))
            return
        row = self.rows[item]
        if row.state == "moved":
            canvas.create_text(width // 2, height // 2, text=f"Moved to {row.moved_to}", fill="#505050",
                               font=("Segoe UI", 13), width=width - 40)
            return
        sides = [("Front", row.photo.front, "front")]
        if row.photo.back:
            sides.append(("Back", row.photo.back, "back"))
        slot = (width - 10 * (len(sides) + 1)) / len(sides)
        clue = row.clue
        for number, (label, path, where) in enumerate(sides):
            small, full = self.load_preview(path)
            left = 10 + number * (slot + 10)
            if small is None:
                canvas.create_text(left + slot / 2, height / 2, text=f"{label}: can't be shown",
                                   fill="#b00020", font=("Segoe UI", 12))
                continue
            room_h = height - 34
            scale = min(slot / small.width, room_h / small.height)
            if where == "front" and self.print_date.get() and row.date:
                small = self.stamp.draw_stamp(small, self.stamp_words(row))   # how it will look
            shown = small.resize((max(1, int(small.width * scale)), max(1, int(small.height * scale))))
            photo = ImageTk.PhotoImage(shown)
            self.preview_images.append(photo)
            x = left + (slot - shown.width) / 2
            y = 26 + (room_h - shown.height) / 2
            canvas.create_text(left + slot / 2, 12, text=label, font=("Segoe UI", 11, "bold"))
            canvas.create_image(x, y, image=photo, anchor="nw")
            if clue is not None and clue.box and clue.where == where:
                ratio = shown.width / full[0]
                l, t, r, b = clue.box
                pad = 6
                canvas.create_rectangle(x + l * ratio - pad, y + t * ratio - pad, x + r * ratio + pad,
                                        y + b * ratio + pad, outline="#ff1a1a", width=3)

    # -- Step 4

    def rename_and_move(self):
        if why := self.busy():
            self.say(why)
            return
        if not self.folder:
            self.say("First pick a folder in Step 1, then find the dates in Step 2.")
            return
        if not self.rename.get() and not self.print_date.get():
            self.say("Tick “Rename each photo by its date” or “Print the date on each photo” "
                     "(or both) first.")
            return
        chosen = [i for i in self.order if self.rows[i].ticked and self.rows[i].date
                  and self.rows[i].state == "read"]
        if not chosen:
            self.say("No photo is ticked to move. Find the dates in Step 2, or type a date for a photo "
                     "and click Use this date.")
            return
        if self.log_busy(self.folder):
            return
        stay =sum(1 for i in self.order if self.rows[i].state != "moved") - len(chosen)
        backs = sum(1 for i in chosen if self.rows[i].photo.back)
        doing = {"Rename and move": "Rename and move", "Print dates and move": "Print the date on and move",
                 "Rename, print dates, move": "Rename, print the date on and move"}[self.go_words()]
        question = (f"{doing} {len(chosen)} photo{'s' if len(chosen) != 1 else ''}"
                    + (f" (and {backs} back{'s' if backs != 1 else ''})" if backs else "")
                    + f" into \u201c{self.files.DEST}\u201d?")
        if stay:
            question += f"\n\n{stay} photo{'s' if stay != 1 else ''} without a tick stay{'s' if stay == 1 else ''} " \
                        "where they are."
        if self.print_date.get():
            question += (f"\n\nThe date is printed in orange like {self.print_style.get()}; an untouched copy "
                         f"of each photo is kept in “{self.files.ORIGINALS}”.")
        if self.set_taken.get():
            question += "\n\nEach photo's Date taken will be set to its date."
        question += "\n\nUndo the last move puts them all back if you change your mind."
        if not messagebox.askyesno(TITLE, question, parent=self.root):
            self.say("Nothing was moved.")
            return
        try:
            moves = self.files.new_names([(self.rows[i].photo, self.rows[i].date) for i in chosen],
                                         self.folder, self.keep_name.get(), self.year_folders.get(),
                                         rename=self.rename.get())
        except OSError as error:
            self.say(f"The new names couldn't be worked out: {error}")
            return
        self.moving_items = {self.rows[i].photo.front: i for i in chosen}
        self.planned = moves
        clues = {self.rows[i].photo.front: self.rows[i].clue for i in chosen}
        printed = {self.rows[i].photo.front: self.stamp_words(self.rows[i]) for i in chosen} \
            if self.print_date.get() else {}
        self.moving = Moving(self.folder, moves, clues, self.set_taken.get(), printed)
        self.moving.start()
        self.say(f"Moving {len(chosen)} photo{'s' if len(chosen) != 1 else ''}...")

    def stamp_words(self, row):
        """What would be printed on this photo: only the parts of its date that are known."""
        return self.stamp.stamp_text(row.date, row.kind or self.dates.DAY, self.print_style.get())

    def moving_done(self):
        moving, self.moving = self.moving, None
        if moving.error:
            messagebox.showerror(TITLE, "The photos couldn't all be moved:\n\n" + moving.error[-1200:],
                                 parent=self.root)
            return
        done = moving.result
        failed = {name for name, _ in done.problems}
        for move in self.planned:
            item = self.moving_items.get(move.photo.front)
            if item is None or not os.path.exists(move.pairs[0][1]):
                continue
            row = self.rows[item]
            row.state = "moved"
            row.ticked = False
            row.moved_to = os.path.relpath(move.pairs[0][1], self.folder)
            self.refresh_row(item)
        self.show_counts()
        self.show_selected(force=True)
        words = (f"Moved {done.moved} photo{'s' if done.moved != 1 else ''} into "
                 f"\u201c{self.files.DEST}\u201d ({done.files} file{'s' if done.files != 1 else ''}).")
        if done.problems:
            words += "\n\nProblems:\n" + "\n".join(f"{name}: {why}" for name, why in done.problems[:15])
            if len(done.problems) > 15:
                words += f"\n...and {len(done.problems) - 15} more"
        words += "\n\nA list of every move is kept in Dated Photographs\\" + self.files.LOG_NAME + "."
        self.say(words.split("\n")[0] + (" Some had problems." if failed else ""))
        messagebox.showinfo(TITLE, words, parent=self.root)

    def undo(self):
        if why := self.busy():
            self.say(why)
            return
        folder = self.folder
        if not folder:
            folder = filedialog.askdirectory(parent=self.root, mustexist=True,
                                             title="Pick the folder that holds \u201cDated Photographs\u201d",
                                             initialdir=self.settings.get("folder") or pictures_folder())
            if not folder:
                self.say("Nothing was undone.")
                return
            folder = os.path.normpath(folder)
            folder = self.above_dated(folder) or folder      # Dated Photographs itself was picked
        batch = self.files.last_batch(folder)
        if batch is None:
            self.say("There's nothing to undo in that folder.")
            messagebox.showinfo(TITLE, f"There's no move to undo in\n{folder}", parent=self.root)
            return
        if self.log_busy(folder):
            return
        rows = [r for r in self.files.read_log(folder) if self.files.batch_of(r) == batch]
        when = rows[0].get("When", "") if rows else ""
        if not messagebox.askyesno(TITLE, f"Put back the {len(rows)} file{'s' if len(rows) != 1 else ''} moved "
                                   f"on {when}, with their old names and dates?", parent=self.root):
            self.say("Nothing was undone.")
            return
        try:
            back, problems = self.files.undo_last(folder)
        except Exception as error:
            back, problems = 0, [("Undo", f"stopped with a problem ({error})")]
        words = f"Put back {back} file{'s' if back != 1 else ''} where they were."
        if problems:
            words += "\n\nProblems:\n" + "\n".join(f"{n}: {w}" for n, w in problems[:15])
            if len(problems) > 15:
                words += f"\n...and {len(problems) - 15} more"
        self.folder = folder
        self.list_photos()
        self.say(words.split("\n")[0] + (" Some had problems." if problems else ""))
        messagebox.showinfo(TITLE, words, parent=self.root)

    def log_busy(self, folder):
        """If the list of moves is open in Excel, say so (a move or Undo couldn't be
        recorded) and return True."""
        if not self.files.log_in_use(folder):
            return False
        self.say(self.files.LOG_BUSY)
        messagebox.showinfo(TITLE, self.files.LOG_BUSY, parent=self.root)
        return True

    def open_dated(self):
        if not self.folder:
            self.say("First pick a folder in Step 1.")
            return
        dest = os.path.join(self.folder, self.files.DEST)
        if not os.path.isdir(dest):
            self.say("There's no Dated Photographs folder yet: nothing has been moved from this folder.")
            return
        os.startfile(dest)
        self.say("Dated Photographs is opening in File Explorer.")

    # -- closing

    def close(self):
        if self.moving:
            self.say("Wait a moment: the photos are being moved. Then close.")
            return
        if self.reading:
            if not messagebox.askyesno(TITLE, "The photos are still being read. Stop and close?",
                                       parent=self.root):
                self.say("Still reading. Not closed.")
                return
            self.reading.stopping.set()
        ticked = sum(1 for i in self.order if self.rows[i].ticked and self.rows[i].state == "read")
        if ticked and not messagebox.askyesno(
                TITLE, f"{ticked} photo{'s are' if ticked != 1 else ' is'} ticked but not moved yet. "
                       "Close anyway? (Nothing has been changed.)", parent=self.root):
            self.say("Not closed. Click Rename and move to move the ticked photos.")
            return
        self.stop_polling()
        self.root.destroy()

    def stop_polling(self):
        try:
            self.root.after_cancel(self.poll_waiting)
        except tk.TclError:
            pass


# ------------------------------------------------------------------ checks for the built program

def self_test(folder):
    """For the Setup's test on a clean PC: makes a photo with a date written on it,
    reads it with only the parts packed into the program, renames it and undoes
    that, then writes what happened to self-test.txt in the folder."""
    lines = []
    try:
        from PIL import Image, ImageDraw, ImageFont
        import dating_photos_dates as dates
        import dating_photos_files as files
        import dating_photos_find as find
        import photo_text_reader
        os.makedirs(folder, exist_ok=True)
        picture = Image.new("RGB", (1400, 900), (236, 232, 220))
        font_path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segoepr.ttf")
        font = ImageFont.truetype(font_path if os.path.exists(font_path) else "arial.ttf", 110)
        ImageDraw.Draw(picture).text((180, 360), "Oct-98", font=font, fill=(25, 35, 110))
        path = os.path.join(folder, "self-test photo.jpg")
        picture.save(path, quality=90)
        started = time.time()
        reader = photo_text_reader.Reader()
        lines.append(f"Tesseract found: {reader.use_tesseract}")
        photo = files.group_photos([path])[0]
        finding = find.find(reader, photo, folder)
        best = finding.best
        lines.append(f"Read in {time.time() - started:.1f} s: {best.describe() if best else None}")
        lines.append(f"Date found: {best is not None and best.date.isoformat() == '1998-10-01'}")
        if best:
            import dating_photos_stamp as stamp
            with open(path, "rb") as f:
                before = f.read()
            moves = files.new_names([(photo, best.date)], folder)
            words = stamp.stamp_text(best.date, best.kind)
            done = files.apply_moves(folder, moves, {photo.front: best}, printed={photo.front: words})
            target = moves[0].pairs[0][1]
            lines.append(f"Renamed: {done.moved == 1 and os.path.exists(target)} ({os.path.relpath(target, folder)})")
            lines.append(f"Date taken set: {files.read_date_taken(target) == '1998:10:01 12:00:00'}")
            with Image.open(target) as moved:
                corner = moved.convert("RGB").crop((moved.width * 3 // 4, moved.height * 3 // 4,
                                                    moved.width, moved.height)).tobytes()
            orange = sum(1 for i in range(0, len(corner), 3)
                         if corner[i] > 200 and 90 < corner[i + 1] < 190 and corner[i + 2] < 110)
            lines.append(f"Date printed ({words}): {orange > 100}")
            back, problems = files.undo_last(folder)
            with open(path, "rb") as f:
                restored = f.read() == before
            lines.append(f"Undo worked: {back == 1 and restored and not problems}")
        lines.append(f"Typed date works: {dates.typed_date('Oct 98') is not None}")
    except Exception:
        lines.append("Self-test crashed:\n" + traceback.format_exc())
    with open(os.path.join(folder, "self-test.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def main():
    no_lost_output()
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        self_test(sys.argv[2])
        return
    sys.excepthook = show_crash
    sharp_on_big_screens()
    install_missing()
    import dating_photos_dates as dates
    import dating_photos_files as files

    root = tk.Tk()
    root.report_callback_exception = show_crash
    set_icon(root)
    app = App(root, dates, files)
    app.fit_to_screen()
    app.bring_to_front()
    root.mainloop()


if __name__ == "__main__":
    main()
