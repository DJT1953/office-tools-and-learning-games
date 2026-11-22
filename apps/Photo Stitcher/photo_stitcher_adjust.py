"""Photo Stitcher's Adjust window: turn, straighten, square up, crop and resize a
finished panorama, then save it as a separate "adjusted" file. The plain
panorama is never changed.

The picture math (Adjustments and the functions above it) has no window code,
so it can be tested on its own (test_photo_stitcher.py, class Adjusting).
"""

from __future__ import annotations

import base64
import os
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import numpy as np

import photo_stitcher_engine as engine

STRAIGHTEN_LIMIT = 15.0    # the slider goes from -15 to +15 degrees
SIZES = ((1.0, "Full size"), (0.75, "3/4"), (0.5, "1/2"), (0.25, "1/4"))
SMALLEST = 20              # a squared-up or cropped picture is at least 20 x 20 pixels


class AdjustError(Exception):
    """An adjustment can't be made. The message says why in plain words."""


# ------------------------------------------------------------------ the picture math

def turn_quarter(image: np.ndarray, clockwise: bool) -> np.ndarray:
    return cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE if clockwise
                      else cv2.ROTATE_90_COUNTERCLOCKWISE)


def straighten(image: np.ndarray, degrees: float) -> np.ndarray:
    """Turns the picture clockwise by `degrees` (counterclockwise if negative),
    then cuts away the black corners the turn leaves."""
    if abs(degrees) < 1e-6:
        return image
    height, width = image.shape[:2]
    turn = cv2.getRotationMatrix2D((width / 2, height / 2), -degrees, 1.0)
    cos, sin = abs(turn[0, 0]), abs(turn[0, 1])
    new_width = int(np.ceil(height * sin + width * cos))
    new_height = int(np.ceil(height * cos + width * sin))
    turn[0, 2] += new_width / 2 - width / 2
    turn[1, 2] += new_height / 2 - height / 2
    turned = cv2.warpAffine(image, turn, (new_width, new_height), flags=cv2.INTER_LINEAR,
                            borderMode=cv2.BORDER_CONSTANT, borderValue=0)
    # Where the picture is, worked out from a white stand-in, so black parts of
    # the picture itself are kept. One pixel in from its edge, where the turned
    # edge is blended with the black outside.
    inside = cv2.warpAffine(np.full((height, width), 255, np.uint8), turn, (new_width, new_height),
                            flags=cv2.INTER_NEAREST, borderValue=0)
    inside = cv2.erode(inside, np.ones((3, 3), np.uint8)) > 0
    return np.ascontiguousarray(engine.crop(turned, engine.trim_rectangle(turned, inside)))


def order_corners(points) -> list[tuple[float, float]]:
    """The 4 points as top-left, top-right, bottom-right, bottom-left."""
    pts = [tuple(float(v) for v in p) for p in points]
    if len(pts) != 4:
        raise AdjustError("Square up needs 4 corner dots.")
    by_sum = sorted(range(4), key=lambda i: pts[i][0] + pts[i][1])
    by_diff = sorted(range(4), key=lambda i: pts[i][1] - pts[i][0])
    top_left, bottom_right = by_sum[0], by_sum[-1]
    top_right, bottom_left = by_diff[0], by_diff[-1]
    order = [top_left, top_right, bottom_right, bottom_left]
    if len(set(order)) != 4:
        raise AdjustError("The 4 dots need to be at 4 different corners. Move them to the "
                          "corners of something that should be a rectangle.")
    return [pts[i] for i in order]


def _check_shape(corners: list[tuple[float, float]]) -> None:
    """The dots must make a proper four-sided shape: not crossed, not bent
    inward, not flat, not tiny."""
    pts = np.array(corners, dtype=np.float64)
    turns = []
    for i in range(4):
        a, b, c = pts[i], pts[(i + 1) % 4], pts[(i + 2) % 4]
        turns.append((b[0] - a[0]) * (c[1] - b[1]) - (b[1] - a[1]) * (c[0] - b[0]))
    area = abs(cv2.contourArea(pts.astype(np.float32)))
    if not (all(t > 0 for t in turns) or all(t < 0 for t in turns)) or area < SMALLEST * SMALLEST:
        raise AdjustError("The 4 dots don't make a proper four-sided shape. Put one dot on "
                          "each corner of something that should be a rectangle, going round it.")


def square_up(image: np.ndarray, corners) -> np.ndarray:
    """The part inside the 4 corners (in pixels), straightened into a rectangle."""
    tl, tr, br, bl = order_corners(corners)
    _check_shape([tl, tr, br, bl])

    def length(a, b):
        return float(np.hypot(a[0] - b[0], a[1] - b[1]))

    width = int(round(max(length(tl, tr), length(bl, br))))
    height = int(round(max(length(tl, bl), length(tr, br))))
    if width < SMALLEST or height < SMALLEST:
        raise AdjustError("The 4 dots are too close together. Spread them out over the part "
                          "that should be a rectangle.")
    target = np.float32([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]])
    change = cv2.getPerspectiveTransform(np.float32([tl, tr, br, bl]), target)
    return cv2.warpPerspective(image, change, (width, height), flags=cv2.INTER_LINEAR,
                               borderMode=cv2.BORDER_REPLICATE)


def crop(image: np.ndarray, box) -> np.ndarray:
    """The part inside the box, given as fractions (left, top, right, bottom)."""
    height, width = image.shape[:2]
    left, top, right, bottom = box
    x0, x1 = sorted((int(round(left * width)), int(round(right * width))))
    y0, y1 = sorted((int(round(top * height)), int(round(bottom * height))))
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(width, x1), min(height, y1)
    if x1 - x0 < SMALLEST or y1 - y0 < SMALLEST:
        raise AdjustError("That crop is too small. Drag a bigger box over the part to keep.")
    return np.ascontiguousarray(image[y0:y1, x0:x1])


class Adjustments:
    """The changes made so far, in order, plus the size to save at. They are
    applied afresh to whichever copy is shown or saved, so the plain panorama
    is never changed and Undo is exact."""

    def __init__(self):
        self.steps: list[tuple] = []
        self.size = 1.0

    @property
    def changed(self) -> bool:
        return bool(self.steps) or self.size != 1.0

    def turn(self, clockwise: bool) -> None:
        self.steps.append(("turn", bool(clockwise)))

    def set_straighten(self, degrees: float) -> None:
        """The slider: half-degree steps; moving it again changes its own step."""
        degrees = max(-STRAIGHTEN_LIMIT, min(STRAIGHTEN_LIMIT, round(degrees * 2) / 2))
        if self.steps and self.steps[-1][0] == "straighten":
            if degrees == 0:
                self.steps.pop()
            else:
                self.steps[-1] = ("straighten", degrees)
        elif degrees != 0:
            self.steps.append(("straighten", degrees))

    def straighten_now(self) -> float:
        """What the slider shows: the last step's angle if it is a straighten."""
        if self.steps and self.steps[-1][0] == "straighten":
            return self.steps[-1][1]
        return 0.0

    def square(self, corners) -> None:
        """Corners as fractions (0 to 1) of the picture as it is now."""
        self.steps.append(("square", tuple(tuple(float(v) for v in c) for c in corners)))

    def crop(self, box) -> None:
        self.steps.append(("crop", tuple(float(v) for v in box)))

    def undo(self) -> bool:
        if not self.steps:
            return False
        self.steps.pop()
        return True

    def start_over(self) -> None:
        self.steps.clear()
        self.size = 1.0

    def state(self) -> tuple:
        return tuple(self.steps), self.size

    def apply(self, image: np.ndarray, with_size: bool = True) -> np.ndarray:
        for kind, value in self.steps:
            if kind == "turn":
                image = turn_quarter(image, value)
            elif kind == "straighten":
                image = straighten(image, value)
            elif kind == "square":
                height, width = image.shape[:2]
                image = square_up(image, [(x * width, y * height) for x, y in value])
            elif kind == "crop":
                image = crop(image, value)
        if with_size and self.size < 1.0:
            image = cv2.resize(image, None, fx=self.size, fy=self.size, interpolation=cv2.INTER_AREA)
        return image


# ------------------------------------------------------------------ the Adjust window

TITLE = "Photo Stitcher"
PREVIEW_LONG_SIDE = 1600   # the window works on a copy this long, so changes show at once,
PREVIEW_SHORT_SIDE = 300   # but at least this tall (or wide) ...
PREVIEW_MOST = 12000       # ... and never longer than this
SQUARE_START = ((0.1, 0.1), (0.9, 0.1), (0.9, 0.9), (0.1, 0.9))


def preview_copy(picture: np.ndarray) -> np.ndarray:
    """The smaller copy the window works on: 1,600 pixels on the long side, but
    at least 300 on the short side (a very wide panorama would otherwise be a
    strip too thin to crop or square up), and never over 12,000 long. Never
    enlarged."""
    height, width = picture.shape[:2]
    long_side, short_side = max(height, width), min(height, width)
    scale = min(1.0, PREVIEW_MOST / long_side,
                max(PREVIEW_LONG_SIDE / long_side, PREVIEW_SHORT_SIDE / short_side))
    if scale >= 1.0:
        return picture
    size = (max(1, round(width * scale)), max(1, round(height * scale)))
    return cv2.resize(picture, size, interpolation=cv2.INTER_AREA)


class AdjustWindow:
    """The Adjust window. `picture` is the plain panorama (never changed),
    `plain_path` where it was saved; the adjusted copy is suggested beside it."""

    def __init__(self, master, picture, plain_path, save_choices, save_type=".jpg"):
        self.picture, self.plain_path = picture, plain_path
        self.save_choices, self.save_type = save_choices, save_type
        self.steps = Adjustments()
        self.saved_state = self.steps.state()
        self.saved_path = None
        self.saving = False
        self.results = []                   # things the saving thread asks the window to do
        self.mode = None                    # None, "square" or "crop"
        self.dots = list(SQUARE_START)      # square-up dots, as fractions of the picture shown
        self.box = None                     # crop box, fractions (left, top, right, bottom)
        self.drag = None
        self.base = preview_copy(picture)
        self.shown = self.base
        self.view = (0, 0, 1, 1)            # where the picture sits on the canvas: x, y, w, h
        self.photo = None
        self.redraw_waiting = None
        self.slider_busy = False

        win = self.window = tk.Toplevel(master)
        win.title("Adjust the panorama - Photo Stitcher")
        win.protocol("WM_DELETE_WINDOW", self.close)
        # A set size, so the window doesn't grow and shrink with the length of the
        # words under the picture (that moved the picture away from the mouse).
        # Then it fills the screen.
        screen_w, screen_h = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{int(screen_w * 0.9)}x{int(screen_h * 0.85)}+{int(screen_w * 0.05)}+20")
        self.fill_waiting = win.after(250, self.fill_screen)    # once the window is on screen
        self.dot = int(13 * win.winfo_fpixels("1i") / 96)    # dot radius, bigger on big screens
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)

        left = ttk.Frame(win, padding=(12, 12, 6, 12))
        left.grid(row=0, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(left, bg="#d4d4d4", highlightthickness=0, width=600, height=400)
        self.canvas.grid(sticky="nsew")
        self.canvas.bind("<Configure>", lambda event: self.schedule_draw())
        self.canvas.bind("<ButtonPress-1>", self.press)
        self.canvas.bind("<B1-Motion>", self.move)
        self.canvas.bind("<ButtonRelease-1>", self.release)
        self.size_note = ttk.Label(left, style="Note.TLabel")
        self.size_note.grid(sticky="w", pady=(6, 0))
        self.status = ttk.Label(left, style="Note.TLabel", wraplength=1000, justify="left",
                                text="Turn, square up, crop or resize the panorama, then click "
                                     "Save as. Your plain panorama stays as it is.")
        self.status.grid(sticky="w", pady=(4, 0))

        side = ttk.Frame(win, padding=(6, 12, 12, 12))
        side.grid(row=0, column=1, sticky="ns")
        ttk.Label(side, text="Adjust the panorama", font=("Segoe UI", 18, "bold")).grid(
            sticky="w", pady=(0, 8))

        turn = self.group(side, "Step 1 - Turn")
        self.left_button = self.big(turn, "Turn left", lambda: self.turn(False), 0, 0)
        self.right_button = self.big(turn, "Turn right", lambda: self.turn(True), 0, 1)
        ttk.Label(turn, text="Straighten:", font=("Segoe UI", 12)).grid(row=1, column=0, sticky="w",
                                                                        pady=(8, 0))
        self.angle_label = ttk.Label(turn, text="0.0°", font=("Segoe UI", 12, "bold"))
        self.angle_label.grid(row=1, column=1, sticky="e", pady=(8, 0))
        self.slider = ttk.Scale(turn, from_=-STRAIGHTEN_LIMIT, to=STRAIGHTEN_LIMIT,
                                orient="horizontal", command=self.slid)
        self.slider.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(2, 0))

        square = self.group(side, "Step 2 - Square up")
        self.square_button = self.big(square, "Square up", self.start_square, 0, 0)
        self.square_apply = self.big(square, "Apply", self.apply_square, 0, 1)
        self.square_cancel = self.big(square, "Cancel", self.cancel_mode, 0, 2)

        cut = self.group(side, "Step 3 - Crop")
        self.crop_button = self.big(cut, "Draw a crop", self.start_crop, 0, 0)
        self.crop_apply = self.big(cut, "Apply", self.apply_crop, 0, 1)
        self.crop_cancel = self.big(cut, "Cancel", self.cancel_mode, 0, 2)

        size = self.group(side, "Step 4 - Size")
        self.size_choice = tk.DoubleVar(value=1.0)
        for column, (fraction, words) in enumerate(SIZES):
            ttk.Radiobutton(size, text=words, value=fraction, variable=self.size_choice,
                            command=self.size_changed).grid(row=0, column=column, sticky="w",
                                                            padx=(0, 10))

        save = self.group(side, "Step 5 - Save")
        self.undo_button = self.big(save, "Undo", self.undo, 0, 0)
        self.start_over_button = self.big(save, "Start over", self.start_over, 0, 1)
        self.save_button = self.big(save, "Save as", self.save, 1, 0)
        self.exit_button = self.big(save, "Exit", self.close, 1, 1)

        self.update_buttons()
        self.show_size()
        self.poll_waiting = win.after(150, self.poll)

    # -- building the window

    @staticmethod
    def group(parent, title):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.grid(sticky="ew", pady=6)
        frame.columnconfigure((0, 1, 2), weight=1)
        return frame

    @staticmethod
    def big(parent, text, command, row, column):
        button = ttk.Button(parent, text=text, style="Big.TButton", command=command)
        button.grid(row=row, column=column, sticky="ew", padx=3, pady=3)
        return button

    def say(self, text):
        self.status.configure(text=text)

    def update_buttons(self):
        def ready(button, ok):
            button.state(["!disabled"] if ok else ["disabled"])
        ready(self.square_apply, self.mode == "square")
        ready(self.square_cancel, self.mode == "square")
        ready(self.crop_apply, self.mode == "crop")
        ready(self.crop_cancel, self.mode == "crop")
        ready(self.undo_button, bool(self.steps.steps) or self.mode is not None)
        ready(self.start_over_button, self.steps.changed)

    @property
    def unsaved(self) -> bool:
        return self.steps.changed and self.steps.state() != self.saved_state

    def exists(self) -> bool:
        try:
            return bool(self.window.winfo_exists())
        except tk.TclError:
            return False

    def bring_to_front(self):
        self.window.deiconify()
        self.window.lift()
        self.window.focus_force()

    def fill_screen(self):
        try:
            self.window.state("zoomed")
        except tk.TclError:
            pass

    # -- showing the picture

    def refresh(self):
        """Applies the changes to the window's copy and shows it."""
        self.redraw_waiting = None
        self.shown = self.steps.apply(self.base, with_size=False)
        self.draw()
        self.show_size()

    def schedule_draw(self, delay=60):
        # The words under the picture wrap at the picture area's width.
        self.status.configure(wraplength=max(300, self.canvas.winfo_width() - 10))
        if self.redraw_waiting:
            self.window.after_cancel(self.redraw_waiting)
        self.redraw_waiting = self.window.after(delay, self.draw)

    def draw(self):
        self.redraw_waiting = None
        canvas = self.canvas
        room_w, room_h = max(60, canvas.winfo_width()), max(60, canvas.winfo_height())
        height, width = self.shown.shape[:2]
        scale = min((room_w - 2 * self.dot - 8) / width, (room_h - 2 * self.dot - 8) / height)
        w, h = max(1, round(width * scale)), max(1, round(height * scale))
        x, y = (room_w - w) // 2, (room_h - h) // 2
        small = cv2.resize(self.shown, (w, h),
                           interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
        png = cv2.imencode(".png", small, [cv2.IMWRITE_PNG_COMPRESSION, 1])[1]
        self.photo = tk.PhotoImage(master=self.window,
                                   data=base64.b64encode(png.tobytes()).decode("ascii"))
        canvas.delete("all")
        canvas.create_image(x, y, image=self.photo, anchor="nw")
        self.view = (x, y, w, h)
        self.draw_marks()

    def to_canvas(self, fx, fy):
        x, y, w, h = self.view
        return x + fx * w, y + fy * h

    def to_fraction(self, cx, cy):
        x, y, w, h = self.view
        return min(1.0, max(0.0, (cx - x) / w)), min(1.0, max(0.0, (cy - y) / h))

    def draw_marks(self):
        canvas = self.canvas
        canvas.delete("marks")
        if self.mode == "square":
            points = [self.to_canvas(*dot) for dot in self.dots]
            flat = [v for point in points + points[:1] for v in point]
            canvas.create_line(*flat, fill="black", width=5, tags="marks")
            canvas.create_line(*flat, fill="#ffd400", width=3, tags="marks")
            r = self.dot
            for px, py in points:
                canvas.create_oval(px - r, py - r, px + r, py + r, fill="white", outline="black",
                                   width=3, tags="marks")
        elif self.mode == "crop" and self.box:
            x, y, w, h = self.view
            left, top = self.to_canvas(min(self.box[0], self.box[2]), min(self.box[1], self.box[3]))
            right, bottom = self.to_canvas(max(self.box[0], self.box[2]), max(self.box[1], self.box[3]))
            for area in ((x, y, x + w, top), (x, bottom, x + w, y + h),
                         (x, top, left, bottom), (right, top, x + w, bottom)):
                canvas.create_rectangle(*area, fill="black", stipple="gray50", width=0, tags="marks")
            canvas.create_rectangle(left, top, right, bottom, outline="black", width=5, tags="marks")
            canvas.create_rectangle(left, top, right, bottom, outline="#ffd400", width=3, tags="marks")

    def show_size(self):
        grow = self.picture.shape[1] / self.base.shape[1]
        height, width = self.shown.shape[:2]
        size = self.steps.size
        self.size_note.configure(
            text=f"It will be saved at about {round(width * grow * size):,} x "
                 f"{round(height * grow * size):,} pixels.")

    # -- the mouse on the picture

    def press(self, event):
        fx, fy = self.to_fraction(event.x, event.y)
        if self.mode == "square":
            def gap(i):
                px, py = self.to_canvas(*self.dots[i])
                return (px - event.x) ** 2 + (py - event.y) ** 2
            nearest = min(range(4), key=gap)
            if gap(nearest) <= (3 * self.dot) ** 2:
                self.drag = nearest
            else:
                self.say("Grab one of the 4 white dots and drag it onto a corner.")
        elif self.mode == "crop":
            self.drag = "box"
            self.box = (fx, fy, fx, fy)
            self.draw_marks()
        else:
            self.say("To mark the picture, click Square up (Step 2) or Draw a crop (Step 3) first.")

    def move(self, event):
        fx, fy = self.to_fraction(event.x, event.y)      # dots and boxes stay on the picture
        if self.mode == "square" and isinstance(self.drag, int):
            self.dots[self.drag] = (fx, fy)
            self.draw_marks()
        elif self.mode == "crop" and self.drag == "box" and self.box:
            self.box = (self.box[0], self.box[1], fx, fy)
            self.draw_marks()

    def release(self, event):
        self.move(event)
        if self.mode == "crop" and self.drag == "box" and self.box:
            self.say("Click Apply to cut the picture to the box, or drag again to draw a new box.")
        self.drag = None

    # -- the buttons

    def busy(self):
        if self.saving:
            self.say("Wait a moment: the adjusted panorama is being saved.")
            return True
        return False

    def end_mode(self):
        self.mode, self.drag, self.box = None, None, None
        self.draw_marks()
        self.update_buttons()

    def turn(self, clockwise):
        if self.busy():
            return
        self.end_mode()
        self.steps.turn(clockwise)
        self.sync_slider()
        self.refresh()
        self.update_buttons()
        self.say("Turned a quarter turn " + ("right." if clockwise else "left."))

    def slid(self, value):
        if self.slider_busy:
            return
        if self.saving:
            self.sync_slider()
            self.say("Wait a moment: the adjusted panorama is being saved.")
            return
        degrees = round(float(value) * 2) / 2
        if self.mode:
            self.end_mode()
        self.steps.set_straighten(degrees)
        self.angle_label.configure(text=f"{degrees:+.1f}°" if degrees else "0.0°")
        self.update_buttons()
        if self.redraw_waiting:
            self.window.after_cancel(self.redraw_waiting)
        self.redraw_waiting = self.window.after(80, self.refresh)
        if degrees:
            way = "clockwise" if degrees > 0 else "counterclockwise"
            self.say(f"Straightened by {abs(degrees):.1f} degrees {way}.")
        else:
            self.say("Not straightened.")

    def sync_slider(self):
        """Puts the slider where the current straighten step is."""
        self.slider_busy = True
        try:
            degrees = self.steps.straighten_now()
            self.slider.set(degrees)
            self.angle_label.configure(text=f"{degrees:+.1f}°" if degrees else "0.0°")
        finally:
            self.slider_busy = False

    def start_square(self):
        if self.busy():
            return
        self.mode, self.dots, self.box = "square", list(SQUARE_START), None
        self.draw_marks()
        self.update_buttons()
        self.say("Drag the 4 white dots onto the corners of something that should be a "
                 "rectangle (a building, a document, a window), then click Apply.")

    def start_crop(self):
        if self.busy():
            return
        self.mode, self.box = "crop", None
        self.draw_marks()
        self.update_buttons()
        self.say("Drag across the picture to draw a box around the part to keep, then click Apply.")

    def cancel_mode(self):
        if self.mode is None:
            self.say("Nothing to cancel.")
            return
        self.end_mode()
        self.say("Cancelled. Nothing was changed.")

    def apply_square(self):
        if self.mode != "square":
            self.say("Click Square up first, then drag the 4 dots.")
            return
        height, width = self.shown.shape[:2]
        try:
            square_up(self.shown, [(x * width, y * height) for x, y in self.dots])
        except AdjustError as error:
            self.say(str(error))
            messagebox.showwarning(TITLE, str(error), parent=self.window)
            return
        self.steps.square(self.dots)
        self.end_mode()
        self.sync_slider()
        self.refresh()
        self.say("Squared up. Undo takes it back.")

    def apply_crop(self):
        if self.mode != "crop":
            self.say("Click Draw a crop first, then drag a box over the part to keep.")
            return
        if not self.box:
            self.say("Drag across the picture to draw a box first.")
            return
        box = (min(self.box[0], self.box[2]), min(self.box[1], self.box[3]),
               max(self.box[0], self.box[2]), max(self.box[1], self.box[3]))
        try:
            crop(self.shown, box)
        except AdjustError as error:
            self.say(str(error))
            messagebox.showwarning(TITLE, str(error), parent=self.window)
            return
        self.steps.crop(box)
        self.end_mode()
        self.sync_slider()
        self.refresh()
        self.say("Cropped. Undo takes it back.")

    def size_changed(self):
        self.steps.size = float(self.size_choice.get())
        words = dict(SIZES)[self.steps.size]
        self.show_size()
        self.update_buttons()
        self.say("It will be saved at full size." if self.steps.size == 1.0
                 else f"It will be saved at {words} size.")

    def undo(self):
        if self.busy():
            return
        if self.mode:
            self.end_mode()
            self.say("Cancelled. Nothing was changed.")
            return
        if not self.steps.undo():
            self.say("There is nothing to undo.")
            return
        self.sync_slider()
        self.refresh()
        self.update_buttons()
        self.say("Took back the last change.")

    def start_over(self):
        if self.busy():
            return
        if not self.steps.changed:
            self.say("Nothing has been changed yet.")
            return
        if not messagebox.askyesno(TITLE, "Go back to the plain panorama? All the adjustments "
                                          "are taken off.", parent=self.window):
            self.say("Nothing was changed.")
            return
        self.end_mode()
        self.steps.start_over()
        self.size_choice.set(1.0)
        self.sync_slider()
        self.refresh()
        self.update_buttons()
        self.say("Back to the plain panorama.")

    def save(self):
        if self.busy():
            return
        if self.mode:
            self.say("Click Apply or Cancel first (Step 2 or Step 3), then Save as.")
            return
        folder = os.path.dirname(self.plain_path)
        stem = os.path.splitext(os.path.basename(self.plain_path))[0]
        ending = self.save_type
        kinds = [(name, f"*{e}") for e, _, name in self.save_choices]
        chosen = [k for k in kinds if k[1] == f"*{ending}"]
        suggested = engine.free_name(os.path.join(folder, f"{stem} adjusted{ending}"))
        path = filedialog.asksaveasfilename(
            parent=self.window, title="Save the adjusted panorama", initialdir=folder,
            initialfile=os.path.basename(suggested), defaultextension=ending,
            filetypes=chosen + [k for k in kinds if k not in chosen])
        if not path:
            self.say("Not saved. Click Save as when you're ready.")
            return
        path = os.path.normpath(path)
        if os.path.normcase(path) == os.path.normcase(os.path.normpath(self.plain_path)):
            self.say("That's the plain panorama's file. Pick another name, so you keep both.")
            messagebox.showwarning(TITLE, "That's the plain panorama's file. Pick another name, "
                                          "so you keep both.", parent=self.window)
            return
        grow = self.picture.shape[1] / self.base.shape[1] * self.steps.size
        biggest = max(self.shown.shape[:2]) * grow
        if os.path.splitext(path)[1].lower() in (".jpg", ".jpeg") and biggest > engine.JPEG_MAX:
            if not messagebox.askyesno(
                    TITLE, f"This picture is too big for JPG ({engine.JPEG_MAX:,} pixels on a side "
                           "at most).\n\nSave it as PNG instead? PNG keeps full quality but makes a "
                           "bigger file.", parent=self.window):
                self.say("Not saved. Pick a smaller size in Step 4, or save as PNG.")
                return
            path = engine.free_name(os.path.splitext(path)[0] + ".png")
        self.saving = True
        state = self.steps.state()
        self.say(f"Saving {os.path.basename(path)} ...")
        threading.Thread(target=self.save_in_background, args=(path, state), daemon=True).start()

    def save_in_background(self, path, state):
        steps = Adjustments()
        steps.steps, steps.size = list(state[0]), state[1]
        try:
            engine.save_image(steps.apply(self.picture), path)
        except (AdjustError, engine.SaveError) as error:
            text = str(error)
            self.results.append(lambda: self.save_failed(text))
        except MemoryError:
            self.results.append(lambda: self.save_failed(
                "There wasn't enough memory to make the full-size picture. Pick a smaller size "
                "in Step 4, or close other programs and try again."))
        except Exception:
            text = "Something went wrong while saving:\n\n" + traceback.format_exc()[-1200:]
            self.results.append(lambda: self.save_failed(text))
        else:
            self.results.append(lambda: self.save_done(path, state))

    def save_done(self, path, state):
        self.saving = False
        self.saved_path, self.saved_state = path, state
        self.say(f"Saved {os.path.basename(path)} in {os.path.dirname(path)}. "
                 "Your plain panorama is kept too.")

    def save_failed(self, text):
        self.saving = False
        self.say("The adjusted panorama wasn't saved. The message says why.")
        messagebox.showerror(TITLE, text, parent=self.window)

    def poll(self):
        try:
            while self.results:
                self.results.pop(0)()
        finally:
            if self.exists():
                self.poll_waiting = self.window.after(150, self.poll)

    def close(self) -> bool:
        """Exit. Asks first if there are adjustments that aren't saved; returns
        False if the window stays open."""
        if self.saving:
            self.say("Wait a moment: the adjusted panorama is being saved. Then exit.")
            return False
        if self.unsaved and not messagebox.askyesno(
                TITLE, "The adjustments aren't saved yet. Close without saving them?",
                parent=self.window):
            self.say("Not closed. Click Save as to keep the adjusted panorama.")
            return False
        for waiting in (self.poll_waiting, self.redraw_waiting, self.fill_waiting):
            if waiting:
                try:
                    self.window.after_cancel(waiting)
                except tk.TclError:
                    pass
        self.window.destroy()
        return True
