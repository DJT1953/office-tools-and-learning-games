"""Shared look and pop-up boxes for the Book Reader.

Big black words (at least twice Windows' usual size), tall buttons with the
picture above the word, and pop-up boxes with big words. The pop-up functions
(ask, inform, ask_text, save_file_name, open_url ...) are looked up at call
time, so tests can swap them for stand-ins that never wait for a person.
"""

from __future__ import annotations

from PySide6.QtCore import QEvent, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont, QIcon, QPainter, QPalette, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

WINDOW_TEXT = 2.0          # times Windows' usual word size
POPUP_TEXT = 2.0
LIST_TEXT = 2.2
MAIN_ROW = 76              # pixels: the top row of buttons
SECOND_ROW = 60            # pixels: the next row (page turning, panels)


def base_points() -> float:
    size = QApplication.font().pointSizeF()
    return size if size > 0 else 10.0


def words(scale: float = 1.0) -> float:
    """Point size of the window's words times scale."""
    return base_points() * WINDOW_TEXT * scale


def black_words(app: QApplication) -> None:
    """Every word black rather than Windows' soft grey-black. Disabled buttons stay grey."""
    palette = app.palette()
    roles = (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText,
             QPalette.ColorRole.PlaceholderText, QPalette.ColorRole.ToolTipText, QPalette.ColorRole.BrightText)
    for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive):
        for role in roles:
            palette.setColor(group, role, QColor("black"))
    app.setPalette(palette)


def window_style(screen_width: int) -> str:
    """Style sheet for the main window: its words, and every pop-up box parented to it."""
    big = words()
    pop = base_points() * POPUP_TEXT
    width = max(420, screen_width // 2 - 100)
    return (f"QWidget {{ font-size: {big:g}pt; }}"
            f" QToolTip {{ font-size: {words(0.75):g}pt; }}"
            f" QMessageBox QLabel, QInputDialog QLabel, QInputDialog QLineEdit,"
            f" QMessageBox QPushButton, QInputDialog QPushButton {{ font-size: {pop:g}pt; }}"
            f" QMessageBox QLabel#qt_msgbox_label {{ min-width: {width}px; }}"
            f" QInputDialog QLineEdit {{ min-width: {width}px; padding: 4px; }}"
            " QMessageBox QPushButton, QInputDialog QPushButton, QDialogButtonBox QPushButton {"
            " min-width: 150px; min-height: 52px; padding: 4px 24px; }")


def glyph_icon(glyph: str, colour: str = "black", size: int = 48) -> QIcon:
    """A picture for a button drawn from one symbol character."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    font = QFont("Segoe UI Symbol")
    font.setPixelSize(int(size * 0.78))
    painter.setFont(font)
    painter.setPen(QColor(colour))
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, glyph)
    painter.end()
    return QIcon(pixmap)


def tool_button(text: str, glyph: str, tip: str, slot, height: int = MAIN_ROW) -> QToolButton:
    """A tall button: picture above the word."""
    b = QToolButton()
    b.setText(text)
    b.setIcon(glyph_icon(glyph))
    b.setIconSize(QSize(34, 34))
    b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
    b.setToolTip(tip)
    b.setMinimumHeight(height)
    b.setMinimumWidth(96)
    b.setAutoRaise(False)
    b.clicked.connect(lambda _checked=False: slot())   # never pass Qt's checked flag on
    return b


def wide_button(text: str, slot, height: int = SECOND_ROW, tip: str = "") -> QPushButton:
    b = QPushButton(text)
    b.setMinimumHeight(height)
    if tip:
        b.setToolTip(tip)
    b.clicked.connect(lambda _checked=False: slot())   # never pass Qt's checked flag on
    return b


# ---------------------------------------------------------------- pop-up boxes (tests swap these)

def ask(parent: QWidget | None, title: str, text: str, yes: str = "Yes", no: str = "No") -> bool:
    box = QMessageBox(QMessageBox.Icon.Question, title, text, QMessageBox.StandardButton.NoButton, parent)
    yes_button = box.addButton(yes, QMessageBox.ButtonRole.YesRole)
    box.addButton(no, QMessageBox.ButtonRole.NoRole)
    box.setDefaultButton(yes_button)
    box.exec()
    return box.clickedButton() is yes_button


def inform(parent: QWidget | None, title: str, text: str) -> None:
    box = QMessageBox(QMessageBox.Icon.Information, title, text, QMessageBox.StandardButton.Ok, parent)
    box.exec()


def ask_text(parent: QWidget | None, title: str, label: str, text: str = "", multiline: bool = True) -> str | None:
    """The words typed, or None when Cancel is pressed."""
    if not multiline:
        value, ok = QInputDialog.getText(parent, title, label, text=text)
        return value if ok else None
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    layout = QVBoxLayout(dialog)
    caption = QLabel(label)
    caption.setWordWrap(True)
    layout.addWidget(caption)
    edit = QPlainTextEdit(text)
    edit.setMinimumSize(700, 260)
    layout.addWidget(edit)
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    edit.setFocus()
    return edit.toPlainText() if dialog.exec() == QDialog.DialogCode.Accepted else None


def save_file_name(parent: QWidget | None, title: str, suggested: str, filter_text: str) -> str:
    name, _ = QFileDialog.getSaveFileName(parent, title, suggested, filter_text)
    return name


def open_file_names(parent: QWidget | None, title: str, folder: str, filter_text: str) -> list[str]:
    names, _ = QFileDialog.getOpenFileNames(parent, title, folder, filter_text)
    return names


def choose_folder(parent: QWidget | None, title: str, folder: str) -> str:
    return QFileDialog.getExistingDirectory(parent, title, folder)


def open_url(url: str) -> None:
    QDesktopServices.openUrl(QUrl(url))


def show_in_folder(path: str) -> None:
    """Open the file's folder in File Explorer with the file picked."""
    import os
    from PySide6.QtCore import QDir, QProcess
    QProcess.startDetached("explorer.exe", ["/select,", QDir.toNativeSeparators(os.path.abspath(path))])


# ---------------------------------------------------------------- rows that fit narrow windows

def holder(layout) -> QWidget:
    w = QWidget()
    layout.setContentsMargins(0, 0, 0, 0)
    w.setLayout(layout)
    return w


class Bar(QFrame):
    """A bar that shows and hides together with the SideScroll holding it."""

    holder_area = None

    def setVisible(self, visible: bool) -> None:  # noqa: N802
        super().setVisible(visible)
        if self.holder_area is not None:
            self.holder_area.setVisible(visible)


class SideScroll(QScrollArea):
    """Holds a row of big buttons; on a window too narrow for them, the row scrolls sideways
    instead of making the whole window wider than the screen."""

    def __init__(self, inner: QWidget) -> None:
        super().__init__()
        self.inner = inner
        # Keep it hidden only if it was hidden on purpose: a widget never shown yet also
        # counts as "hidden", and those rows (the main buttons) must show.
        was_hidden = inner.isHidden() and inner.testAttribute(Qt.WidgetAttribute.WA_WState_ExplicitShowHide)
        self.setWidget(inner)                 # this shows the widget, so hide it again if it was hidden
        if was_hidden:
            inner.hide()
        inner.installEventFilter(self)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        if isinstance(inner, Bar):
            inner.holder_area = self
            self.setVisible(not was_hidden)
        self._fit()

    def eventFilter(self, watched, event) -> bool:  # noqa: N802
        # Re-measure when the row's contents or words change (big words apply only once shown).
        if watched is self.inner and event.type() in (QEvent.Type.LayoutRequest, QEvent.Type.Show,
                                                      QEvent.Type.StyleChange, QEvent.Type.FontChange):
            QTimer.singleShot(0, self, self._fit)       # tied to this row: never runs after it is gone
        return False

    def _fit(self) -> None:
        for w in [self.inner] + self.inner.findChildren(QWidget):
            w.ensurePolished()
        if self.inner.layout() is not None:
            self.inner.layout().activate()
        narrow = self.inner.minimumSizeHint().width() > self.viewport().width() + 2
        bar = self.horizontalScrollBar().sizeHint().height() if narrow else 0
        self.setFixedHeight(max(self.inner.sizeHint().height(), self.inner.minimumSizeHint().height()) + bar)

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._fit()

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        return QSize(120, super().minimumSizeHint().height())

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(self.inner.sizeHint().width(), self.height())


def choose(parent: QWidget | None, title: str, text: str, options: list[str]) -> str | None:
    """One big button per option, plus Cancel. Returns the option pressed, or None."""
    box = QMessageBox(QMessageBox.Icon.Question, title, text, QMessageBox.StandardButton.NoButton, parent)
    buttons = {box.addButton(option, QMessageBox.ButtonRole.AcceptRole): option for option in options}
    box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
    box.exec()
    return buttons.get(box.clickedButton())
