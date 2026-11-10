"""Reading aloud with Windows voices, and looking up words.

ReadAloud reads one sentence at a time, marks the sentence being read on the
page, turns pages by itself and carries on into the next chapter. The voice
comes from a speaker_factory (Windows voices through Qt by default), so tests
can use a stand-in.

lookup_word() asks the free dictionary at api.dictionaryapi.dev about one word
(only the word is sent) and hands back its meanings, or says why it couldn't.
"""

from __future__ import annotations

import html
import json
import re
from typing import Callable, NamedTuple

from PySide6.QtCore import QObject, Qt, QUrl, Signal
from PySide6.QtNetwork import QNetworkAccessManager, QNetworkReply, QNetworkRequest
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QScrollArea, QVBoxLayout, QWidget

import book_reader_ui as ui

DICTIONARY_URL = "https://api.dictionaryapi.dev/api/v2/entries/en/{}"
WEB_LOOKUP_URL = "https://en.wiktionary.org/wiki/{}"


# ---------------------------------------------------------------- voices

class QtSpeaker(QObject):
    """Windows voices through Qt; says one piece of text at a time."""

    finished = Signal()          # the text given to say() has been read

    def __init__(self) -> None:
        super().__init__()
        from PySide6.QtTextToSpeech import QTextToSpeech
        self._states = QTextToSpeech.State
        engines = QTextToSpeech.availableEngines()
        engine = next((e for e in ("winrt", "sapi") if e in engines), "")
        self.tts = QTextToSpeech(engine, self) if engine else QTextToSpeech(self)
        self._speaking = False
        self.tts.stateChanged.connect(self._state)

    def _state(self, state) -> None:
        if state == self._states.Speaking:
            self._speaking = True
        elif state in (self._states.Ready, self._states.Error) and self._speaking:
            self._speaking = False
            self.finished.emit()

    def say(self, text: str) -> None:
        self.tts.say(text)

    def stop(self) -> None:
        self._speaking = False
        self.tts.stop()

    def voice_names(self) -> list[str]:
        return [v.name() for v in self.tts.availableVoices()]

    def set_voice(self, name: str) -> None:
        for voice in self.tts.availableVoices():
            if voice.name() == name:
                self.tts.setVoice(voice)
                return

    def set_rate(self, rate: float) -> None:
        self.tts.setRate(max(-1.0, min(1.0, rate)))


class ReadAloud(QObject):
    """Reads the open chapter aloud, sentence by sentence, on into the next."""

    state_changed = Signal(str)       # "reading", "paused", "stopped", "finished"

    def __init__(self, parent: QObject | None = None,
                 speaker_factory: Callable[[], object] | None = None) -> None:
        super().__init__(parent)
        self.speaker_factory = speaker_factory or QtSpeaker
        self.speaker = None
        self.view = None
        self.sentences: list[dict] = []
        self.index = 0
        self.state = "stopped"
        self._generation = 0

    def _ensure_speaker(self) -> bool:
        if self.speaker is None:
            try:
                self.speaker = self.speaker_factory()
            except Exception:
                return False
            self.speaker.finished.connect(self._done)
        return bool(self.speaker.voice_names())

    def voice_names(self) -> list[str]:
        return self.speaker.voice_names() if self._ensure_speaker() else []

    def start(self, view, from_offset: int | None = None, voice: str = "", rate: float = 0.0) -> str | None:
        """Start reading; returns a plain message if it can't."""
        if not self._ensure_speaker():
            return ("This computer has no reading voices that the Book Reader can use. Voices can be "
                    "added in Windows Settings, under Time & language, then Speech.")
        self.stop(quiet=True)
        if voice:
            self.speaker.set_voice(voice)
        self.speaker.set_rate(rate)
        self.view = view
        self.state = "reading"
        self._generation += 1
        self._fetch(view.offset if from_offset is None else from_offset)
        self.state_changed.emit("reading")
        return None

    def _fetch(self, offset: int) -> None:
        generation = self._generation
        self.view.sentences_from(max(0, offset), lambda sentences: self._got(generation, sentences))

    def _got(self, generation: int, sentences) -> None:
        if generation != self._generation or self.state == "stopped":
            return
        sentences = [s for s in sentences or [] if isinstance(s, dict) and s.get("text")]
        if not sentences:
            self._next_chapter()
            return
        self.sentences, self.index = sentences, 0
        if self.state == "reading":
            self._speak()

    def _speak(self) -> None:
        sentence = self.sentences[self.index]
        self.view.mark_sentence(int(sentence["start"]), int(sentence["end"]))
        self.speaker.say(str(sentence["text"]))

    def _done(self) -> None:
        if self.state != "reading":
            return
        self.index += 1
        if self.index < len(self.sentences):
            self._speak()
        else:
            self._next_chapter()

    def _next_chapter(self) -> None:
        view = self.view
        if view is None or view.section + 1 >= view.section_count():
            self.stop(quiet=True)
            self.state_changed.emit("finished")
            return
        generation = self._generation

        def when_ready() -> None:
            view.ready.disconnect(when_ready)
            if generation == self._generation and self.state == "reading":
                self._fetch(0)
        view.ready.connect(when_ready)
        view.go_to(view.section + 1, 0)

    def pause(self) -> None:
        if self.state == "reading":
            self.state = "paused"
            self.speaker.stop()
            self.state_changed.emit("paused")

    def resume(self) -> None:
        if self.state == "paused" and self.sentences:
            self.state = "reading"
            self._speak()                 # the interrupted sentence again, from its start
            self.state_changed.emit("reading")

    def stop(self, quiet: bool = False) -> None:
        was = self.state
        self._generation += 1
        self.state = "stopped"
        if self.speaker is not None:
            self.speaker.stop()
        if self.view is not None:
            self.view.clear_speech()
        if not quiet and was != "stopped":
            self.state_changed.emit("stopped")

    def set_rate(self, rate: float) -> None:
        if self.speaker is not None:
            self.speaker.set_rate(rate)

    def set_voice(self, name: str) -> None:
        if self.speaker is not None:
            self.speaker.set_voice(name)


# ---------------------------------------------------------------- dictionary

class LookupResult(NamedTuple):
    word: str
    meanings: list[tuple[str, str]]      # (part of speech, meaning)
    error: str | None                     # None, "not found" or "offline"


def clean_word(text: str) -> str:
    """The one word to look up (only that word is ever sent), without punctuation around it."""
    m = re.search(r"[^\W_]+(?:['\u2019-][^\W_]+)*", str(text))
    return m.group(0)[:60] if m else ""


_manager: QNetworkAccessManager | None = None


def lookup_word(word: str, done: Callable[[LookupResult], None], url_template: str = DICTIONARY_URL,
                timeout_ms: int = 10000) -> None:
    """Ask the dictionary about word; done(result) is called on the window's thread."""
    global _manager
    if _manager is None:
        _manager = QNetworkAccessManager()
    encoded = bytes(QUrl.toPercentEncoding(word.lower())).decode("ascii")
    request = QNetworkRequest(QUrl(url_template.format(encoded)))
    request.setTransferTimeout(timeout_ms)
    request.setHeader(QNetworkRequest.KnownHeaders.UserAgentHeader, "BookReader/1.0")
    reply = _manager.get(request)

    def finished() -> None:
        status = reply.attribute(QNetworkRequest.Attribute.HttpStatusCodeAttribute)
        data = bytes(reply.readAll())
        error = reply.error()
        reply.deleteLater()
        if status == 404:
            done(LookupResult(word, [], "not found"))
            return
        if error != QNetworkReply.NetworkError.NoError and not status:
            done(LookupResult(word, [], "offline"))
            return
        meanings: list[tuple[str, str]] = []
        try:
            entries = json.loads(data.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            entries = []
        for entry in entries if isinstance(entries, list) else []:
            for meaning in entry.get("meanings", []) if isinstance(entry, dict) else []:
                part = str(meaning.get("partOfSpeech", "")) if isinstance(meaning, dict) else ""
                for definition in meaning.get("definitions", []) if isinstance(meaning, dict) else []:
                    text = definition.get("definition") if isinstance(definition, dict) else None
                    if isinstance(text, str) and text.strip() and len(meanings) < 10:
                        meanings.append((part, text.strip()))
        done(LookupResult(word, meanings, None if meanings else "not found"))

    reply.finished.connect(finished)


class LookupWindow(QDialog):
    """Shows a word's meanings in big words, with a button to look further on the web."""

    def __init__(self, word: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Look up a word")
        self.word = word
        self.has_answer = False
        self.setMinimumSize(760, 480)
        layout = QVBoxLayout(self)
        self.heading = QLabel(word)
        self.heading.setStyleSheet(f"font-size: {ui.words(1.4):g}pt; font-weight: bold;")
        layout.addWidget(self.heading)
        self.body = QLabel(f"Looking up “{html.escape(word)}”…")
        self.body.setWordWrap(True)
        self.body.setTextFormat(Qt.TextFormat.RichText)
        self.body.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.body)
        layout.addWidget(scroll, 1)
        row = QHBoxLayout()
        self.web_button = ui.wide_button("Look it up on the web", self.open_web, 56,
                                         "Open the word in Wiktionary in your web browser")
        row.addWidget(self.web_button)
        row.addStretch(1)
        row.addWidget(ui.wide_button("Close", self.close, 56))
        layout.addLayout(row)

    def show_result(self, result: LookupResult) -> None:
        self.has_answer = True
        word = html.escape(result.word)
        if result.error == "offline":
            self.body.setText("The dictionary couldn't be reached. Is the computer connected to the internet?")
        elif result.error:
            self.body.setText(f"Sorry, the dictionary couldn't find “{word}”. "
                              "It knows single English words best. You can still look on the web.")
        else:
            rows = "".join(f"<p><b>{html.escape(part)}</b>: {html.escape(text)}</p>" for part, text in result.meanings)
            self.body.setText(rows)

    def open_web(self) -> None:
        ui.open_url(WEB_LOOKUP_URL.format(bytes(QUrl.toPercentEncoding(self.word.replace(" ", "_"))).decode()))
