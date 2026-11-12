"""
proofreader_word.py

Microsoft Word's part of the Proofreader. Word opens the .doc, .docx and .rtf
copies, hands over the text of every paragraph, and puts each correction back
in place so fonts, bold, tables and layout stay exactly as they were. Word's
own spelling checker gives the second opinion.

Word runs hidden and separately from any Word window the user has open.
"""

import re

import comtypes
import comtypes.client

from proofreader_engine import Note, Skip, _whole_words

# Parts of a document that hold the writer's words (Word's story numbers):
# main text, footnotes, endnotes, text boxes, headers and footers.
STORIES = {1, 2, 3, 5, 6, 7, 8, 9, 10, 11}
CELL_END = "\r\x07"
WD_FORMAT_DOCX = 16
MAX_GRAMMAR_NOTES = 60
WD_PARAGRAPH = 4
NO_PASSWORD = "not-a-real-password"
MACROS_OFF = 3              # msoAutomationSecurityForceDisable


def _clean(raw, first_position):
    """Text fit for a checker, and where each of its characters sits in the
    document (None when positions can't be trusted)."""
    text, places = [], []
    for offset, char in enumerate(raw):
        if char == "\x0b":
            char = "\n"                                # a line break inside a paragraph
        elif char < " " and char != "\t":
            continue                                   # pictures, note marks and the like
        text.append(char)
        places.append(None if first_position is None else first_position + offset)
    return "".join(text), places


class _Paragraph:
    def __init__(self, story, start, end, raw, exact):
        self.story = story
        self.start = start
        self.end = end
        self.exact = exact                              # can corrections go in by position?
        self.text, self.places = _clean(raw, start if exact else None)


class WordDocument:
    def __init__(self, word, path):
        self.word = word
        try:
            # The made-up password is ignored by ordinary documents; a locked
            # one then fails at once instead of waiting on a hidden prompt.
            self.doc = word.app.Documents.Open(path, False, False, False, NO_PASSWORD)
        except Exception as error:
            raise Skip("Word could not open it (" + _reason(error) + ")")
        try:
            if self.doc.ReadOnly:
                raise Skip("it is open in another program or marked read-only")
            if self.doc.ProtectionType != -1:
                raise Skip("it is protected against changes")
            if self.doc.Revisions.Count:
                raise Skip("it has tracked changes waiting; accept or reject them in Word first")
        except Skip:
            self.close()
            raise
        self.paragraphs = []
        self.writers = {}

    # ------------------------------------------------------------ reading
    def _stories(self):
        for story in self.doc.StoryRanges:
            while story is not None:
                if story.StoryType in STORIES:
                    yield story
                try:
                    story = story.NextStoryRange
                except Exception:
                    story = None

    def read(self):
        """The text of every paragraph, in document order."""
        self.paragraphs = []
        self.writers = {}
        for story in self._stories():
            self._cut(story, self._handle(story), story.Start, story.End)
        return [p.text for p in self.paragraphs]

    @staticmethod
    def _handle(story):
        """A movable range in this story that reads hidden text too, so every
        character it returns stands for one place in the document."""
        handle = story.Duplicate
        mode = handle.TextRetrievalMode
        mode.IncludeHiddenText = True
        mode.IncludeFieldCodes = False
        return handle

    def _cut(self, story, handle, start, end):
        """Take a stretch whole when its characters can be counted. When a
        field (page number, link, contents) throws the count off, halve the
        stretch at a paragraph, so only the paragraphs that hold fields are
        handled one at a time. Each call to Word is slow; this keeps them few."""
        if start >= end:
            return
        handle.SetRange(start, end)
        raw = handle.Text or ""
        if len(raw) - raw.count(CELL_END) == end - start:
            self._split(story, raw, start)
            return
        middle = (start + end) // 2
        handle.SetRange(middle, middle)
        handle.Expand(WD_PARAGRAPH)
        first, last = max(handle.Start, start), min(handle.End, end)
        if not first <= middle < last:              # Word gave no usable paragraph
            handle.SetRange(start, end)
            self._one_by_one(story, handle)
        elif first == start and last == end:        # a single paragraph with a field in it
            body = raw[:-2] if raw.endswith(CELL_END) else raw.rstrip("\r\x0c")
            self.paragraphs.append(_Paragraph(story, start, end, body, False))
        else:
            self._cut(story, handle, start, first)
            self._cut(story, handle, first, last)
            self._cut(story, handle, last, end)

    def _split(self, story, raw, position):
        """Every character sits where it is counted, so paragraphs can be cut
        from the story's text without asking Word for each one."""
        for piece in re.split(r"(\r\x07|\r|\x0c)", raw):
            if piece in (CELL_END, "\r", "\x0c"):
                position += 1
            elif piece:
                self.paragraphs.append(
                    _Paragraph(story, position, position + len(piece), piece, True))
                position += len(piece)

    def _one_by_one(self, story, stretch):
        """The slow, sure way: ask Word for each paragraph in the stretch."""
        for paragraph in stretch.Paragraphs:
            rng = paragraph.Range
            rng.TextRetrievalMode.IncludeHiddenText = True
            raw, start, end = rng.Text or "", rng.Start, rng.End
            body = raw[:-2] if raw.endswith(CELL_END) else raw.rstrip("\r\x0c")
            marks = 1 if len(raw) > len(body) else 0
            exact = len(body) + marks == end - start
            self.paragraphs.append(_Paragraph(story, start, end, body, exact))

    # ------------------------------------------------------------ writing
    def apply(self, edits):
        """Put corrections into the document, last first so earlier places
        stay put. Returns the ones that went in."""
        done = []
        for edit in sorted(edits, key=lambda e: (e.para, e.start), reverse=True):
            paragraph = self.paragraphs[edit.para]
            try:
                went_in = (self._by_position(paragraph, edit) if paragraph.exact
                           else self._by_search(paragraph, edit))
            except Exception:
                went_in = False
            if went_in:
                done.append(edit)
        done.reverse()
        return done

    def _writer(self, story):
        """One handle per part of the document, reused for every correction."""
        if id(story) not in self.writers:
            self.writers[id(story)] = self._handle(story)
        return self.writers[id(story)]

    def _by_position(self, paragraph, edit):
        places, text = paragraph.places, paragraph.text
        if edit.start == edit.end:
            if edit.start < len(places):
                first = places[edit.start]
                if edit.start and places[edit.start - 1] + 1 != first:
                    return False
            elif places:
                first = places[-1] + 1
            else:
                return False
            last = first
        else:
            first, last = places[edit.start], places[edit.end - 1] + 1
            if last - first != edit.end - edit.start:
                return False                            # something unseen sits in between
        rng = self._writer(paragraph.story)
        # Look before writing: the document must hold exactly the expected
        # characters at this place (for an insertion, the character beside it).
        if first != last:
            check, expected = (first, last), text[edit.start:edit.end]
        elif edit.start:
            check, expected = (first - 1, first), text[edit.start - 1]
        else:
            check, expected = (first, first + 1), text[:1]
        rng.SetRange(*check)
        if (rng.Text or "").replace("\x0b", "\n") != expected:
            return False
        rng.SetRange(first, last)
        if first == last:
            rng.InsertAfter(edit.new)
        else:
            rng.Text = edit.new
        return True

    def _by_search(self, paragraph, edit):
        text = paragraph.text
        a, b = _whole_words(text, edit.start, edit.end)
        if edit.start == edit.end and a == b:           # between words: take a neighbour along
            a, b = _whole_words(text, max(0, edit.start - 1), min(len(text), edit.end + 1))
        old = text[a:b]
        new = text[a:edit.start] + edit.new + text[edit.end:b]
        if not old.strip() or len(old) > 250 or "\n" in old or "\t" in old:
            return False
        earlier = [m.start() for m in re.finditer(re.escape(old), text)]
        if a not in earlier:
            return False
        rng = self._writer(paragraph.story)
        rng.SetRange(paragraph.start, paragraph.end)
        for _ in range(earlier.index(a) + 1):
            find = rng.Find
            find.ClearFormatting()
            if not find.Execute(old.replace("^", "^^"), True, False, False, False, False, True, 0):
                return False
            if rng.End > paragraph.end:
                return False
            found_start, found_end = rng.Start, rng.End
            rng.SetRange(found_end, paragraph.end)
        rng.SetRange(found_start, found_end)
        if rng.Text != old:
            return False
        rng.Text = new
        return True

    # ------------------------------------------------------------ Word's grammar opinion
    def grammar_notes(self):
        """Sentences Word's grammar check questions. Word offers no fix through
        automation, so these are only listed."""
        notes = []
        try:
            errors = self.doc.Content.GrammaticalErrors
            for index in range(1, min(errors.Count, MAX_GRAMMAR_NOTES) + 1):
                sentence = " ".join((errors.Item(index).Text or "").split())
                if sentence:
                    notes.append(Note(-1, 0, 0, "grammar",
                                      "Word's grammar check questions: \"" + sentence[:160] + "\""))
        except Exception:
            pass
        return notes

    def save(self):
        self.doc.Save()

    def close(self):
        try:
            self.doc.Close(0)
        except Exception:
            pass


class WordSpeller:
    """Word's dictionary: None for a word it accepts, else its suggestions."""

    def __init__(self, word):
        self.word = word
        self.known = {}

    def suggest(self, text):
        if text not in self.known:
            self.known[text] = self._ask(text)
        return self.known[text]

    def _ask(self, text):
        app = self.word.app
        try:
            if app.CheckSpelling(text):
                return None
            self.word.need_a_document()
            return [s.Name for s in app.GetSpellingSuggestions(text)]
        except Exception:
            return None                                 # no opinion is safer than a wrong one


class Word:
    """A hidden copy of Word used only by the Proofreader."""

    def __init__(self):
        comtypes.CoInitialize()
        try:
            self.app = comtypes.client.CreateObject("Word.Application", dynamic=True)
        except Exception as error:
            raise RuntimeError("Microsoft Word could not be started (" + _reason(error) + ")")
        # never take over, or later close, a Word window the user is working in
        self.ours = not self.app.Visible and self.app.Documents.Count == 0
        if self.ours:
            self.app.DisplayAlerts = 0
            # Word opened by a program runs a document's macros unless told not to.
            self.app.AutomationSecurity = MACROS_OFF
            self.app.ScreenUpdating = False     # nothing is shown, so don't draw it
        self.scratch = None

    def open(self, path):
        return WordDocument(self, path)

    def speller(self):
        return WordSpeller(self)

    def need_a_document(self):
        """Word only suggests spellings while some document is open."""
        if self.app.Documents.Count == 0:
            self.scratch = self.app.Documents.Add()

    def changes_copy(self, original, corrected, target):
        """A Word file showing every correction as a tracked change, so each
        can be accepted or rejected."""
        before = self.app.Documents.Open(original, False, True, False, NO_PASSWORD)
        try:
            after = self.app.Documents.Open(corrected, False, True, False, NO_PASSWORD)
            try:
                shown = self.app.CompareDocuments(before, after, 2, 1)
                try:
                    shown.SaveAs2(target, WD_FORMAT_DOCX)
                finally:
                    shown.Close(0)
            finally:
                after.Close(0)
        finally:
            before.Close(0)

    def close(self):
        if self.ours:
            # This hidden Word is the Proofreader's alone, so nothing in it is
            # the user's: shut every document without saving and leave.
            try:
                for document in list(self.app.Documents):
                    try:
                        document.Close(0)
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self.app.Quit(0)
            except Exception:
                pass
        elif self.scratch is not None:
            try:
                self.scratch.Close(0)
            except Exception:
                pass
        self.scratch = None
        self.app = None
        comtypes.CoUninitialize()


def _reason(error):
    text = str(error)
    details = getattr(error, "details", None)
    if details and details[0]:
        text = str(details[0])
    return " ".join(text.split())[:200]
