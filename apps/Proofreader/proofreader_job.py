"""
proofreader_job.py

Runs a batch for the Proofreader: finds the documents, makes a copy of each in
a "corrected" folder beside it, corrects the copy, and writes a report of every
change and everything left for the reader to decide.

The originals are only ever read.
"""

import dataclasses
import datetime
import json
import os
import shutil
from dataclasses import dataclass, field

from proofreader_engine import (OFF_CATEGORIES, OFF_RULES, Options, Skip, _whole_words, review,
                                snippet)

NO_WORD = ("Microsoft Word was not found on this computer, so its check was skipped and only "
           "the free checker was used.")
NEEDS_WORD = ("This kind of file needs Microsoft Word. Without Word, save it as .docx "
              "(or .txt) and it can be corrected.")


def word_installed():
    """Is Microsoft Word on this computer? (It registers itself with Windows.)"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except (ImportError, OSError):
        return False

WORD_TYPES = (".doc", ".docx", ".rtf")
TYPES = WORD_TYPES + (".txt",)
OUT_FOLDER = "corrected"
CHANGES_SUFFIX = " - changes shown.docx"
MAX_PASSES = 4          # a fix can uncover another; four looks settle nearly everything
PART_SIZE = 50          # documents done before Word and the checker are started afresh
LABELS = {"spelling": "Spelling", "grammar": "Grammar", "punctuation": "Punctuation",
          "capitals": "Capitals", "runons": "Run-on", "name": "Name"}


@dataclass
class Settings:
    options: Options = field(default_factory=Options)
    use_word: bool = True           # Word's checker as the second opinion
    show_changes: bool = False      # also save a copy with tracked changes
    subfolders: bool = False
    word_grammar: bool = False      # also list sentences Word's grammar check questions (slow)


@dataclass
class Change:
    kind: str
    before: str
    after: str
    where: str


@dataclass
class FileResult:
    source: str
    output: str = ""
    changes_copy: str = ""
    changes: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    error: str = ""
    stopped: bool = False           # cut off part-way; not counted as done


# ---------------------------------------------------------------- finding documents
def is_document(path):
    name = os.path.basename(path)
    return (name.lower().endswith(TYPES) and not name.startswith(("~$", "Proofreading report "))
            and not name.lower().endswith(CHANGES_SUFFIX))


def collect(paths, subfolders=False):
    """The documents in a mix of files and folders, each once, in order."""
    found, seen = [], set()

    def add(path):
        key = os.path.normcase(os.path.abspath(path))
        if key not in seen and is_document(path):
            seen.add(key)
            found.append(os.path.abspath(path))

    for path in paths:
        if os.path.isdir(path):
            for folder, folders, files in os.walk(path):
                folders[:] = sorted(f for f in folders if f.lower() != OUT_FOLDER) \
                    if subfolders else []
                for name in sorted(files, key=str.lower):
                    add(os.path.join(folder, name))
        elif os.path.isfile(path):
            add(path)
    return found


def output_path(source):
    """Where the corrected copy goes; an earlier copy is never written over."""
    folder = os.path.join(os.path.dirname(source), OUT_FOLDER)
    name, ext = os.path.splitext(os.path.basename(source))
    target = os.path.join(folder, name + ext)
    number = 2
    while os.path.exists(target):
        target = os.path.join(folder, f"{name} ({number}){ext}")
        number += 1
    return target


# ---------------------------------------------------------------- plain text files
class TextDocument:
    """A .txt file: every line is a paragraph. Its encoding and line endings
    are kept as they were."""

    def __init__(self, path):
        with open(path, "rb") as f:
            data = f.read()
        self.encoding = self._sniff(data)
        text = data.decode(self.encoding)
        self.lines, self.ends = [], []
        for line in text.splitlines(keepends=True):
            body = line.rstrip("\r\n")
            self.lines.append(body)
            self.ends.append(line[len(body):])

    @staticmethod
    def _sniff(data):
        if data.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        if data.startswith((b"\xff\xfe", b"\xfe\xff")):
            return "utf-16"
        try:
            data.decode("utf-8")
            return "utf-8"
        except UnicodeDecodeError:
            return "cp1252"

    def read(self):
        return list(self.lines)

    def apply(self, edits):
        done = []
        for edit in sorted(edits, key=lambda e: (e.para, e.start), reverse=True):
            try:
                edit.new.encode(self.encoding)
            except UnicodeEncodeError:
                continue                    # this file's alphabet has no such character
            line = self.lines[edit.para]
            self.lines[edit.para] = line[:edit.start] + edit.new + line[edit.end:]
            done.append(edit)
        done.reverse()
        return done

    def grammar_notes(self):
        return []

    def save_as(self, path):
        text = "".join(body + end for body, end in zip(self.lines, self.ends))
        with open(path, "wb") as f:
            f.write(text.encode(self.encoding, errors="replace"))

    def close(self):
        pass


# ---------------------------------------------------------------- correcting one document
def proofread(document, options, checker, speller=None, stop=lambda: False, pause=lambda: None):
    """Correct an open document in place. Returns (changes, notes)."""
    changes, texts, unmade = [], [], []
    notes = {}                      # paragraph number -> its notes, from its latest look
    looked_at = None                # the text of each paragraph when last looked at
    settled = False
    for _ in range(MAX_PASSES):
        pause()
        if stop():
            break
        texts = document.read()
        only = _changed(looked_at, texts)
        edits, found = review(texts, options, checker, speller, only)
        _keep_notes(notes, found, only, len(texts))
        looked_at = list(texts)
        applied = document.apply(edits) if edits else []
        went_in = {id(edit) for edit in applied}
        unmade = [edit for edit in edits if id(edit) not in went_in]
        changes += [Change(edit.kind, *_before_after(texts[edit.para], edit),
                           snippet(texts[edit.para], edit.start, edit.end)) for edit in applied]
        if not applied:
            settled = True
            break
    if not settled and not stop():                  # list what is left as it now reads
        texts = document.read()
        only = _changed(looked_at, texts)
        unmade, found = review(texts, options, checker, speller, only)
        _keep_notes(notes, found, only, len(texts))
    lines = _note_lines([n for para in sorted(notes) for n in notes[para]], texts)
    for edit in unmade:                             # a fix that never went in must not vanish
        before, after = _before_after(texts[edit.para], edit)
        lines.append(f"{LABELS.get(edit.kind, edit.kind)} - \"{before}\" should probably be "
                     f"\"{after}\", but it could not be changed safely in place.")
    return changes, lines


def _changed(before, now):
    """The paragraphs whose words changed since the last look, or None to
    look at them all. Only these need checking again: the checkers' answers
    for the rest have not changed, and checking a whole long document again
    after a handful of fixes was most of the time spent."""
    if before is None or len(before) != len(now):
        return None
    return {i for i, (old, new) in enumerate(zip(before, now)) if old != new}


def _keep_notes(notes, found, only, count):
    """Replace the notes of the paragraphs just looked at."""
    looked = range(count) if only is None else only
    for para in looked:
        notes.pop(para, None)
    for note in found:
        notes.setdefault(note.para, []).append(note)


def _before_after(text, edit):
    a, b = _whole_words(text, edit.start, edit.end)
    return text[a:b], text[a:edit.start] + edit.new + text[edit.end:b]


def _note_lines(notes, texts):
    """Notes as report lines; the same word questioned many times is one line."""
    lines, counts = [], {}
    for note in notes:
        if note.para < 0:
            line = note.text
        else:
            text = texts[note.para]
            words = text[note.start:note.end].strip()
            if note.kind == "runons":
                line = f"{note.text}  \"{snippet(text, note.start, min(note.end, note.start + 90), around=0)}\""
            elif note.kind in ("name", "spelling"):
                line = f"\"{words}\": {note.text}"
            else:
                line = f"\"{snippet(text, note.start, note.end, around=20)}\": {note.text}"
            if note.choices:
                line += "  Maybe: " + ", ".join(note.choices)
        line = f"{LABELS.get(note.kind, note.kind)} - {line}"
        if line in counts:
            counts[line] += 1
        else:
            counts[line] = 1
            lines.append(line)
    return [line + (f"  ({counts[line]} times)" if counts[line] > 1 else "") for line in lines]


# ---------------------------------------------------------------- a batch saved as it goes
class Journal:
    """A batch's progress, written down after every document, so a batch that
    was stopped, closed or cut off (a power cut, a crash) can be picked up
    where it left off instead of starting again.

    One JSON line per event, only ever added to, so saving stays just as quick
    for the thousandth document as for the first:
        {"batch": ...}      what to do: the documents, the choices, when
        {"started": ...}    a document begun, and the copy it was writing to
        {"done": ...}       a document finished, with everything for the report
    """

    def __init__(self, path):
        self.path = path

    # ------------------------------------------------------------ writing
    def begin(self, sources, files, settings, word_missing=False):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        batch = {"sources": list(sources), "files": list(files),
                 "settings": settings_to_dict(settings),
                 "started": datetime.datetime.now().isoformat(timespec="seconds"),
                 "word_missing": word_missing}
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(json.dumps({"batch": batch}) + "\n")

    def _add(self, entry):
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
            f.flush()
            os.fsync(f.fileno())                    # on the disk before the next document

    def started(self, source, target):
        self._add({"started": source, "target": target})

    def done(self, result):
        self._add({"done": dataclasses.asdict(result)})

    def word_was_missing(self):
        self._add({"word_missing": True})

    def finish(self):
        """The whole batch is done: nothing is left to pick up."""
        try:
            os.remove(self.path)
        except FileNotFoundError:
            pass

    # ------------------------------------------------------------ reading
    def load(self):
        """What is left of an earlier batch, or None when there is nothing."""
        try:
            with open(self.path, encoding="utf-8") as f:
                lines = f.read().splitlines()
        except OSError:
            return None
        batch, results, open_copies, word_missing = None, {}, {}, False
        for line in lines:
            try:
                entry = json.loads(line)
            except ValueError:
                continue                            # a line cut short by a power cut
            if "batch" in entry:
                batch = entry["batch"]
                word_missing = bool(batch.get("word_missing"))
            elif "started" in entry:
                open_copies[_key(entry["started"])] = entry.get("target", "")
            elif "done" in entry:
                result = result_from_dict(entry["done"])
                results[_key(result.source)] = result
                open_copies.pop(_key(result.source), None)
            elif entry.get("word_missing"):
                word_missing = True
        if not batch or not batch.get("files"):
            return None
        left = [f for f in batch["files"] if _key(f) not in results]
        if not left:
            return None
        return Unfinished(sources=batch.get("sources") or batch["files"], files=batch["files"],
                          left=left, done=list(results.values()),
                          settings=settings_from_dict(batch.get("settings", {})),
                          started=batch.get("started", ""), word_missing=word_missing,
                          half_done=[t for t in open_copies.values() if t])


@dataclass
class Unfinished:
    sources: list
    files: list
    left: list
    done: list
    settings: Settings
    started: str
    word_missing: bool
    half_done: list

    def when(self):
        try:
            moment = datetime.datetime.fromisoformat(self.started)
        except ValueError:
            return ""
        return f"{moment:%A, %B} {moment.day} at {moment:%I:%M %p}".replace(" 0", " ")


def _key(path):
    return os.path.normcase(os.path.abspath(path))


def settings_to_dict(settings):
    return {"options": dataclasses.asdict(settings.options), "use_word": settings.use_word,
            "show_changes": settings.show_changes, "subfolders": settings.subfolders,
            "word_grammar": settings.word_grammar}


def settings_from_dict(data):
    known = {f.name for f in dataclasses.fields(Options)}
    options = {k: v for k, v in data.get("options", {}).items() if k in known}
    options["keep_words"] = tuple(options.get("keep_words", ()))
    return Settings(Options(**options), bool(data.get("use_word", True)),
                    bool(data.get("show_changes", False)), bool(data.get("subfolders", False)),
                    bool(data.get("word_grammar", False)))


def result_from_dict(data):
    changes = [Change(**c) for c in data.get("changes", [])]
    known = {f.name for f in dataclasses.fields(FileResult)} - {"changes"}
    return FileResult(changes=changes, **{k: v for k, v in data.items() if k in known})


def tidy_half_done(paths):
    """A document cut off part-way leaves the start of its corrected copy
    behind. It is the Proofreader's own unfinished file, never an original, so
    it is removed before the document is done again."""
    for target in paths:
        if os.path.basename(os.path.dirname(target)).lower() != OUT_FOLDER:
            continue                                # only ever inside a "corrected" folder
        for path in (target, os.path.splitext(target)[0] + CHANGES_SUFFIX):
            try:
                os.remove(path)
            except OSError:
                pass


# ---------------------------------------------------------------- a whole batch
def run_in_parts(files, settings, open_tools, close_tools, progress=None, stop=lambda: False,
                 pause=lambda: None, journal=None, part=PART_SIZE):
    """Proofread the files in parts of `part` documents, with Word and the
    checker started afresh for each part: open_tools() returns (checker,
    word) and close_tools(checker, word) shuts them, which hands back all the
    memory they gathered. Word's answers about words already checked are
    kept from part to part, so nothing is asked twice."""
    results, known = [], {}
    for first in range(0, len(files), part):
        piece = files[first:first + part]
        pause()
        if stop():
            break
        checker, word = open_tools()
        try:
            step = None
            if progress:
                def step(done, total, name, first=first):
                    if name:                        # the end of a part is not the end
                        progress(first + done, len(files), name)
            got = run(piece, settings, checker, word, step, stop, pause, journal, known)
        finally:
            close_tools(checker, word)
        results += got
        if len(got) < len(piece):
            break                                   # stopped part-way
    if progress:
        progress(len(results), len(files), "")
    return results


def run(files, settings, checker, word=None, progress=None, stop=lambda: False,
        pause=lambda: None, journal=None, known=None):
    """Proofread each file. word is a proofreader_word.Word, or None.
    progress(index, total, file name). pause() returns once the user lets the
    work carry on. journal, a Journal, notes each document as it is done.
    known holds Word's answers about words already checked."""
    checker.off_rules, checker.off_categories = OFF_RULES, OFF_CATEGORIES
    speller = word.speller(known) if (word and settings.use_word) else None
    results = []
    for index, source in enumerate(files):
        pause()
        if stop():
            break
        if progress:
            progress(index, len(files), os.path.basename(source))
        result = _one(source, settings, checker, word, speller, stop, pause, journal)
        if result.stopped:
            break
        results.append(result)
        if journal:
            journal.done(result)
    if progress:
        progress(len(results), len(files), "")
    return results


def _one(source, settings, checker, word, speller, stop, pause=lambda: None, journal=None):
    result = FileResult(source)
    is_text = source.lower().endswith(".txt")
    if not os.path.isfile(source):
        result.error = "The document is no longer there (moved, renamed or deleted)."
        return result
    if not is_text and word is None and not source.lower().endswith(".docx"):
        result.error = NEEDS_WORD
        return result
    target = output_path(source)
    if journal:
        journal.started(source, target)
    document = None
    try:
        os.makedirs(os.path.dirname(target), exist_ok=True)
        if is_text:
            document = TextDocument(source)
        else:
            shutil.copyfile(source, target)             # the copy is what gets corrected
            document = word.open(target) if word else _without_word(target)
        result.changes, result.notes = proofread(document, settings.options, checker, speller,
                                                 stop, pause)
        if stop():
            result.stopped = True                   # unfinished: done again when picked up
            result.changes = []
        elif settings.use_word and settings.word_grammar:
            # Word works out the grammar of the whole document for this, which
            # took about as long as everything else together, so it is asked for.
            result.notes += _note_lines(document.grammar_notes(), [])
        if result.changes:
            if is_text:
                document.save_as(target)
            else:
                document.save()
            result.output = target
    except Exception as error:
        result.error = str(error) if isinstance(error, Skip) else f"{type(error).__name__}: {error}"
    finally:
        if document is not None:
            document.close()
    if not result.output and not is_text and os.path.exists(target):
        os.remove(target)                               # our own untouched copy
    if result.output and settings.show_changes and word and not is_text and not stop():
        shown = os.path.splitext(target)[0] + CHANGES_SUFFIX
        try:
            word.changes_copy(source, target, shown)
            result.changes_copy = shown
        except Exception as error:
            result.notes.append(f"The copy showing the changes could not be made ({error}).")
    return result


def _without_word(path):
    """A .docx can be corrected directly when there is no Word to do it."""
    from proofreader_docx import DocxDocument
    return DocxDocument(path)


# ---------------------------------------------------------------- the report
def report_text(results, when=None, word_missing=False):
    when = when or datetime.datetime.now()
    total = sum(len(r.changes) for r in results)
    lines = [f"Proofreading report - {when:%B %d, %Y  %I:%M %p}",
             f"{len(results)} document(s) checked, {total} correction(s) made."]
    if word_missing:
        lines.append(NO_WORD)
    lines.append("")
    for r in results:
        lines.append("=" * 72)
        lines.append(os.path.basename(r.source))
        lines.append(f"   in {os.path.dirname(r.source)}")
        if r.error:
            lines += [f"   NOT CORRECTED: {r.error}", ""]
            continue
        if r.output:
            lines.append(f"   corrected copy: {os.path.join(OUT_FOLDER, os.path.basename(r.output))}")
        else:
            lines.append("   Nothing needed correcting, so no copy was made.")
        if r.changes_copy:
            lines.append(f"   changes shown in: {os.path.join(OUT_FOLDER, os.path.basename(r.changes_copy))}")
        if r.changes:
            lines += ["", f"   CORRECTED ({len(r.changes)})"]
            for c in r.changes:
                lines.append(f"     {LABELS.get(c.kind, c.kind):<12} \"{c.before}\"  ->  \"{c.after}\""
                             f"     [{c.where}]")
        if r.notes:
            lines += ["", f"   LEFT FOR YOU TO DECIDE ({len(r.notes)})"]
            lines += [f"     {note}" for note in r.notes]
        lines.append("")
    return "\n".join(lines)


def save_report(results, when=None, word_missing=False):
    """One report per folder that got corrected copies. Returns the paths."""
    when = when or datetime.datetime.now()
    by_folder = {}
    for r in results:
        by_folder.setdefault(os.path.join(os.path.dirname(r.source), OUT_FOLDER), []).append(r)
    paths = []
    for folder, group in by_folder.items():
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"Proofreading report {when:%Y-%m-%d %H.%M.%S}.txt")
        with open(path, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(report_text(group, when, word_missing))
        paths.append(path)
    return paths
