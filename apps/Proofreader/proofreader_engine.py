"""
proofreader_engine.py

The rules of the Proofreader: given the paragraphs of a document and what two
checkers say about them, decide what to change and what only to list.

Nothing here touches files, Word or the network, so it can be tested alone.

    review(paragraphs, options, checker, speller) -> (edits, notes)

checker.check_many(texts) returns LanguageTool's matches for each text.
speller.suggest(word) returns None when Word accepts the word, otherwise
Word's suggestions (possibly an empty list). speller may be None.

Swear words are never changed: no rule about them is switched on, and any
change that would touch one is dropped.
"""

import re
from dataclasses import dataclass, field

# ---------------------------------------------------------------- what the user can switch on
KINDS = ("spelling", "grammar", "punctuation", "capitals", "runons")


@dataclass
class Options:
    spelling: bool = True
    grammar: bool = True
    punctuation: bool = True
    capitals: bool = True
    runons: bool = True
    keep_words: tuple = ()          # words the user never wants changed

    def wants(self, kind):
        return getattr(self, kind, False)


@dataclass
class Edit:
    """Replace paragraph[start:end] with new."""
    para: int
    start: int
    end: int
    new: str
    kind: str
    why: str = ""


@dataclass
class Note:
    """Something left for the reader to decide."""
    para: int
    start: int
    end: int
    kind: str
    text: str
    choices: list = field(default_factory=list)


class Skip(Exception):
    """A document the app should leave alone, with the reason."""


# ---------------------------------------------------------------- swearing is left alone
_SWEAR_ANYWHERE = ("fuck|shit|bitch|cunt|goddam|asshole|dumbass|jackass|badass|smartass|bastard|"
                   "bollock|wank|twat|pussy|slut|whore|bullcrap|dipshit|horseshit|douche|bugger|"
                   "nigger|nigga|faggot|tranny")
_SWEAR_WHOLE = ("ass|asses|arse|arses|arsehole|arseholes|hell|damn|damned|damnit|dammit|darn|"
                "dick|dicks|dickhead|dickheads|cock|cocks|cocksucker|cocksuckers|crap|crappy|craps|"
                "piss|pissed|pisses|pissing|pissy|prick|pricks|tit|tits|titties|screw|screwed|"
                "bloody|bleeding|frigging|friggin|freaking|freakin|effing|feck|fecking|sod|git|"
                "fag|fags|retard|retards|retarded|dyke|dykes|spic|spics|kike|kikes|chink|chinks|"
                "jerk|jerkoff|turd|turds|skank|hoe|hoes|ho|balls|ballsy|knob|tosser|shag|shagged")
SWEAR = re.compile(r"\b(?:\w*(?:%s)\w*|(?:%s))\b" % (_SWEAR_ANYWHERE, _SWEAR_WHOLE), re.IGNORECASE)

WORD = re.compile(r"[^\W\d_]+(?:['’][^\W\d_]+)*")


def protected_spans(text, keep_words=()):
    """Places in the text no change may touch."""
    spans = [m.span() for m in SWEAR.finditer(text)]
    if keep_words:
        keep = {w.casefold() for w in keep_words}
        spans += [m.span() for m in WORD.finditer(text) if m.group().casefold() in keep]
    return spans


def _touches(spans, start, end):
    if start == end:                                # an insertion: inside a word counts
        return any(s < start < e for s, e in spans)
    return any(s < end and start < e for s, e in spans)


# ---------------------------------------------------------------- LanguageTool's answers, sorted
# Opinions about wording and tone are not corrections; they stay switched off.
OFF_CATEGORIES = ("STYLE", "COLLOQUIALISMS", "REDUNDANCY", "PLAIN_ENGLISH", "WORDINESS",
                  "CREATIVE_WRITING", "GENDER_NEUTRALITY", "FALSE_FRIENDS", "REGIONALISMS",
                  "TEXT_ANALYSIS", "BRITISH_ENGLISH", "AMERICAN_ENGLISH_STYLE", "SEMANTICS",
                  "ACADEMIC_WRITING", "REPETITIONS_STYLE")
OFF_RULES = ("PROFANITY", "EN_QUOTES", "DASH_RULE", "ELLIPSIS", "MULTIPLICATION_SIGN",
             "ARROWS", "TRADEMARK", "COPYRIGHT", "PLUS_MINUS", "R_SYMBOL", "APOS_ARE",
             "TOO_LONG_SENTENCE", "TOO_LONG_PARAGRAPH")
TYPOGRAPHY_OK = {"DOUBLE_PUNCTUATION", "COMMA_PARENTHESIS_WHITESPACE", "WHITESPACE_RULE",
                 "SENTENCE_WHITESPACE", "UNLIKELY_OPENING_PUNCTUATION", "PUNCTUATION_PARAGRAPH_END",
                 "APOSTROPHE_IN_DATES", "SPACE_BEFORE_PARENTHESIS", "CONSECUTIVE_SPACES"}


def classify(match):
    """Which tick box a LanguageTool match belongs to; None means ignore it."""
    rule = match.get("rule", {})
    rule_id = rule.get("id", "")
    category = rule.get("category", {}).get("id", "")
    if category in OFF_CATEGORIES or rule_id in OFF_RULES:
        return None
    if rule_id.startswith(("MORFOLOGIK", "HUNSPELL")):
        return "dictionary"
    if category == "TYPOS":
        return "spelling"
    if category == "CASING":
        return "capitals"
    if category == "PUNCTUATION":
        return "punctuation"
    if category == "TYPOGRAPHY":
        if rule.get("issueType") == "whitespace" or rule_id in TYPOGRAPHY_OK:
            return "punctuation"
        return None
    return "grammar"


# ---------------------------------------------------------------- how alike two spellings are
def _plain(word):
    """A word without its case, apostrophes, spaces and hyphens: "alot" and
    "a lot" differ only in those, the smallest slip there is."""
    return re.sub(r"['’\s-]", "", word.casefold())


def distance(a, b):
    """Typing mistakes between two words; swapped neighbours count a little
    less than a wrong letter, because that is the commoner slip."""
    a, b = _plain(a), _plain(b)
    if a == b:
        return 0.0
    rows = [[0.0] * (len(b) + 1) for _ in range(len(a) + 1)]
    for i in range(len(a) + 1):
        rows[i][0] = float(i)
    for j in range(len(b) + 1):
        rows[0][j] = float(j)
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            cost = 0.0 if a[i - 1] == b[j - 1] else 1.0
            best = min(rows[i - 1][j] + 1, rows[i][j - 1] + 1, rows[i - 1][j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                best = min(best, rows[i - 2][j - 2] + 0.9)
            rows[i][j] = best
    return rows[-1][-1]


def sure_spelling(word, suggestions):
    """The one suggestion that is clearly meant, or None."""
    if not suggestions:
        return None
    if len(suggestions) == 1:
        return suggestions[0] if distance(word, suggestions[0]) <= 2 else None
    scored = sorted((distance(word, s), i) for i, s in enumerate(suggestions[:8]))
    (best, index), (second, _) = scored[0], scored[1]
    if index == 0 and best < second and best <= 2:
        return suggestions[0]
    return None


# ---------------------------------------------------------------- names
_SENTENCE_END = re.compile(r"[.!?:]['\")\]”’]*$")


def starts_sentence(text, start):
    before = text[:start].rstrip()
    return not before or bool(_SENTENCE_END.search(before)) or before[-1] in "\"“(["


def find_names(texts):
    """Capitalized words used in the middle of a sentence: most likely names."""
    names = set()
    for text in texts:
        for m in WORD.finditer(text):
            word = m.group()
            if word[0].isupper() and len(word) > 1 and not starts_sentence(text, m.start()):
                names.add(word)
    return names


# ---------------------------------------------------------------- our own rules
_NOT_A_SENTENCE_END = {
    "com", "org", "net", "edu", "gov", "io", "co", "uk", "us", "ca", "txt", "doc", "docx", "rtf",
    "pdf", "exe", "jpg", "jpeg", "png", "gif", "mp3", "mp4", "zip", "html", "htm", "js", "py",
    "css", "csv", "xls", "xlsx", "ppt", "pptx", "json", "xml", "md", "ps1", "bat", "wav", "mov"}
_GLUED = re.compile(r"(?<![\w./@\\:])([A-Za-z][a-z]+)([.!?])([A-Za-z][a-z]*)(?![\w./@\\(=\[])")
# After a full stop, only words that commonly open a sentence count; anything
# else ("self.name", "report.final") may be a name with a dot in it.
_OPENERS = set("""the a an i we he she they it you this that these those there then but and so
    if when my our his her their your me in on at for to what why how who where please thanks
    thank no yes also however now after before because as is are was do did can will let not
    all some one""".split())
_ABBREVIATIONS = set("""mr mrs ms dr vs etc st jr sr inc ltd co no fig approx dept est gov prof
    rev gen col sgt ave blvd""".split())

_ADVERBS_WITH_COMMA = "however|therefore|moreover|furthermore|nevertheless|consequently|meanwhile"
_ADVERBS_BARE = "otherwise|instead|thus|hence"
_SPLICE = re.compile(
    r",\s+(%s|%s),?\s+(?=(?:I|we|he|she|they|it|you|this|that|there|the|my|your|his|her|our|"
    r"their|its|a|an|these|those|nobody|everyone|someone|people)\b)"
    % (_ADVERBS_WITH_COMMA, _ADVERBS_BARE), re.IGNORECASE)

_SENTENCE = re.compile(r"[^.!?]+(?:[.!?]+[\"')\]”’]*|$)")
_SUBORDINATORS = {
    "when", "if", "because", "although", "though", "since", "while", "after", "before", "as",
    "once", "unless", "until", "whenever", "wherever", "whether", "where", "even", "so", "that",
    "in", "on", "at", "by", "for", "from", "with", "without", "during", "despite", "like",
    "to", "having", "being", "yes", "no", "well", "oh", "ok", "okay", "sure", "then", "now",
    "which", "who", "whom", "whose", "what", "why", "how", "and", "but", "or", "nor", "yet"}
_COMMA_PRONOUN = re.compile(r",\s+(?:I|we|he|she|they|it|you)\s+[a-z']+", re.IGNORECASE)
LONG_SENTENCE = 45          # words
LONG_PARAGRAPH = 300        # words


# Regular endings put on irregular verbs. Both checkers offer a look-alike
# word for these ("runned" -> "runner"), so they are settled here instead.
IRREGULAR = {
    "runned": ("ran", "run"), "goed": ("went", "gone"), "eated": ("ate", "eaten"),
    "catched": ("caught",), "teached": ("taught",), "buyed": ("bought",),
    "thinked": ("thought",), "bringed": ("brought",), "drived": ("drove", "driven"),
    "writed": ("wrote", "written"), "comed": ("came", "come"), "falled": ("fell", "fallen"),
    "knowed": ("knew", "known"), "growed": ("grew", "grown"), "throwed": ("threw", "thrown"),
    "drawed": ("drew", "drawn"), "breaked": ("broke", "broken"), "speaked": ("spoke", "spoken"),
    "taked": ("took", "taken"), "gived": ("gave", "given"), "heared": ("heard",),
    "telled": ("told",), "selled": ("sold",), "finded": ("found",), "holded": ("held",),
    "keeped": ("kept",), "sleeped": ("slept",), "feeled": ("felt",), "leaved": ("left",),
    "losed": ("lost",), "maked": ("made",), "sayed": ("said",), "sended": ("sent",),
    "spended": ("spent",), "standed": ("stood",), "swimmed": ("swam", "swum"),
    "winned": ("won",), "hurted": ("hurt",), "cutted": ("cut",), "hitted": ("hit",),
    "builded": ("built",), "choosed": ("chose", "chosen"), "drinked": ("drank", "drunk"),
    "flied": ("flew", "flown"), "forgetted": ("forgot", "forgotten"), "getted": ("got",),
    "hided": ("hid", "hidden"), "meeted": ("met",), "rided": ("rode", "ridden"),
    "rised": ("rose", "risen"), "sitted": ("sat",), "stealed": ("stole", "stolen"),
    "sticked": ("stuck",), "weared": ("wore", "worn"), "beginned": ("began", "begun"),
    "bited": ("bit", "bitten"), "blowed": ("blew", "blown"), "digged": ("dug",),
    "feeded": ("fed",), "fighted": ("fought",), "freezed": ("froze", "frozen"),
    "lended": ("lent",), "shaked": ("shook", "shaken"), "shooted": ("shot",),
    "sinked": ("sank", "sunk"), "slided": ("slid",), "striked": ("struck",),
    "sweeped": ("swept",), "swinged": ("swung",), "teared": ("tore", "torn"),
    "weeped": ("wept",)}
_IRREGULAR = re.compile(r"\b(?:%s)\b" % "|".join(IRREGULAR), re.IGNORECASE)
_HELPERS = {"have", "has", "had", "having", "be", "been", "being", "is", "are", "was", "were",
            "am", "get", "gets", "got", "ve", "d"}


def _same_case(model, word):
    if model.isupper() and len(model) > 1:
        return word.upper()
    return word.capitalize() if model[:1].isupper() else word


def own_matches(text, options):
    """Fixes LanguageTool does not make, in LanguageTool's shape."""
    found = []
    if options.grammar:
        for m in _IRREGULAR.finditer(text):
            forms = IRREGULAR[m.group().lower()]
            before = WORD.findall(text[:m.start()])[-2:]
            helped = any(w.lower().split("'")[-1] in _HELPERS for w in before)
            pick = forms[-1] if helped else forms[0]
            found.append({"offset": m.start(), "length": len(m.group()),
                          "replacements": [{"value": _same_case(m.group(), pick)}],
                          "message": "irregular verb", "kind": "grammar"})
    if options.punctuation:
        for m in _GLUED.finditer(text):
            before, mark, after = m.groups()
            if mark == "." and (after.lower() in _NOT_A_SENTENCE_END
                                or after.lower() not in _OPENERS):
                continue
            if mark != "." or before.lower() not in _ABBREVIATIONS:
                after = after[0].upper() + after[1:]        # "vs.the" stays small
            found.append({"offset": m.start(2), "length": m.end() - m.start(2),
                          "replacements": [{"value": mark + " " + after}],
                          "message": "space after the end of a sentence", "kind": "punctuation"})
    if options.runons:
        for m in _SPLICE.finditer(text):
            sentence_start = max(text.rfind(c, 0, m.start()) for c in ".!?;") + 1
            if len(text[sentence_start:m.start()].split()) < 3:
                continue
            adverb = m.group(1)
            comma = "," if re.fullmatch(_ADVERBS_WITH_COMMA, adverb, re.IGNORECASE) else ""
            found.append({"offset": m.start(), "length": m.end() - m.start(),
                          "replacements": [{"value": "; " + adverb + comma + " "}],
                          "message": "two sentences joined by a comma", "kind": "runons"})
    return found


def runon_notes(index, text):
    """Run-ons that need a person: listed, never rewritten."""
    notes = []
    words = len(text.split())
    if words > LONG_PARAGRAPH:
        notes.append(Note(index, 0, 0, "runons",
                          f"Long paragraph ({words} words). Consider splitting it."))
    for m in _SENTENCE.finditer(text):
        sentence = m.group()
        count = len(sentence.split())
        if count > LONG_SENTENCE:
            notes.append(Note(index, m.start(), m.end(), "runons",
                              f"Long sentence ({count} words). It may be a run-on."))
            continue
        first = WORD.search(sentence)
        if not first or first.group().lower() in _SUBORDINATORS:
            continue
        splice = _COMMA_PRONOUN.search(sentence)
        if not splice:
            continue
        clause = sentence[:splice.start()]
        clause_words = [w.lower() for w in WORD.findall(clause)]
        if len(clause_words) < 4 or "," in clause or _SUBORDINATORS & set(clause_words):
            continue
        notes.append(Note(index, m.start(), m.end(), "runons",
                          "Possible run-on: two sentences joined by only a comma."))
    return notes


# ---------------------------------------------------------------- deciding
def _trim(text, start, end, new):
    """Shrink a change to the characters that really differ, so the rest of
    the words keep their look in a formatted document."""
    old = text[start:end]
    head = 0
    while head < len(old) and head < len(new) and old[head] == new[head]:
        head += 1
    tail = 0
    while (tail < len(old) - head and tail < len(new) - head
           and old[len(old) - 1 - tail] == new[len(new) - 1 - tail]):
        tail += 1
    return start + head, end - tail, new[head:len(new) - tail]


def _spelled_out(text):
    """"could've" and "could have" are one answer, not two."""
    return re.sub(r"['’]ve\b", " have", text).casefold()


def _swear_words(text):
    return sorted(m.group().lower() for m in SWEAR.finditer(text))


def _whole_words(text, start, end):
    """Stretch a span to the edges of the words it cuts through."""
    while start > 0 and (text[start - 1].isalnum() or text[start - 1] in "'’"):
        start -= 1
    while end < len(text) and (text[end].isalnum() or text[end] in "'’"):
        end += 1
    return start, end


class _Reviewer:
    def __init__(self, texts, options, speller):
        self.texts = texts
        self.options = options
        self.speller = speller
        self.names = find_names(texts)
        self.edits = []
        self.notes = []

    # -- one match
    def consider(self, index, match):
        text = self.texts[index]
        start = match["offset"]
        end = start + match["length"]
        old = text[start:end]
        kind = match.get("kind") or classify(match)
        if kind is None or "\n" in old:
            return
        if kind == "spelling" and " " in old.strip():
            kind = "grammar"                    # a phrase, not a misspelled word
        box = "spelling" if kind == "dictionary" else kind
        if not self.options.wants(box):
            return
        if _touches(self.protected[index], start, end):
            return
        label = not re.search(r"[.!?]", text) and len(text.split()) < 6
        if label and match.get("rule", {}).get("id") == "UPPERCASE_SENTENCE_START":
            return                  # a table cell, list item or label, not a sentence
        choices = [r["value"] for r in match.get("replacements", [])]
        if (label and kind == "dictionary" and old[:1].islower() and not text[:start].strip()
                and all(c[:1].isupper() for c in choices)):
            # suggestions were capitalized only because the word comes first
            choices = [c[:1].lower() + c[1:] for c in choices]
        message = match.get("message", "")

        if kind == "dictionary":
            self.spelling(index, start, end, old, choices)
        elif len(choices) == 1 or (choices and len({_spelled_out(c) for c in choices}) == 1):
            self.change(index, start, end, choices[0], box, message)
        elif choices:
            self.notes.append(Note(index, start, end, box, message, choices[:4]))
        elif box != "punctuation":
            self.notes.append(Note(index, start, end, box, message))

    def spelling(self, index, start, end, word, choices):
        if any(c.isdigit() for c in word) or (len(word) > 1 and word.isupper()):
            return                                          # codes and abbreviations
        text = self.texts[index]
        case_only = choices and _plain(choices[0]) == _plain(word)
        if word[:1].isupper() and not case_only:
            if word in self.names or not starts_sentence(text, start):
                self.notes.append(Note(index, start, end, "name",
                                       "Looks like a name; left as it is.", choices[:3]))
                return
        sure = sure_spelling(word, choices)
        if sure:
            self.change(index, start, end, sure, "spelling", "spelling")
            return
        words_opinion = self.speller.suggest(word) if self.speller else None
        if self.speller and words_opinion is None:
            self.notes.append(Note(index, start, end, "spelling",
                                   "The free checker questions this word; Word accepts it.",
                                   choices[:3]))
        elif words_opinion and self.agreed(word, words_opinion, choices):
            self.change(index, start, end, words_opinion[0], "spelling", "both checkers agree")
        else:
            both = list(dict.fromkeys((words_opinion or [])[:2] + choices[:3]))
            self.notes.append(Note(index, start, end, "spelling",
                                   "Not sure which word was meant.", both[:4]))

    @staticmethod
    def agreed(word, from_word, from_checker):
        """Word's first suggestion counts only when no other word both checkers
        offer is just as close ("wich" could be "which" or "witch")."""
        shared = [s for s in from_word[:6] if s in from_checker[:6]]
        if not shared or shared[0] != from_word[0]:
            return False
        best = distance(word, shared[0])
        return best <= 2 and all(distance(word, other) > best for other in shared[1:])

    def change(self, index, start, end, new, kind, why):
        text = self.texts[index]
        if "\n" in new or "\r" in new:
            return
        if "’" in text and "'" not in text:
            new = new.replace("'", "’")        # keep the document's curly apostrophes
        a, b = _whole_words(text, start, end)
        after = text[a:start] + new + text[end:b]
        if _swear_words(text[a:b]) != _swear_words(after):
            return                                          # would add, drop or alter a swear word
        start, end, new = _trim(text, start, end, new)
        if start == end and not new:
            return
        self.edits.append(Edit(index, start, end, new, kind, why))

    # -- Word's second look at words the free checker passed
    def second_opinion(self, flagged):
        if not (self.speller and self.options.spelling):
            return
        seen = set()
        for index, text in enumerate(self.texts):
            for m in WORD.finditer(text):
                word = m.group()
                key = (index, m.start())
                if key in flagged or word in seen or len(word) < 2:
                    continue
                if word.isupper() or _touches(self.protected[index], m.start(), m.end()):
                    continue
                if word[0].isupper() and (word in self.names
                                          or not starts_sentence(text, m.start())):
                    continue
                opinion = self.speller.suggest(word)
                if opinion is None:
                    continue
                seen.add(word)
                self.notes.append(Note(index, m.start(), m.end(), "spelling",
                                       "Word questions this word; the free checker accepts it.",
                                       opinion[:3]))

    def run(self, checker):
        self.protected = [protected_spans(t, self.options.keep_words) for t in self.texts]
        answers = checker.check_many(self.texts)
        flagged = set()
        for index, matches in enumerate(answers):
            ours = own_matches(self.texts[index], self.options)
            for match in ours + list(matches):
                if classify(match) == "dictionary":
                    flagged.add((index, match["offset"]))
                self.consider(index, match)
            if self.options.runons:
                self.notes.extend(runon_notes(index, self.texts[index]))
        self.second_opinion(flagged)
        return self.without_overlaps(), self.notes

    def without_overlaps(self):
        """Two fixes for the same words: keep the first, the next pass sees the rest."""
        kept, last = [], {}
        for edit in sorted(self.edits, key=lambda e: (e.para, e.start, e.end)):
            if edit.start < last.get(edit.para, -1):
                continue
            if edit.start == last.get(edit.para, -1) and edit.start == edit.end:
                continue
            kept.append(edit)
            last[edit.para] = max(edit.end, edit.start + 1)
        return kept


def review(texts, options, checker, speller=None):
    """One look over a document's paragraphs: (edits to make, notes to list)."""
    return _Reviewer(list(texts), options, speller).run(checker)


def apply_edits(text, edits):
    """A paragraph's text after its edits (used for plain text and for reports)."""
    for edit in sorted(edits, key=lambda e: e.start, reverse=True):
        text = text[:edit.start] + edit.new + text[edit.end:]
    return text


def snippet(text, start, end, new=None, around=28):
    """The words round a change, for the report."""
    a, b = max(0, start - around), min(len(text), end + around)
    middle = text[start:end] if new is None else new
    return (("..." if a else "") + text[a:start] + middle + text[end:b]
            + ("..." if b < len(text) else "")).replace("\n", " ").strip()
