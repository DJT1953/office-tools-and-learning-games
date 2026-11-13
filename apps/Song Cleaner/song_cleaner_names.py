"""Song Cleaner: getting song names right, by rule.

Radio recorders leave marks on names: the DJ's name stuck on the end of the
title ("Auto Dj Bud"), copy numbers ("_01", "(2)"), titles cut off at 30
letters, "artist-title" with no spaces, "Surname, First", a record label and
catalogue number where the title should be. This module proposes a clean
"Artist - Title" for each song and says which names it could not be sure of.
It only proposes; nothing here renames a file.
"""

import collections
import dataclasses
import os
import re

FIXED_JUNK = ["Auto Dj Bud", "Dj Bud Francis", "Enhanced Mp3", "Classic Rock Jukebox"]
CUT_OFF_AT = 30                      # Radiotracker cut artists and titles to 30 letters
MAX_NAME = 150

PLACEHOLDER_TITLES = {"single", "album version", "album", "live", "version", "track", "unknown",
                      "untitled", "radio edit", "rare", "nan"}
LABELS = ("columbia|liberty|london|atlantic|capitol|philips|casablanca|stateside|decca|parlophone|"
          "pye|rca|rca victor|mercury|emi|polydor|tamla|tamla motown|motown|reprise|warner|"
          "warner bros|mgm|brunswick|fontana|hmv|coral|chess|stax|top rank|cbs|epic|"
          "united artists|a&m|island|apple|bell|vogue|immediate|track|dot|imperial|sun|"
          "philles|cameo|cameo parkway|scepter|roulette|laurie|swan|vee jay|kapp|abc|buddah")
_CATALOGUE = re.compile(rf"^(?:{LABELS})(?: [a-z]{{1,4}})? ?\d{{3,6}}$", re.I)
_RESERVED = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)),
             *(f"lpt{i}" for i in range(1, 10))}
_NT = "don|can|won|didn|isn|ain|wasn|couldn|wouldn|shouldn|doesn|aren|haven|hasn|weren"
_S = "let|it|that|what|he|she|there|here|who|everybody|nobody|somebody|where"


def squeeze(text):
    return re.sub(r"\s+", " ", text or "").strip()


def split_name(stem):
    """("artist", "title") from a file name without its ending."""
    stem = squeeze(stem.replace("_", " ") if " " not in stem else stem)
    if " - " in stem:
        artist, _, title = stem.partition(" - ")
    elif "-" in stem and not stem.startswith("-") and not stem.endswith("-"):
        artist, _, title = stem.partition("-")
    else:
        return "", stem
    return squeeze(artist), squeeze(title)


def strip_copy_number(text):
    """Without "_01", "(2)" or " 01" at the end. "1941", "Cloud 9" and "TVC 15" stay."""
    text = squeeze(text)
    while True:
        shorter = re.sub(r"(?:_\d{1,3}|\s*\(\d{1,2}\)|\s+0\d)$", "", text).strip()
        if shorter == text or not shorter:
            return text
        text = shorter


def find_junk(pairs):
    """Phrases that DJs' software stuck on the end of titles.

    pairs: (artist, title) for every song. A phrase starting "Dj" or "Auto Dj"
    that ends the titles of three or more different artists is junk.
    """
    seen = collections.defaultdict(set)
    for artist, title in pairs:
        found = re.search(r"\b((?:Auto )?Dj(?: [A-Za-z]+){1,2})$", strip_copy_number(title), re.I)
        if not found:
            continue
        words = found.group(1).split()
        start = 1 if words[0].lower() == "auto" else 0
        for end in range(start + 2, len(words) + 1):
            seen[" ".join(words[start:end]).lower()].add(artist.lower())
    junk = list(FIXED_JUNK)
    for phrase, artists in seen.items():
        if len(artists) >= 3 and phrase not in (j.lower() for j in junk):
            junk.append(phrase.title())
    return sorted(junk, key=len, reverse=True)


def clean_title(title, junk=FIXED_JUNK):
    """The title without junk on the end, even when the junk itself was cut off."""
    title = strip_copy_number(title)
    long_title = len(title) >= CUT_OFF_AT - 2
    pieces = set()
    for phrase in junk:
        p = phrase.lower()
        pieces.add(p)
        # a cut-off ending: "Auto Dj Bu", "Auto D"; one bare word only on a long title
        for n in range(len(p) - 1, 1, -1):
            piece = p[:n].rstrip()
            if " " in piece or (long_title and piece == p.split()[0]):
                pieces.add(piece)
    while True:
        low = title.lower()
        fits = [p for p in pieces if low.endswith(" " + p) and len(title) > len(p) + 1]
        if not fits:
            return title
        title = strip_copy_number(title[:-len(max(fits, key=len))].rstrip(" -,"))


def _cap(word):
    for i, ch in enumerate(word):
        if ch.isalpha():
            return word[:i] + ch.upper() + word[i + 1:].lower()
    return word


def tidy(text):
    """Even spacing, straight apostrophes, and capitals when there were none (or nothing else)."""
    text = squeeze(text.replace("`", "'").replace("’", "'"))
    text = re.sub(rf"\b({_NT}) t\b", r"\1't", text, flags=re.I)
    text = re.sub(rf"\b({_S}) s\b", r"\1's", text, flags=re.I)
    letters = [c for c in text if c.isalpha()]
    if letters and (all(c.islower() for c in letters)
                    or (all(c.isupper() for c in letters) and max(map(len, text.split())) > 4)):
        text = " ".join(_cap(w) for w in text.split(" "))
    return text


def tidy_artist(artist):
    """Also turns "Anka, Paul" into "Paul Anka" and "Box Tops, The" into "Box Tops"."""
    artist = tidy(artist)
    if artist.count(",") == 1:
        left, right = (part.strip() for part in artist.split(","))
        if right.lower() == "the":
            return left
        if left and right and " " not in left and " " not in right \
                and right.lower().rstrip(".") not in ("jr", "sr", "ii", "iii"):
            return f"{right} {left}"
    return artist


def letters_only(text):
    text = (text or "").lower().replace("&", " and ")
    text = re.sub(r"['`’]", "", text)
    words = [w for w in re.split(r"[^a-z0-9]+", text) if w and w not in ("the", "a")]
    return "".join(words)


def loose_key(artist, title):
    """The same for names that differ only in capitals, spacing, punctuation, "the" and "a"."""
    return letters_only(artist) + "|" + letters_only(title)


def edit_distance(a, b, limit=2):
    """How many single-letter changes turn a into b; limit + 1 once it is past the limit."""
    if abs(len(a) - len(b)) > limit:
        return limit + 1
    row = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        new = [i]
        for j, cb in enumerate(b, 1):
            new.append(min(row[j] + 1, new[j - 1] + 1, row[j - 1] + (ca != cb)))
        if min(new) > limit:
            return limit + 1
        row = new
    return row[-1]


def spelling_fixes(artists):
    """{spelling: better spelling} for a Counter of artist names.

    Spellings that differ only in capitals, spacing or punctuation take the
    most used form. A rare spelling one or two letters away from one used at
    least five times as often (both six letters or more) takes that one.
    """
    groups = collections.defaultdict(collections.Counter)
    for name, count in artists.items():
        if name:
            groups[letters_only(name)][name] += count

    def best(spellings):
        # the wording used most; then, among its ways of using capitals, not
        # SHOUTED or all small, and the one with most capitals ("ZZ Top", "McCartney")
        wordings = collections.Counter()
        for name, count in spellings.items():
            wordings[name.lower()] += count
        wording = max(wordings.items(), key=lambda kv: (kv[1], kv[0]))[0]

        def score(item):
            name, count = item
            shouted = name.isupper() and max(map(len, name.split())) >= 4
            return (not shouted and not name.islower(), sum(c.isupper() for c in name), count, name)
        return max(((n, c) for n, c in spellings.items() if n.lower() == wording), key=score)[0]

    totals = {key: sum(c.values()) for key, c in groups.items()}
    target = {key: key for key in groups}
    common = [k for k, n in totals.items() if n >= 5 and len(k) >= 6]
    for key, n in totals.items():
        if len(key) < 6:
            continue
        limit = 1 if len(key) < 8 else 2
        choice = None
        for other in common:
            if other != key and totals[other] >= 5 * n and edit_distance(key, other, limit) <= limit:
                if choice is None or totals[other] > totals[choice]:
                    choice = other
        if choice:
            target[key] = choice
    fixes = {}
    for key, spellings in groups.items():
        good = best(groups[target[key]])
        for name in spellings:
            if name != good:
                fixes[name] = good
    return fixes


def safe_file_name(artist, title, ending=".mp3"):
    """A file name Windows accepts for this artist and title."""
    name = f"{artist} - {title}" if artist and title else (title or artist or "Unnamed song")
    name = name.replace("/", " ").replace("\\", " ").replace(":", " -").replace("?", "")
    name = name.replace('"', "'").replace("*", "").replace("|", " ")
    name = name.replace("<", "(").replace(">", ")")
    name = squeeze("".join(c for c in name if ord(c) >= 32))
    name = name[:MAX_NAME].rstrip(" .")
    if not name:
        name = "Unnamed song"
    if name.split(".")[0].strip().lower() in _RESERVED:
        name = "Song " + name
    return name + ending


def is_placeholder(title):
    low = squeeze(title).lower()
    return low in PLACEHOLDER_TITLES or bool(_CATALOGUE.match(low))


@dataclasses.dataclass
class NameChange:
    path: str
    old_name: str
    artist: str
    title: str
    new_name: str
    doubtful: bool = False
    reason: str = ""

    @property
    def changes(self):
        return self.new_name != self.old_name


def _longer_from_tag(from_name, from_tag):
    """The tag's wording when it is the file name's wording carried on further."""
    a, b = letters_only(from_name), letters_only(from_tag)
    return from_tag if a and len(b) > len(a) and b.startswith(a) else from_name


def propose_names(songs):
    """A NameChange for every song whose name would change or cannot be trusted."""
    raw = []
    for song in songs:
        artist, title = split_name(song.stem)
        if not artist and song.artist and song.title:
            artist, title = song.artist, song.title
        raw.append((artist, title))
    junk = find_junk(raw)

    drafts = []
    for song, (artist, title) in zip(songs, raw):
        cut_off = len(title) == CUT_OFF_AT or len(artist) == CUT_OFF_AT
        title = clean_title(title, junk)
        tag_title = clean_title(song.title or "", junk)
        fuller = _longer_from_tag(title, tag_title)
        if fuller != title:
            title, cut_off = fuller, len(artist) == CUT_OFF_AT
        elif len(title) != CUT_OFF_AT and len(artist) != CUT_OFF_AT:
            cut_off = False                       # the long part was junk, now gone
        artist = _longer_from_tag(artist, squeeze(song.artist or ""))
        drafts.append([song, tidy_artist(artist), tidy(title), cut_off])

    # "Bohemian Rhapsody - Queen": the two halves the wrong way round
    counts = collections.Counter(letters_only(d[1]) for d in drafts if d[1])
    for draft in drafts:
        _, artist, title, _ = draft
        other = re.sub(r"\s+(19|20)\d\d$", "", title)
        if artist and counts[letters_only(other)] >= 3 and counts[letters_only(artist)] == 1:
            draft[1], draft[2] = tidy_artist(other), artist

    fixes = spelling_fixes(collections.Counter(d[1] for d in drafts if d[1]))
    changes = []
    for song, artist, title, cut_off in drafts:
        artist = fixes.get(artist, artist)
        reason = ""
        if not title or not artist:
            reason = "No artist or no title"
        elif is_placeholder(title):
            reason = "The title looks like a label or a placeholder, not a song name"
        elif cut_off:
            reason = "The name may be cut off"
        ending = os.path.splitext(song.name)[1]
        new_name = safe_file_name(artist, title, ending)
        if new_name[:-len(ending)] != song.stem:
            new_name = new_name[:-len(ending)] + ".mp3"
        else:
            new_name = song.name
        change = NameChange(song.path, song.name, artist, title, new_name, bool(reason), reason)
        if change.changes or change.doubtful:
            changes.append(change)
    return changes
