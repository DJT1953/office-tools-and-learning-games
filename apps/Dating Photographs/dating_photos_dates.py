"""Dating Photographs: finds dates in words read from old photographs.

"Oct-98", "98-Oct", "October 1998", "June 12, 1975", "3/15/91", "Summer '82",
"Christmas 1987", a camera's "'98 10 12" and many more each become a date.
A date with no day gets the 1st ("Oct-98" is 1998-10-01); a year alone gets
January 1st. Letters and numbers a reader often mixes up (0 and O, 1 and l,
"Ang" for "Aug") are put right before the words are looked at.

Nothing here reads pictures; dating_photos_read does that and hands the
words over.
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass, replace

DAY, HOLIDAY, MONTH, SEASON, YEAR, DECADE = "day", "holiday", "month", "season", "year", "decade"
EXACTNESS = {DAY: 5, HOLIDAY: 4, MONTH: 3, SEASON: 2, YEAR: 1, DECADE: 0}
FIRST_PHOTO_YEAR = 1839          # nothing older can be a photograph

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6, "july": 7,
    "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7, "aug": 8, "sep": 9,
    "sept": 9, "oct": 10, "nov": 11, "dec": 12, "febr": 2,
}
MONTH_ABBR = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
SEASONS = {"spring": 4, "summer": 7, "fall": 10, "autumn": 10, "winter": 1}

# Letter pairs a reader confuses in handwriting: in a three-letter month, one may stand for the other.
LOOK_ALIKE = {frozenset(pair) for pair in ("un", "uv", "ao", "ec", "lt", "tf", "il", "oc", "mn",
                                           "gq", "yv", "hb", "ae", "ce", "lj", "ij", "rn", "hn")}
# Inside a number: letters read in place of digits ("9B" is 98, "l998" is 1998).
AS_DIGIT = str.maketrans({"o": "0", "O": "0", "D": "0", "Q": "0", "l": "1", "I": "1", "i": "1",
                          "|": "1", "!": "1", "S": "5", "$": "5", "B": "8", "Z": "2",
                          "z": "2", "g": "9", "q": "9", "b": "6", "G": "6", "T": "7"})
# Inside a word: digits and marks read in place of letters ("0ct" is Oct, "Oc+" is Oct).
AS_LETTER = str.maketrans({"0": "o", "1": "l", "|": "l", "!": "l", "5": "s", "$": "s", "+": "t",
                           "8": "b", "6": "b", "9": "g", "2": "z"})

# How much each kind of date is trusted, before the reader's own confidence.
WEIGHT = {DAY: 1.0, HOLIDAY: 0.95, MONTH: 0.85, SEASON: 0.7, YEAR: 0.5, DECADE: 0.3}
# Where the words were found: writing on the photo beats a file name.
PLACE_WEIGHT = {"back": 1.0, "front": 1.0, "typed": 1.0, "file name": 0.9, "folder name": 0.6}
# A birth or death date written under a photo ("b: 11-02-1886", "(04/22/1833 to 09/14/1899)") is
# rarely the day it was taken: trusted this much less, so "~1895" on the line above it wins.
LIFE_WEIGHT = 0.4
LIFE_MARK = re.compile(r"(?<![a-z])(?:b|d|born|died|dob|dod|birth|death|deceased|buried|bur)\s*[.:]?\s*\(?\s*$",
                       re.I)
# "born" or "married" a few words before: "married to Frank B. Hale May 18, 1948" (not past
# a comma, a number or the end of a sentence: "where Dad was born, June 1952" and "born in Ohio
# 1902 Christmas 1952" end with the photo's)
LIFE_WORD = re.compile(r"(?<![a-z])(?:born|died|married|marr|deceased|buried)(?:(?![,;!?\d]|[a-z]{2}\.).){0,45}$",
                       re.I)
RANGE_GAP = re.compile(r"\s*\)?\s*(?:to|-|until|till|thru|through)\s*\(?\s*", re.I)
ORDINAL = r"(?:st|nd|rd|th)"


@dataclass(frozen=True)
class Clue:
    """One date found in some words."""
    date: dt.date
    kind: str                 # DAY, HOLIDAY, MONTH, SEASON, YEAR or DECADE
    words: str                # the words it came from, as read
    score: float              # higher is more believable, at most 1
    where: str = ""           # "front", "back", "file name", "folder name" or "typed"
    box: tuple | None = None  # (left, top, right, bottom) in the picture, if read from one
    life: bool = False        # looks like a birth, death or wedding date ("b: 1886", "(1833 to 1899)")

    def describe(self):
        """Plain words for the Check list, e.g. '"Oct-98" (no day written, so the 1st)'."""
        note = {MONTH: " (no day written, so the 1st)", SEASON: " (a season)",
                YEAR: " (only a year, so January 1st)", DECADE: " (only a decade)"}.get(self.kind, "")
        if self.life:
            note += " (looks like a birth, death or wedding date)"
        return f"“{self.words.strip()}”{note}"


# ------------------------------------------------------------------ dates of holidays

def easter(year):
    """Easter Sunday (Western), by the anonymous Gregorian method."""
    a, b, c = year % 19, year // 100, year % 100
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month, day = divmod(h + l - 7 * m + 114, 31)
    return dt.date(year, month, day + 1)


def nth_weekday(year, month, weekday, n):
    """The nth weekday (Monday = 0) of a month; n = -1 is the last one."""
    if n > 0:
        first = dt.date(year, month, 1)
        return first + dt.timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))
    last = dt.date(year + month // 12, month % 12 + 1, 1) - dt.timedelta(days=1)
    return last - dt.timedelta(days=(last.weekday() - weekday) % 7)


# Written against cleaned, lowercased words, where a month name is a <mnn> mark.
JULY = r"<m07~?>"
HOLIDAYS = [   # pattern, the date it means in a given year
    (r"christmas\s*eve|xmas\s*eve|x-mas\s*eve", lambda y: dt.date(y, 12, 24)),
    (r"christmas(?:\s*day)?|xmas|x-mas|x'mas", lambda y: dt.date(y, 12, 25)),
    (r"new\s*year'?s?\s*eve", lambda y: dt.date(y, 12, 31)),
    (r"new\s*year'?s?(?:\s*day)?", lambda y: dt.date(y, 1, 1)),
    (r"easter(?:\s*sunday)?", easter),
    (r"thanksgiving(?:\s*day)?", lambda y: nth_weekday(y, 11, 3, 4)),
    (r"hallowe'?en", lambda y: dt.date(y, 10, 31)),
    (r"(?:4(?:th)?|fourth)\s*of\s*" + JULY + r"|independence\s*day", lambda y: dt.date(y, 7, 4)),
    (r"valentine'?s?(?:\s*day)?", lambda y: dt.date(y, 2, 14)),
    (r"mother'?s?\s*day", lambda y: nth_weekday(y, 5, 6, 2)),
    (r"father'?s?\s*day", lambda y: nth_weekday(y, 6, 6, 3)),
    (r"memorial\s*day", lambda y: nth_weekday(y, 5, 0, -1) if y >= 1971 else dt.date(y, 5, 30)),
    (r"labou?r\s*day", lambda y: nth_weekday(y, 9, 0, 1)),
    (r"st\.?\s*patrick'?s?(?:\s*day)?", lambda y: dt.date(y, 3, 17)),
]


# ------------------------------------------------------------------ cleaning up the words

def edit_distance(a, b):
    previous = list(range(len(b) + 1))
    for i, x in enumerate(a, 1):
        current = [i]
        for j, y in enumerate(b, 1):
            current.append(min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (x != y)))
        previous = current
    return previous[-1]


def month_of(word):
    """The month a word names, allowing for a reader's slips: (month, exact) or None."""
    word = word.lower().rstrip(".")
    if word in MONTHS:
        return MONTHS[word], True
    def swapped(names):         # one letter swapped for a look-alike: "Ang" for "Aug"
        return {MONTHS[name] for name in names if len(name) == len(word)
                and sum(x != y for x, y in zip(word, name)) == 1
                and all(x == y or frozenset((x, y)) in LOOK_ALIKE for x, y in zip(word, name))}
    if len(word) == 3:
        found = swapped(MONTH_ABBR)
        return (found.pop(), False) if len(found) == 1 else None
    if len(word) >= 5:          # "Febuary", "Septmber", "Agust"
        allowed = 1 if len(word) < 8 else 2
        # five letters near March or April may be a name ("Marcy"), so only a
        # look-alike slip counts for those ("Marcn", "Aprll")
        smallest = 6 if len(word) == 5 else 5
        found = {number for name, number in MONTHS.items()
                 if len(name) >= smallest and edit_distance(word, name) <= allowed}
        if len(word) == 5:
            found |= swapped(("march", "april"))
        return (found.pop(), False) if len(found) == 1 else None
    return None


def clean_token(token):
    """One word or number with a reader's look-alikes put right."""
    ending = re.fullmatch(r"(\d{1,2})(st|nd|rd|th)", token, re.I)
    if ending:            # kept: a day can be 12th, a year can't be ("39th Birthday" isn't '39)
        return ending.group(1) + ending.group(2).lower()
    digits = sum(ch.isdigit() for ch in token)
    letters = sum(ch.isalpha() for ch in token)
    if digits and digits >= letters:
        if len(token) <= 4:
            number = token.translate(AS_DIGIT)
            if number.isdigit():
                return number
        return token.lower()            # 1950s, 35mm: leave alone
    word = token.translate(AS_LETTER).lower()
    month = month_of(word)
    if month:
        return f"<M{month[0]:02}{'' if month[1] else '~'}>"
    return word


def clean(text):
    """Text with look-alike characters put right and each month name replaced by
    a <Mnn> mark (<Mnn~> when guessed from a misspelling), ready for the patterns."""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[‘’‛′`´]", "'", text)
    text = re.sub(r"[‐-―−]", "-", text)
    out = []
    for token in re.split(r"(\s+|[-/.,:;'()\[\]&])", text):
        if not token or not token.strip() or re.fullmatch(r"[-/.,:;'()\[\]&]", token):
            out.append(token)
            continue
        # "May1969", "Oct98", "1998Oct": a month stuck to a number
        stuck = re.fullmatch(r"([A-Za-z+|]{3,})(\d{2,4})|(\d{2,4})([A-Za-z+|]{3,})", token)
        if stuck and month_of((stuck.group(1) or stuck.group(4)).translate(AS_LETTER)):
            parts = [stuck.group(1), stuck.group(2)] if stuck.group(1) else [stuck.group(3), stuck.group(4)]
            out.append(" ".join(clean_token(p) for p in parts))
            continue
        out.append(clean_token(token))
    return "".join(out)


# ------------------------------------------------------------------ the patterns

Y4 = r"(18[3-9]\d|19\d\d|20\d\d)"
NOT_ORDINAL = r"(?!\d|st|nd|rd|th)"   # a year is never "39th"
YEAR_ANY = r"('?\d{2}|18[3-9]\d|19\d\d|20\d\d)" + NOT_ORDINAL
SURE_YEAR = r"('\d{2}|18[3-9]\d|19\d\d|20\d\d|3[2-9]|[4-9]\d)" + NOT_ORDINAL   # can't be a day of the month
DAYLIKE_YEAR = r"([0-2]\d|3[01])" + NOT_ORDINAL                             # 12 could be a day or 2012
NOTHING_MORE = r"(?![\s,.\-/']*\d)"
M = r"<M(\d\d)(~?)>"
D = r"(\d{1,2})" + ORDINAL + r"?(?!\d)"   # a day: 12 or 12th
B = r"(?<![\dM])"     # not in the middle of a number or a month mark
E = r"(?!\d)"


def full_year(text, today):
    """'98 -> 1998, 1998 -> 1998, '05 -> 2005; two digits after this year's are 1900s."""
    digits = str(text).strip("'")
    year = int(digits)
    if len(digits) == 2:
        year += 2000 if year <= today.year % 100 else 1900
    return year


OUT_OF_RANGE = "out of range"   # a real date, but before photography or after today


def make_date(year, month, day, today):
    """The date, None if there's no such day (Feb 30), or OUT_OF_RANGE."""
    try:
        date = dt.date(int(year), int(month), int(day))
    except (ValueError, TypeError):
        return None
    return date if FIRST_PHOTO_YEAR <= date.year and date <= today else OUT_OF_RANGE


def dated(year_text, month, day, today):
    """make_date for a year as written: a two-digit year that would be later than
    today is the 1900s one ("Thanksgiving '26" written in November 2026 is 1926)."""
    year = full_year(year_text, today)
    date = make_date(year, month, day, today)
    if date == OUT_OF_RANGE and len(str(year_text).strip("'")) == 2 and year >= 2000:
        date = make_date(year - 100, month, day, today)
    return date


def numeric_order(a, b, month_first):
    """For 10/12/98 or a camera's 10 12 '98: (month, day). In the US the month
    comes first; anywhere else the day does. A number over 12 settles it."""
    a, b = int(a), int(b)
    if month_first:
        return (a, b) if a <= 12 else (b, a)
    return (b, a) if b <= 12 else (a, b)


def squashed_stamp(digits):
    """A camera's '7 4 '86' read without its spaces as 74'86: (month, day) or None."""
    splits = {1: [], 2: [(digits[0], digits[1])], 3: [(digits[:1], digits[1:]), (digits[:2], digits[2:])],
              4: [(digits[:2], digits[2:])]}[len(digits)]
    for month, day in splits:
        if 1 <= int(month) <= 12 and 1 <= int(day) <= 31:
            return month, day
    return None


def date_patterns(month_first):
    """(pattern, kind, how much it's trusted, groups -> (year, month, day) or None), most exact first."""
    def numeric(g):
        month, day = numeric_order(g[0], g[1], month_first)
        return g[2], month, day

    def squashed(g):
        found = squashed_stamp(g[0])
        return (g[1], *found) if found else None

    return [
        # 1998-10-12, 1998/10/12, 1998.10.12
        (B + Y4 + r"[-/.]" + D + r"[-/.]" + D, DAY, 1.0, lambda g: (g[0], g[1], g[2])),
        # 19981012, and 1998 10 12 (a file name's underscores become spaces)
        (B + Y4 + r"(\d\d)(\d\d)" + E, DAY, 0.9, lambda g: (g[0], g[1], g[2])),
        (B + Y4 + r"\s+" + D + r"\s+" + D, DAY, 0.85, lambda g: (g[0], g[1], g[2])),
        # 1881 October 14, 1998 Oct. 12th
        (B + Y4 + r"[\s,.\-/]*" + M + r"[\s,.\-/]*" + D, DAY, 0.95, lambda g: (g[0], g[1], g[3])),
        # June 12, 1975 / Dec. 25,1990 / June 12th 75
        (M + r"[\s,.\-/]*" + D + r"[\s,.\-/']*(?:of\s*)?" + YEAR_ANY, DAY, 1.0,
         lambda g: (g[3], g[0], g[2])),
        # 12 June 1975 / 12th of June, 1975
        (B + D + r"[\s,.\-/]*(?:of\s*)?" + M + r"[\s,.\-/']*(?:of\s*)?" + YEAR_ANY, DAY, 1.0,
         lambda g: (g[3], g[1], g[0])),
        # 10/12/98, 10-12-1998
        (B + D + r"[-/.]" + D + r"[-/.]" + r"(\d{2}|18[3-9]\d|19\d\d|20\d\d)" + E, DAY, 0.95, numeric),
        # a camera's imprint: '98 10 12, and 10 12 '98
        (r"'(\d\d)\s+" + D + r"\s+" + D, DAY, 0.95, lambda g: (g[0], g[1], g[2])),
        (B + D + r"\s+" + D + r"\s+'(\d\d)" + E, DAY, 0.95, numeric),
        (B + r"(\d{2,4})'(\d\d)" + E, DAY, 0.75, squashed),
        (r"'(\d\d)(\d{2,4})" + E, DAY, 0.75, lambda g: squashed([g[1], g[0]])),
        # October 1998, October of 1998 ("June 1975, 12 people" too)
        (M + r"[\s,.\-/']*(?:of\s*)?" + Y4 + E, MONTH, 1.0, lambda g: (g[2], g[0], 1)),
        # Oct-98, Oct '98 (not "Feb 30 1990": a number follows)
        (M + r"[\s,.\-/']*(?:of\s*)?" + SURE_YEAR + NOTHING_MORE, MONTH, 1.0, lambda g: (g[2], g[0], 1)),
        # "Oct 12" or "Oct 31" may be a day with no year, so a year like 12 is trusted
        # too little to be ticked: it is marked "check this"
        (M + r"[\s,.\-/]*" + DAYLIKE_YEAR + NOTHING_MORE, MONTH, 0.5, lambda g: (g[2], g[0], 1)),
        # 98-Oct, 1998 October; again "12 Oct" may be a day
        (B + SURE_YEAR + r"[\s,.\-/]*" + M, MONTH, 1.0, lambda g: (g[0], g[1], 1)),
        (B + DAYLIKE_YEAR + r"[\s,.\-/]*" + M, MONTH, 0.5, lambda g: (g[0], g[1], 1)),
        # 1998-10, and "1959 07" or "1979 05 - 2 yrs old" (a year, then the month's number)
        (B + Y4 + r"[-/.]" + D, MONTH, 0.8, lambda g: (g[0], g[1], 1)),
        (B + Y4 + r"\s+(0[1-9]|1[0-2])(?![\d']|\s*\d)", MONTH, 0.85, lambda g: (g[0], g[1], 1)),
        # 10/98, 10-1998; "3/15" or "12-25" could be a day with no year, so trusted less
        (B + D + r"[-/.]" + r"('\d{2}|3[2-9]|[4-9]\d|18[3-9]\d|19\d\d|20\d\d)" + E, MONTH, 0.8,
         lambda g: (g[1], g[0], 1)),
        (B + D + r"[-/.]" + DAYLIKE_YEAR, MONTH, 0.5, lambda g: (g[1], g[0], 1)),
    ]


def find_dates(text, where="", confidence=1.0, box=None, today=None, month_first=True):
    """Every date in some words, best first. confidence is the reader's own (0 to 1)."""
    today = today or dt.date.today()
    cleaned = clean(text)
    lowered = cleaned.lower()
    found = []
    taken = []           # character spans already used by a more exact pattern
    trust_reader = max(0.05, min(1.0, confidence)) * PLACE_WEIGHT.get(where, 1.0)

    def add(match, date, kind, trust, fuzzy=False):
        start, end = match.span()
        if date is None or any(start < e and s < end for s, e in taken):
            return       # no such day: a vaguer pattern may still use these words (1998-99)
        taken.append((start, end))
        if date == OUT_OF_RANGE:
            return       # a future date: its words mean nothing else either
        words = text if len(text) <= 60 else cleaned[start:end]
        score = WEIGHT[kind] * trust * trust_reader * (0.85 if fuzzy else 1.0)
        found.append(((start, end), Clue(date, kind, words, round(score, 4), where, box)))

    def run(kinds):
        for pattern, kind, trust, parts in date_patterns(month_first):
            if kind not in kinds:
                continue
            for m in re.finditer(pattern, cleaned):
                pieces = parts(m.groups())
                if pieces:
                    year, month, day = pieces
                    add(m, dated(year, month, day, today), kind, trust, fuzzy="~" in m.groups())

    run({DAY})
    # holidays before month and year, so "Fourth of July 1986" isn't just July
    for pattern, when in HOLIDAYS:
        for m in re.finditer(r"(?<![a-z])(?:" + pattern + r")[\s,.\-/']*(?:of\s*)?" + YEAR_ANY, lowered):
            add(m, holiday(when, m.group(m.lastindex), today), HOLIDAY, 1.0)
        for m in re.finditer(B + Y4 + r"[\s,.\-/']*(?:" + pattern + r")", lowered):
            add(m, holiday(when, m.group(1), today), HOLIDAY, 1.0)
    run({MONTH})

    for m in re.finditer(r"(?<![a-z])(spring|summer|fall|autumn|winter)[\s,.\-/']*(?:of\s*)?" + YEAR_ANY,
                         lowered):
        add(m, dated(m.group(2), SEASONS[m.group(1)], 1, today), SEASON, 1.0)
    for m in re.finditer(B + Y4 + r"[\s,.\-/']*(spring|summer|fall|autumn|winter)(?![a-z])", lowered):
        add(m, make_date(m.group(1), SEASONS[m.group(2)], 1, today), SEASON, 1.0)      # 1956 summer

    # 1950s, the '50s, 50's (on a print, "the 20s" is the 1920s; "the 00s" the 2000s)
    for m in re.finditer(B + r"(18[4-9]|19\d|20[0-2])0'?s(?![a-z])", lowered):
        add(m, make_date(int(m.group(1)) * 10, 1, 1, today), DECADE, 1.0)
    for m in re.finditer(r"(?<![\d'])'?(\d)0'?s(?![a-z])", lowered):
        decade = 2000 if m.group(1) == "0" else 1900 + 10 * int(m.group(1))
        add(m, make_date(decade, 1, 1, today), DECADE, 0.8)

    # a year on its own: 1998, '98, circa 1950, 1998-99
    for m in re.finditer(B + Y4 + E, cleaned):
        add(m, make_date(m.group(1), 1, 1, today), YEAR, 1.0)
    for m in re.finditer(r"(?<![\w'])'(\d\d)" + E, cleaned):
        add(m, dated(m.group(1), 1, 1, today), YEAR, 0.9)

    return merge(mark_life_dates(cleaned, found, taken))


def mark_life_dates(cleaned, found, taken):
    """The clues, with birth and death dates marked and trusted less: a date after
    b:, d:, born or died, at either end of "1833 to 1899", or in a bracket the
    line cut off ("(04/22/1833 to" ... "09/14/1899)": the end of "(... to ...)").
    A date alone in brackets, "Grandma's house (1952)", is the photo's. taken holds
    every date's place, those too early to be a photo's too (1833 in "1833 to 1899")."""
    spans = sorted(taken)
    ranged = set()
    for (_, end), (start, _) in zip(spans, spans[1:]):
        if RANGE_GAP.fullmatch(cleaned[end:start]):
            ranged.update({end, start})
    out = []
    for (start, end), clue in found:
        before, after = cleaned[:start], cleaned[end:]
        opened = before.count("(") > before.count(")")
        first_after = re.search(r"[()]", after)
        closes = bool(first_after and first_after.group() == ")")
        cut_bracket = (opened and not closes) or (closes and not opened)
        if end in ranged or start in ranged or cut_bracket or LIFE_MARK.search(before[-20:]) \
                or LIFE_WORD.search(before[-60:]):
            clue = replace(clue, life=True, score=round(clue.score * LIFE_WEIGHT, 4))
        out.append(clue)
    return out


def holiday(when, year_text, today):
    """The holiday's date in the year written, or OUT_OF_RANGE; like dated(), a
    two-digit year that would be later than today is the 1900s one."""
    def on(year):
        if not FIRST_PHOTO_YEAR <= year <= today.year:
            return OUT_OF_RANGE
        date = when(year)
        return date if date <= today else OUT_OF_RANGE
    year = full_year(year_text, today)
    date = on(year)
    if date == OUT_OF_RANGE and len(str(year_text).strip("'")) == 2 and year >= 2000:
        date = on(year - 100)
    return date


def merge(clues):
    """One clue per date, the best-scored kept, best first (then the more exact). A date
    seen anywhere as a birth or death date stays one, even where read without its b: or d:;
    so does just its year or month read on its own ("21,1985" of "Born July 21,1985")."""
    life = {clue.date for clue in clues if clue.life}
    life_years = {date.year for date in life}
    life_months = {(date.year, date.month) for date in life}

    def part_of_life_date(clue):
        return clue.date in life or (clue.kind == YEAR and clue.date.year in life_years) \
            or (clue.kind == MONTH and (clue.date.year, clue.date.month) in life_months)

    by_date = {}
    for clue in clues:
        if not clue.life and part_of_life_date(clue):
            clue = replace(clue, life=True, score=round(clue.score * LIFE_WEIGHT, 4))
        kept = by_date.get(clue.date)
        if kept is None or clue.score > kept.score:
            by_date[clue.date] = clue
    return sorted(by_date.values(), key=lambda c: (-c.score, -EXACTNESS[c.kind]))


def best(clues):
    """The most believable date: the highest score, then the most exact, then the first found."""
    ranked = sorted(enumerate(clues), key=lambda ic: (-ic[1].score, -EXACTNESS[ic[1].kind], ic[0]))
    return ranked[0][1] if ranked else None


def typed_clue(text, today=None, month_first=True):
    """What someone typed in the date box (1998-10-01, Oct 98, 10/1998, 1998 or 98)
    as a Clue, which says how exact it is; None if it isn't a date."""
    today = today or dt.date.today()
    text = text.strip()
    if re.fullmatch(r"'?\d\d", text):
        date = dated(text, 1, 1, today)
        return Clue(date, YEAR, text, 1.0, "typed") if isinstance(date, dt.date) else None
    return best(find_dates(text, where="typed", today=today, month_first=month_first))


def typed_date(text, today=None, month_first=True):
    """typed_clue's date, or None."""
    clue = typed_clue(text, today, month_first)
    return clue.date if clue else None


COUNTER = re.compile(r"(?i)\b(?:img|dsc|dscn|dscf|scan|pict?|photo|image|p|file)\s*0*(\d{4})\b")
COUNTER_AND_NUMBER = re.compile(COUNTER.pattern + r"\s+(\d\d)\b")      # IMG_1998_05: picture 1998, copy 05


def from_name(name, where="file name", today=None, month_first=True):
    """Dates in a file or folder name ("IMG_20190305_1234", "Christmas 1987")."""
    words = re.sub(r"_+", " ", name)
    words = re.sub(r"(?<=\d)(?=[A-Za-z]{3})|(?<=[A-Za-z]{3})(?=\d)", " ", words)
    counters = {match.group(1) for match in COUNTER.finditer(words)}
    numbered = {(match.group(1), match.group(2)) for match in COUNTER_AND_NUMBER.finditer(words)}
    clues = find_dates(words, where=where, today=today, month_first=month_first)
    # in names, IMG_1998 is a picture's number, not a year, and IMG_1998_05 isn't May 1998
    return [c for c in clues if not (c.kind == YEAR and str(c.date.year) in counters)
            and not (c.kind == MONTH and (str(c.date.year), f"{c.date.month:02}") in numbered)]
