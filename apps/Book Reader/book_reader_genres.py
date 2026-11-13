"""Plain genres (Mystery, Westerns, History ...) from a book's subject tags.

Books carry library-catalogue headings ("Frontier and pioneer life -- West (U.S.) -- Fiction")
or shop categories ("FICTION / Mystery & Detective / General"). There are hundreds of different
ones, so each is put into a short list of plain genres by the words it contains.
"""

from __future__ import annotations

import re
from functools import lru_cache

NO_GENRE = "No genre"            # the book has no subject tags
OTHER = "Other"                  # it has tags, but none says what kind of book it is
GENERAL_FICTION = "General fiction"
HISTORICAL_FICTION = "Historical fiction"

# A heading with one of these words is about made-up stories. Calibre cuts headings at their
# commas ("... History -- Tudors" and "1485-1603 -- Fiction"), so this is decided for the whole book.
FICTION = re.compile(r"\b(?:fiction|stories|novels?|tales)\b", re.I)
HISTORY = re.compile(r"\b(?:history|historical)\b", re.I)

ANY, NONFICTION = "any", "nonfiction"
# Genre, the words that put a heading in it (whole words; \w* lets a word run on), and whether
# it is only for books that aren't stories. A heading can fall in several; this order is the
# order a book's genres are listed in, so its first one is the one it is sorted by.
RULES = [
    ("Mystery and detective", r"detectives?|myster(?:y|ies)|crime|murder|sherlock|private investigators", ANY),
    ("Thrillers and suspense", r"thrillers?|suspense|espionage|spy stories|spies", ANY),
    ("Science fiction", r"science fiction|time travel|interplanetary|outer space|life on other planets"
                        r"|robots|dystopias?|space opera", ANY),
    ("Fantasy", r"fantasy|fantastic fiction|magic|dragons|wizards", ANY),
    ("Horror and ghost stories", r"horror|ghosts?|supernatural|vampires?|werewolves|haunted", ANY),
    ("Westerns", r"western stories|westerns|frontier and pioneer life|cowboys|west \(u\.s\.\)", ANY),
    ("Adventure", r"adventure\w*|sea stories|pirates|shipwrecks|survival", ANY),
    ("Romance", r"love stories|romance|romantic|man-woman relationships|courtship", ANY),
    ("War", r"wars?|military|soldiers|battles|prisoners? of war|prisoner-of-war", ANY),
    ("Erotica", r"erotic\w*|pornograph\w*", ANY),
    ("Humor", r"humou?r\w*|satire|wit and|comic\w*|parodies", ANY),
    ("Short stories", r"short stories", ANY),
    ("Poetry", r"poetry|poems|verse|ballads|sonnets", ANY),
    ("Plays", r"drama|plays|tragedies|comedies|theater|theatre", ANY),
    ("Children's books", r"juvenile|children's|fairy tales|nursery rhymes|picture books|young adult", ANY),
    ("Folklore and myths", r"folklore|mytholog\w*|legends|fables", ANY),
    (HISTORICAL_FICTION, r"historical fiction", ANY),
    ("History", r"history|historical", NONFICTION),
    ("Biography and memoirs", r"biograph\w*|autobiograph\w*|memoirs?|personal narratives|diaries"
                              r"|correspondence", NONFICTION),
    ("Religion", r"religio\w*|bible|christian\w*|church\w*|theology|sermons|spiritual\w*|god|prayers?"
                 r"|islam|buddhis\w*|judaism|saints", NONFICTION),
    ("Philosophy", r"philosoph\w*|ethics|metaphysics", NONFICTION),
    ("Science and nature", r"science|natural history|nature|biology|physics|chemistry|astronomy"
                           r"|mathematics|technology|medicine|medical|animals|botany|zoology|geology"
                           r"|birds|plants|engineering", NONFICTION),
    ("Travel", r"travels?|voyages|description and travel|explorers|discovery and exploration", NONFICTION),
    ("Cooking and home", r"cooking|cookery|recipes|gardening|household|handicraft|crafts", NONFICTION),
    ("Hobbies and how-to", r"camping|construction|engines|brewing|tanning|fireplaces|carpentry|woodwork\w*"
                           r"|repairing|manuals?|handbooks|do-it-yourself|hunting|fishing|farm\w*"
                           r"|agricultur\w*", NONFICTION),
    ("Health and self-help", r"health|self-help|fitness|diet|psychology|success|conduct of life"
                             r"|nutrition", NONFICTION),
    ("Business and money", r"business|economics|economic conditions|finance|money|investments?"
                           r"|investing|management", NONFICTION),
    ("Politics and society", r"politics|political|government|social conditions|social life and customs"
                             r"|sociology|slavery|women's rights|feminism|law|indians of north america"
                             r"|african americans?|race relations|lynching", NONFICTION),
    ("Education and reference", r"dictionaries|language|grammar|encyclopedias|reference|etiquette"
                                r"|education|study and teaching|textbooks", NONFICTION),
    ("Essays and criticism", r"essays|criticism|books and reading|literary", NONFICTION),
]
COMPILED = [(genre, re.compile(r"\b(?:%s)\b" % words, re.I), kind) for genre, words, kind in RULES]


@lru_cache(maxsize=4096)
def _genres(subjects: tuple[str, ...]) -> tuple[str, ...]:
    if not subjects:
        return (NO_GENRE,)
    fiction = any(FICTION.search(s) for s in subjects)
    found: set[str] = set()
    for subject in subjects:
        for genre, pattern, kind in COMPILED:
            if kind == NONFICTION and fiction:
                continue
            if pattern.search(subject):
                found.add(genre)
        if fiction and HISTORY.search(subject):
            found.add(HISTORICAL_FICTION)       # "England -- History -- Fiction" is a story set in the past
    ordered = tuple(genre for genre, _words, _kind in RULES if genre in found)
    if ordered:
        return ordered
    return (GENERAL_FICTION,) if fiction else (OTHER,)


def genres(subjects: list[str]) -> list[str]:
    """The plain genres a book with these subject tags belongs to, its main one first.
    Always at least one: No genre (no tags) or Other (tags that say nothing about its kind)."""
    return list(_genres(tuple(s for s in subjects if s.strip())))


def sort_key(genre: str) -> tuple[int, str]:
    """Genres in A to Z order, with Other and then No genre at the end."""
    return ({OTHER: 1, NO_GENRE: 2}.get(genre, 0), genre.casefold())
