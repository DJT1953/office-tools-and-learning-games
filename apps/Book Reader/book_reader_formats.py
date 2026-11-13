"""One book model for every reflowable book type the Book Reader opens.

open_book(path) gives a Book: its title, authors, cover, sections (chapters,
each an HTML document), contents list, and resource(href) for anything a page
asks for. peek(path) gives a cheap BookInfo for the library list.

EPUB, Kindle (MOBI/AZW/AZW3/PRC), plain text, saved web pages (HTM/HTML/MHT)
are handled here. PDFs are shown by Qt's PDF view; only their library details
come from here. Locked books raise BookLocked with a plain reason; nothing here
removes copy protection.
"""

from __future__ import annotations

import email
import email.policy
import html as htmllib
import mimetypes
import os
import posixpath
import re
import threading
import zipfile
from dataclasses import dataclass, field

from lxml import etree
from lxml import html as lxml_html

import book_reader_mobi as mobi

READABLE = {".epub", ".mobi", ".azw", ".azw3", ".prc", ".pdf", ".txt", ".htm", ".html", ".mht", ".mhtml"}
KINDLE = {".mobi", ".azw", ".azw3", ".prc"}
SPLIT_LIMIT = 200_000          # characters of text a section may hold before it is split
SPLIT_TARGET = 100_000         # characters per piece when splitting
FONT_ALGORITHMS = {"http://www.idpf.org/2008/embedding", "http://ns.adobe.com/pdf/enc#RC"}
AMAZON_LOCK = "Locked by Amazon. It opens only in Amazon's Kindle app or at read.amazon.com."
XML_PARSER = etree.XMLParser(recover=True, resolve_entities=False, no_network=True)
STRICT_PARSER = etree.XMLParser(recover=False, resolve_entities=False, no_network=True)


@dataclass
class Section:
    href: str
    title: str = ""
    linear: bool = True


@dataclass
class TocEntry:
    title: str
    section: int
    anchor: str = ""
    depth: int = 0


@dataclass
class LockInfo:
    owner: str
    message: str
    free_copy_url: str | None = None


@dataclass
class BookInfo:
    path: str
    format: str
    title: str
    authors: list[str] = field(default_factory=list)
    cover: bytes | None = None
    lock: LockInfo | None = None
    error: str | None = None
    subjects: list[str] = field(default_factory=list)     # its subject tags, for its genre


class BookError(Exception):
    """The book can't be opened; the message says why in plain words."""


class BookLocked(Exception):
    def __init__(self, info: LockInfo) -> None:
        super().__init__(info.message)
        self.info = info


# ---------------------------------------------------------------- helpers

LATIN1_LABELS = {"iso-8859-1", "iso8859-1", "latin-1", "latin1", "l1", "ascii", "us-ascii", "cp819", "ibm819"}


def decode_html(data: bytes) -> str:
    """HTML/XHTML bytes as text, using the BOM, XML declaration or meta charset."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", "replace")
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", "replace")
    head = data[:2048]
    m = (re.search(rb"""<\?xml[^>]*encoding\s*=\s*["']([\w.-]+)""", head, re.I)
         or re.search(rb"""<meta[^>]*charset\s*=\s*["']?([\w.-]+)""", head, re.I))
    names = [m.group(1).decode("ascii", "replace")] if m else []
    # Browsers read these labels as Windows-1252, whose curly quotes and dashes sit where
    # Latin-1 has invisible control characters.
    names = ["cp1252" if n.lower().replace("_", "-") in LATIN1_LABELS else n for n in names]
    for name in names + ["utf-8"]:
        try:
            return data.decode(name)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("cp1252", "replace")


def utf8_html(text: str) -> bytes:
    """Text HTML as UTF-8 bytes with <meta charset="utf-8"> first in the head."""
    text = re.sub(r"^\s*<\?xml[^>]*>", "", text)
    text = re.sub(r"""<meta[^>]*charset[^>]*>""", "", text, flags=re.I)
    meta = '<meta charset="utf-8">'
    m = re.search(r"<head\b[^>]*>", text, re.I)
    if m:
        text = text[:m.end()] + meta + text[m.end():]
    else:
        m = re.search(r"<html\b[^>]*>", text, re.I)
        text = (text[:m.end()] + "<head>" + meta + "</head>" + text[m.end():]) if m else (
            "<!DOCTYPE html><html><head>" + meta + "</head><body>" + text + "</body></html>")
    return text.encode("utf-8")


def parse_html(data: bytes):
    """An lxml HTML document from bytes, or None when there is nothing in it."""
    text = re.sub(r"^\s*<\?xml[^>]*>", "", decode_html(data))
    if not text.strip():
        return None
    try:
        return lxml_html.document_fromstring(text)
    except (etree.ParserError, ValueError):
        return None


def html_text(data: bytes) -> str:
    """The words in an HTML document's body (no scripts, styles or head)."""
    doc = parse_html(data)
    if doc is None:
        return ""
    for bad in doc.xpath("//script|//style|//noscript|//head"):
        bad.drop_tree()
    return doc.text_content()


def clean_text(text: str) -> str:
    return " ".join(text.split())


def safe_join(base: str, href: str) -> str | None:
    """base/href as a tidy relative path, or None if it climbs out of base."""
    joined = posixpath.normpath(posixpath.join(base, href)) if base else posixpath.normpath(href)
    if joined.startswith("../") or joined == ".." or joined.startswith("/") or ":" in joined.split("/")[0]:
        return None
    return joined


def looks_like_junk(title: str) -> bool:
    return not title.strip() or bool(re.fullmatch(r"[0-9a-f]{16,}( \d+)?", title.strip(), re.I))


def title_from_file_name(path: str) -> tuple[str, list[str]]:
    """'Title - Author.mobi' / 'Title-asin_B00..-type_EBOK-v_0.azw' -> (title, authors)."""
    name = os.path.splitext(os.path.basename(path))[0]
    name = re.sub(r"-asin_[A-Z0-9]+-type_\w+-v_\d+$", "", name)
    name = re.sub(r"_EBOK$|_EBSP$", "", name)
    name = name.replace("_ ", ": ").replace("_", " ")
    authors: list[str] = []
    if " - " in name:
        title, _, author = name.rpartition(" - ")
        if author.strip().lower() not in ("unknown", ""):
            authors = [author.strip()]
        name = title
    return name.strip() or os.path.basename(path), authors


def calibre_metadata(path: str) -> tuple[str, list[str]]:
    """Title and authors from a Calibre library's metadata.opf beside the book."""
    opf = os.path.join(os.path.dirname(path), "metadata.opf")
    try:
        with open(opf, "rb") as f:
            root = etree.fromstring(f.read(), XML_PARSER)
    except (OSError, etree.XMLSyntaxError, ValueError):
        return "", []
    if root is None:
        return "", []
    title = "".join(root.xpath("//*[local-name()='title']/text()")[:1]).strip()
    authors = [a.strip() for a in root.xpath("//*[local-name()='creator']/text()") if a.strip()]
    authors = [a for a in authors if a.lower() != "unknown"]
    return ("" if looks_like_junk(title) else title), authors


def calibre_subjects(path: str) -> list[str]:
    """Subject tags from a Calibre library's metadata.opf beside the book."""
    try:
        with open(os.path.join(os.path.dirname(path), "metadata.opf"), "rb") as f:
            root = etree.fromstring(f.read(), XML_PARSER)
    except (OSError, etree.XMLSyntaxError, ValueError):
        return []
    return epub_subjects(root) if root is not None else []


def headings_toc(book: "Book") -> list[TocEntry]:
    """A contents list from each section's first heading, for books without one."""
    found = []
    for i in range(min(len(book.sections), 3000)):
        data = book.section_html(i)
        m = re.search(rb"<h([1-3])\b[^>]*>(.*?)</h\1\s*>", data, re.S | re.I)
        if m:
            title = clean_text(htmllib.unescape(re.sub(r"<[^>]*>", "", decode_html(m.group(2)))))
            if title:
                found.append(TocEntry(title[:200], i, "", 0))
    return found


ID_ATTR = re.compile(rb"""\s(?:id|name)\s*=\s*["']([^"'<>]+)["']""", re.I)


def split_html(data: bytes) -> list[bytes] | None:
    """Split a long HTML document into pieces of about SPLIT_TARGET characters,
    at the edges of its top-level blocks. None when it needn't or can't be split."""
    doc = parse_html(data)
    if doc is None:
        return None
    body = doc.find("body")
    if body is None or len(body.text_content()) <= SPLIT_LIMIT:
        return None
    # Step into the one block that holds most of the words; what comes before it goes
    # into the first piece and what comes after it into the last.
    container, levels = body, []
    while True:
        kids = list(container)
        big = [k for k in kids if len(k.text_content()) > SPLIT_TARGET]
        if len(big) != 1 or len(big[0]) < 2 or len(big[0].text_content()) < 0.6 * len(container.text_content()):
            break
        inner = big[0]
        at = kids.index(inner)
        before = htmllib.escape(container.text or "") + "".join(
            lxml_html.tostring(k, encoding="unicode") for k in kids[:at])
        after = htmllib.escape(inner.tail or "") + "".join(
            lxml_html.tostring(k, encoding="unicode") for k in kids[at + 1:])
        levels.append((inner, before, after))
        container = inner
    if len(container) < 2:
        return None
    head = doc.find("head")
    head_html = "".join(lxml_html.tostring(child, encoding="unicode") for child in head) if head is not None else ""

    def open_tag(el) -> str:
        attrs = "".join(f' {k}="{htmllib.escape(v)}"' for k, v in el.attrib.items())
        return f"<{el.tag}{attrs}>"

    pieces, chunk, size = [], [htmllib.escape(container.text or "")], 0
    for child in container:
        child_html = lxml_html.tostring(child, encoding="unicode")
        chunk.append(child_html)
        size += len(child.text_content()) + len(child.tail or "")
        if size >= SPLIT_TARGET:
            pieces.append("".join(chunk))
            chunk, size = [], 0
    if "".join(chunk).strip():
        pieces.append("".join(chunk))
    if len(pieces) < 2:
        return None
    def wrap(i: int, piece: str) -> bytes:
        first, last = i == 0, i == len(pieces) - 1
        opening = open_tag(body) + "".join((b if first else "") + open_tag(el) for el, b, _a in levels)
        closing = "".join(f"</{el.tag}>" + (a if last else "") for el, _b, a in reversed(levels)) + "</body>"
        return utf8_html(f"<!DOCTYPE html><html><head>{head_html}</head>{opening}{piece}{closing}</html>")

    return [wrap(i, p) for i, p in enumerate(pieces)]


# ---------------------------------------------------------------- the model

class Book:
    """An open reflowable book. Sections are served from section_html(i)."""

    def __init__(self, path: str, fmt: str) -> None:
        self.path = path
        self.format = fmt
        self.title = ""
        self.authors: list[str] = []
        self.cover: bytes | None = None
        self.sections: list[Section] = []
        self.toc: list[TocEntry] = []
        self._html: list[bytes] = []
        self._mime: list[str] = []
        self._by_href: dict[str, int] = {}
        self._origin: dict[str, list[int]] = {}
        self._texts: dict[int, str] = {}
        self._ids: dict[int, set[str]] = {}
        self._resource_reader = lambda href: None
        self._lock = threading.Lock()
        self._closers: list = []

    # building
    def add_section(self, href: str, data: bytes, mime: str, title: str = "", linear: bool = True) -> None:
        pieces = split_html(data) if len(data) > SPLIT_LIMIT else None
        if pieces is None:
            pieces, mimes = [data], [mime]
        else:
            mimes = ["text/html"] * len(pieces)
        indices = []
        for k, (piece, piece_mime) in enumerate(zip(pieces, mimes)):
            piece_href = href if k == 0 else f"{href}~{k + 1}"
            self._by_href[piece_href] = len(self.sections)
            indices.append(len(self.sections))
            self.sections.append(Section(piece_href, title if k == 0 else "", linear))
            self._html.append(piece)
            self._mime.append(piece_mime)
        self._origin[href] = indices

    # reading
    def section_html(self, i: int) -> bytes:
        return self._html[i]

    def section_text(self, i: int) -> str:
        if i not in self._texts:
            self._texts[i] = html_text(self._html[i])
        return self._texts[i]

    def section_ids(self, i: int) -> set[str]:
        if i not in self._ids:
            self._ids[i] = {m.decode("utf-8", "replace") for m in ID_ATTR.findall(self._html[i])}
        return self._ids[i]

    def resource(self, href: str) -> tuple[bytes, str] | None:
        """A section or anything a page asks for, by its path inside the book."""
        href = href.split("#", 1)[0].split("?", 1)[0]
        tidy = safe_join("", href)
        if tidy is None:
            return None
        if tidy in self._by_href:
            i = self._by_href[tidy]
            return self._html[i], self._mime[i]
        with self._lock:
            try:
                return self._resource_reader(tidy)
            except (OSError, KeyError, ValueError, zipfile.BadZipFile, mobi.NotABook):
                return None

    def find_href(self, href: str) -> tuple[int, str] | None:
        """(section number, anchor) for a link target path#anchor inside the book."""
        path, _, anchor = href.partition("#")
        path = safe_join("", path) if path else ""
        if not path:
            return None
        indices = self._origin.get(path)
        if indices is None and path in self._by_href:
            indices = [self._by_href[path]]
        if indices is None:
            return None
        if anchor:
            for i in indices:
                if anchor in self.section_ids(i):
                    return i, anchor
        return indices[0], anchor

    def close(self) -> None:
        for closer in self._closers:
            try:
                closer()
            except OSError:
                pass
        self._closers.clear()

    def finish(self) -> None:
        if not self.toc:
            self.toc = headings_toc(self)
        if looks_like_junk(self.title) or not self.authors:
            title, authors = calibre_metadata(self.path)
            if looks_like_junk(self.title):
                self.title = title
            if not self.authors:
                self.authors = authors
        if looks_like_junk(self.title):
            title, authors = title_from_file_name(self.path)
            self.title = title
            self.authors = self.authors or authors


def open_book(path: str) -> Book:
    """Open any reflowable book. Raises BookLocked or BookError."""
    ext = os.path.splitext(path)[1].lower()
    if not os.path.isfile(path):
        raise BookError("The file isn't there any more (it may have been moved or deleted).")
    try:
        if ext == ".epub":
            book = open_epub(path)
        elif ext in KINDLE:
            book = open_kindle(path)
        elif ext == ".txt":
            book = open_text(path)
        elif ext in (".htm", ".html", ".xhtml"):
            book = open_web_page(path)
        elif ext in (".mht", ".mhtml"):
            book = open_mht(path)
        elif ext == ".pdf":
            raise BookError("PDF files open in the PDF view.")
        else:
            raise BookError("This kind of file isn't a book this reader can open.")
    except mobi.NotABook as e:
        raise BookError(str(e)) from e
    except (zipfile.BadZipFile, EOFError) as e:
        raise BookError("This EPUB file is damaged or incomplete (perhaps cut short while downloading).") from e
    except OSError as e:
        raise BookError(f"The file can't be read ({e.strerror or e}).") from e
    except (BookLocked, BookError):
        raise
    except Exception as e:             # any other damage: a plain message, never a crash
        raise BookError(f"This book is damaged, so it can't be opened ({type(e).__name__}).") from e
    if not book.sections:
        book.close()
        raise BookError("This book has no pages in it.")
    book.finish()
    return book


# ---------------------------------------------------------------- EPUB

def epub_lock(z: zipfile.ZipFile) -> LockInfo | None:
    names = set(z.namelist())
    if "META-INF/rights.xml" in names:
        return LockInfo("Adobe", "Locked with Adobe copy protection (from a library or a bookstore). "
                                 "It opens in Adobe Digital Editions.")
    if "META-INF/license.lcpl" in names:
        return LockInfo("Readium LCP", "Locked with library copy protection (LCP). "
                                       "It opens in the library's own reading app.")
    if "META-INF/sinf.xml" in names:
        return LockInfo("Apple", "Locked by Apple Books. It opens only in Apple Books.")
    if "META-INF/encryption.xml" not in names:
        return None
    data = z.read("META-INF/encryption.xml")
    algorithms = {a.decode("utf-8", "replace") for a in re.findall(rb'Algorithm\s*=\s*"([^"]+)"', data)}
    others = {a for a in algorithms if a not in FONT_ALGORITHMS and "xmldsig" not in a}
    if not others:
        return None
    if any("samsung" in a.lower() for a in others):
        number = (re.search(rb'URI\s*=\s*"(\d{2,6})/', data)
                  or re.search(rb"@files@(\d{2,6})@", data))
        url = f"https://www.gutenberg.org/ebooks/{int(number.group(1))}" if number else None
        message = "Locked by Samsung Books. It opens only in the Samsung Books app."
        if url:
            message += " It is a free Project Gutenberg book, so you can get an unlocked copy."
        return LockInfo("Samsung", message, url)
    if any("adobe" in a.lower() for a in others):
        return LockInfo("Adobe", "Locked with Adobe copy protection. It opens in Adobe Digital Editions.")
    return LockInfo("copy protection", "Locked with copy protection. It opens only in the app "
                                       "or store it came from.")


def locked_title(info: LockInfo) -> str:
    if info.free_copy_url:
        return "Locked book (Project Gutenberg number %s)" % info.free_copy_url.rsplit("/", 1)[-1]
    return ""


def epub_package(z: zipfile.ZipFile):
    """(OPF path, parsed OPF root)."""
    opf_path = None
    try:
        container = etree.fromstring(z.read("META-INF/container.xml"), XML_PARSER)
        if container is not None:
            paths = container.xpath("//*[local-name()='rootfile']/@full-path")
            opf_path = paths[0] if paths else None
    except KeyError:
        pass
    if not opf_path:
        candidates = [n for n in z.namelist() if n.lower().endswith(".opf")]
        opf_path = candidates[0] if candidates else None
    if not opf_path:
        raise BookError("This EPUB file is damaged (its list of contents is missing).")
    root = etree.fromstring(z.read(opf_path), XML_PARSER)
    if root is None:
        raise BookError("This EPUB file is damaged (its list of contents can't be read).")
    return opf_path, root


def epub_metadata(root) -> tuple[str, list[str]]:
    title = clean_text("".join(root.xpath("//*[local-name()='metadata']/*[local-name()='title'][1]//text()")))
    authors = [clean_text("".join(c.itertext())) for c in root.xpath("//*[local-name()='metadata']/*[local-name()='creator']")]
    return title, [a for a in authors if a and a.lower() != "unknown"]


def epub_subjects(root) -> list[str]:
    """The dc:subject tags of an OPF package (an EPUB's, or Calibre's metadata.opf)."""
    found = [clean_text("".join(s.itertext())) for s in root.xpath("//*[local-name()='subject']")]
    return list(dict.fromkeys(s for s in found if s))


def epub_manifest(root, base: str) -> dict[str, tuple[str, str, str]]:
    items = {}
    for item in root.xpath("//*[local-name()='manifest']/*[local-name()='item']"):
        href = safe_join(base, item.get("href", "").split("#")[0])
        if href and item.get("id"):
            from urllib.parse import unquote
            items[item.get("id")] = (unquote(href), item.get("media-type", ""), item.get("properties", ""))
    return items


def epub_cover(z: zipfile.ZipFile, root, manifest) -> bytes | None:
    href = None
    for _id, (h, mime, props) in manifest.items():
        if "cover-image" in props.split():
            href = h
    if href is None:
        ids = root.xpath("//*[local-name()='metadata']/*[local-name()='meta'][@name='cover']/@content")
        if ids and ids[0] in manifest:
            href = manifest[ids[0]][0]
    if href is None:
        for item_id, (h, mime, _props) in manifest.items():
            if mime.startswith("image/") and "cover" in (item_id + h).lower():
                href = h
                break
    if href is None:
        return None
    try:
        data = z.read(href)
    except KeyError:
        return None
    return data if mobi.image_mime(data) else None


def open_epub(path: str) -> Book:
    z = zipfile.ZipFile(path)
    book = Book(path, "EPUB")
    book._closers.append(z.close)
    try:
        lock = epub_lock(z)
        if lock:
            raise BookLocked(lock)
        opf_path, root = epub_package(z)
        base = posixpath.dirname(opf_path)
        book.title, book.authors = epub_metadata(root)
        manifest = epub_manifest(root, base)
        book.cover = epub_cover(z, root, manifest)
        names = set(z.namelist())
        spine = root.xpath("//*[local-name()='spine']")
        for ref in root.xpath("//*[local-name()='spine']/*[local-name()='itemref']"):
            item = manifest.get(ref.get("idref", ""))
            if not item or item[0] not in names:
                continue
            href, mime, _props = item
            data = z.read(href)
            book.add_section(href, *epub_section(data, mime), linear=ref.get("linear", "yes") != "no")

        def read(href: str):
            if href not in names:
                return None
            data = z.read(href)
            guessed = mimetypes.guess_type(href)[0] or "application/octet-stream"
            for _id, (h, mime, _p) in manifest.items():
                if h == href and mime:
                    guessed = mime
            if guessed in ("application/xhtml+xml", "text/html"):
                data, guessed = epub_section(data, guessed)
            return data, guessed

        book._resource_reader = read
        toc_id = spine[0].get("toc") if spine else None
        nav = [h for h, _m, props in manifest.values() if "nav" in props.split()]
        if nav and nav[0] in names:
            book.toc = nav_toc(book, z.read(nav[0]), posixpath.dirname(nav[0]))
        if not book.toc:
            ncx = manifest.get(toc_id or "", (None,))[0] or next(
                (h for h, m, _p in manifest.values() if m == "application/x-dtbncx+xml"), None)
            if ncx and ncx in names:
                book.toc = ncx_toc(book, z.read(ncx), posixpath.dirname(ncx))
    except BaseException:
        book.close()
        raise
    return book


def epub_section(data: bytes, mime: str) -> tuple[bytes, str]:
    """Well-formed XHTML is served as XHTML; anything else as forgiving HTML."""
    if mime == "application/xhtml+xml" or data.lstrip()[:5] == b"<?xml":
        try:
            etree.fromstring(data, STRICT_PARSER)
            return data, "application/xhtml+xml"
        except etree.XMLSyntaxError:
            pass
    return utf8_html(decode_html(data)), "text/html"


def nav_toc(book: Book, data: bytes, base: str) -> list[TocEntry]:
    doc = parse_html(data)
    if doc is None:
        return []
    navs = [n for n in doc.iter("nav") if "toc" in (n.get("epub:type") or n.get("type") or "").split()]
    navs = navs or list(doc.iter("nav"))
    if not navs:
        return []
    found: list[TocEntry] = []

    def walk(ol, depth: int) -> None:
        for li in ol.iterchildren("li"):
            link = li.find("a")
            label_el = link if link is not None else li.find("span")
            if label_el is not None:
                title = clean_text(label_el.text_content())
                href = link.get("href") if link is not None else None
                target = book.find_href(resolve(base, href)) if href else None
                if title and target:
                    found.append(TocEntry(title, target[0], target[1], depth))
            for sub in li.iterchildren("ol"):
                walk(sub, depth + 1)

    for ol in navs[0].iter("ol"):
        walk(ol, 0)
        break
    return found


def ncx_toc(book: Book, data: bytes, base: str) -> list[TocEntry]:
    root = etree.fromstring(data, XML_PARSER)
    if root is None:
        return []
    found: list[TocEntry] = []

    def walk(point, depth: int) -> None:
        label = clean_text("".join(point.xpath("./*[local-name()='navLabel']//*[local-name()='text']//text()")))
        src = point.xpath("./*[local-name()='content']/@src")
        target = book.find_href(resolve(base, src[0])) if src else None
        if label and target:
            found.append(TocEntry(label, target[0], target[1], depth))
        for child in point.xpath("./*[local-name()='navPoint']"):
            walk(child, depth + 1)

    for point in root.xpath("//*[local-name()='navMap']/*[local-name()='navPoint']"):
        walk(point, 0)
    return found


def resolve(base: str, href: str) -> str:
    from urllib.parse import unquote
    path, _, anchor = href.partition("#")
    joined = safe_join(base, unquote(path)) if path else ""
    return (joined or "") + ("#" + anchor if anchor else "")


def peek_epub(path: str) -> BookInfo:
    with zipfile.ZipFile(path) as z:
        lock = epub_lock(z)
        title, authors, cover, subjects = "", [], None, []
        try:
            _opf, root = epub_package(z)
            title, authors = epub_metadata(root)
            subjects = epub_subjects(root)
            cover = epub_cover(z, root, epub_manifest(root, posixpath.dirname(_opf)))
        except (BookError, KeyError, etree.XMLSyntaxError, ValueError):
            if not lock:
                raise
        if lock and looks_like_junk(title):
            title = locked_title(lock)
        return BookInfo(path, "EPUB", title, authors, cover, lock, subjects=subjects)


# ---------------------------------------------------------------- Kindle

def open_kindle(path: str) -> Book:
    kindle = mobi.MobiFile.open(path)
    if kindle.lock:
        raise BookLocked(LockInfo("Amazon", AMAZON_LOCK))
    book = Book(path, "Kindle")
    book.title, book.authors, book.cover = kindle.title, kindle.authors, kindle.cover()
    parts = kindle.book_parts()
    for part in parts:
        book.add_section(part.name, part.html, "text/html")
    book._resource_reader = kindle.resource
    for title, part, anchor, depth in kindle.toc():
        target = book.find_href(parts[part].name + ("#" + anchor if anchor else ""))
        if target and title:
            book.toc.append(TocEntry(title, target[0], target[1], depth))
    return book


# ---------------------------------------------------------------- plain text

CHAPTER = re.compile(
    r"^\s*(?:(?:CHAPTER|Chapter|BOOK|Book|PART|Part|ACT|Act|SECTION|Section|STAVE|Stave)\s+"
    r"(?:[0-9]+|[IVXLCDM]+|[ivxlcdm]+|[A-Z][A-Za-z]+)\b[^\n]{0,70}"
    r"|PREFACE|INTRODUCTION|PROLOGUE|EPILOGUE|CONTENTS|Preface|Introduction|Prologue|Epilogue)\s*$")


def decode_text(data: bytes) -> str:
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", "replace")
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", "replace").lstrip("﻿")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1252", "replace")


def gutenberg_details(text: str) -> tuple[str, list[str]]:
    head = text[:8000]
    title = re.search(r"^\s*Title:\s*(.+(?:\n {6,}.+)*)", head, re.M)
    author = re.search(r"^\s*Author:\s*(.+)", head, re.M)
    return (clean_text(title.group(1)) if title else "",
            [clean_text(author.group(1))] if author else [])


def open_text(path: str) -> Book:
    with open(path, "rb") as f:
        text = decode_text(f.read()).replace("\r\n", "\n").replace("\r", "\n")
    book = Book(path, "Text")
    book.title, book.authors = gutenberg_details(text)
    lines = text.split("\n")
    lengths = sorted(len(l.rstrip()) for l in lines if l.strip())
    wrap = max(lengths[int(len(lengths) * 0.9)] if lengths else 70, 40)
    blocks: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        if line.strip():
            current.append(line.rstrip())
        elif current:
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    sections: list[tuple[str | None, list[str]]] = [(None, [])]
    count = 0
    for block in blocks:
        if len(block) == 1 and CHAPTER.match(block[0]):
            count += 1
            heading = clean_text(block[0])
            sections.append((heading, [f'<h2 id="c{count}">{htmllib.escape(heading)}</h2>']))
            continue
        if len(block) == 1:
            para = htmllib.escape(block[0].strip())
        else:
            long = sum(1 for l in block[:-1] if len(l) >= 0.75 * wrap)
            joiner = " " if long >= max(1, (len(block) - 1) // 2) else "<br>"
            para = joiner.join(htmllib.escape(l.strip() if joiner == " " else l) for l in block)
        sections[-1][1].append(f"<p>{para}</p>")
    # Headings with nothing after them (a contents listing) join the next section.
    merged: list[tuple[list[str], list[str]]] = []
    pending_html: list[str] = []
    pending_heads: list[str] = []
    for heading, html_parts in sections:
        if heading is not None:
            pending_heads.append(heading)
        pending_html += html_parts
        if any(p.startswith("<p>") for p in html_parts):
            merged.append((pending_html, pending_heads))
            pending_html, pending_heads = [], []
    if pending_html:
        if merged:
            merged[-1][0].extend(pending_html)
            merged[-1][1].extend(pending_heads)
        else:
            merged.append((pending_html, pending_heads))
    pieces: list[tuple[list[str], list[str]]] = []
    for html_parts, heads in merged:
        if not heads and sum(len(p) for p in html_parts) > 60_000:
            chunk: list[str] = []
            for p in html_parts:
                chunk.append(p)
                if sum(len(c) for c in chunk) > 50_000:
                    pieces.append((chunk, []))
                    chunk = []
            if chunk:
                pieces.append((chunk, []))
        else:
            pieces.append((html_parts, heads))
    title = htmllib.escape(book.title or os.path.basename(path))
    for n, (html_parts, heads) in enumerate(pieces, 1):
        href = "text%04d.html" % n
        doc = (f'<!DOCTYPE html><html><head><meta charset="utf-8"><title>{title}</title></head><body>'
               + "\n".join(html_parts) + "</body></html>")
        book.add_section(href, doc.encode("utf-8"), "text/html", heads[0] if heads else "")
    for i in range(len(book.sections)):
        for m in re.finditer(rb'<h2 id="(c\d+)">(.*?)</h2>', book.section_html(i)):
            book.toc.append(TocEntry(htmllib.unescape(m.group(2).decode("utf-8")), i, m.group(1).decode(), 0))
    return book


def peek_text(path: str) -> BookInfo:
    with open(path, "rb") as f:
        text = decode_text(f.read(65536))
    title, authors = gutenberg_details(text)
    return BookInfo(path, "Text", title, authors)


# ---------------------------------------------------------------- web pages

def open_web_page(path: str) -> Book:
    with open(path, "rb") as f:
        text = decode_html(f.read())
    book = Book(path, "Web page")
    m = re.search(r"<title[^>]*>(.*?)</title>", text, re.S | re.I)
    book.title = clean_text(htmllib.unescape(m.group(1))) if m else ""
    folder = os.path.dirname(os.path.abspath(path))
    book.add_section(os.path.basename(path).replace(" ", "_"), utf8_html(text), "text/html")

    def read(href: str):
        from urllib.parse import unquote
        rel = safe_join("", unquote(href))
        if rel is None:
            return None
        full = os.path.abspath(os.path.join(folder, *rel.split("/")))
        if os.path.commonpath([folder, full]) != folder or not os.path.isfile(full):
            return None
        with open(full, "rb") as f:
            return f.read(), mimetypes.guess_type(full)[0] or "application/octet-stream"

    book._resource_reader = read
    return book


def open_mht(path: str) -> Book:
    with open(path, "rb") as f:
        message = email.message_from_bytes(f.read(), policy=email.policy.default)
    book = Book(path, "Web page")
    book.title = clean_text(str(message.get("Subject", "") or ""))
    main = None
    resources: dict[str, tuple[bytes, str]] = {}
    names: dict[str, str] = {}
    for n, part in enumerate(message.walk()):
        if part.is_multipart():
            continue
        mime = part.get_content_type()
        payload = part.get_payload(decode=True) or b""
        if main is None and mime == "text/html":
            charset = part.get_content_charset() or "utf-8"
            try:
                main = payload.decode(charset, "replace")
            except LookupError:
                main = payload.decode("utf-8", "replace")
            continue
        ext = mimetypes.guess_extension(mime) or ""
        name = "mht/%03d%s" % (n, ext)
        resources[name] = (payload, mime)
        for key in ("Content-Location", "Content-ID"):
            value = part.get(key)
            if value:
                value = str(value).strip()
                names["cid:" + value.strip("<>") if key == "Content-ID" else value] = name
    if main is None:
        raise BookError("This saved web page has no page in it.")
    for original in sorted(names, key=len, reverse=True):
        main = main.replace(original, names[original]).replace(htmllib.escape(original), names[original])
    if not book.title:
        m = re.search(r"<title[^>]*>(.*?)</title>", main, re.S | re.I)
        book.title = clean_text(htmllib.unescape(m.group(1))) if m else ""
    book.add_section("page.html", utf8_html(main), "text/html")
    book._resource_reader = resources.get
    return book


def peek_web_page(path: str, fmt: str = "Web page") -> BookInfo:
    with open(path, "rb") as f:
        data = f.read(65536)
    title = ""
    if path.lower().endswith((".mht", ".mhtml")):
        m = re.search(rb"^Subject:\s*(.+)$", data, re.M | re.I)
        if m:
            title = str(email.message_from_bytes(b"Subject: " + m.group(1).strip() + b"\n\n",
                                                 policy=email.policy.default).get("Subject", ""))
    if not title:
        m = re.search(r"<title[^>]*>(.*?)</title>", decode_html(data), re.S | re.I)
        title = clean_text(htmllib.unescape(m.group(1))) if m else ""
    return BookInfo(path, fmt, clean_text(title))


# ---------------------------------------------------------------- PDF details

def pdf_string(raw: bytes) -> str:
    """A PDF literal '(...)' or hex '<...>' string as text."""
    if raw.startswith(b"<"):
        digits = re.sub(rb"[^0-9A-Fa-f]", b"", raw)
        data = bytes.fromhex((digits + b"0" * (len(digits) % 2)).decode())
    else:
        body, out, i = raw[1:-1], bytearray(), 0
        escapes = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}
        while i < len(body):
            c = body[i]
            if c == 0x5C and i + 1 < len(body):
                nxt = body[i + 1]
                if nxt in escapes:
                    out.append(escapes[nxt])
                    i += 2
                elif 0x30 <= nxt <= 0x37:
                    m = re.match(rb"[0-7]{1,3}", body[i + 1:i + 4])
                    out.append(int(m.group(0), 8) & 0xFF)
                    i += 1 + len(m.group(0))
                elif nxt in (10, 13):
                    i += 2
                else:
                    out.append(nxt)
                    i += 2
            else:
                out.append(c)
                i += 1
        data = bytes(out)
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", "replace")
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", "replace")
    return data.decode("cp1252", "replace")


def peek_pdf(path: str) -> BookInfo:
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        head = f.read(1 << 20)
        tail = b""
        if size > 2 << 20:
            f.seek(size - (1 << 20))
            tail = f.read()
    if not head.startswith(b"%PDF"):
        raise BookError("This PDF file is damaged.")
    data = head + tail
    title = ""
    if b"/Encrypt" not in data:
        found = re.findall(rb"/Title\s*(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]*>)", data, re.S)
        if found:
            title = clean_text(pdf_string(found[-1]))
    return BookInfo(path, "PDF", title)


# ---------------------------------------------------------------- peek

def peek(path: str) -> BookInfo:
    """Library details for any book file. Never raises: problems go in .error."""
    ext = os.path.splitext(path)[1].lower()
    fmt = {".epub": "EPUB", ".pdf": "PDF", ".txt": "Text"}.get(ext, "Kindle" if ext in KINDLE else "Web page")
    try:
        if ext == ".epub":
            info = peek_epub(path)
        elif ext in KINDLE:
            kindle = mobi.MobiFile.open(path)
            lock = LockInfo("Amazon", AMAZON_LOCK) if kindle.lock else None
            info = BookInfo(path, "Kindle", kindle.title, kindle.authors,
                            None if kindle.lock else kindle.cover(), lock, subjects=kindle.subjects)
        elif ext == ".pdf":
            info = peek_pdf(path)
        elif ext == ".txt":
            info = peek_text(path)
        elif ext in (".htm", ".html", ".xhtml", ".mht", ".mhtml"):
            info = peek_web_page(path)
        else:
            info = BookInfo(path, fmt, "", error="This kind of file isn't a book this reader can open.")
    except (mobi.NotABook, BookError) as e:
        info = BookInfo(path, fmt, "", error=str(e))
    except (zipfile.BadZipFile, EOFError):
        info = BookInfo(path, fmt, "", error="This EPUB file is damaged or incomplete (perhaps cut short while downloading).")
    except OSError as e:
        info = BookInfo(path, fmt, "", error=f"The file can't be read ({e.strerror or e}).")
    except Exception as e:             # any other damage: listed as can't-be-opened
        info = BookInfo(path, fmt, "", error=f"This file is damaged ({type(e).__name__}).")
    if looks_like_junk(info.title) or not info.authors:
        title, authors = calibre_metadata(path)
        if looks_like_junk(info.title) and title:
            info.title = title
        info.authors = info.authors or authors
    if looks_like_junk(info.title):
        title, authors = title_from_file_name(path)
        info.title = title
        info.authors = info.authors or authors
    if not info.subjects:
        info.subjects = calibre_subjects(path)
    return info
