"""Save a copy of a book as EPUB, plain text, one web page (for PDF), or Word.

save_epub and save_text work from any open Book. single_html writes the whole
book as one web page with its pictures beside it; Chromium prints that to PDF
(book_reader_copy) and Word turns it into a .docx (save_word). Locked books
are never copied.
"""

from __future__ import annotations

import html as htmllib
import os
import posixpath
import re
import threading
import time
import uuid
import zipfile

from lxml import etree
from lxml import html as lxml_html

import book_reader_formats as formats
import book_reader_mobi as mobi

XHTML_NS = "http://www.w3.org/1999/xhtml"
EXTENSIONS = {"EPUB": ".epub", "PDF": ".pdf", "Word": ".docx", "Plain text": ".txt"}


class Cancelled(Exception):
    """Cancel was pressed."""


def copy_kinds(fmt: str, locked: bool = False) -> list[str]:
    """The kinds of copy offered for a book of this format."""
    if locked:
        return []
    if fmt == "PDF":
        return ["Plain text"]
    if fmt == "EPUB":
        return ["PDF", "Word", "Plain text"]
    return ["EPUB", "PDF", "Word", "Plain text"]


def _check(stop: threading.Event | None) -> None:
    if stop is not None and stop.is_set():
        raise Cancelled()


def _doc(book, i: int):
    doc = formats.parse_html(book.section_html(i))
    if doc is None:
        doc = lxml_html.document_fromstring("<html><head></head><body></body></html>")
    for bad in doc.xpath("//script|//noscript|//iframe|//object|//embed"):
        bad.drop_tree()
    return doc


def _is_local(value: str) -> bool:
    return bool(value) and not re.match(r"^(#|[a-z][a-z0-9+.-]*:)", value, re.I)


# ---------------------------------------------------------------- plain text

def save_text(book, out: str) -> None:
    lines = [book.title or "Book"]
    if book.authors:
        lines.append("by " + ", ".join(book.authors))
    lines.append("")
    for i in range(len(book.sections)):
        text = book.section_text(i).replace("\r", "")
        text = "\n".join(line.strip() for line in text.split("\n"))
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        if text:
            lines += [text, ""]
    with open(out, "w", encoding="utf-8", newline="\r\n") as f:
        f.write("\n".join(lines))


# ---------------------------------------------------------------- EPUB

def _file_name(href: str) -> str:
    name = re.sub(r"[^\w./-]", "_", href.replace("~", "_"))
    if not name.lower().endswith((".xhtml", ".html", ".htm")):
        name += ".xhtml"
    return name


def _xhtml(doc, title: str) -> bytes:
    for el in doc.iter():
        if not isinstance(el.tag, str):
            continue
        for name in list(el.attrib):
            if name.startswith("xlink:"):
                el.attrib[name.split(":", 1)[1]] = el.attrib.pop(name)
            elif ":" in name and not name.startswith("xml:") or name == "xmlns" or name.startswith("on"):
                del el.attrib[name]
    for meta in doc.xpath("//meta[@charset or @http-equiv]"):
        meta.drop_tree()
    head = doc.find("head")
    if head is None:
        head = etree.SubElement(doc, "head")
        doc.insert(0, head)
    if head.find("title") is None:
        etree.SubElement(head, "title").text = title
    text = etree.tostring(doc, method="xml", encoding="unicode")
    text = re.sub(r"^<html\b", f'<html xmlns="{XHTML_NS}"', text, count=1)
    return ('<?xml version="1.0" encoding="utf-8"?>\n' + text).encode("utf-8")


def save_epub(book, out: str, stop: threading.Event | None = None) -> None:
    """Write the book as an EPUB 3 file (with an EPUB 2 contents file for older readers)."""
    names = {s.href: "OEBPS/" + _file_name(s.href) for s in book.sections}
    for original, pieces in getattr(book, "_origin", {}).items():
        names.setdefault(original, names[book.sections[pieces[0]].href])
    sections: list[tuple[str, bytes]] = []
    resources: dict[str, tuple[bytes, str]] = {}
    for i, section in enumerate(book.sections):
        _check(stop)
        doc = _doc(book, i)
        base = posixpath.dirname(section.href)
        mine = names[section.href]
        for el in doc.iter():
            if not isinstance(el.tag, str):
                continue
            for attr in ("src", "href", "xlink:href"):
                value = el.get(attr)
                if not value or not _is_local(value):
                    continue
                target = formats.resolve(base, value)
                path, _, anchor = target.partition("#")
                if path in names:
                    rel = posixpath.relpath(names[path], posixpath.dirname(mine))
                    el.set(attr, rel + ("#" + anchor if anchor else ""))
                elif path:
                    found = book.resource(path)
                    if found and not found[1].startswith(("text/html", "application/xhtml")):
                        resources["OEBPS/" + path] = found
                    elif el.tag == "a":
                        el.attrib.pop(attr)
        sections.append((mine, _xhtml(doc, book.title or "Book")))
    cover_name = None
    if book.cover:
        ext = mobi.MIME_EXT.get(mobi.image_mime(book.cover) or "", "jpg")
        cover_name = f"OEBPS/cover-image.{ext}"
        resources[cover_name] = (book.cover, mobi.image_mime(book.cover) or "image/jpeg")
    _check(stop)
    toc = [(e.title, names[book.sections[e.section].href] + ("#" + e.anchor if e.anchor else ""))
           for e in book.toc if 0 <= e.section < len(book.sections)]
    if not toc:
        toc = [(f"Part {n + 1}", name) for n, (name, _d) in enumerate(sections)]
    ident = "urn:uuid:" + str(uuid.uuid4())
    title = htmllib.escape(book.title or "Book")
    items, spine = [], []
    for n, (name, _data) in enumerate(sections):
        items.append(f'<item id="s{n}" href="{htmllib.escape(name[6:])}" media-type="application/xhtml+xml"/>')
        spine.append(f'<itemref idref="s{n}"/>')
    for n, (name, (_data, mime)) in enumerate(sorted(resources.items())):
        props = ' properties="cover-image"' if name == cover_name else ""
        items.append(f'<item id="r{n}" href="{htmllib.escape(name[6:])}" media-type="{mime}"{props}/>')
    creators = "".join(f"<dc:creator>{htmllib.escape(a)}</dc:creator>" for a in book.authors)
    opf = (f'<?xml version="1.0" encoding="utf-8"?>\n<package xmlns="http://www.idpf.org/2007/opf" version="3.0" '
           f'unique-identifier="bookid"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/">'
           f'<dc:identifier id="bookid">{ident}</dc:identifier><dc:title>{title}</dc:title>{creators}'
           f'<dc:language>en</dc:language><meta property="dcterms:modified">'
           f'{time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}</meta></metadata><manifest>'
           f'<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>'
           f'<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>{"".join(items)}</manifest>'
           f'<spine toc="ncx">{"".join(spine)}</spine></package>')
    links = "".join(f'<li><a href="{htmllib.escape(href[6:])}">{htmllib.escape(t)}</a></li>' for t, href in toc)
    nav = (f'<?xml version="1.0" encoding="utf-8"?>\n<html xmlns="{XHTML_NS}" xmlns:epub="http://www.idpf.org/2007/ops">'
           f'<head><title>{title}</title></head><body><nav epub:type="toc"><h1>Contents</h1><ol>{links}</ol></nav>'
           f'</body></html>')
    points = "".join(f'<navPoint id="p{n}" playOrder="{n + 1}"><navLabel><text>{htmllib.escape(t)}</text></navLabel>'
                     f'<content src="{htmllib.escape(href[6:])}"/></navPoint>' for n, (t, href) in enumerate(toc))
    ncx = (f'<?xml version="1.0" encoding="utf-8"?>\n<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">'
           f'<head><meta name="dtb:uid" content="{ident}"/></head><docTitle><text>{title}</text></docTitle>'
           f'<navMap>{points}</navMap></ncx>')
    temp = out + ".part"
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(zipfile.ZipInfo("mimetype"), "application/epub+zip", compress_type=zipfile.ZIP_STORED)
            z.writestr("META-INF/container.xml",
                       '<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:'
                       'xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" '
                       'media-type="application/oebps-package+xml"/></rootfiles></container>')
            z.writestr("OEBPS/content.opf", opf)
            z.writestr("OEBPS/nav.xhtml", nav)
            z.writestr("OEBPS/toc.ncx", ncx)
            for name, data in sections:
                _check(stop)
                z.writestr(name, data)
            for name, (data, _mime) in resources.items():
                z.writestr(name, data)
        os.replace(temp, out)
    finally:
        if os.path.exists(temp):
            os.remove(temp)


# ---------------------------------------------------------------- one web page (for PDF and Word)

def single_html(book, folder: str, stop: threading.Event | None = None) -> str:
    """The whole book as one web page, pictures in folder/images. Returns the page's path."""
    os.makedirs(os.path.join(folder, "images"), exist_ok=True)
    parts, count = [], 0
    for i, section in enumerate(book.sections):
        _check(stop)
        doc = _doc(book, i)
        body = doc.find("body")
        if body is None:
            continue
        for bad in body.xpath(".//style|.//link"):
            bad.drop_tree()
        base = posixpath.dirname(section.href)
        for el in body.iter("img", "image"):
            src = el.get("src") or el.get("href") or el.get("xlink:href") or ""
            found = book.resource(formats.resolve(base, src).split("#")[0]) if _is_local(src) else None
            if not found:
                el.drop_tree()
                continue
            count += 1
            name = f"images/{count:04d}.{mobi.MIME_EXT.get(found[1], 'img')}"
            with open(os.path.join(folder, name), "wb") as f:
                f.write(found[0])
            if el.tag == "image":
                el.attrib.pop("xlink:href", None)
                el.set("href", name)
            else:
                el.set("src", name)
        for a in body.iter("a"):
            if a.get("href") and _is_local(a.get("href")):
                del a.attrib["href"]
        inner = htmllib.escape(body.text or "") + "".join(lxml_html.tostring(c, encoding="unicode") for c in body)
        parts.append(f'<div class="chapter">{inner}</div>')
    title = htmllib.escape(book.title or "Book")
    by = f"<p class='by'>by {htmllib.escape(', '.join(book.authors))}</p>" if book.authors else ""
    page = ('<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + title + '</title><style>'
            'body { font-family: Georgia, serif; font-size: 12pt; line-height: 1.45; color: black; }'
            ' img, svg { max-width: 100%; height: auto; } .chapter { page-break-before: always; }'
            ' h1, h2, h3 { page-break-after: avoid; } .booktitle { text-align: center; margin-top: 30%; }'
            ' .by { text-align: center; }</style></head><body>'
            f'<h1 class="booktitle">{title}</h1>{by}' + "".join(parts) + "</body></html>")
    path = os.path.join(folder, "book.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(page)
    return path


# ---------------------------------------------------------------- Word

def save_word(html_path: str, out: str, stop: threading.Event | None = None, timeout_s: float = 300) -> None:
    """Have Microsoft Word open the one-page book and save it as .docx, pictures inside."""
    _check(stop)
    import comtypes
    import comtypes.client
    comtypes.CoInitialize()
    word = None
    try:
        word = comtypes.client.CreateObject("Word.Application", dynamic=True)
        word.Visible = False
        word.DisplayAlerts = 0
        _check(stop)
        # Word's late-bound methods take their arguments in order, not by name:
        # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
        # Revert, WritePasswordDocument, WritePasswordTemplate, Format (0 = auto), Encoding (UTF-8), Visible.
        doc = word.Documents.Open(str(os.path.abspath(html_path)), False, True, False, "", "", False, "", "",
                                  0, 65001, False)
        try:
            shapes = doc.InlineShapes
            for n in range(1, shapes.Count + 1):
                try:
                    link = shapes.Item(n).LinkFormat
                    if link is not None:
                        link.SavePictureWithDocument = True
                        link.BreakLink()
                except Exception:   # a picture that isn't linked: nothing to do
                    pass
            _check(stop)
            doc.SaveAs2(str(os.path.abspath(out)), 16)          # 16 = Word document (.docx)
        finally:
            doc.Close(0)
    finally:
        if word is not None:
            try:
                word.Quit(0)
            except Exception:
                pass
        word = None
        comtypes.CoUninitialize()
