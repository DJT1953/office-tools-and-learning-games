"""Reads Kindle MOBI, AZW, AZW3 and PRC books (unlocked ones only).

A Kindle book is a Palm database: a header, then a list of records. Record 0
describes the book (PalmDOC header, MOBI header, EXTH metadata). Then come the
compressed text records, then pictures and other resources. Older books
("MOBI 6") hold one long HTML text. Newer ones ("KF8", also the second half of
"combined" files) hold the text as skeletons and fragments that must be put
back together, described by index records.

Locked books (Amazon's copy protection) are recognised and reported, never
opened. Nothing here removes copy protection.
"""

from __future__ import annotations

import bisect
import re
import struct
import zlib
from dataclasses import dataclass

ENCODINGS = {65001: "utf-8", 1252: "cp1252"}
NO_INDEX = 0xFFFFFFFF


class NotABook(Exception):
    """The file can't be read as a book; the message says why in plain words."""


def palmdoc_decompress(data: bytes) -> bytes:
    """Undo PalmDOC compression (LZ77 with space+letter pairs)."""
    out = bytearray()
    i, n = 0, len(data)
    while i < n:
        c = data[i]
        i += 1
        if c == 0 or 0x09 <= c <= 0x7F:
            out.append(c)
        elif c <= 0x08:
            out += data[i:i + c]
            i += c
        elif c >= 0xC0:
            out.append(0x20)
            out.append(c ^ 0x80)
        else:
            if i >= n:
                break
            pair = (c << 8) | data[i]
            i += 1
            distance = (pair >> 3) & 0x7FF
            length = (pair & 0x07) + 3
            if distance == 0 or distance > len(out):
                continue  # damaged code: skip it rather than fail the book
            start = len(out) - distance
            for k in range(length):
                out.append(out[start + k])
    return bytes(out)


def trailing_size(record: bytes, flags: int) -> int:
    """How many bytes at the end of a text record are extra data, not text."""
    size = len(record)
    total = 0
    bits = flags >> 1
    while bits:
        if bits & 1:
            value, shift, pos = 0, 0, size - total
            while pos > 0:
                byte = record[pos - 1]
                value |= (byte & 0x7F) << shift
                shift += 7
                pos -= 1
                if byte & 0x80 or shift >= 28:
                    break
            total += value
        bits >>= 1
    if flags & 1 and size - total - 1 >= 0:
        total += (record[size - total - 1] & 0x03) + 1
    return min(total, size)


def u16(data: bytes, offset: int) -> int:
    return struct.unpack_from(">H", data, offset)[0] if offset + 2 <= len(data) else 0


def u32(data: bytes, offset: int, default: int = NO_INDEX) -> int:
    return struct.unpack_from(">I", data, offset)[0] if offset + 4 <= len(data) else default


IMAGE_MAGIC = (b"\xff\xd8\xff", b"\x89PNG", b"GIF8", b"BM")


def image_mime(data: bytes) -> str | None:
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data.startswith(b"\x89PNG"):
        return "image/png"
    if data.startswith(b"GIF8"):
        return "image/gif"
    if data.startswith(b"BM"):
        return "image/bmp"
    if data.lstrip().startswith(b"<svg") or data.lstrip().startswith(b"<?xml"):
        return "image/svg+xml"
    return None


class Header:
    """Record 0 of a MOBI 6 book, or of the KF8 half of a book."""

    def __init__(self, record: bytes, start: int) -> None:
        self.start = start                       # record number of this record 0
        self.compression = u16(record, 0)
        self.text_length = u32(record, 4, 0)
        self.text_records = u16(record, 8)
        self.record_size = u16(record, 10)
        self.encryption = u16(record, 12)
        self.has_mobi = record[16:20] == b"MOBI"
        self.length = u32(record, 20, 0) if self.has_mobi else 0
        mobi = self.has_mobi
        self.encoding = ENCODINGS.get(u32(record, 28, 1252) if mobi else 1252, "cp1252")
        self.version = u32(record, 36, 0) if mobi else 0
        self.full_name = ""
        if mobi:
            offset, length = u32(record, 84, 0), u32(record, 88, 0)
            if 0 < length < 1024 and offset + length <= len(record):
                self.full_name = record[offset:offset + length].decode(self.encoding, "replace")
        self.first_image = u32(record, 108) if mobi else NO_INDEX
        self.huff_record = u32(record, 112) if mobi else NO_INDEX
        self.huff_count = u32(record, 116, 0) if mobi else 0
        has_exth = mobi and bool(u32(record, 128, 0) & 0x40)
        long_enough = self.length >= 0xE4
        self.extra_flags = u16(record, 0xF2) if mobi and long_enough else 0
        self.ncx = u32(record, 0xF4) if mobi and long_enough else NO_INDEX
        self.fdst = u32(record, 0xC0) if mobi and self.version >= 8 else NO_INDEX
        self.fragments = u32(record, 0xF8) if mobi and self.length >= 0xEC else NO_INDEX
        self.skeletons = u32(record, 0xFC) if mobi and self.length >= 0xF0 else NO_INDEX
        self.guide = u32(record, 0x104) if mobi and self.length >= 0xF8 else NO_INDEX
        self.exth: dict[int, list[bytes]] = {}
        if has_exth:
            self.exth = read_exth(record, 16 + self.length)

    def index(self, value: int) -> int:
        """A record number from this header, as an absolute record number."""
        return NO_INDEX if value == NO_INDEX else value + self.start


def read_exth(record: bytes, offset: int) -> dict[int, list[bytes]]:
    found: dict[int, list[bytes]] = {}
    if record[offset:offset + 4] != b"EXTH":
        return found
    count = u32(record, offset + 8, 0)
    pos = offset + 12
    for _ in range(min(count, 10000)):
        kind, length = u32(record, pos, 0), u32(record, pos + 4, 0)
        if length < 8 or pos + length > len(record):
            break
        found.setdefault(kind, []).append(record[pos + 8:pos + length])
        pos += length
    return found


class MobiFile:
    """One Kindle book file. Use MobiFile.open(path)."""

    def __init__(self, path: str) -> None:
        self.path = path
        self.lock: str | None = None
        self.title = ""
        self.authors: list[str] = []
        self.exth: dict[int, list[bytes]] = {}
        self.encoding = "cp1252"
        self.is_kf8 = False
        self.records: list[bytes] = []
        self.mobi6: Header | None = None
        self.kf8: Header | None = None
        self._text: bytes | None = None
        self._huff_cache: dict[int, HuffCdic] = {}
        self._parts: list[MobiPart] | None = None
        self._toc: list[tuple[str, int, str, int]] = []
        self._flows: list[bytes] = []
        self._image_base: Header | None = None

    # ---------------------------------------------------------------- opening
    @classmethod
    def open(cls, path: str) -> "MobiFile":
        book = cls(path)
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError as e:
            raise NotABook(f"The file can't be read ({e.strerror or e}).") from e
        book._parse(data)
        return book

    def _parse(self, data: bytes) -> None:
        if data.startswith(b"\xeaDRMION\xee"):
            self.lock = "amazon"
            return
        if data.startswith(b"CONT"):
            raise NotABook("This is Amazon's newest Kindle format (KFX), which this reader can't open.")
        if data.startswith(b"TPZ"):
            self.lock = "amazon"
            return
        if len(data) < 86 or data[60:68] not in (b"BOOKMOBI", b"TEXtREAd"):
            raise NotABook("This isn't a Kindle book file, or it is damaged.")
        count = u16(data, 76)
        table_end = 78 + 8 * count
        if count < 2 or table_end > len(data):
            raise NotABook("This Kindle book is damaged (its list of parts is cut short).")
        offsets = [u32(data, 78 + 8 * i, 0) for i in range(count)] + [len(data)]
        if max(offsets[:-1]) > len(data) and offsets[:-1] == sorted(offsets[:-1]):
            raise NotABook("This Kindle book file is incomplete (it was cut short, probably while "
                           "downloading). Download it again.")
        for a, b in zip(offsets, offsets[1:]):
            if not table_end <= a <= b <= len(data):
                raise NotABook("This Kindle book is damaged (its parts point outside the file).")
        self.records = [data[offsets[i]:offsets[i + 1]] for i in range(count)]
        header = Header(self.records[0], 0)
        self.mobi6 = header
        self.encoding = header.encoding
        self.exth = header.exth
        palm_name = data[:32].split(b"\0", 1)[0].decode("cp1252", "replace")
        self.title = (self._exth_text(503) or header.full_name or palm_name).strip()
        self.authors = [a for a in (self._decode(v).strip() for v in self.exth.get(100, [])) if a]
        if header.encryption in (1, 2):
            self.lock = "amazon"
            return
        if header.version >= 8:
            self.kf8, self.is_kf8 = header, True
        else:
            boundary = self.exth.get(121)
            if boundary:
                start = u32(boundary[0], 0)
                if 0 < start < len(self.records) and self.records[start][16:20] == b"MOBI":
                    kf8 = Header(self.records[start], start)
                    if kf8.version >= 8 and kf8.encryption == 0:
                        self.kf8, self.is_kf8 = kf8, True

    def _decode(self, value: bytes) -> str:
        return value.decode(self.encoding, "replace")

    def _exth_text(self, kind: int) -> str:
        values = self.exth.get(kind)
        return self._decode(values[0]) if values else ""

    # ---------------------------------------------------------------- text
    def text(self, header: Header | None = None) -> bytes:
        """The whole book text, decompressed (KF8 half when there is one)."""
        if self.lock:
            raise NotABook("This book is locked.")
        header = header or self.kf8 or self.mobi6
        if header is self.kf8 or header is self.mobi6:
            if self._text is not None and header is (self.kf8 or self.mobi6):
                return self._text
        parts = []
        last = min(header.start + header.text_records, len(self.records) - 1)
        for number in range(header.start + 1, last + 1):
            record = self.records[number]
            record = record[:len(record) - trailing_size(record, header.extra_flags)]
            parts.append(self._decompress(record, header))
        text = b"".join(parts)
        if header.text_length and len(text) > header.text_length:
            text = text[:header.text_length]
        if header is (self.kf8 or self.mobi6):
            self._text = text
        return text

    def _decompress(self, record: bytes, header: Header) -> bytes:
        if header.compression == 1:
            return record
        if header.compression == 2:
            return palmdoc_decompress(record)
        if header.compression == 17480:
            return self._huffcdic(header).unpack(record)
        raise NotABook("This Kindle book uses a kind of compression this reader doesn't know.")

    def _huffcdic(self, header: Header) -> "HuffCdic":
        cached = self._huff_cache.get(header.start)
        if cached is None:
            first = header.index(header.huff_record)
            if first == NO_INDEX or header.huff_count < 2 or first + header.huff_count > len(self.records):
                raise NotABook("This Kindle book is damaged (its compression tables are missing).")
            tables = self.records[first:first + header.huff_count]
            cached = HuffCdic(tables[0], tables[1:])
            self._huff_cache[header.start] = cached
        return cached

    # ---------------------------------------------------------------- pictures
    def image(self, n: int, header: Header | None = None) -> bytes | None:
        """Picture number n (1-based, as in recindex="00001")."""
        header = header or self.mobi6
        if header is None or header.first_image == NO_INDEX or n < 1:
            return None
        number = header.first_image + n - 1
        if not 0 < number < len(self.records):
            return None
        data = self.records[number]
        return data if image_mime(data) else None

    def cover(self) -> bytes | None:
        for kind in (201, 202):
            values = self.exth.get(kind)
            if values and len(values[0]) >= 4:
                offset = u32(values[0], 0)
                if offset != NO_INDEX:
                    picture = self.image(offset + 1)
                    if picture:
                        return picture
        return None

    # ---------------------------------------------------------------- parts
    def book_parts(self) -> list["MobiPart"]:
        """The book as separate HTML documents (UTF-8), in reading order."""
        if self.lock:
            raise NotABook("This book is locked.")
        if self._parts is None:
            parts = None
            if self.kf8 is not None:
                try:
                    parts = self._kf8_parts(self.kf8)
                except (NotABook, ValueError, IndexError, KeyError, struct.error, zlib.error):
                    if self.kf8 is self.mobi6:
                        raise NotABook("This Kindle book is damaged (its chapters can't be put together).")
                    parts = None
            if parts is None:
                self._toc, self._flows, self._image_base = [], [], None
                self.is_kf8 = False
                parts = self._mobi6_parts(self.mobi6)
            self._parts = parts
        return self._parts

    def toc(self) -> list[tuple[str, int, str, int]]:
        """(title, part number, anchor id or '', depth) from the book's own contents index."""
        self.book_parts()
        return list(self._toc)

    def resource(self, name: str) -> tuple[bytes, str] | None:
        """A picture ('img/00001.jpg'), font ('img/00003.ttf') or style sheet ('flow/0001.css')."""
        if self.lock:
            return None
        self.book_parts()
        m = re.fullmatch(r"img/(\d{5})\.\w+", name)
        if m:
            data = self._resource_record(int(m.group(1)))
            if data is None:
                return None
            if data.startswith(b"FONT"):
                font = font_data(data)
                return (font, font_mime(font)) if font else None
            mime = image_mime(data)
            return (data, mime) if mime else None
        m = re.fullmatch(r"flow/(\d{4})\.(css|svg)", name)
        if m:
            n = int(m.group(1))
            if 0 < n < len(self._flows):
                return self._flows[n], "text/css" if m.group(2) == "css" else "image/svg+xml"
        return None

    def _resource_record(self, n: int) -> bytes | None:
        header = self._image_base or self.mobi6
        if header is None or header.first_image == NO_INDEX or n < 1:
            return None
        number = header.first_image + n - 1
        return self.records[number] if 0 < number < len(self.records) else None

    def _resource_name(self, n: int) -> str:
        data = self._resource_record(n) or b""
        if data.startswith(b"FONT"):
            font = font_data(data) or b""
            ext = {"font/otf": "otf", "font/woff": "woff"}.get(font_mime(font), "ttf")
        else:
            ext = MIME_EXT.get(image_mime(data) or "", "jpg")
        return "img/%05d.%s" % (n, ext)

    # ---------------------------------------------------------------- MOBI 6
    def _mobi6_parts(self, header: Header) -> list["MobiPart"]:
        raw = self.text(header)
        ncx = self._ncx_entries(header)
        targets = {int(v) for v in FILEPOS.findall(raw)} | {pos for _, pos, _ in ncx if pos is not None}
        raw = insert_anchors(raw, sorted(t for t in targets if 0 <= t <= len(raw)))
        guide_toc = None
        for ref in REFERENCE.findall(raw):
            if TOC_TYPE.search(ref):
                m = FILEPOS.search(ref)
                if m:
                    guide_toc = int(m.group(1))
        raw = GUIDE.sub(b"", raw)
        head, body = split_head(raw)
        head = META_CHARSET.sub(b"", head)
        chunks = merge_empty(PAGEBREAK.split(body))
        where: dict[int, int] = {}
        for number, chunk in enumerate(chunks):
            for fp in ANCHOR_FP.findall(chunk):
                where[int(fp)] = number

        def link(m: re.Match) -> bytes:
            fp = int(m.group(1))
            if fp in where:
                return b'href="part%04d.html#fp%d"' % (where[fp], fp)
            return b'href="#"'

        head_text = head.decode(self.encoding, "replace")
        parts = []
        for number, chunk in enumerate(chunks):
            chunk = MBP_TAG.sub(b"", chunk)
            chunk = FILEPOS_ATTR.sub(link, chunk)
            chunk = IMG_TAG.sub(self._mobi6_image, chunk)
            html = ('<!DOCTYPE html><html><head><meta charset="utf-8">' + head_text
                    + "</head><body>" + chunk.decode(self.encoding, "replace") + "</body></html>")
            parts.append(MobiPart("part%04d.html" % number, html.encode("utf-8")))
        toc = [(label, where[pos], "fp%d" % pos, depth) for label, pos, depth in ncx if pos in where]
        if not toc and guide_toc in where:
            toc = links_in(parts[where[guide_toc]].html, after=b'id="fp%d"' % guide_toc)
        self._toc = toc
        return parts

    def _mobi6_image(self, m: re.Match) -> bytes:
        tag = m.group(0)
        found = None
        for pattern in (RECINDEX, HIRECINDEX, LOWRECINDEX):
            a = pattern.search(tag)
            if a:
                found = int(a.group(1))
                break
        if found is None:
            return tag
        tag = ANY_RECINDEX.sub(b"", tag)
        tag = SRC_ATTR.sub(b"", tag)
        return tag[:4] + b' src="' + self._resource_name(found).encode() + b'"' + tag[4:]

    def _ncx_entries(self, header: Header) -> list[tuple[str, int | None, int]]:
        """(label, position, depth) from a MOBI 6 contents index."""
        first = header.index(header.ncx)
        if first == NO_INDEX or first >= len(self.records):
            return []
        try:
            entries, cncx = read_index(self.records, first)
        except (NotABook, IndexError, struct.error):
            return []
        found = []
        for _name, tags in entries:
            label = cncx.get(tags.get(3, [-1])[0])
            pos = tags.get(1, [None])[0]
            if label:
                found.append((label.decode(self.encoding, "replace").strip(), pos, tags.get(4, [0])[0]))
        return found

    # ---------------------------------------------------------------- KF8
    def _kf8_parts(self, header: Header) -> list["MobiPart"]:
        raw = self.text(header)
        # Pictures are shared: in combined files the KF8 half numbers them from
        # the MOBI 6 half's first picture (its own field points 12 or so further on).
        self._image_base = self.mobi6
        flows = [raw]
        fdst = header.index(header.fdst)
        if fdst != NO_INDEX and fdst < len(self.records) and self.records[fdst][:4] == b"FDST":
            record = self.records[fdst]
            count = u32(record, 8, 0)
            pairs = [struct.unpack_from(">II", record, 12 + 8 * i) for i in range(count)]
            if pairs:
                flows = [raw[a:b] for a, b in pairs]
        skeletons, _ = read_index(self.records, header.index(header.skeletons))
        fragments_raw, _ = read_index(self.records, header.index(header.fragments))
        fragments = [(int(name), tags[6][0], tags[6][1]) for name, tags in fragments_raw]
        flow0 = flows[0]
        assembled, starts = [], []
        f = 0
        for _name, tags in skeletons:
            count, start, length = tags[1][0], tags[6][0], tags[6][1]
            html = bytearray(flow0[start:start + length])
            cursor = start + length
            for _ in range(count):
                if f >= len(fragments):
                    break
                insert, _fstart, flen = fragments[f]
                at = min(max(insert - start, 0), len(html))
                html[at:at] = flow0[cursor:cursor + flen]
                cursor += flen
                f += 1
            assembled.append(bytes(html))
            starts.append(start)
        if not assembled:
            raise NotABook("This Kindle book has no chapters.")
        anchors = [anchor_table(p) for p in assembled]
        wanted: list[set[bytes]] = [set() for _ in assembled]

        def target(pos: int) -> tuple[int, str]:
            k = max(bisect.bisect_right(starts, pos) - 1, 0)
            anchor = nearest_anchor(assembled[k], anchors[k], pos - starts[k])
            if anchor:
                if anchor[1]:
                    wanted[k].add(anchor[1])
                return k, anchor[0]
            return k, ""

        def by_fid(fid: int, off: int) -> tuple[int, str]:
            return target(fragments[fid][0] + off if fid < len(fragments) else 0)

        def link(m: re.Match) -> bytes:
            k, anchor = by_fid(int(m.group(1), 32), int(m.group(2), 32))
            return b"part%04d.html" % k + (b"#" + anchor.encode() if anchor else b"")

        def embed(m: re.Match) -> bytes:
            return self._resource_name(int(m.group(1), 32)).encode()

        def flow(m: re.Match) -> bytes:
            kind = b"svg" if b"svg" in (m.group(2) or b"") else b"css"
            return b"flow/%04d.%s" % (int(m.group(1), 32), kind)

        toc = []
        first = header.index(header.ncx)
        if first != NO_INDEX and first < len(self.records):
            try:
                entries, cncx = read_index(self.records, first)
            except (NotABook, IndexError, struct.error):
                entries, cncx = [], {}
            for _name, tags in entries:
                label = cncx.get(tags.get(3, [-1])[0])
                if not label:
                    continue
                if 6 in tags and len(tags[6]) >= 2:
                    k, anchor = by_fid(tags[6][0], tags[6][1])
                elif 1 in tags:
                    k, anchor = target(tags[1][0])
                else:
                    continue
                toc.append((label.decode("utf-8", "replace").strip(), k, anchor, tags.get(4, [0])[0]))
        parts = []
        for k, html in enumerate(assembled):
            html = KINDLE_POS.sub(link, html)
            html = KINDLE_EMBED.sub(embed, html)
            html = KINDLE_FLOW.sub(flow, html)
            parts.append(html)
        for k in range(len(parts)):
            if wanted[k]:
                parts[k] = add_ids(parts[k], wanted[k])
        self._flows = [flows[0]] + [KINDLE_EMBED.sub(lambda m: b"../" + embed(m), fl) for fl in flows[1:]]
        self._toc = toc
        return [MobiPart("part%04d.html" % k, p) for k, p in enumerate(parts)]


@dataclass
class MobiPart:
    name: str
    html: bytes


MIME_EXT = {"image/jpeg": "jpg", "image/png": "png", "image/gif": "gif", "image/bmp": "bmp",
            "image/svg+xml": "svg"}
FILEPOS = re.compile(rb"filepos\s*=\s*[\"']?0*(\d+)", re.I)
FILEPOS_ATTR = re.compile(rb"filepos\s*=\s*[\"']?0*(\d+)[\"']?", re.I)
REFERENCE = re.compile(rb"<reference\b[^>]*>", re.I)
TOC_TYPE = re.compile(rb"type\s*=\s*[\"']?toc\b", re.I)
GUIDE = re.compile(rb"<guide\b.*?</guide\s*>", re.S | re.I)
META_CHARSET = re.compile(rb"<meta\b[^>]*(charset|http-equiv)[^>]*>", re.I)
PAGEBREAK = re.compile(rb"<mbp:pagebreak\b[^>]*>", re.I)
ANCHOR_FP = re.compile(rb'<a id="fp(\d+)"></a>')
MBP_TAG = re.compile(rb"</?mbp:[^>]*>", re.I)
IMG_TAG = re.compile(rb"<img\b[^>]*>", re.I)
RECINDEX = re.compile(rb"\srecindex\s*=\s*[\"']?(\d+)", re.I)
HIRECINDEX = re.compile(rb"\shirecindex\s*=\s*[\"']?(\d+)", re.I)
LOWRECINDEX = re.compile(rb"\slowrecindex\s*=\s*[\"']?(\d+)", re.I)
ANY_RECINDEX = re.compile(rb"\s(?:hi|low)?recindex\s*=\s*[\"']?\d+[\"']?", re.I)
SRC_ATTR = re.compile(rb"\ssrc\s*=\s*(\"[^\"]*\"|'[^']*'|[^\s>]+)", re.I)
KINDLE_POS = re.compile(rb"kindle:pos:fid:([0-9A-Va-v]+):off:([0-9A-Va-v]+)")
KINDLE_EMBED = re.compile(rb"kindle:embed:([0-9A-Va-v]+)(?:\?mime=[\w/+.-]*)?")
KINDLE_FLOW = re.compile(rb"kindle:flow:([0-9A-Va-v]+)(\?mime=[\w/+.-]*)?")
TAG_WITH_ID = re.compile(rb"<[^<>]*?\s(id|name|aid)\s*=\s*[\"']([^\"'<>]+)[\"'][^<>]*>", re.I)
HAS_ID = re.compile(rb"\sid\s*=", re.I)
TAG_TEXT = re.compile(rb"<[^>]*>")
BODY_TAG = re.compile(rb"<body\b[^>]*>", re.I)


def insert_anchors(raw: bytes, positions: list[int]) -> bytes:
    """Put <a id="fpN"></a> at each byte position, moved before a tag it would split."""
    pieces, last = [], 0
    for pos in positions:
        at = pos
        lt, gt = raw.rfind(b"<", 0, pos), raw.rfind(b">", 0, pos)
        if lt > gt:
            at = lt
        at = max(at, last)
        pieces.append(raw[last:at])
        pieces.append(b'<a id="fp%d"></a>' % pos)
        last = at
    pieces.append(raw[last:])
    return b"".join(pieces)


def split_head(raw: bytes) -> tuple[bytes, bytes]:
    """(what is inside <head>, what is inside <body>)."""
    head = b""
    m = re.search(rb"<head\b[^>]*>(.*?)</head\s*>", raw, re.S | re.I)
    if m:
        head = m.group(1)
    b = BODY_TAG.search(raw)
    body = raw[b.end():] if b else raw[m.end():] if m else raw
    body = re.sub(rb"</body\s*>\s*(</html\s*>)?\s*$", b"", body, flags=re.I)
    return head, body


def merge_empty(chunks: list[bytes]) -> list[bytes]:
    """Join chunks with nothing to see onto the next one, so no part is blank."""
    out: list[bytes] = []
    carry = b""
    for chunk in chunks:
        chunk = carry + chunk
        visible = re.sub(rb"<[^>]*>|&nbsp;|\s", b"", chunk)
        if visible or re.search(rb"<(img|svg|hr|table)\b", chunk, re.I):
            out.append(chunk)
            carry = b""
        else:
            carry = chunk
    if carry:
        if out:
            out[-1] += carry
        else:
            out.append(carry)
    return out or [b""]


def links_in(html: bytes, after: bytes) -> list[tuple[str, int, str, int]]:
    """Contents entries from the links on a contents page, after a marker."""
    start = max(html.find(after), 0)
    found = []
    for m in re.finditer(rb'<a\b[^>]*href="part(\d{4})\.html#([^"]+)"[^>]*>(.*?)</a>', html[start:], re.S | re.I):
        label = " ".join(TAG_TEXT.sub(b"", m.group(3)).decode("utf-8", "replace").split())
        if label:
            found.append((label, int(m.group(1)), m.group(2).decode(), 0))
    return found


def anchor_table(html: bytes) -> list[tuple[int, int, str, bytes]]:
    """(tag start, tag end, kind, value) of tags with id/name/aid, plus <body> as a stop."""
    table = []
    for m in TAG_WITH_ID.finditer(html):
        tag = m.group(0)
        kind = m.group(1).lower().decode()
        idm = re.search(rb"\sid\s*=\s*[\"']([^\"'<>]+)", tag, re.I)
        if idm:                                   # a tag with both aid and id: use the id
            table.append((m.start(), m.end(), "id", idm.group(1)))
        else:
            table.append((m.start(), m.end(), kind, m.group(2)))
    for m in BODY_TAG.finditer(html):
        table.append((m.start(), m.end(), "body", b""))
    table.sort()
    return table


def nearest_anchor(html: bytes, table, local: int) -> tuple[str, bytes] | None:
    """(id to link to, aid/name value that needs an id or b'') for a position."""
    lt, gt = html.rfind(b"<", 0, local), html.rfind(b">", 0, local)
    if lt > gt:                       # inside a tag: count that tag as before the position
        end = html.find(b">", local)
        local = end + 1 if end >= 0 else local
    i = bisect.bisect_right([row[0] for row in table], local - 1) - 1
    while i >= 0:
        _start, _end, kind, value = table[i]
        if kind == "body":
            return None
        text = value.decode("utf-8", "replace")
        if kind == "id":
            return text, b""
        if kind == "aid":
            return "aid-" + text, value
        return text, value
    return None


def add_ids(html: bytes, values: set[bytes]) -> bytes:
    """Give tags with a wanted aid="x" an id="aid-x" (and name="x" an id="x")."""
    def fix(m: re.Match) -> bytes:
        tag = m.group(0)
        if HAS_ID.search(tag):
            return tag
        kind, value = m.group(1).lower(), m.group(2)
        if value not in values:
            return tag
        new_id = (b"aid-" + value) if kind == b"aid" else value
        end = len(tag) - (2 if tag.endswith(b"/>") else 1)
        return tag[:end] + b' id="' + new_id + b'"' + tag[end:]
    return TAG_WITH_ID.sub(fix, html)


def font_data(record: bytes) -> bytes | None:
    """An embedded font from a FONT record (zlib-packed, its start scrambled with a key)."""
    if record[:4] != b"FONT" or len(record) < 24:
        return None
    _size, flags, start, key_len, key_start = struct.unpack_from(">IIIII", record, 4)
    data = bytearray(record[start:])
    if flags & 2 and key_len:
        key = record[key_start:key_start + key_len]
        for i in range(min(1040, len(data))):
            data[i] ^= key[i % key_len]
    if flags & 1:
        try:
            data = bytearray(zlib.decompress(bytes(data)))
        except zlib.error:
            return None
    return bytes(data)


def font_mime(font: bytes) -> str:
    if font.startswith(b"OTTO"):
        return "font/otf"
    if font.startswith(b"wOFF"):
        return "font/woff"
    return "font/ttf"


def forward_varint(data: bytes, pos: int) -> tuple[int, int]:
    """(value, bytes used) of a forward variable-width number (its last byte has 0x80)."""
    value, used = 0, 0
    while pos + used < len(data):
        byte = data[pos + used]
        used += 1
        value = (value << 7) | (byte & 0x7F)
        if byte & 0x80 or used >= 5:
            break
    return value, used


def read_index(records: list[bytes], first: int) -> tuple[list[tuple[bytes, dict[int, list[int]]]], dict[int, bytes]]:
    """Read an INDX index: its entries (name, {tag: values}) and its CNCX strings."""
    if first == NO_INDEX or first >= len(records) or records[first][:4] != b"INDX":
        raise NotABook("A Kindle index is missing.")
    primary = records[first]
    header_length = u32(primary, 4, 0)
    index_records = u32(primary, 24, 0)
    cncx_records = u32(primary, 52, 0)
    tags: list[tuple[int, int, int, int]] = []
    control_bytes = 0
    if primary[header_length:header_length + 4] == b"TAGX":
        tagx_length = u32(primary, header_length + 4, 0)
        control_bytes = u32(primary, header_length + 8, 0)
        for i in range(12, tagx_length, 4):
            tags.append(tuple(primary[header_length + i:header_length + i + 4]))
    entries = []
    for r in range(first + 1, min(first + 1 + index_records, len(records))):
        record = records[r]
        if record[:4] != b"INDX":
            break
        idxt, count = u32(record, 20, 0), u32(record, 24, 0)
        if record[idxt:idxt + 4] != b"IDXT":
            break
        offsets = [u16(record, idxt + 4 + 2 * j) for j in range(count)] + [idxt]
        for j in range(count):
            start, end = offsets[j], offsets[j + 1]
            name_length = record[start]
            name = record[start + 1:start + 1 + name_length]
            entries.append((name, read_tags(record, start + 1 + name_length, end, control_bytes, tags)))
    cncx: dict[int, bytes] = {}
    for k in range(cncx_records):
        r = first + 1 + index_records + k
        if r >= len(records):
            break
        record, pos = records[r], 0
        while pos < len(record):
            length, used = forward_varint(record, pos)
            if length == 0 and not any(record[pos:pos + 4]):
                break
            cncx[pos + k * 0x10000] = record[pos + used:pos + used + length]
            pos += used + length
    return entries, cncx


def read_tags(record: bytes, start: int, end: int, control_bytes: int,
              tags: list[tuple[int, int, int, int]]) -> dict[int, list[int]]:
    """The tag values of one index entry, decoded with its control bytes."""
    pos = start + control_bytes
    control = 0
    pending = []
    for tag, per_entry, mask, end_flag in tags:
        if end_flag == 1:
            control += 1
            continue
        if start + control >= len(record):
            break
        value = record[start + control] & mask
        if not value:
            continue
        if value == mask and bin(mask).count("1") > 1:
            size, used = forward_varint(record, pos)
            pos += used
            pending.append((tag, None, size, per_entry))
        else:
            while not mask & 1:
                mask >>= 1
                value >>= 1
            pending.append((tag, value, None, per_entry))
    result: dict[int, list[int]] = {}
    for tag, count, size, per_entry in pending:
        values = []
        if count is not None:
            for _ in range(count * per_entry):
                v, used = forward_varint(record, pos)
                pos += used
                values.append(v)
        else:
            consumed = 0
            while consumed < size and pos < end:
                v, used = forward_varint(record, pos)
                pos += used
                consumed += used
                values.append(v)
        result[tag] = values
    return result


class HuffCdic:
    """Amazon's HUFF/CDIC text compression: Huffman codes standing for dictionary phrases."""

    def __init__(self, huff: bytes, cdics: list[bytes]) -> None:
        if huff[:4] != b"HUFF" or len(huff) < 24:
            raise NotABook("This Kindle book is damaged (bad compression table).")
        cache_at, base_at = struct.unpack_from(">II", huff, 8)
        self.cache = []
        for v in struct.unpack_from(">256I", huff, cache_at):
            length, terminal, top = v & 0x1F, bool(v & 0x80), v >> 8
            if length == 0:
                raise NotABook("This Kindle book is damaged (bad compression table).")
            self.cache.append((length, terminal, ((top + 1) << (32 - length)) - 1))
        base = struct.unpack_from(">64I", huff, base_at)
        self.min_code = [0] * 33
        self.max_code = [0] * 33
        for length in range(1, 33):
            self.min_code[length] = base[2 * (length - 1)] << (32 - length)
            self.max_code[length] = ((base[2 * (length - 1) + 1] + 1) << (32 - length)) - 1
        self.phrases: list = []
        for cdic in cdics:
            if cdic[:4] != b"CDIC":
                raise NotABook("This Kindle book is damaged (bad phrase table).")
            total, bits = struct.unpack_from(">II", cdic, 8)
            count = min(1 << bits, total - len(self.phrases))
            for off in struct.unpack_from(">%dH" % count, cdic, 16):
                blen = u16(cdic, 16 + off)
                self.phrases.append([cdic[18 + off:18 + off + (blen & 0x7FFF)], bool(blen & 0x8000)])

    def unpack(self, data: bytes, depth: int = 0) -> bytes:
        if depth > 32:
            raise NotABook("This Kindle book is damaged (a phrase refers to itself).")
        bits_left = len(data) * 8
        data = data + b"\0" * 8
        pos, shift = 0, 32
        window = struct.unpack_from(">Q", data, 0)[0]
        out = bytearray()
        cache, min_code, max_code, phrases = self.cache, self.min_code, self.max_code, self.phrases
        while True:
            if shift <= 0:
                pos += 4
                window = struct.unpack_from(">Q", data, pos)[0]
                shift += 32
            code = (window >> shift) & 0xFFFFFFFF
            length, terminal, top = cache[code >> 24]
            if not terminal:
                while length < 32 and code < min_code[length]:
                    length += 1
                top = max_code[length]
            shift -= length
            bits_left -= length
            if bits_left < 0:
                break
            index = (top - code) >> (32 - length)
            if index >= len(phrases):
                raise NotABook("This Kindle book is damaged (unknown phrase).")
            phrase = phrases[index]
            if not phrase[1]:
                phrase[0] = self.unpack(phrase[0], depth + 1)
                phrase[1] = True
            out += phrase[0]
        return bytes(out)
