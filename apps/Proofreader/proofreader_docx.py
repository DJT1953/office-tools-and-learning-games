"""
proofreader_docx.py

Corrects a .docx file without Microsoft Word, for computers that don't have
it. The file is opened directly (python-docx) and each correction is written
into the piece of text it belongs to, so fonts, bold and layout stay as they
were.

Covers the main text, tables, text boxes, headers and footers. Footnotes and
endnotes are not reached this way, and the results of fields (page numbers,
contents lists) are left alone.
"""

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

from proofreader_engine import Skip

W_P, W_T, W_TAB, W_BR = qn("w:p"), qn("w:t"), qn("w:tab"), qn("w:br")
W_FLDCHAR, W_FLDSIMPLE, W_TYPE = qn("w:fldChar"), qn("w:fldSimple"), qn("w:fldCharType")
W_INS, W_DEL = qn("w:ins"), qn("w:del")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class _Piece:
    """A stretch of a paragraph's text and the element that holds it."""

    def __init__(self, element, text, locked):
        self.element = element
        self.text = text
        self.locked = locked


def _pieces(paragraph):
    """The text of one paragraph, in order, as the pieces Word stores it in."""
    pieces, in_field = [], 0
    for element in paragraph.iter():
        tag = element.tag
        if tag == W_FLDCHAR:
            kind = element.get(W_TYPE)
            in_field += 1 if kind == "begin" else -1 if kind == "end" else 0
            continue
        if tag not in (W_T, W_TAB, W_BR):
            continue
        if next(element.iterancestors(W_P)) is not paragraph:
            continue                                    # belongs to a text box inside
        if tag == W_T:
            locked = in_field > 0 or next(element.iterancestors(W_FLDSIMPLE), None) is not None
            pieces.append(_Piece(element, element.text or "", locked))
        else:
            pieces.append(_Piece(None, "\t" if tag == W_TAB else "\n", True))
    return pieces


class DocxDocument:
    def __init__(self, path):
        self.path = path
        try:
            self.doc = Document(path)
        except Exception as error:
            raise Skip(f"it could not be opened ({type(error).__name__})")
        self.roots = [self.doc.element.body]
        for rel in self.doc.part.rels.values():
            if rel.reltype in (RT.HEADER, RT.FOOTER) and not rel.is_external:
                self.roots.append(rel.target_part.element)
        for root in self.roots:
            if next(root.iter(W_INS, W_DEL), None) is not None:
                raise Skip("it has tracked changes waiting; accept or reject them in Word first")
        self.paragraphs = []

    def read(self):
        self.paragraphs = [p for root in self.roots for p in root.iter(W_P)]
        return ["".join(piece.text for piece in _pieces(p)) for p in self.paragraphs]

    def apply(self, edits):
        """Write corrections in, last first. Returns the ones that went in."""
        done = []
        for edit in sorted(edits, key=lambda e: (e.para, e.start), reverse=True):
            if self._write(_pieces(self.paragraphs[edit.para]), edit):
                done.append(edit)
        done.reverse()
        return done

    @staticmethod
    def _write(pieces, edit):
        touched, at = [], 0                 # (piece, from, to) inside each piece
        for piece in pieces:
            end = at + len(piece.text)
            if edit.start == edit.end:
                # new text joins the word before it, so it takes that word's look
                inside = at < edit.start <= end or (edit.start == 0 and at == 0 and not touched)
            else:
                inside = at < edit.end and edit.start < end
            if inside and piece.text:
                touched.append((piece, max(edit.start, at) - at, min(edit.end, end) - at))
                if edit.start == edit.end:
                    break
            at = end
        if not touched or any(piece.locked for piece, _, _ in touched):
            return False
        for index, (piece, start, end) in enumerate(touched):
            new = edit.new if index == 0 else ""
            piece.element.text = piece.text[:start] + new + piece.text[end:]
            piece.element.set(XML_SPACE, "preserve")
        return True

    def grammar_notes(self):
        return []

    def save(self):
        self.doc.save(self.path)

    def close(self):
        pass
