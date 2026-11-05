"""
combine_docx.py

Pick a folder in a Windows Explorer dialog. Every .docx in it is joined,
in natural filename order (page 2 before page 10), into one Word file:

    <FolderName>_combined.docx   (saved in the same folder)

A horizontal divider line is placed before each new file's content.
Each file keeps its own fonts, spacing, page size, orientation and margins, as
well as its pictures, tables, lists, footnotes, endnotes and comments. A file
with a different page size or orientation starts on a new page.

Headers and footers come from the first file only; the closing message says if
a later file had its own. Page numbers run on through the whole document.

If the combined file already exists you are asked before it is replaced.

Requirements (one-time):
    python-docx (1.2 or newer) and docxcompose. If the Python running this
    lacks them, it offers to install them.
"""

import importlib.util
import os
import re
import site
import subprocess
import sys
import traceback
import zipfile
from copy import deepcopy
from io import BytesIO

import tkinter as tk
from tkinter import filedialog, messagebox

TITLE = "Combine Word files"
REQUIRED = {"docx": "python-docx", "docxcompose": "docxcompose"}


def dialog_root():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)        # keep dialogs in front of other windows
    return root


def show_crash(kind, value, tb):
    """A double-clicked script's window closes the moment it fails, taking the
    error with it, so show anything unexpected in a dialog."""
    text = "".join(traceback.format_exception(kind, value, tb))
    print(text)
    dialog_root()
    messagebox.showerror(TITLE, "Something went wrong:\n\n" + text[-1500:])


def install_missing():
    """Double-clicking can start a different Python from the one pip installed
    into, so offer to add whatever this one lacks."""
    missing = [pip for module, pip in REQUIRED.items()
               if importlib.util.find_spec(module) is None]
    if not missing:
        return
    root = dialog_root()
    if not messagebox.askyesno(
            TITLE,
            f"This copy of Python ({sys.version.split()[0]}) still needs: "
            f"{', '.join(missing)}.\n\n"
            "Download and install now? This only happens once and takes about a minute."):
        sys.exit()
    root.destroy()

    print("Installing", ", ".join(missing), "...", flush=True)
    result = subprocess.run([sys.executable, "-m", "pip", "install"] + missing,
                            stderr=subprocess.PIPE, text=True)
    if result.returncode:
        raise RuntimeError("Could not install " + ", ".join(missing) + ":\n" + result.stderr[-800:])
    site.addsitedir(site.getusersitepackages())   # a first install may have just created it
    importlib.invalidate_caches()


sys.excepthook = show_crash
install_missing()

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt
from docxcompose.composer import Composer
from lxml import etree

# ---------------------------------------------------------------- settings
DIVIDER_BEFORE_FIRST = False    # True = also put a line above the first file
SHOW_FILENAME = False           # True = print the file's name under each line
NEW_PAGE_FOR_EACH_FILE = False  # True = start each file on a new page as well

# ---------------------------------------------------------------- word's markup
W_ID = qn("w:id")
W_VAL = qn("w:val")
W_NS = W_ID[:W_ID.index("}") + 1]
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
THEME_FONT_ATTRS = {qn("w:asciiTheme"): qn("w:ascii"), qn("w:hAnsiTheme"): qn("w:hAnsi"),
                    qn("w:eastAsiaTheme"): qn("w:eastAsia"), qn("w:cstheme"): qn("w:cs")}

# what Word assumes for anything a file doesn't set
BASE_LOOK = parse_xml(
    f"<w:style {nsdecls('w')}>"
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
    ' w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
    '<w:kern w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>')

# the order Word wants these settings in
PPR_ORDER = ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers "
             "pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE "
             "autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents "
             "suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle").split()
RPR_ORDER = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint "
             "noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u "
             "effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath").split()

# parts of a table style that don't affect how it looks
TABLE_STYLE_LABELS = {qn(t) for t in ("w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:rsid",
                                      "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed",
                                      "w:qFormat", "w:locked")}

NOTE_KINDS = (
    # relationship, content type, part name, marks in the text
    (RT.FOOTNOTES, CT.WML_FOOTNOTES, "/word/footnotes.xml", ("w:footnoteReference",)),
    (RT.ENDNOTES, CT.WML_ENDNOTES, "/word/endnotes.xml", ("w:endnoteReference",)),
    (RT.COMMENTS, CT.WML_COMMENTS, "/word/comments.xml",
     ("w:commentReference", "w:commentRangeStart", "w:commentRangeEnd")),
)


def natural_key(name):
    return [int(t) if t.isdecimal() else t.lower() for t in re.split(r"(\d+)", name)]


class FileComposer(Composer):
    """docxcompose, with the jobs that the functions below do instead switched off."""

    def add_footnotes(self, doc, element):
        pass    # see carry_notes

    def add_styles_from_other_parts(self, doc):
        pass    # see carry_notes

    def fix_section_types(self, doc):
        pass    # see open_section

    def fix_header_and_footers(self, doc):
        pass    # see open_section

    def save(self, filename):
        # SmartArt from two files arrives with the same internal part names, and
        # Word calls a file with two parts of one name corrupted
        seen = set()
        for part in self.pkg.iter_parts():
            name, base, n = part.partname, part.partname[:-len(part.partname.ext) - 1], 1
            while name.lower() in seen:
                n += 1
                name = PackURI(f"{base}_{n}.{part.partname.ext}")
            part.partname = name
            seen.add(name.lower())
        super().save(filename)


# ---------------------------------------------------------------- each file's look
def theme_of(doc):
    """A file's theme, and the real fonts behind its theme font names."""
    try:
        blob = doc.part.part_related_by(RT.THEME).blob
    except KeyError:
        return None, {}
    fonts = {}
    root = etree.fromstring(blob)
    for group in ("major", "minor"):
        for name, tag in (("Ascii", "latin"), ("HAnsi", "latin"), ("EastAsia", "ea"), ("Bidi", "cs")):
            face = root.find(f".//{A_NS}{group}Font/{A_NS}{tag}")
            if face is not None:
                fonts[group + name] = face.get("typeface")
    return blob, fonts


def spell_out_theme(root, fonts):
    """Theme fonts and colors follow whichever theme the combined file has, so
    swap them for what they mean in this file. Word stores the plain color
    beside the theme color, so that part only needs removing."""
    for el in root.iter():
        for attr in list(el.attrib):
            if attr in THEME_FONT_ATTRS:
                face = fonts.get(el.attrib.pop(attr))
                if face:
                    el.set(THEME_FONT_ATTRS[attr], face)
            elif attr.startswith(W_NS + "theme"):
                del el.attrib[attr]


def style_chain(owner, style):
    """A style and the styles it is based on, starting from the bottom."""
    chain = []
    while style is not None and style not in chain:
        chain.insert(0, style)
        based_on = style.basedOn_val
        style = owner.styles.element.get_by_id(based_on) if based_on else None
    return chain


def overlay(look, layer, fonts):
    """Lay one layer of settings over the ones beneath it."""
    if layer is None:
        return
    for kind, props in look.items():
        found = layer.find(qn(kind))
        if found is None:
            continue
        for prop in found:
            prop = deepcopy(prop)
            spell_out_theme(prop, fonts)
            prop.attrib.pop(qn("w:hint"), None)
            old = props.get(prop.tag)
            if old is not None and not len(prop):
                # spacing, indents and fonts inherit one attribute at a time
                for attr, value in old.attrib.items():
                    if attr not in prop.attrib:
                        prop.set(attr, value)
            props[prop.tag] = prop


def paragraph_look(owner, fonts, style, table_layers=()):
    """Everything a paragraph style ends up meaning, in the order Word applies it."""
    look = {"w:pPr": {}, "w:rPr": {}}
    overlay(look, BASE_LOOK, {})
    defaults = owner.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        overlay(look, defaults.find(qn("w:pPrDefault")), fonts)
        overlay(look, defaults.find(qn("w:rPrDefault")), fonts)
    for layer in table_layers:
        overlay(look, layer, fonts)
    for layer in style_chain(owner, style):
        overlay(look, layer, fonts)
    return look


def run_look(owner, fonts, style):
    look = {"w:pPr": {}, "w:rPr": {}}
    for layer in style_chain(owner, style):
        overlay(look, layer, fonts)
    return look


def signature(el):
    return (el.tag, tuple(sorted(el.attrib.items())), tuple(signature(c) for c in el))


def fingerprint(look):
    """Equal for two looks that print the same. Language doesn't change the look,
    and list numbers differ from file to file, so only whether there is a list counts."""
    marks = []
    for kind, props in look.items():
        for tag, prop in props.items():
            if tag != qn("w:lang"):
                marks.append((kind, (tag,) if tag == qn("w:numPr") else signature(prop)))
    return tuple(sorted(marks))


def table_fingerprint(style, fonts):
    style = deepcopy(style)
    spell_out_theme(style, fonts)
    return tuple(signature(c) for c in style if c.tag not in TABLE_STYLE_LABELS)


class StyleKeeper:
    """docxcompose treats styles with the same name as the same style, so a later
    file's Normal or Heading 1 takes on the first file's look. Where the look
    really differs, this gives the file a copy of the style under its own name,
    such as "Normal (2)", with every setting written into it because the copy
    can't lean on the file's defaults or base styles. Files that look alike share
    one copy, and files that match the first file need no copies at all."""

    def __init__(self, master):
        self.master = master
        self.master_theme, self.master_fonts = theme_of(master)
        self.copies = {}    # (style name, fingerprint) -> (name, id) of its copy

    def apply(self, doc, roots):
        """Point the given pieces of a file's text at styles that keep its look."""
        self.doc = doc
        self.styles = doc.styles.element
        theme, self.fonts = theme_of(doc)
        if theme != self.master_theme:
            for root in (self.styles, *roots):
                spell_out_theme(root, self.fonts)
        self.names = {s.style_id: s.name for s in doc.styles}
        self.master_styles = {s.name: s.element for s in self.master.styles}

        chosen = {}
        for root in roots:
            for p in root.iter(qn("w:p")):
                key = (p.style, self.table_style_of(p))
                if key not in chosen:
                    chosen[key] = self.paragraph_style(*key)
                if chosen[key] != p.style:
                    p.style = chosen[key]
            for tag, pick in (("w:rStyle", self.run_style), ("w:tblStyle", self.table_style)):
                for el in root.iter(qn(tag)):
                    key = (tag, el.get(W_VAL))
                    if key not in chosen:
                        chosen[key] = pick(key[1])
                    el.set(W_VAL, chosen[key])

    @staticmethod
    def table_style_of(p):
        """None outside a table, "" in a table using the default table style."""
        table = next(p.iterancestors(qn("w:tbl")), None)
        return None if table is None else table.tblPr.style or ""

    def paragraph_style(self, style_id, table_id):
        style = self.styles.get_by_id(style_id) if style_id else None
        named = style is not None
        if not named:
            default = self.doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
            if default is None:
                return style_id
            style = default.element

        # a table's style sits between the defaults and the paragraph's style
        table = None
        if table_id:
            table = self.styles.get_by_id(table_id)
        elif table_id == "" and self.doc.styles.default(WD_STYLE_TYPE.TABLE) is not None:
            table = self.doc.styles.default(WD_STYLE_TYPE.TABLE).element
        layers = []
        for layer in style_chain(self.doc, table):
            layer = deepcopy(layer)
            spell_out_theme(layer, self.fonts)
            layers.append(layer)

        mine = paragraph_look(self.doc, self.fonts, style, layers)
        if named:
            theirs = self.master_styles.get(self.names[style.styleId])
        else:
            default = self.master.styles.default(WD_STYLE_TYPE.PARAGRAPH)
            theirs = default.element if default is not None else None
        if theirs is not None and (
                fingerprint(paragraph_look(self.master, self.master_fonts, theirs, layers))
                == fingerprint(mine)):
            return style_id
        return self.copy_with_look(style, mine)

    def run_style(self, style_id):
        style = self.styles.get_by_id(style_id)
        if style is None:
            return style_id
        mine = run_look(self.doc, self.fonts, style)
        theirs = self.master_styles.get(self.names[style_id])
        if theirs is not None and (
                fingerprint(run_look(self.master, self.master_fonts, theirs)) == fingerprint(mine)):
            return style_id
        return self.copy_with_look(style, mine)

    def table_style(self, style_id):
        style = self.styles.get_by_id(style_id)
        if style is None:
            return style_id
        name = self.names[style_id]
        mine = table_fingerprint(style, self.fonts)
        theirs = self.master_styles.get(name)
        if theirs is None or table_fingerprint(theirs, self.master_fonts) == mine:
            return style_id      # nothing to clash with, or the same anyway
        new_name, new_id = self.copy_name(name, mine)
        if self.styles.get_by_id(new_id) is None:
            copy = deepcopy(style)
            spell_out_theme(copy, self.fonts)
            self.add_copy(copy, style, new_name, new_id)
        return new_id

    def copy_with_look(self, style, look):
        name = self.names[style.styleId]
        new_name, new_id = self.copy_name(name, fingerprint(look))
        if self.styles.get_by_id(new_id) is None:
            copy = deepcopy(style)
            for tag in ("w:aliases", "w:basedOn", "w:next", "w:link", "w:pPr", "w:rPr"):
                for el in copy.findall(qn(tag)):
                    copy.remove(el)
            for kind, order in (("w:pPr", PPR_ORDER), ("w:rPr", RPR_ORDER)):
                if look[kind]:
                    rank = {qn("w:" + tag): n for n, tag in enumerate(order)}
                    props = OxmlElement(kind)
                    props.extend(sorted(look[kind].values(), key=lambda p: rank.get(p.tag, len(rank))))
                    copy.append(props)
            self.add_copy(copy, style, new_name, new_id)
        return new_id

    def add_copy(self, copy, style, new_name, new_id):
        copy.set(qn("w:styleId"), new_id)
        copy.attrib.pop(qn("w:default"), None)
        if new_name != self.names[style.styleId]:
            copy.name_val = new_name
            copy.set(qn("w:customStyle"), "1")
        self.styles.append(copy)

    def copy_name(self, name, look_key):
        key = (name, look_key)
        if key not in self.copies:
            names = {s.name for s in self.master.styles} | {n for n, _ in self.copies.values()}
            ids = ({s.style_id for s in self.master.styles} | set(self.names)
                   | {i for _, i in self.copies.values()})
            new_name, n = name, 1
            while new_name in names:
                n += 1
                new_name = f"{name} ({n})"
            new_id = base_id = re.sub(r"\W", "", new_name)
            n = 1
            while new_id in ids:
                n += 1
                new_id = f"{base_id}x{n}"
            self.copies[key] = (new_name, new_id)
        return self.copies[key]

# ---------------------------------------------------------------- notes and comments
def read_part(part):
    return part.element if isinstance(part, XmlPart) else parse_xml(part.blob)


def find_notes(doc, kind):
    """(kind, part, {id: note}, marks in the text) for one kind of note in one file."""
    marks = [m for tag in kind[3] for m in doc.element.body.iter(qn(tag))]
    if not marks:
        return None
    try:
        part = doc.part.part_related_by(kind[0])
    except KeyError:
        return kind, None, {}, marks
    return kind, part, {note.get(W_ID): note for note in read_part(part)}, marks


def carry_links(composer, src_part, dst_part, note):
    """Links and pictures inside a note are numbered per part, so they need re-making."""
    composer.add_referenced_parts(src_part, dst_part, note)
    for blip in note.iter(A_NS + "blip"):
        rid = blip.get(qn("r:embed"))
        if rid:
            image = composer.pkg.get_or_add_image_part(BytesIO(src_part.rels[rid].target_part.blob))
            blip.set(qn("r:embed"), dst_part.relate_to(image, RT.IMAGE))


def carry_notes(composer, doc, found):
    """docxcompose copies the comment and endnote marks in a file's text but not
    what they point at, and Word then calls the whole file corrupted. Give each
    note a number that is free in the combined file and copy it across. Footnotes
    are done here too, because docxcompose loses the links inside them."""
    (reltype, content_type, partname, _), src_part, notes, marks = found
    master = composer.doc
    dst_part = None
    used = {}
    for mark in marks:
        old_id = mark.get(W_ID)
        if old_id not in notes:     # already broken in this file; don't pass it on
            mark.getparent().remove(mark)
            continue
        if dst_part is None:
            try:
                dst_part = master.part.part_related_by(reltype)
            except KeyError:
                root = deepcopy(notes[old_id].getparent())
                for note in list(root):
                    if note.get(qn("w:type")) is None:    # keep the separator lines
                        root.remove(note)
                dst_part = Part(PackURI(partname), content_type, serialize_part_xml(root), composer.pkg)
                master.part.relate_to(dst_part, reltype)
            dst_root = read_part(dst_part)
            next_id = max([int(n.get(W_ID)) for n in dst_root if n.get(W_ID)] + [0]) + 1
        if old_id not in used:
            used[old_id] = str(next_id)
            next_id += 1
            note = deepcopy(notes[old_id])
            note.set(W_ID, used[old_id])
            carry_links(composer, src_part, dst_part, note)
            composer.add_styles(doc, note)
            dst_root.append(note)
        mark.set(W_ID, used[old_id])
    if dst_part is not None and not isinstance(dst_part, XmlPart):
        dst_part._blob = serialize_part_xml(dst_root)


# ---------------------------------------------------------------- page setup
def page_setup(section):
    """Whatever decides where text sits on the page."""
    return (section.page_width, section.page_height,
            section.left_margin, section.right_margin, section.top_margin, section.bottom_margin,
            section.header_distance, section.footer_distance, section.gutter,
            *section._sectPr.xpath("w:cols/@w:num|w:vAlign/@w:val|w:pgBorders/*/@w:val"))


def open_section(master, doc):
    """docxcompose pours a later file into the page setup already in use, which
    squeezes a landscape page onto a portrait one. When the setup changes, close
    the current section and start one with this file's setup. Without a change
    there is no section break, so the files can share a page."""
    for sectPr in doc.element.body.iter(qn("w:sectPr")):
        for tag in ("w:headerReference", "w:footerReference", "w:titlePg", "w:pgNumType"):
            for el in sectPr.findall(qn(tag)):
                sectPr.remove(el)

    sections = doc.sections
    current = master.sections[-1]
    if not len(sections):
        return
    if len(sections) == 1 and page_setup(current) == page_setup(sections[0]):
        return

    same_size = ((current.page_width, current.page_height)
                 == (sections[0].page_width, sections[0].page_height))
    start = WD_SECTION.NEW_PAGE if NEW_PAGE_FOR_EACH_FILE or not same_size else WD_SECTION.CONTINUOUS
    last = master.add_section(start)._sectPr

    # add_section puts the section's end mark in a paragraph of its own. Keep it in
    # the file's real last paragraph, so no blank line can spill onto an empty page.
    holder = last.getprevious()
    before = holder.getprevious()
    if (before is not None and before.tag == qn("w:p")
            and before.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None):
        before.set_sectPr(holder.pPr.sectPr)
        holder.getparent().remove(holder)

    for child in list(last):
        last.remove(child)
    last.extend(deepcopy(child) for child in doc.element.body.sectPr)
    if len(sections) == 1:
        last.start_type = start
    else:
        sections[0].start_type = start


# ---------------------------------------------------------------- joining
def divider_doc(label=None, line=True, new_page=False):
    """A tiny document holding just the divider line (and optional label)."""
    d = Document()
    body = d.element.body
    for p in list(body.iterchildren(qn("w:p"))):
        body.remove(p)

    if line:
        p = d.add_paragraph()
        # bottom border on an empty paragraph = full-width horizontal line
        pPr = p._p.get_or_add_pPr()
        borders = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")        # 1.5 pt thick
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "000000")
        borders.append(bottom)
        pPr.append(borders)                 # before the spacing below: Word wants this order
        pf = p.paragraph_format
        pf.page_break_before = new_page or None
        pf.keep_with_next = True            # never strand a line at the foot of a page
        pf.space_before = Pt(12)
        pf.space_after = Pt(12)
        pf.line_spacing = 1.0
        run = p.add_run(" ")
        run.font.size = Pt(2)
        mark = OxmlElement("w:rPr")         # the paragraph mark sets the line height too
        size = OxmlElement("w:sz")
        size.set(qn("w:val"), "4")
        mark.append(size)
        pPr.append(mark)

    if label:
        lp = d.add_paragraph()
        r = lp.add_run(label)
        r.bold = True
        r.font.size = Pt(9)
        lp.paragraph_format.space_after = Pt(6)
        lp.paragraph_format.keep_with_next = True
    return d


def start_with(doc, fname):
    """The first file becomes the combined document."""
    doc.element.body.get_or_add_sectPr()
    if DIVIDER_BEFORE_FIRST or SHOW_FILENAME:
        first = divider_doc(fname if SHOW_FILENAME else None, line=DIVIDER_BEFORE_FIRST)
        body = doc.element.body
        for el in reversed([e for e in first.element.body if e.tag != qn("w:sectPr")]):
            body.insert(0, el)
    return FileComposer(doc)


def add_file(composer, keeper, doc, fname):
    notes = [n for n in (find_notes(doc, kind) for kind in NOTE_KINDS) if n]
    roots = [doc.element.body]
    for kind, _, items, marks in notes:
        if kind[0] != RT.COMMENTS:     # comments aren't printed, so their look can stay simple
            used = {m.get(W_ID) for m in marks}
            roots += [note for note_id, note in items.items() if note_id in used]
    keeper.apply(doc, roots)

    composer.reset_reference_mapping()
    composer._create_style_id_mapping(doc)
    for found in notes:
        carry_notes(composer, doc, found)

    open_section(composer.doc, doc)
    composer.append(divider_doc(fname if SHOW_FILENAME else None, new_page=NEW_PAGE_FOR_EACH_FILE))
    composer.append(doc)


def has_own_header(doc):
    for rel in doc.part.rels.values():
        if rel.reltype in (RT.HEADER, RT.FOOTER):
            root = parse_xml(rel.target_part.blob)
            has_picture = next(root.iter(qn("w:drawing"), qn("w:pict")), None) is not None
            if has_picture or "".join(root.itertext()).strip():
                return True
    return False


class BadFile(Exception):
    def __init__(self, fname, error):
        if isinstance(error, (PackageNotFoundError, zipfile.BadZipFile)):
            error = "not a readable .docx: damaged or password-protected"
        super().__init__(f"{fname} ({error})")
        self.fname = fname


def join_files(folder, files):
    """Returns the composer and how many later files had their own headers.
    Raises BadFile for the first file that can't be added."""
    composer = keeper = None
    own_headers = 0
    for n, fname in enumerate(files, 1):
        print(f"[{n}/{len(files)}] {fname} ...", flush=True)
        try:
            doc = Document(os.path.join(folder, fname))
            if composer is None:
                composer = start_with(doc, fname)
                keeper = StyleKeeper(doc)
            else:
                own_headers += has_own_header(doc)
                add_file(composer, keeper, doc, fname)
        except Exception as e:
            traceback.print_exc()
            raise BadFile(fname, e) from e
    return composer, own_headers


def main():
    dialog_root()

    folder = filedialog.askdirectory(title="Pick the folder of Word files to combine")
    if not folder:
        return

    label = os.path.basename(os.path.normpath(folder)) or os.path.splitdrive(folder)[0].rstrip(":")
    out_name = label + "_combined.docx"
    out_path = os.path.join(folder, out_name)

    names = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    files = sorted((f for f in names
                    if f.lower().endswith(".docx")
                    and not f.startswith("~$")                 # Word's temp lock files
                    and not f.lower().endswith("_combined.docx")),
                   key=natural_key)
    other_formats = [f for f in names
                     if f.lower().endswith((".doc", ".docm", ".rtf", ".odt")) and not f.startswith("~$")]
    older_combined = [f for f in names
                      if f.lower().endswith("_combined.docx") and f.lower() != out_name.lower()]

    if not files:
        messagebox.showinfo(TITLE, "No .docx files found in that folder.")
        return

    if os.path.exists(out_path) and not messagebox.askyesno(
            TITLE,
            f"{out_name} is already in this folder. Anything you changed in it "
            "would be lost if it is replaced.\n\nReplace it?",
            default=messagebox.NO):
        return

    # A file can fail partway in, leaving pieces of it behind, so start over without it.
    failed = []
    while files:
        try:
            composer, own_headers = join_files(folder, files)
            break
        except BadFile as e:
            failed.append(str(e))
            files.remove(e.fname)
    if not files:
        messagebox.showerror(TITLE, "None of the files could be read:\n" + "\n".join(failed[:15]))
        return

    # Save beside the real file first, so a failed save can't ruin an existing one.
    temp_path = out_path + ".tmp"
    try:
        composer.save(temp_path)
        os.replace(temp_path, out_path)
    except PermissionError:
        messagebox.showerror(TITLE, f"{out_name} is open in Word or locked. Close it and run again.")
        return
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

    msg = f"Combined {len(files)} file(s) into:\n{out_path}"
    if failed:
        msg += "\n\nCould not add:\n" + "\n".join(failed[:15])
        if len(failed) > 15:
            msg += f"\n...and {len(failed) - 15} more"
    if own_headers:
        msg += (f"\n\n{own_headers} later file(s) had their own headers or footers. "
                "Those were left out; the first file's are used throughout.")
    if other_formats:
        msg += (f"\n\nLeft out {len(other_formats)} file(s) in other formats, such as "
                f"{other_formats[0]}. Save them as .docx to include them.")
    if older_combined:
        msg += f"\n\nLeft out {len(older_combined)} earlier combined file(s), such as {older_combined[0]}."
    messagebox.showinfo(TITLE, msg)


if __name__ == "__main__":
    main()