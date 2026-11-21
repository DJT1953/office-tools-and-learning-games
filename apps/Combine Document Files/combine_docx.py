"""
combine_docx.py  -  Combine Document Files

A window lists every Word file (.docx, .doc and .rtf) in a folder, and any files
added one by one. Each file has its own tick box: tick the ones to join, put them
in order with the Move buttons, pick what to save as (.docx, .doc, .rtf or PDF)
and press Combine. The result is saved in the folder as:

    <FolderName>_combined.<type>

A file that can't be used (locked with a password, damaged, empty) is listed in
grey with the reason, and can't be ticked. Word's own ~$ temp files are left out.

A horizontal divider line is placed before each new file's content.
Each file keeps its own fonts, spacing, page size, orientation and margins, as
well as its pictures, tables, lists, footnotes, endnotes and comments. A file
with a different page size or orientation starts on a new page.

Headers and footers come from the first file only; the closing message says if
a later file had its own. Page numbers run on through the whole document.

.docx files are read and joined without Microsoft Word. Reading .doc and .rtf
files, and saving as .doc, .rtf or PDF, uses Word, hidden in the background.

If the combined file already exists you are asked before it is replaced.

Requirements (one-time):
    python-docx (1.2 or newer), docxcompose and comtypes. If the Python running
    this lacks them, it offers to install them.
"""

import importlib.util
import os
import queue
import re
import shutil
import site
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser
import zipfile
from copy import deepcopy
from io import BytesIO

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

TITLE = "Combine Document Files"
# The "Check for updates" button opens this page, which has the newest version of every program.
UPDATES_PAGE = "https://djt1953.github.io/office-tools-and-learning-games/"
REQUIRED = {"docx": "python-docx", "docxcompose": "docxcompose"}
if os.name == "nt":
    REQUIRED["comtypes"] = "comtypes"      # talks to Microsoft Word


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
    into, so offer to add whatever this one lacks. A built copy (the .exe from
    Setup) already has everything inside it and has no pip, so it never asks."""
    if getattr(sys, "frozen", False):
        return
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
    def __init__(self, entry, error):
        if isinstance(error, (PackageNotFoundError, zipfile.BadZipFile)):
            error = "not a readable .docx: damaged or password-protected"
        super().__init__(f"{entry.name} ({error})")
        self.entry = entry


class Cancelled(Exception):
    pass


def join_files(entries, report=None, cancel=None):
    """Join the files in the order given. Returns the composer and how many later
    files had their own headers. Raises BadFile for the first file that can't be added."""
    composer = keeper = None
    own_headers = 0
    for n, entry in enumerate(entries, 1):
        if cancel is not None and cancel.is_set():
            raise Cancelled
        if report:
            report(n - 1, len(entries), f"Joining {entry.name} ({n} of {len(entries)})...")
        try:
            doc = Document(entry.docx_path)
            if composer is None:
                composer = start_with(doc, entry.name)
                keeper = StyleKeeper(doc)
            else:
                own_headers += has_own_header(doc)
                add_file(composer, keeper, doc, entry.name)
        except Exception as e:
            traceback.print_exc()
            raise BadFile(entry, e) from e
    return composer, own_headers


# ---------------------------------------------------------------- the files on the list
KINDS = (".docx", ".doc", ".rtf")
OLE_FILE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # how .doc files, and .docx files locked with a password, begin
LOCKED = "Can't use: locked with a password"
NEEDS_WORD = "Can't use: needs Microsoft Word, which is not on this PC"
EARLIER = "An earlier combined file (not ticked)"


class SaveType:
    def __init__(self, key, label, ext, word_code):
        self.key, self.label, self.ext, self.word_code = key, label, ext, word_code


WD_FORMAT_DOCX = 16
SAVE_TYPES = (
    SaveType("docx", "Word (.docx)", ".docx", None),        # no Word needed
    SaveType("doc", "Older Word (.doc)", ".doc", 0),        # Word's own numbers for each type
    SaveType("rtf", "Rich Text (.rtf)", ".rtf", 6),
    SaveType("pdf", "PDF", ".pdf", 17),
)
SAVE_TYPE = {t.key: t for t in SAVE_TYPES}


def find_documents(folder):
    """Every .docx, .doc and .rtf file in a folder, in number order (page 2 before page 10).
    Word's ~$ files are the temporary lock files it keeps beside an open document."""
    names = [f for f in os.listdir(folder)
             if os.path.splitext(f)[1].lower() in KINDS and not f.startswith("~$")
             and os.path.isfile(os.path.join(folder, f))]
    return [os.path.join(folder, f) for f in sorted(names, key=natural_key)]


class Entry:
    """One file on the list."""

    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.name = os.path.basename(self.path)
        self.kind = os.path.splitext(self.name)[1].lower()
        self.earlier = os.path.splitext(self.name)[0].lower().endswith("_combined")
        self.ticked = False
        self.checked = False
        self.problem = None         # why it can't be used, once checked
        self.docx_path = None       # the .docx to join: the file itself, or Word's copy of a .doc or .rtf

    @property
    def usable(self):
        return self.checked and self.problem is None

    def note(self):
        if not self.checked:
            return "Checking..."
        if self.problem:
            return self.problem
        return EARLIER if self.earlier else ""


class FileList:
    """The files on the list, in the order they will be joined."""

    def __init__(self):
        self.entries = []

    def add(self, paths):
        """Add files not already listed; returns the new entries."""
        known = {os.path.normcase(e.path) for e in self.entries}
        new = []
        for path in paths:
            key = os.path.normcase(os.path.abspath(path))
            if key not in known:
                known.add(key)
                new.append(Entry(path))
        self.entries += new
        return new

    def clear(self):
        self.entries = []

    def ticked(self):
        return [e for e in self.entries if e.ticked and e.usable]

    def tick_all(self, on):
        for e in self.entries:
            if e.usable:
                e.ticked = on

    def move(self, entry, where):
        """where: "up", "down", "top" or "bottom". False if it is already there."""
        old = self.entries.index(entry)
        new = {"up": old - 1, "down": old + 1, "top": 0, "bottom": len(self.entries) - 1}[where]
        new = max(0, min(len(self.entries) - 1, new))
        if new == old:
            return False
        self.entries.insert(new, self.entries.pop(old))
        return True

    def name_order(self):
        self.entries.sort(key=lambda e: natural_key(e.name))


def check_entry(entry, word, workdir):
    """Find out whether a file can be joined, and tick it if so (unless it is an
    earlier combined file). Word, when there is one, makes a .docx copy of a .doc
    or .rtf file in workdir; the file itself is only ever read."""
    entry.problem, entry.docx_path = _check(entry, word, workdir)
    entry.checked = True
    entry.ticked = entry.problem is None and not entry.earlier


def _check(entry, word, workdir):
    try:
        with open(entry.path, "rb") as f:
            head = f.read(8)
    except OSError as e:
        return f"Can't use: could not be opened ({e.strerror or e})", None
    if not head:
        return "Can't use: the file is empty", None
    if entry.kind == ".docx" and head != OLE_FILE:
        try:
            Document(entry.path)
            return None, entry.path
        except Exception:
            if word is None:
                return "Can't use: damaged, or not really a .docx file", None
    if word is None:
        if entry.kind == ".docx":
            return LOCKED + " (or it is really an older .doc file)", None
        return NEEDS_WORD, None

    # Word reads a copy, so a file open in Word right now is never touched or closed.
    fd, copy = tempfile.mkstemp(suffix=entry.kind, dir=workdir)
    os.close(fd)
    target = copy[:-len(entry.kind)] + "-as.docx"
    try:
        shutil.copyfile(entry.path, copy)
        word.convert(copy, target, WD_FORMAT_DOCX)
        Document(target)
    except WordError as e:
        return f"Can't use: {e}", None
    except Exception:
        return "Can't use: damaged, or Word could not read it", None
    finally:
        try:
            os.remove(copy)
        except OSError:
            pass
    return None, target


# ---------------------------------------------------------------- Microsoft Word
WORD_WRONG_PASSWORD = -2146822880
NO_PASSWORD = "not-a-real-password"     # a locked file then fails at once, instead of a hidden prompt
MACROS_OFF = 3


class WordError(Exception):
    pass


def word_installed():
    """Is Microsoft Word on this computer? (It registers itself with Windows.)"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except (ImportError, OSError):
        return False


def _reason(error):
    """Word's own words for what went wrong, or "" when it gives none."""
    details = getattr(error, "details", None)
    if details is not None:
        text = str(details[0]) if details and details[0] else ""
    else:
        text = str(error)
    return " ".join(text.split())[:200]


class Word:
    """A hidden copy of Word, started the first time it is needed. Use it from
    one thread only, and close it in that same thread."""

    def __init__(self):
        self.app = None

    def _start(self):
        import comtypes
        import comtypes.client
        comtypes.CoInitialize()
        self.started = True
        try:
            self.app = comtypes.client.CreateObject("Word.Application", dynamic=True)
        except Exception as error:
            self.app = None
            raise WordError("Microsoft Word could not be started (" + (_reason(error) or str(error)) + ")")
        # a Word the user is working in is left open at the end
        self.ours = not self.app.Visible and self.app.Documents.Count == 0
        self.before = (self.app.DisplayAlerts, self.app.AutomationSecurity)
        self.app.DisplayAlerts = 0
        self.app.AutomationSecurity = MACROS_OFF      # never run a document's macros

    def convert(self, source, target, file_format):
        """Open source (read-only, no window) and save it as target in Word's file_format."""
        if self.app is None:
            self._start()
        import comtypes.automation
        missing = comtypes.automation.VARIANT.missing
        try:
            doc = self.app.Documents.Open(source, False, True, False, NO_PASSWORD, "", False,
                                          "", "", 0, missing, False)
        except Exception as error:
            code = getattr(error, "hresult", None) or (error.args[0] if error.args else None)
            if code == WORD_WRONG_PASSWORD:
                raise WordError("locked with a password")
            why = _reason(error)
            raise WordError("damaged, or not a file Word can open" + (f" ({why})" if why else ""))
        try:
            if file_format == WD_FORMAT_DOCX:
                # the last 15 means today's Word, so an old file doesn't stay in Compatibility Mode
                doc.SaveAs2(target, file_format, False, "", False, "", False, False, False, False,
                            False, missing, False, False, missing, False, 15)
            else:
                doc.SaveAs2(target, file_format, False, "", False)
        except Exception as error:
            why = _reason(error)
            raise WordError("Word could not save it" + (f" ({why})" if why else ""))
        finally:
            try:
                doc.Close(0)
            except Exception:
                pass

    def close(self):
        if not getattr(self, "started", False):
            return
        if self.app is not None:
            try:
                if self.ours and self.app.Documents.Count == 0:
                    self.app.Quit(0)
                else:
                    self.app.DisplayAlerts, self.app.AutomationSecurity = self.before
            except Exception:
                pass                    # Word already gone
        self.app = None
        self.started = False
        import comtypes
        comtypes.CoUninitialize()


# ---------------------------------------------------------------- combining
def output_path(folder, save_type):
    label = os.path.basename(os.path.normpath(folder)) or os.path.splitdrive(folder)[0].rstrip(":")
    return os.path.join(folder, label + "_combined" + save_type.ext)


def combine(entries, out_path, save_type, word, workdir, report=None, cancel=None):
    """Join the entries, in the order given, into out_path.
    Returns (the entries joined, messages for any that could not be, how many had own headers)."""
    entries = list(entries)
    failed = []
    # A file can fail partway in, leaving pieces of it behind, so start over without it.
    while entries:
        try:
            composer, own_headers = join_files(entries, report, cancel)
            break
        except BadFile as e:
            failed.append(str(e))
            entries.remove(e.entry)
    if not entries:
        return entries, failed, 0

    if report:
        report(len(entries), len(entries), f"Saving as {save_type.label}...")
    made = os.path.join(workdir, "combined.docx")
    composer.save(made)
    if save_type.word_code is not None:
        if word is None:
            raise WordError(f"saving as {save_type.label} needs Microsoft Word")
        final = os.path.join(workdir, "combined" + save_type.ext)
        word.convert(made, final, save_type.word_code)
        made = final
    # Copy beside the real file first, so a failed save can't ruin an existing one.
    temp = out_path + ".tmp"
    try:
        shutil.copyfile(made, temp)
        os.replace(temp, out_path)
    finally:
        if os.path.exists(temp):
            os.remove(temp)
    return entries, failed, own_headers


# ---------------------------------------------------------------- the window
BOX_ON, BOX_OFF = "\u2611", "\u2610"


def work_area(root):
    """(left, top, width, height) of the screen space not covered by the taskbar."""
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - 60


class App:
    def __init__(self, root):
        self.root = root
        root.title(TITLE)
        self.files = FileList()
        self.folder = None
        self.has_word = word_installed()
        self.workdir = tempfile.mkdtemp(prefix="Combine Document Files ")
        self.q = queue.Queue()
        self.worker = None
        self.busy = None                # None, "check" or "combine"
        self.cancel = threading.Event()
        self.closing = False
        self.last_picked = None

        self.save_var = tk.StringVar(value="docx")
        self.folder_var = tk.StringVar(value="No folder picked yet.")
        self.out_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value="Step 1: press \"1. Pick folder...\", or \"Add files...\" to pick files one at a time.")

        self._fit_window()
        self._style()
        self._build()
        self._refresh()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.poll_id = root.after(100, self._poll)
        root.bind("<Destroy>", self._destroyed, add="+")
        # Alt+1 and Alt+4 press the two main buttons (the Sandbox test uses them too)
        root.bind("<Alt-Key-1>", lambda e: self._press(self.btn_folder))
        root.bind("<Alt-Key-4>", lambda e: self._press(self.btn_combine))

    def _press(self, button):
        if str(button.cget("state")) == "disabled":
            self.say("Wait until the work under way is finished.")
        else:
            button.invoke()
        return "break"

    def check_for_updates(self):
        """Opens UPDATES_PAGE in the web browser, or says where to go when that can't be done."""
        try:
            opened = webbrowser.open(UPDATES_PAGE)
        except Exception:
            opened = False
        if not opened:
            messagebox.showinfo(TITLE, "The web browser could not be opened. To check for a newer "
                                f"version, go to:\n\n{UPDATES_PAGE}", parent=self.root)

    def _destroyed(self, event):
        if event.widget is self.root:
            self.root.after_cancel(self.poll_id)

    # ------------------------------------------------------------ layout
    def _fit_window(self):
        # Fit above the taskbar, so the Combine button is never hidden on short screens.
        left, top, width, height = work_area(self.root)
        h = max(500, min(720, height - 50))       # 50 = title bar and window edges
        w = min(860, width - 20)
        x = left + max(0, (width - w) // 2)
        y = top + max(0, (height - h - 40) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(min(760, w), min(500, h))

    def _style(self):
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=11)
        tkfont.nametofont("TkTextFont").configure(size=11)
        family = base.actual("family")
        style = ttk.Style()
        style.configure("TButton", font=(family, 11), padding=(10, 4))
        style.configure("Big.TButton", font=(family, 12, "bold"), padding=(14, 8))
        style.configure("TRadiobutton", font=(family, 11))
        style.configure("TLabelframe.Label", font=(family, 11, "bold"))
        style.configure("Treeview", rowheight=26, font=(family, 11))
        style.configure("Treeview.Heading", font=(family, 11, "bold"))

    def _build(self):
        pad = {"padx": 10, "pady": 4}
        root = self.root

        # Bottom blocks first, so a short window squeezes the list, not the buttons.
        bottom = ttk.Frame(root)
        bottom.pack(side="bottom", fill="x", **pad)
        row = ttk.Frame(bottom)
        row.pack(fill="x")
        self.btn_combine = ttk.Button(row, text="4. Combine", style="Big.TButton", command=self.combine_click)
        self.btn_combine.pack(side="left")
        self.btn_stop = ttk.Button(row, text="Stop", command=self.stop_click)
        self.btn_stop.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.status_lbl = ttk.Label(bottom, textvariable=self.status_var, anchor="w", justify="left")
        self.status_lbl.pack(fill="x", pady=(6, 0))
        bottom.bind("<Configure>", lambda e: self.status_lbl.configure(wraplength=max(300, e.width - 10)))

        step3 = ttk.LabelFrame(root, text="Step 3: Save as")
        step3.pack(side="bottom", fill="x", **pad)
        radios = ttk.Frame(step3)
        radios.pack(fill="x", padx=8, pady=(4, 0))
        self.radios = []
        for t in SAVE_TYPES:
            r = ttk.Radiobutton(radios, text=t.label, value=t.key, variable=self.save_var,
                                command=self.save_type_click)
            r.pack(side="left", padx=(0, 18))
            self.radios.append(r)
        ttk.Label(step3, textvariable=self.out_var, anchor="w").pack(fill="x", padx=8, pady=(2, 6))

        step2 = ttk.LabelFrame(root, text="Step 2: Set the order (click a file in the list first)")
        step2.pack(side="bottom", fill="x", **pad)
        moves = ttk.Frame(step2)
        moves.pack(fill="x", padx=8, pady=6)
        self.move_buttons = []
        for text, where in (("Move up", "up"), ("Move down", "down"),
                            ("Move to top", "top"), ("Move to bottom", "bottom")):
            b = ttk.Button(moves, text=text, command=lambda w=where: self.move_click(w))
            b.pack(side="left", padx=(0, 6))
            self.move_buttons.append(b)
        b = ttk.Button(moves, text="Back to name order", command=self.name_order_click)
        b.pack(side="left", padx=(12, 0))
        self.move_buttons.append(b)

        # A thin bar across the top with "Check for updates" in the right corner
        bar = ttk.Frame(root)
        bar.pack(side="top", fill="x", padx=10, pady=(8, 0))
        self.btn_updates = ttk.Button(bar, text="Check for updates", command=self.check_for_updates)
        self.btn_updates.pack(side="right")

        step1 = ttk.LabelFrame(root, text="Step 1: Pick the files")
        step1.pack(side="top", fill="both", expand=True, **pad)
        row = ttk.Frame(step1)
        row.pack(fill="x", padx=8, pady=(6, 4))
        self.btn_folder = ttk.Button(row, text="1. Pick folder...", style="Big.TButton",
                                     command=self.pick_folder_click)
        self.btn_folder.pack(side="left")
        self.btn_add = ttk.Button(row, text="Add files...", command=self.add_files_click)
        self.btn_add.pack(side="left", padx=8)
        ttk.Label(row, textvariable=self.folder_var, anchor="w").pack(side="left", fill="x", expand=True)

        row = ttk.Frame(step1)
        row.pack(side="bottom", fill="x", padx=8, pady=(4, 6))
        self.btn_tick = ttk.Button(row, text="Tick all", command=lambda: self.tick_all_click(True))
        self.btn_tick.pack(side="left")
        self.btn_untick = ttk.Button(row, text="Untick all", command=lambda: self.tick_all_click(False))
        self.btn_untick.pack(side="left", padx=6)
        self.btn_remove = ttk.Button(row, text="Take off the list", command=self.remove_click)
        self.btn_remove.pack(side="left", padx=6)
        ttk.Label(row, text="Click a file's box (or press Space) to tick or untick it.",
                  anchor="w").pack(side="left", padx=8)

        listbox = ttk.Frame(step1)
        listbox.pack(fill="both", expand=True, padx=8)
        cols = ("use", "num", "file", "type", "note")
        self.tree = ttk.Treeview(listbox, columns=cols, show="headings", selectmode="browse", height=6)
        for col, text, width, anchor, stretch in (
                ("use", "Use", 50, "center", False), ("num", "#", 40, "center", False),
                ("file", "File", 300, "w", True), ("type", "Type", 60, "center", False),
                ("note", "Note", 280, "w", True)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, minwidth=width if not stretch else 80,
                             anchor=anchor, stretch=stretch)
        self.tree.tag_configure("bad", foreground="#6b6b6b")
        scroll = ttk.Scrollbar(listbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self.tree_click)
        self.tree.bind("<Double-Button-1>", self.tree_double)
        self.tree.bind("<space>", lambda e: self.toggle(self.selected()) or "break")
        self.tree.bind("<<TreeviewSelect>>", self.tree_select)

    # ------------------------------------------------------------ showing the list
    def _refresh(self):
        chosen = self.selected()
        self.tree.delete(*self.tree.get_children())
        number = 0
        for e in self.files.entries:
            if e.ticked and e.usable:
                number += 1
                use, num = BOX_ON, str(number)
            else:
                use, num = (BOX_OFF if e.usable else "-"), ""
            self.tree.insert("", "end", iid=str(id(e)), values=(use, num, e.name, e.kind, e.note()),
                             tags=() if e.usable or not e.checked else ("bad",))
        if chosen is not None and chosen in self.files.entries:
            self.tree.selection_set(str(id(chosen)))
            self.tree.see(str(id(chosen)))

        ticked = len(self.files.ticked())
        if ticked == 1:
            self.btn_combine.configure(text="4. Combine the 1 ticked file")
        elif ticked:
            self.btn_combine.configure(text=f"4. Combine the {ticked} ticked files")
        else:
            self.btn_combine.configure(text="4. Combine")
        self._show_out()
        self._set_states()

    def _out_folder(self):
        if self.folder:
            return self.folder
        ticked = self.files.ticked() or self.files.entries
        return os.path.dirname(ticked[0].path) if ticked else None

    def _show_out(self):
        folder = self._out_folder()
        save_type = SAVE_TYPE[self.save_var.get()]
        if folder is None:
            self.out_var.set(f"Saves as <folder name>_combined{save_type.ext}, in the folder you pick.")
        else:
            name = os.path.basename(output_path(folder, save_type))
            where = "the same folder" if self.folder else folder
            self.out_var.set(f"Saves as {name}, in {where}.")

    def _set_states(self):
        # Grey only while work is under way; everything turns black again when it ends.
        lists = "disabled" if self.busy else "normal"
        edits = "disabled" if self.busy == "combine" else "normal"
        for b in (self.btn_folder, self.btn_add, self.btn_combine, self.btn_remove):
            b.configure(state=lists)
        for b in (self.btn_tick, self.btn_untick, *self.move_buttons, *self.radios):
            b.configure(state=edits)
        self.btn_stop.configure(state="normal" if self.busy else "disabled")

    def say(self, text):
        self.status_var.set(text)

    def selected(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return next((e for e in self.files.entries if str(id(e)) == sel[0]), None)

    def _entry_at(self, event):
        row = self.tree.identify_row(event.y)
        return next((e for e in self.files.entries if str(id(e)) == row), None) if row else None

    # ------------------------------------------------------------ clicks in the list
    def tree_click(self, event):
        entry = self._entry_at(event)
        if entry is not None and self.tree.identify_column(event.x) == "#1":
            self.tree.selection_set(str(id(entry)))
            self.toggle(entry)
            return "break"

    def tree_double(self, event):
        entry = self._entry_at(event)
        if entry is not None:
            self.toggle(entry)
            return "break"

    def tree_select(self, _event=None):
        entry = self.selected()
        if entry is None or entry is self.last_picked:
            return              # the list was only redrawn: keep what the status line says
        self.last_picked = entry
        if entry.checked and entry.problem:
            self.say(f"{entry.name}: {entry.problem}.")
        else:
            self.say(f"Picked {entry.name}. Use the Move buttons to change its place, "
                     "or click its box to tick or untick it.")

    def toggle(self, entry):
        if entry is None:
            self.say("Click a file in the list first.")
            return
        if self.busy == "combine":
            self.say("Wait until the files are combined.")
            return
        if not entry.checked:
            self.say(f"{entry.name} is still being checked. Wait a moment.")
            return
        if entry.problem:
            self.say(f"{entry.name} can't be ticked. {entry.problem}.")
            return
        entry.ticked = not entry.ticked
        self._refresh()
        self.say(f"{'Ticked' if entry.ticked else 'Unticked'} {entry.name}. "
                 f"{len(self.files.ticked())} file(s) will be combined.")

    # ------------------------------------------------------------ step 1
    def pick_folder_click(self):
        folder = filedialog.askdirectory(parent=self.root, title="Pick the folder of documents to combine")
        if not folder:
            self.say("No folder picked. Press \"1. Pick folder...\" to try again.")
            return
        self.open_folder(folder)

    def open_folder(self, folder):
        folder = os.path.normpath(folder)
        try:
            paths = find_documents(folder)
        except OSError as e:
            messagebox.showerror(TITLE, f"That folder can't be read:\n{e}", parent=self.root)
            return
        self.folder = folder
        self.folder_var.set(folder)
        self.files.clear()
        self.files.add(paths)
        self._refresh()
        if not paths:
            self.say("No .docx, .doc or .rtf files in that folder. Pick another folder, "
                     "or use \"Add files...\".")
            messagebox.showinfo(TITLE, "No .docx, .doc or .rtf files found in that folder.", parent=self.root)
            return
        self.start_checking()

    def open_paths(self, paths):
        """Folders and files given when the program starts (such as dropped on its icon)."""
        folders = [p for p in paths if os.path.isdir(p)]
        if folders:
            self.open_folder(folders[0])
        files = [os.path.normpath(p) for p in paths
                 if os.path.isfile(p) and os.path.splitext(p)[1].lower() in KINDS]
        if files and self.files.add(files):
            self._refresh()
            self.start_checking()

    def add_files_click(self):
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Pick documents to add to the list",
            filetypes=[("Documents (.docx, .doc, .rtf)", "*.docx *.doc *.rtf"), ("All files", "*.*")])
        if not paths:
            self.say("No files picked. Press \"Add files...\" to try again.")
            return
        wrong = [p for p in paths if os.path.splitext(p)[1].lower() not in KINDS]
        new = self.files.add(os.path.normpath(p) for p in paths if p not in wrong)
        self._refresh()
        if wrong:
            messagebox.showinfo(TITLE, f"Left out {len(wrong)} file(s), such as "
                                f"{os.path.basename(wrong[0])}: only .docx, .doc and .rtf files "
                                "can be combined.", parent=self.root)
        if new:
            self.tree.selection_set(str(id(new[-1])))
            self.tree.see(str(id(new[-1])))
            self.start_checking()
        elif not wrong:
            self.say("Those files are already on the list.")

    def tick_all_click(self, on):
        if not self.files.entries:
            self.say("The list is empty. Pick a folder or add files first (step 1).")
            return
        self.files.tick_all(on)
        self._refresh()
        if on:
            self.say(f"Ticked every file that can be used: {len(self.files.ticked())}.")
        else:
            self.say("Unticked every file. Click a file's box to tick it.")

    def remove_click(self):
        entry = self.selected()
        if entry is None:
            self.say("Click a file in the list first, then press \"Take off the list\".")
            return
        self.files.entries.remove(entry)
        self._refresh()
        self.say(f"Took {entry.name} off the list. The file itself is not changed.")

    # ------------------------------------------------------------ step 2
    def move_click(self, where):
        entry = self.selected()
        if entry is None:
            self.say("Click a file in the list first, then press a Move button.")
            return
        if self.files.move(entry, where):
            self._refresh()
            self.say(f"Moved {entry.name} {'to the ' if where in ('top', 'bottom') else ''}{where}.")
        else:
            self.say(f"{entry.name} is already at the {'top' if where in ('up', 'top') else 'bottom'}.")

    def name_order_click(self):
        if not self.files.entries:
            self.say("The list is empty. Pick a folder or add files first (step 1).")
            return
        self.files.name_order()
        self._refresh()
        self.say("Put the files back in name order (page 2 before page 10).")

    # ------------------------------------------------------------ step 3
    def save_type_click(self):
        save_type = SAVE_TYPE[self.save_var.get()]
        if save_type.word_code is not None and not self.has_word:
            self.save_var.set("docx")
            self._show_out()
            messagebox.showinfo(TITLE, f"Saving as {save_type.label} needs Microsoft Word, which is not "
                                "on this PC.\n\nWord (.docx) works without it.", parent=self.root)
            self.say("Saving as Word (.docx).")
            return
        self._show_out()
        self.say(f"Will save as {save_type.label}.")

    # ------------------------------------------------------------ step 4
    def combine_click(self):
        if not self.files.entries:
            self.say("Pick a folder or add files first (step 1).")
            return
        if any(not e.checked for e in self.files.entries):
            self.say("Wait: the files are still being checked.")
            return
        entries = self.files.ticked()
        if not entries:
            self.say("Tick at least one file first: click its box in the Use column.")
            return
        save_type = SAVE_TYPE[self.save_var.get()]
        out_path = output_path(self._out_folder(), save_type)
        out_name = os.path.basename(out_path)
        if os.path.exists(out_path) and not messagebox.askyesno(
                TITLE, f"{out_name} is already in this folder. Anything you changed in it "
                "would be lost if it is replaced.\n\nReplace it?", default=messagebox.NO, parent=self.root):
            self.say("Nothing was saved. Change the list, or say Yes to replace the old file.")
            return
        self._start("combine", self._combine_job, entries, out_path, save_type)

    def stop_click(self):
        if not self.busy:
            self.say("Nothing is running.")
            return
        self.cancel.set()
        self.say("Stopping after the file in hand...")

    # ------------------------------------------------------------ work in the background
    def start_checking(self):
        todo = [e for e in self.files.entries if not e.checked]
        if todo:
            self._start("check", self._check_job, todo)

    def _start(self, kind, job, *args):
        self.busy = kind
        self.cancel.clear()
        self._set_states()
        self.progress.configure(value=0, maximum=1)
        self.worker = threading.Thread(target=self._work, args=(kind, job) + args, daemon=True)
        self.worker.start()

    def _work(self, kind, job, *args):
        word = Word() if self.has_word else None
        try:
            job(word, *args)
        except Cancelled:
            self.q.put(("stopped",))
        except PermissionError as e:
            self.q.put(("locked", getattr(e, "filename", None) or str(e)))
        except WordError as e:
            self.q.put(("error", f"Microsoft Word could not make the file: {e}."))
        except Exception:
            self.q.put(("crash", traceback.format_exc()))
        finally:
            if word is not None:
                word.close()
            self.q.put(("idle", kind))

    def _check_job(self, word, todo):
        for n, entry in enumerate(todo):
            if self.cancel.is_set():
                raise Cancelled
            self.q.put(("progress", n, len(todo), f"Checking {entry.name} ({n + 1} of {len(todo)})..."))
            check_entry(entry, word, self.workdir)
            self.q.put(("checked",))

    def _combine_job(self, word, entries, out_path, save_type):
        def report(n, total, text):
            self.q.put(("progress", n, total, text))
        joined, failed, own_headers = combine(entries, out_path, save_type, word, self.workdir,
                                              report, self.cancel)
        self.q.put(("done", joined, failed, own_headers, out_path))

    def _poll(self):
        try:
            while True:
                self._handle(*self.q.get_nowait())
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after_cancel(self.poll_id)
            self.poll_id = self.root.after(100, self._poll)

    def _handle(self, what, *info):
        if what == "progress":
            n, total, text = info
            self.progress.configure(maximum=max(1, total), value=n)
            self.say(text)
        elif what == "checked":
            self._refresh()
        elif what == "idle":
            self.busy = None
            if self.closing:
                self.root.destroy()
                return
            self.progress.configure(value=self.progress.cget("maximum"))
            self._refresh()
            if info[0] == "check" and not self.cancel.is_set():
                self._after_check()
        elif what == "stopped":
            self.say("Stopped. Files not yet checked can't be ticked; pick the folder again to check them."
                     if self.busy == "check" else "Stopped. Nothing was saved.")
        elif what == "locked":
            messagebox.showerror(TITLE, f"The combined file is open in Word or another program, or the "
                                 f"folder is locked:\n{info[0]}\n\nClose it and press 4 again.",
                                 parent=self.root)
            self.say("Nothing was saved. Close the file and press 4 again.")
        elif what == "error":
            messagebox.showerror(TITLE, info[0], parent=self.root)
            self.say("Nothing was saved. " + info[0])
        elif what == "crash":
            messagebox.showerror(TITLE, "Something went wrong:\n\n" + info[0][-1500:], parent=self.root)
            self.say("Something went wrong. Nothing was saved.")
        elif what == "done":
            self._finished(*info)

    def _after_check(self):
        good = sum(e.usable for e in self.files.entries)
        bad = sum(e.checked and not e.usable for e in self.files.entries)
        text = (f"{good} file(s) can be combined; {len(self.files.ticked())} are ticked. "
                "Untick any you don't want, set the order in step 2, pick the type in step 3, then press 4.")
        if bad:
            text += f" {bad} can't be used (shown in grey; click one to see why)."
        self.say(text)

    def _finished(self, joined, failed, own_headers, out_path):
        if not joined:
            messagebox.showerror(TITLE, "None of the ticked files could be joined:\n" + "\n".join(failed[:15]),
                                 parent=self.root)
            self.say("Nothing was saved: none of the ticked files could be joined.")
            return
        msg = f"Combined {len(joined)} file(s) into:\n{out_path}"
        if failed:
            msg += "\n\nCould not add:\n" + "\n".join(failed[:15])
            if len(failed) > 15:
                msg += f"\n...and {len(failed) - 15} more"
        if own_headers:
            msg += (f"\n\n{own_headers} later file(s) had their own headers or footers. "
                    "Those were left out; the first file's are used throughout.")
        self.say(f"Saved {os.path.basename(out_path)}.")
        if messagebox.askyesno(TITLE, msg + "\n\nOpen it now?", parent=self.root):
            try:
                os.startfile(out_path)
            except OSError as e:
                messagebox.showerror(TITLE, f"It could not be opened:\n{e}", parent=self.root)

    def on_close(self):
        if self.worker is not None and self.worker.is_alive():
            self.closing = True
            self.cancel.set()
            self.say("Closing after the file in hand...")
            return
        self.root.destroy()


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)     # sharp text on high-DPI screens
        except Exception:
            pass
    root = tk.Tk()
    root.report_callback_exception = show_crash
    app = App(root)
    if sys.argv[1:]:
        root.after(300, app.open_paths, sys.argv[1:])
    try:
        root.mainloop()
    finally:
        shutil.rmtree(app.workdir, ignore_errors=True)


if __name__ == "__main__":
    main()
