#!/usr/bin/env python3
"""Build the reference DOCX containing the manuscript's named styles."""

from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "reference.docx"
FONT = "Times New Roman"
BLUE = RGBColor(0x00, 0x00, 0x99)


def set_style_font(style, size, bold=False, italic=False, color=None):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    if color:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{key}"), FONT)


def set_paragraph(style, *, align, before, after, line, left=0, first=0, keep=False):
    pf = style.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.left_indent = Pt(left)
    pf.first_line_indent = Pt(first)
    pf.keep_with_next = keep
    pf.widow_control = True


def get_or_add(doc, name, base="Normal"):
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles[base] if base else None
    return style


def get_or_add_character(doc, name):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, WD_STYLE_TYPE.CHARACTER)


def set_doc_defaults(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, 12)
    set_paragraph(
        normal,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        before=7.8,
        after=7.8,
        line=1.5,
    )
    for name in ("Body Text", "First Paragraph", "Compact"):
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        style.base_style = normal
        set_style_font(style, 12)
        set_paragraph(
            style,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            before=7.8,
            after=7.8,
            line=1.5,
        )


def build():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(23)
    section.bottom_margin = Mm(23)
    section.left_margin = Mm(24)
    section.right_margin = Mm(24)
    section.header_distance = Mm(15.01)
    section.footer_distance = Mm(14.23)

    set_doc_defaults(doc)

    definitions = [
        ("Manuscript Title", 13, True, False, None, WD_ALIGN_PARAGRAPH.LEFT, 7.8, 7.8, 1.5, 0, 0, False),
        ("Authors", 12, False, False, None, WD_ALIGN_PARAGRAPH.LEFT, 24.95, 21.8, 1.2791667, 0, 0, False),
        ("Affiliation", 11, False, False, None, WD_ALIGN_PARAGRAPH.LEFT, 7.8, 7.8, 1.2, 7, -7, False),
        ("Equal Contribution", 12, False, False, None, WD_ALIGN_PARAGRAPH.JUSTIFY, 23.4, 7.8, 1.3208333, 6.4, -6.4, False),
        ("Correspondence", 12, False, False, None, WD_ALIGN_PARAGRAPH.LEFT, 15.6, 7.8, 1.3208333, 6.4, -6.4, False),
        ("Table Caption", 12, False, False, None, WD_ALIGN_PARAGRAPH.JUSTIFY, 6, 6, 1.3208333, 0, 0, True),
        ("Table Note", 12, False, False, None, WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 0, 1.3208333, 0, 0, False),
        ("Figure Caption", 12, False, False, None, WD_ALIGN_PARAGRAPH.JUSTIFY, 9.6, 6, 1.3208333, 0, 0, False),
        ("Reference Entry", 12, False, False, None, WD_ALIGN_PARAGRAPH.LEFT, 7.8, 7.8, 1.1, 0, 0, False),
    ]
    for name, size, bold, italic, color, align, before, after, line, left, first, keep in definitions:
        style = get_or_add(doc, name)
        set_style_font(style, size, bold, italic, color)
        set_paragraph(style, align=align, before=before, after=after, line=line, left=left, first=first, keep=keep)

    superscript = get_or_add_character(doc, "Superscript")
    set_style_font(superscript, 12)
    superscript.font.superscript = True
    subscript = get_or_add_character(doc, "Subscript")
    set_style_font(subscript, 12)
    subscript.font.subscript = True
    underline = get_or_add_character(doc, "Underline")
    set_style_font(underline, 12)
    underline.font.underline = True

    for name, size, bold in (("Heading 1", 14, True), ("Heading 2", 12, True), ("Heading 3", 12, False)):
        style = doc.styles[name]
        style.base_style = doc.styles["Normal"]
        set_style_font(style, size, bold, False, BLUE)
        set_paragraph(
            style,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            before=7.8,
            after=7.8,
            line=1.5,
            keep=True,
        )

    for name in ("Title", "Subtitle", "Author", "Date", "Caption"):
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        style.base_style = doc.styles["Normal"]
        set_style_font(style, 12)

    # Match the source's 15.6 pt document grid and 21 pt default tab stop.
    sect_pr = section._sectPr
    doc_grid = sect_pr.find(qn("w:docGrid"))
    if doc_grid is None:
        doc_grid = OxmlElement("w:docGrid")
        sect_pr.append(doc_grid)
    doc_grid.set(qn("w:linePitch"), "312")
    settings = doc.settings.element
    tab = settings.find(qn("w:defaultTabStop"))
    if tab is None:
        tab = OxmlElement("w:defaultTabStop")
        settings.insert(0, tab)
    tab.set(qn("w:val"), "420")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
