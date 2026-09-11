#!/usr/bin/env python3
"""Build the reference DOCX containing Supporting Information styles."""

from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "reference-supporting.docx"
FONT = "Times New Roman"
BLUE = RGBColor(0x00, 0x00, 0x99)


def set_style_font(style, size, bold=False, italic=False, color=None):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{key}"), FONT)


def set_paragraph(style, *, align, before, after, line, keep=False):
    pf = style.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep
    pf.widow_control = True


def get_or_add(doc, name, style_type=WD_STYLE_TYPE.PARAGRAPH, base="Normal"):
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, style_type)
    if style_type == WD_STYLE_TYPE.PARAGRAPH:
        style.base_style = doc.styles[base] if base else None
    return style


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

    definitions = [
        ("SI Cover Label", 12, True, True, None, 7.8, 23.4, 1.2, False),
        ("SI Cover Title", 14, True, False, None, 7.8, 7.8, 1.2, False),
        ("SI Section Heading", 12, False, False, BLUE, 7.8, 15.6, 1.5, True),
        ("SI Subheading", 12, False, False, None, 7.8, 7.8, 1.5, True),
        ("SI Figure Caption", 12, False, False, None, 7.8, 7.8, 1.3208333, False),
        ("SI Table Caption", 12, False, False, None, 6, 6, 1.5, True),
        ("SI Table Note", 12, False, False, None, 0, 0, 1.5, False),
    ]
    for name, size, bold, italic, color, before, after, line, keep in definitions:
        style = get_or_add(doc, name)
        set_style_font(style, size, bold, italic, color)
        set_paragraph(
            style,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            before=before,
            after=after,
            line=line,
            keep=keep,
        )

    for name, attr in (("Superscript", "superscript"), ("Subscript", "subscript")):
        style = get_or_add(doc, name, WD_STYLE_TYPE.CHARACTER)
        set_style_font(style, 12)
        setattr(style.font, attr, True)
    underline = get_or_add(doc, "Underline", WD_STYLE_TYPE.CHARACTER)
    set_style_font(underline, 12)
    underline.font.underline = True

    for name in ("Heading 1", "Heading 2", "Heading 3"):
        style = doc.styles[name]
        style.base_style = normal
        set_style_font(style, 12, name != "Heading 3", False, BLUE)
        set_paragraph(
            style,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            before=7.8,
            after=7.8,
            line=1.5,
            keep=True,
        )

    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is None:
        grid = OxmlElement("w:docGrid")
        sect_pr.append(grid)
    grid.set(qn("w:linePitch"), "312")
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
