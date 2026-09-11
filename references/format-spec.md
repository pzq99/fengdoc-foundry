# Manuscript format specification

This specification records the formatting contract implemented by the bundled reference asset and converter. Values with unusual decimals are retained for reproducible layout rather than rounded editorial conventions.

## Page and document

| Parameter | Required value |
|---|---:|
| Page size | A4, 210 × 297 mm, portrait |
| Top margin | 23.00 mm |
| Bottom margin | 23.00 mm |
| Left margin | 24.00 mm |
| Right margin | 24.00 mm |
| Text width | 162.00 mm |
| Header distance | 15.01 mm |
| Footer distance | 14.23 mm |
| Gutter | 0 mm |
| Columns | One |
| Document grid pitch | 15.6 pt (312 twips) |
| Default tab stop | 21 pt (420 twips) |
| Header | Empty |
| Footer | Centered Arabic page number, continuous from page 1 |
| Theme body font | Times New Roman for Latin text |
| Accent color | `#000099` |

The source uses nine explicit next-page section breaks: after front matter, after abstract, after references/table material, and between full-page figures. The converter represents these as controlled page breaks because all sections share the same page geometry.

## Front matter

| Element | Font | Paragraph formatting |
|---|---|---|
| Title | Times New Roman 13 pt, bold; inline method/platform names and scientific terms may be bold italic | Left; 1.5 lines; 7.8 pt before/after |
| Authors | Times New Roman 12 pt; affiliation numbers bold superscript | Left; 1.279 lines; 24.95 pt before; 21.8 pt after |
| Affiliations | Times New Roman 11 pt; affiliation number bold superscript | Left; 1.2 lines; 7.8 pt before/after; 7 pt hanging indent |
| Equal-contribution note | Times New Roman 12 pt | Justified; 1.321 lines; 23.4 pt before; 7.8 pt after; 6.4 pt hanging indent |
| Correspondence | Times New Roman 12 pt; hyperlinks dark blue when Pandoc emits them | Left; 1.321 lines; 15.6 pt before; 7.8 pt after; 6.4 pt hanging indent |

## Headings and body

| Element | Font | Paragraph formatting |
|---|---|---|
| Abstract and numbered top-level heading | Times New Roman 14 pt, bold, `#000099` | Justified; 1.5 lines; 7.8 pt before/after at page start; explicit breaks precede Results and Methods |
| Later numbered top-level heading | Same | 31.2 pt before; 7.8 pt after |
| Unnumbered back-matter heading | Same | 15.6 pt before; 7.8 pt after |
| Second-level heading | Times New Roman 12 pt, bold, `#000099` | Left/justified; 1.5 lines; 7.8 pt before/after; keep with next |
| Third-level heading | Times New Roman 12 pt, regular, `#000099`; scientific terms may be italic | Left; 1.5 lines; 7.8 pt before/after; keep with next |
| Body paragraph | Times New Roman 12 pt, black | Justified; 1.5 lines; 7.8 pt before/after; no first-line indent |
| Inline procedure label | Times New Roman 12 pt, underlined | Continues in body paragraph, followed by period |
| Display equation | Cambria Math/Word OMML | Centered; body-equivalent vertical spacing |

No automatic hyphenation or multi-column layout is used. The source does not use first-line indents in ordinary body paragraphs.

## Table

| Parameter | Required value |
|---|---|
| Placement | New page in the source template |
| Caption | Above table; 12 pt; justified; 1.321 lines; 6 pt before/after; first title sentence blue, with `Table N.` bold; later explanation text black |
| Width | Approximately full text width (100.32% in source OOXML) |
| Layout | Fixed |
| Cell font | Times New Roman 12 pt; 11 pt for dense tables with six or more columns |
| Six-column proportions | Content-safe portrait profile; keep method names and single-word headers intact |
| Cell alignment | Horizontal and vertical center |
| Cell spacing | Header about 2.4 pt before/after; data rows 6 pt before/after; single line |
| Header | Bold; 1 pt top and bottom borders |
| Body shading | Alternating rows, `#F2F2F2` beginning with first data row |
| Body rules | Light gray row separators where present; no vertical borders |
| Table bottom | 0.5 pt outer bottom rule |
| Seven-column proportions | 18.53%, 16.97%, 13.87%, 10.80%, 13.85%, 10.80%, 15.39% |
| Notes | 12 pt; 1.321 lines; first note 12 pt before (one blank line below the table); zero after; arrow glyph may be 10 pt bold |

Bold denotes a column-best value and underline denotes second best where the manuscript says so.

## Figure captions

| Parameter | Required value |
|---|---|
| Source placement | Below each figure, with each figure starting a new page |
| Caption font | Times New Roman 12 pt |
| Lead sentence | First title sentence `#000099`, with `Figure N.` bold |
| Body | Later explanation text black; justified |
| Spacing | 9.6 pt before, 6 pt after |
| Line spacing | 1.321 lines |
| Panel markers | Enumerator letters are bold italic; enclosing parentheses and separators remain regular |

Complete in-text figure references, including subfigure suffixes such as `Figure 3a` or `Figure 3a,c`, are bold. Ordered Roman lists use dedicated Unicode numerals `(ⅰ)`–`(ⅹ)`; Latin-letter spellings are normalized only when the same paragraph makes the Roman sequence unambiguous. Lettered lists use `(a)`–`(i)`. For either system, styling applies only when at least two distinct ordered items establish a sequence, including grouped or ranged forms such as `(h, i)` and `(d–f)`. Only the enumerator characters are bold italic; parentheses and separators remain regular. Sentence-led transition labels `First,`, `Second,`, and `Third,` are bold only as a coordinated same-paragraph sequence, never as isolated words in names, headings, or affiliations.

When the requested document excludes images, omit the image object but keep the complete caption and any requested page break.

## References limitation

The observed source reference list uses Times New Roman 12 pt, left alignment, about 1.1-line spacing, tabs after numbering, and 7.8 pt before/after. This skill intentionally does not promise reference-format fidelity; it preserves reference text and applies a readable fallback only.
