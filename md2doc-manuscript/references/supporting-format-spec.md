# Supporting Information format specification

This specification records the Supporting Information formatting contract implemented by the bundled reference asset and converter.

## Page system

| Parameter | Portrait sections | Landscape wide-table section |
|---|---:|---:|
| Page size | A4, 210 × 297 mm | US Letter landscape, 279.4 × 215.9 mm |
| Top / bottom margin | 23.00 / 23.00 mm | 24.00 / 24.00 mm |
| Left / right margin | 24.00 / 24.00 mm | 24.00 / 24.00 mm |
| Text width | about 162 mm | about 231 mm |
| Header distance | 15.01 mm | 12.70 mm |
| Footer distance | about 14.23 mm | 12.70 mm |
| Columns | One | One |

- Header is empty.
- Footer contains a centered Arabic page number, continuous from page 1 across sections.
- Default font is Times New Roman; body size is 12 pt.
- Accent color is `#000099`.
- Document grid pitch is 15.6 pt (312 twips); default tab stop is 21 pt (420 twips).
- The source contains 16 new-page sections. The converter uses semantic page/section markers and avoids duplicating a hard page break next to a section break.

## Cover and body

| Role | Font | Paragraph formatting |
|---|---|---|
| Cover label | Times New Roman 12 pt, bold italic | Justified; 1.2 lines; 7.8 pt before, 23.4 pt after |
| Cover title | Times New Roman 14 pt, bold | Justified; 1.2 lines; 7.8 pt before/after |
| Body | Times New Roman 12 pt, black | Justified; 1.5 lines; 7.8 pt before/after; no first-line indent |
| `Discussion S#` / `Method S#` | Times New Roman 12 pt; entire heading `#000099`; label through S-number bold | Justified; 1.5 lines; 7.8 pt before, 15.6 pt after; keep with next |
| Lettered subheading | Times New Roman 12 pt; letter bold italic, title bold | Justified; 1.5 lines; 7.8 pt before/after; keep with next |
| Synthesis subheading | Outline square plus Times New Roman 12 pt bold | Justified; 1.5 lines; 7.8 pt before/after; keep with next |
| Procedure step | Times New Roman 12 pt; `Step-N` bold | Body paragraph rhythm |
| Display equation | Native Word OMML / Cambria Math | Centered; body-equivalent vertical spacing |

## Figures

| Parameter | Required value |
|---|---|
| Image alignment | Centered; original aspect ratio retained |
| Figure S1–S7 width | Approximately 458–459 pt, near the 162 mm portrait text width |
| Synthesis-route width | Approximately 458.3–458.6 pt |
| NMR spectrum width | Approximately 375.95–376.9 pt, about 132.6–133.0 mm |
| Caption font | Times New Roman 12 pt |
| Caption lead | First title sentence `#000099`, with `Figure S#.` bold |
| Caption body | Later explanation black; justified |
| Caption spacing | 7.8 pt before/after; approximately 1.321 lines |
| Panel markers | Preserve bold italic emphasis where supplied |

Images may be omitted when requested; keep the complete caption and intentional page break.

## Supplementary table

| Parameter | Required value |
|---|---|
| Placement | New landscape section for the measured wide table |
| Caption | Above table; 12 pt; justified; 1.5 lines; 6 pt before/after; lead sentence `#000099` with `Table S#` bold, matching the figure-caption convention |
| Width | Approximately full landscape text width (100.5% in source OOXML) |
| Layout | Fixed |
| Five-column grid | 1701 / 6067 / 1643 / 1645 / 2128 twips |
| Five-column percentages | 645 / 2301 / 623 / 624 / 807 out of 5000 |
| Cell font | Times New Roman 12 pt |
| Alignment | Columns 1–2 left; columns 3–5 centered; all vertically centered |
| Header | Bold; `#F2F2F2`; strong black top and bottom rules; repeat on page breaks |
| Body shading | First data row white, alternating `#F2F2F2` from the second data row |
| Body borders | Light horizontal separators; no vertical borders; black table-bottom rule |
| Note | 12 pt; 1.5 lines; left/justified; superscript marker preserved; 12 pt of space before the first note (one blank line below the table) |

The exact grid applies to a five-column table matching the reference. Other column counts retain the same typography, zebra pattern and border logic but should use widths appropriate to their content.
