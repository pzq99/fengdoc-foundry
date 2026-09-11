---
name: md2doc-manuscript
description: Convert scientific Markdown into consistently formatted Word documents in either manuscript or Supporting Information mode. Use for scholarly manuscripts and supplementary/supporting-information documents that need controlled page geometry, typography, equations, figures, captions, tables, page breaks, orientation changes, and page numbers; reference-list style normalization remains out of scope.
---

# Scientific Markdown to DOCX

## Select one mode

- **Manuscript**: use for the main article. Read [references/markdown-schema.md](references/markdown-schema.md) and, when exact values matter, [references/format-spec.md](references/format-spec.md).
- **Supporting Information**: use for supplementary information, supporting information, SI, supplementary methods, Figure S#, or Table S#. Read [references/supporting-markdown-schema.md](references/supporting-markdown-schema.md) and, when exact values matter, [references/supporting-format-spec.md](references/supporting-format-spec.md).

Do not mix the two schemas in one DOCX. Infer the mode from the user's request and content; `Supplementary Information for:`, `Method S#`, `Figure S#`, and `Table S#` indicate Supporting Information.

## Convert

Normalize the source into the selected schema. Preserve Markdown emphasis, `<u>`, `<sup>`, `<sub>`, links, visible citation text, pipe tables, TeX math, and supported images.

Run one of:

```bash
python3 scripts/md2doc.py --type manuscript input.md output.docx
python3 scripts/md2doc.py --type supporting input.md output.docx
```

Omitting `--type` keeps the original manuscript behavior for backward compatibility.

## Shared rules

- Treat Markdown as semantic input. Never fake layout with repeated spaces or blank paragraphs.
- Keep pipe-table rows contiguous: a blank line inside a table truncates it to its header row and dumps the remaining rows as literal pipe-delimited text.
- The converter fixes every table to the exact text width of its containing section (twips, fixed layout), preserving the Markdown dash-ratio column proportions; the measured five-column reference grid applies only to five-column tables inside landscape sections.
- Portrait six-column model-comparison tables use content-safe column proportions and an 11 pt dense-table font so method names and compact headers are not split mid-word.
- The converter applies text conventions automatically: complete internal cross-references ("Section 2.3", "Figure 3a", "Figure 4a,c", "Table S1", "Supplementary Figure S7") are bolded, including subfigure letters. Ordered Roman lists use dedicated Unicode numerals `(ⅰ)`, `(ⅱ)`, `(ⅲ)` rather than Latin-letter spellings; unambiguous ASCII sequences such as `(i)`, `(ii)`, `(iii)` are normalized automatically, while an isolated `(i)` remains an alphabetic marker. Compact Roman or letter markers are emphasized only when the paragraph establishes an actual sequence (at least two distinct items, including grouped/ranged forms such as `(h, i)` and `(d–f)`): only the numeral/letter is bold italic, while parentheses, commas, spaces, and range dashes remain regular. Sentence-led **First,** **Second,** and **Third,** are bold only when at least two occur as a coordinated sequence in the same paragraph; isolated lexical uses such as "Second Affiliated Hospital" remain regular. Platform names ("LinkeReady"), including occurrences in manuscript and SI titles, Latin loci ("in silico", "in vitro", "in vivo"), and statistical letters (P, q, n before a comparison operator) are italicized; URLs and e-mail addresses are rendered dark blue and underlined as plain text, without hyperlink fields.
- Use `{{PAGE_BREAK}}` only for an intentional new page.
- In manuscript mode, put `{{PAGE_BREAK}}` immediately before the top-level Results and Materials and Methods headings. This creates the same prior-page break used between Abstract and Introduction and keeps page-start heading spacing consistent.
- In Supporting Information, use `{{SECTION_BREAK_LANDSCAPE}}` and `{{SECTION_BREAK_PORTRAIT}}` for orientation transitions; a section marker already starts a new page, so do not pair it with `{{PAGE_BREAK}}`.
- Convert display math to native Word OMML through Pandoc.
- Preserve supplied images, center them, keep their aspect ratio, and cap them at the active text width. If images are intentionally omitted, keep complete figure captions.
- Preserve supplied reference text, but do not claim template-faithful manuscript reference-list formatting.
- Main-manuscript DOCX output excludes trailing `Supplementary Figure Legends` and `Supplementary Tables` sections by default; those inventories may remain in the editable Markdown and belong in the Supporting Information deliverable.
- Keep model-name cells in comparison tables citation-free when the surrounding prose or caption already carries the method citations; do not add empty citation placeholders inside tables.
- Use ASCII hyphens in generated prose unless scientific notation requires another character already present in the source.

## Dependencies and assets

- Require `pandoc` on `PATH`.
- Require Python 3 with `python-docx` and `lxml`.
- Manuscript mode uses [assets/reference.docx](assets/reference.docx).
- Supporting Information mode uses [assets/reference-supporting.docx](assets/reference-supporting.docx).
- Regenerate an asset only after deliberately changing its corresponding format specification:

```bash
python3 scripts/build_reference.py
python3 scripts/build_supporting_reference.py
```

## Validation

Confirm conversion succeeds and the result opens as a DOCX. Render every output page and inspect it at 100% zoom. Check page geometry, hierarchy, tables, image proportions, captions, equations, orientation transitions, overflow, orphaned headings, and continuous centered page numbers. Verify every table kept all of its data rows and shading—no row may survive as literal pipe-delimited text. Correct the Markdown or converter and repeat until clean.
