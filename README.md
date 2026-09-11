# fengdoc-foundry

A reusable Codex skill for converting scientific Markdown to formatted Word documents.

The skill supports two modes:

- `manuscript`: front matter, abstract, scholarly headings, equations, tables, captions, and page numbers.
- `supporting`: Supporting Information covers, `Discussion S#` / `Method S#` sections, figures, synthesis procedures, wide landscape tables, and continuous page numbers.

## Requirements

- Python 3 with `python-docx` and `lxml`
- Pandoc available on `PATH`

## Minimal usage

Run from `md2doc-manuscript`:

```bash
python3 scripts/md2doc.py --type manuscript examples/minimal-manuscript.md manuscript.docx
python3 scripts/md2doc.py --type supporting examples/minimal-supporting.md supporting.docx
```

The bundled examples are fictional and contain no real research data, author identity, institution, or contact information.
