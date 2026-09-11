# Markdown schema

Use Pandoc Markdown. The custom style names below are exact and case-sensitive.

## Front matter

```markdown
::: {custom-style="Manuscript Title"}
An Informative Manuscript Title with *Italic Method Names*
:::

::: {custom-style="Authors"}
First Author<sup>1,†</sup>, Second Author<sup>2,*</sup>
:::

::: {custom-style="Affiliation"}
<sup>1</sup> First institution, City Postal Code, Country
:::

::: {custom-style="Affiliation"}
<sup>2</sup> Second institution, City Postal Code, Country
:::

::: {custom-style="Equal Contribution"}
† These authors contributed equally to this work.
:::

::: {custom-style="Correspondence"}
\* Correspondence to: Name (email@example.org)
:::

{{PAGE_BREAK}}
```

Keep all authors in one paragraph. Use superscripts for affiliations and contribution markers. Use one affiliation block per institution.

## Abstract and body

```markdown
# Abstract

Abstract text.

{{PAGE_BREAK}}

# 1. Introduction

Body text with *italics*, **bold**, ***bold italics***, <u>underline</u>, H<sub>2</sub>O, and x<sup>2</sup>.

{{PAGE_BREAK}}

# 2. Results and Discussion

## 2.1 Second-level heading

### 2.1.1 Third-level heading

Body text.
```

Use a forced page break after the front matter, after the abstract, and immediately before the top-level Results (`2.`) and Materials and Methods (`3.`) headings. Placing these markers in the Markdown keeps the break at the end of the preceding page and gives all page-start headings the same top spacing. The converter retains a backward-compatible automatic break when a normalized input omits one.

Write Roman enumerators with dedicated Unicode numerals `(ⅰ)`–`(ⅹ)`, not Latin letters. For backward compatibility, the converter normalizes an unambiguous same-paragraph sequence such as `(i)`, `(ii)`, `(iii)`, but preserves an isolated `(i)` as an alphabetic marker. Lettered sequences remain `(a)`–`(i)`, including grouped or ranged forms such as `(h, i)` and `(d–f)`. Enumerator styling is contextual: the paragraph must contain at least two distinct ordered items. Only the numeral or letter is bold italic; enclosing parentheses and separators remain regular. Capitalized transition labels `First,`, `Second,`, and `Third,` are bold only when at least two appear sentence-initially as a coordinated sequence in the same paragraph; ordinary lexical uses, author names, headings, and affiliations remain regular. Full figure cross-references such as `Figure 3a` and `Figure 3a,c` are bolded including the subfigure letters.

## Equations

Use TeX display math. Pandoc converts it to native Word math:

````markdown
``` math
E = mc^2 \#(1)
```
````

Use inline math as `$E = mc^2$`. Keep equation numbers inside the math source when exact right-edge numbering is not required.

## Table

```markdown
{{PAGE_BREAK}}

::: {custom-style="Table Caption"}
**Table 1.** Caption text. Explain **BOLD** and <u>UNDERLINED</u> values if used.
:::

| Method | Metric A ↓ | Metric B ↑ |
|:--|--:|--:|
| Method 1 | <u>1.23</u> | 4.56 |
| **Method 2** | **1.11** | **4.78** |

::: {custom-style="Table Note"}
<sup>**a**</sup> Explanatory footnote text.
:::

::: {custom-style="Table Note"}
**↑** A higher value corresponds to better performance.
:::
```

The converter gives the header strong top and bottom rules, alternates light-gray data rows, centers cell content, and gives the table an outer bottom rule. A seven-column table uses the measured template column proportions; other tables use proportional automatic widths.

Carry the full table legend—title sentence plus explanatory text—in the caption paragraph, matching the figure-caption convention; place only footnote-style notes below the table, with one blank line (12 pt) of space before the first note.

## Figure captions without figures

```markdown
{{PAGE_BREAK}}

::: {custom-style="Figure Caption"}
**Figure 1.** Complete figure caption, including *(a)* panel descriptions.
:::
```

When images are omitted, retain one caption block per figure. Use `{{PAGE_BREAK}}` before every figure caption only when reproducing the template's separate figure-pages convention.

## Back matter and references

Use level-1 headings for unnumbered sections:

```markdown
# Data and Code Availability

# Acknowledgements

# Author Contributions

# References
```

Reference-list formatting is outside this skill's fidelity contract. Supplied entries remain visible and receive a readable Times New Roman fallback.

Keep Supplementary Figure Legends and Supplementary Tables in the SI source. If those inventories are retained at the end of an editable manuscript Markdown file, manuscript-mode conversion excludes them from the main DOCX by default.
