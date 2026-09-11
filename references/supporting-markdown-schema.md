# Supporting Information Markdown schema

Use Pandoc Markdown. Custom style names and control markers are case-sensitive.

## Cover

```markdown
::: {custom-style="SI Cover Label"}
Supplementary Information for:
:::

::: {custom-style="SI Cover Title"}
A Generic Study with Reproducible Supporting Information
:::

{{PAGE_BREAK}}
```

## Discussion and method sections

```markdown
::: {custom-style="SI Section Heading"}
**Discussion S1.** Discussion title.
:::

Body text.

{{PAGE_BREAK}}

::: {custom-style="SI Section Heading"}
**Method S1.** Detailed method title.
:::
```

Start each major `Discussion S#` or `Method S#` unit on a new page when following the source template. The converter colors the entire heading blue and makes the label bold.

## Subheadings and procedure steps

```markdown
::: {custom-style="SI Subheading"}
***a.*** **example metric (EM)**
:::

::: {custom-style="SI Subheading"}
□ **Highlighted procedure**
:::

**Step-1.** Procedure text.
```

Use an outline square `□` for emphasized procedure subheadings. Keep the letter marker bold italic and the descriptive title bold.

## Figures and images

```markdown
{{PAGE_BREAK}}

![](images/figure-s1.png){width=16.2cm}

::: {custom-style="SI Figure Caption"}
**Figure S1.** Short title sentence. Full explanation with ***(a)*** panel descriptions.
:::
```

Use a width attribute when exact image sizing matters. The portrait text width is 16.2 cm; the secondary-width figure profile is approximately 13.3 cm. Images are optional, but captions are not.

## Equations

Use inline math as `$x^2$` and display math as:

```markdown
$$
T(A,B)=\frac{|A\cap B|}{|A\cup B|}
$$
```

The converter emits native Word math.

## Portrait tables

```markdown
::: {custom-style="SI Table Caption"}
**Table S1.** Table title.
:::

| Item | Description | Identifier | Value^a^ | Date |
|:--|:--|:--:|--:|:--:|
| Example | Placeholder row | ID-001 | 0.328 | 2000/01/01 |

::: {custom-style="SI Table Note"}
***^a^*** Explanatory note.
:::
```

## Landscape tables and return to portrait

```markdown
{{SECTION_BREAK_LANDSCAPE}}

::: {custom-style="SI Table Caption"}
**Table S1.** Wide-table title.
:::

| ... |
| ... |

::: {custom-style="SI Table Note"}
***^a^*** Explanatory note.
:::

{{SECTION_BREAK_PORTRAIT}}
```

A section marker starts a new page. Do not place `{{PAGE_BREAK}}` immediately before or after it. Omit the return-to-portrait marker when the wide table is the document's final content.

Carry the full table legend—title sentence plus explanatory text—in the caption paragraph, matching the figure-caption convention; place only footnote-style notes below the table, with one blank line (12 pt) of space before the first note.

## Inline scientific formatting

Preserve `*italic*`, `**bold**`, `***bold italic***`, `<u>underline</u>`, H<sub>2</sub>O, x<sup>2</sup>, Pandoc `~subscript~`, Pandoc `^superscript^`, links, and chemical symbols. Use Markdown emphasis inside captions and table cells as needed.
