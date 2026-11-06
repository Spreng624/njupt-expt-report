---
name: njupt-expt-report
description: Author NJUPT experiment reports in Markdown and convert them to the retained DOCX template using placeholder replacement. Use when a request concerns Nanjing University of Posts and Telecommunications experiment report formatting, YAML report metadata, Markdown-to-DOCX conversion, or preserving the supplied Word report template.
---

# NJUPT Experiment Report

## Outcome

Produce a complete Markdown experiment report first, validate it, then convert the approved Markdown into a DOCX using the fixed template at `templates/实验报告模板.docx`.

The template is authoritative. Conversion must replace placeholders and insert body content at existing anchors. Do not rebuild the cover, tables, rows, merges, page breaks, or scoring area.

## Mandatory Workflow

1. Author the experiment report in Markdown.
2. Put all cover metadata in YAML front matter.
3. Finish all experiment content before creating the DOCX.
4. Run:
   ```powershell
   python scripts/validate_markdown.py path/to/report.md
   ```
5. Convert only after Markdown validation passes:
   ```powershell
   python scripts/build_report.py path/to/report.md -o output/report.docx
   ```
6. Validate the generated document:
   ```powershell
   python scripts/validate_docx.py output/report.docx
   ```
7. Render and inspect the DOCX before delivery:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/render_preview.ps1 `
     -InputDocx output/report.docx `
     -OutputDirectory output/preview
   ```

The one-command equivalent, which performs Markdown validation, conversion, and DOCX validation, is:

```powershell
powershell -ExecutionPolicy Bypass -File convert.ps1 `
  -Markdown path/to/report.md `
  -Output output/report.docx
```

## Non-Negotiable Source Rules

- Use exactly six second-level sections.
- The first two and last three headings are fixed.
- The third section title is supplied through `{{T3}}`.
- Sections 4 and 5 each require exactly three third-level subsections.
- Section 6 remains empty in Markdown; the template supplies the instructor evaluation table.
- Do not invent experimental results, screenshots, flags, or measurements.
- Read [references/markdown-authoring.md](references/markdown-authoring.md) before writing or repairing a report.

## Conversion Boundaries

- Work only in this project directory unless the user explicitly requests another destination.
- Use only the project snapshot `templates/实验报告模板.docx` for conversion.
- Read [references/template-placeholders.md](references/template-placeholders.md) before changing the template or placeholder mapping.
- Never delete or clone body rows.
- Never modify table widths, row heights, cell merges, page breaks, section breaks, or scoring rows during conversion.
- Preserve every unchanged OOXML package part byte-for-byte from the template. For text-only
  reports, `word/document.xml` is the only changed part. Images may additionally add media,
  content-type, and relationship entries.
- Do not apply Word `Title`, `Subtitle`, or `Heading 1–9` styles.
- Replace text across runs while preserving the placeholder paragraph’s `pPr` and first run’s `rPr`.
- Insert Markdown-generated content only at `{{B1}}`, `{{B2}}`, `{{B3}}`,
  `{{B4_1}}`–`{{B4_3}}`, and `{{B5_1}}`–`{{B5_3}}`.
- If the DOCX has a defect, edit the Markdown and regenerate. Do not patch the generated DOCX in place.

## Validation Gates

Markdown is ready only when:

- all required YAML metadata is present;
- the report has exactly six allowed second-level sections;
- sections 1–5 contain substantive content;
- sections 4 and 5 each contain exactly three third-level subsections;
- section 6 is empty;
- fenced code blocks and image paths are valid.

DOCX is ready only when:

- `validate_docx.py` passes;
- all placeholders are gone;
- the document retains the template’s 13-column body table and scoring area;
- the rendered pages have no clipping, overlap, blank pages, or broken borders.
