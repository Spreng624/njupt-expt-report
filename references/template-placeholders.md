# Template Placeholder Contract

## Principle

The DOCX template controls layout. Conversion replaces placeholder text and inserts content only at the marked body anchors.

Do not change table structure, cell merges, row heights, section breaks, page breaks, borders, scoring rows, or page margins during conversion.

Package preservation is strict: reuse every template package part byte-for-byte except
`word/document.xml`. If Markdown contains images, the added media part and the corresponding
relationship/content-type entries are the only additional changes.

## Metadata Placeholders

| Placeholder | YAML field |
|---|---|
| `{{COURSE}}` | `course_name` |
| `{{EXPERIMENT}}` | `experiment_name` |
| `{{Y}}` / `{{M}}` / `{{D}}` | `experiment_date` |
| `{{YEAR}}` / `{{TERM}}` | `academic_year` / `term` |
| `{{DEPT}}` | `department` |
| `{{TEACHER}}` | `instructor` |
| `{{NAME}}` | `student_name` |
| `{{ID}}` | `student_id` |
| `{{COLLEGE}}` | `college` |
| `{{MAJOR}}` | `major` |
| `{{TYPE}}` | `experiment_type` |
| `{{HOURS}}` | `experiment_hours` |
| `{{TIME}}` | `experiment_time` |

## Body Anchors

| Placeholder | Destination |
|---|---|
| `{{B1}}` | Section 1 body |
| `{{B2}}` | Section 2 body |
| `{{T3}}` | Section 3 title |
| `{{B3}}` | Section 3 body |
| `{{B4_1}}`–`{{B4_3}}` | Three subsection bodies in Section 4 |
| `{{B5_1}}`–`{{B5_3}}` | Three subsection bodies in Section 5 |

Section 6 and the scoring table are preserved without a body placeholder.

## Formatting Inheritance

For metadata replacement:

- preserve the placeholder paragraph properties;
- preserve the first placeholder run’s font, size, bold and italic formatting;
- replace text spanning multiple Word runs as one logical token.

For body insertion:

- do not alter the enclosing cell or table row;
- insert generated elements immediately after the body anchor;
- use five-point body text and bold five-point third-level headings;
- remove only the placeholder token or placeholder paragraph being consumed.
