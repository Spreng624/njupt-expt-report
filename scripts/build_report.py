#!/usr/bin/env python3
"""Convert a YAML-fronted Markdown report into the NJUPT experiment DOCX template."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = PROJECT_ROOT / "templates" / "实验报告模板.docx"

BODY_FONT = "宋体"
CODE_FONT = "Consolas"
WESTERN_FONT = "Times New Roman"
BODY_SIZE = 10.5
FRAME_CONTENT_WIDTH = 6.2

REQUIRED_KEYS = (
    "course_name",
    "experiment_name",
    "experiment_date",
    "academic_year",
    "term",
    "department",
    "instructor",
    "student_name",
    "student_id",
    "class_name",
    "major",
)

TAIL_HEADINGS = (
    "四、实验小结",
    "五、支撑毕业要求指标点",
    "六、指导教师评语",
)

INLINE_RE = re.compile(r"(\*\*.*?\*\*|`[^`]*`|\[[^\]]+\]\([^)]+\))")
HEADING_NUMBER_RE = re.compile(r"^[一二三四五六七八九十]+、")


class ReportError(ValueError):
    """Raised when the Markdown source does not satisfy the report contract."""


def parse_front_matter(source: str, source_path: Path) -> tuple[dict[str, str], str]:
    lines = source.splitlines()
    if not lines or lines[0].strip() != "---":
        raise ReportError(
            f"{source_path}: Markdown must start with a YAML front-matter block delimited by ---"
        )

    try:
        end = next(i for i in range(1, len(lines)) if lines[i].strip() == "---")
    except StopIteration as exc:
        raise ReportError(f"{source_path}: missing closing --- for YAML front matter") from exc

    metadata: dict[str, str] = {}
    for line_no, raw in enumerate(lines[1:end], start=2):
        stripped = raw.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if ":" not in raw:
            raise ReportError(
                f"{source_path}:{line_no}: metadata must use flat key: value syntax"
            )
        key, value = raw.split(":", 1)
        key = key.strip()
        value = value.strip()
        if value[:1] in {"'", '"'} and value[-1:] == value[:1]:
            value = value[1:-1]
        elif " #" in value:
            value = value.split(" #", 1)[0].rstrip()
        metadata[key] = value

    missing = [key for key in REQUIRED_KEYS if not metadata.get(key)]
    if missing:
        raise ReportError(
            f"{source_path}: missing required metadata keys: {', '.join(missing)}"
        )

    try:
        year, month, day = metadata["experiment_date"].split("-")
        if len(year) != 4 or len(month) != 2 or len(day) != 2:
            raise ValueError
        int(year), int(month), int(day)
    except (ValueError, TypeError) as exc:
        raise ReportError(
            f"{source_path}: experiment_date must use YYYY-MM-DD"
        ) from exc

    metadata.setdefault("experiment_type", "验证")
    metadata.setdefault("experiment_hours", "2")
    metadata.setdefault("experiment_time", metadata["experiment_date"])

    body = "\n".join(lines[end + 1 :])
    return metadata, body


def normalize_heading(text: str) -> str:
    text = text.replace("**", "").strip()
    text = " ".join(text.split())
    text = re.sub(r"(?<=[、])\s+", "", text)
    text = re.sub(r"\s+(?=[）)])", "", text)
    text = re.sub(r"(?<=[（(])\s+", "", text)
    return text


def standalone_bold_text(line: str) -> str | None:
    stripped = line.strip()
    if len(stripped) >= 4 and stripped.startswith("**") and stripped.endswith("**"):
        return normalize_heading(stripped)
    return None


def parse_sections(
    body: str, source_path: Path
) -> list[tuple[str, list[str]]]:
    lines = body.splitlines()
    sections: list[tuple[str, list[str]]] = []
    current_title: str | None = None
    current_lines: list[str] = []
    preamble: list[str] = []

    for line_no, line in enumerate(lines, start=1):
        stripped = line.strip()
        h2 = None
        if re.match(r"^##\s+", stripped):
            h2 = normalize_heading(re.sub(r"^##\s+", "", stripped))
        else:
            bold = standalone_bold_text(stripped)
            if bold and HEADING_NUMBER_RE.match(bold):
                h2 = bold

        if h2:
            if current_title is not None:
                sections.append((current_title, current_lines))
            elif any(value.strip() for value in preamble):
                raise ReportError(
                    f"{source_path}: content appears before the first section heading"
                )
            current_title = h2
            current_lines = []
            continue

        if current_title is None:
            preamble.append(line)
        else:
            current_lines.append(line)

    if current_title is None:
        raise ReportError(f"{source_path}: no section headings found")
    sections.append((current_title, current_lines))

    if not sections[0][0].startswith("一、实验目的"):
        raise ReportError("first section must be 一、实验目的和要求")
    sections[0] = ("一、实验目的和要求", sections[0][1])

    if not sections[1][0].startswith("二、实验环境"):
        raise ReportError("second section must be 二、实验环境(实验设备)")
    sections[1] = ("二、实验环境(实验设备)", sections[1][1])

    tail = tuple(title for title, _ in sections[-3:])
    if tail != TAIL_HEADINGS:
        raise ReportError(
            "the final three headings must be: "
            + " / ".join(TAIL_HEADINGS)
        )

    return sections


def set_run_font(
    run,
    family: str = BODY_FONT,
    size: float = BODY_SIZE,
    bold: bool | None = None,
    italic: bool | None = None,
    color: RGBColor = RGBColor(0, 0, 0),
) -> None:
    run.font.name = family
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    if family == CODE_FONT:
        rfonts.set(qn("w:ascii"), CODE_FONT)
        rfonts.set(qn("w:hAnsi"), CODE_FONT)
        rfonts.set(qn("w:eastAsia"), CODE_FONT)
    else:
        rfonts.set(qn("w:ascii"), WESTERN_FONT)
        rfonts.set(qn("w:hAnsi"), WESTERN_FONT)
        rfonts.set(qn("w:eastAsia"), BODY_FONT)


def add_inline(paragraph, text: str, size: float = BODY_SIZE, bold: bool = False) -> None:
    pos = 0
    for match in INLINE_RE.finditer(text):
        if match.start() > pos:
            set_run_font(paragraph.add_run(text[pos : match.start()]), size=size, bold=bold)
        token = match.group(0)
        if token.startswith("**"):
            set_run_font(paragraph.add_run(token[2:-2]), size=size, bold=True)
        elif token.startswith("`"):
            set_run_font(paragraph.add_run(token[1:-1]), family=CODE_FONT, size=size, bold=bold)
        else:
            label, url = re.match(r"\[([^\]]+)\]\(([^)]+)\)", token).groups()
            set_run_font(paragraph.add_run(f"{label}（{url}）"), size=size, bold=bold)
        pos = match.end()
    if pos < len(text):
        set_run_font(paragraph.add_run(text[pos:]), size=size, bold=bold)


def clear_cell(cell):
    for child in list(cell._tc):
        if child.tag != qn("w:tcPr"):
            cell._tc.remove(child)
    return cell.add_paragraph()


def unique_cells(row) -> list:
    result = []
    seen: set[int] = set()
    for cell in row.cells:
        marker = id(cell._tc)
        if marker not in seen:
            seen.add(marker)
            result.append(cell)
    return result


def set_table_borders(table, color: str = "000000", size: str = "4") -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def set_cell_margins(cell, top: int = 30, start: int = 70, bottom: int = 30, end: int = 70) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def add_code_block(cell, lines: list[str]) -> None:
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.left_indent = Inches(0.12)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "F2F2F2")
    paragraph._p.get_or_add_pPr().append(shd)
    for index, line in enumerate(lines):
        run = paragraph.add_run(line)
        set_run_font(run, family=CODE_FONT, size=BODY_SIZE)
        if index < len(lines) - 1:
            run.add_break()


def add_markdown_table(cell, headers: list[str], rows: list[list[str]]) -> None:
    table = cell.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table)
    for col, value in enumerate(headers):
        header_cell = table.rows[0].cells[col]
        set_cell_margins(header_cell)
        paragraph = clear_cell(header_cell)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(paragraph, value, size=BODY_SIZE, bold=True)
    for row_index, row in enumerate(rows, start=1):
        for col, value in enumerate(row):
            table_cell = table.rows[row_index].cells[col]
            set_cell_margins(table_cell)
            paragraph = clear_cell(table_cell)
            add_inline(paragraph, value, size=BODY_SIZE)
    spacer = cell.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(2)
    spacer.add_run("").font.size = Pt(1)


def add_markdown_image(cell, source: Path, alt: str) -> None:
    if not source.exists():
        raise ReportError(f"image not found: {source}")
    with Image.open(source) as image:
        width_px, height_px = image.size
    width = FRAME_CONTENT_WIDTH
    height = width * height_px / width_px
    if height > 7.2:
        height = 7.2
        width = height * width_px / height_px
    paragraph = cell.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.keep_with_next = True
    picture = paragraph.add_run().add_picture(str(source), width=Inches(width))
    picture._inline.docPr.set("descr", alt)
    picture._inline.docPr.set("title", alt)
    cap = cell.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(6)
    add_inline(cap, alt, size=BODY_SIZE)
    for run in cap.runs:
        run.font.italic = True


def render_section_content(cell, lines: list[str], source_dir: Path) -> None:
    index = 0
    while index < len(lines):
        stripped = lines[index].strip()
        if not stripped or stripped.startswith("<!--"):
            index += 1
            continue

        if stripped.startswith("```"):
            index += 1
            code_lines = []
            while index < len(lines) and not lines[index].strip().startswith("```"):
                code_lines.append(lines[index])
                index += 1
            index += 1
            add_code_block(cell, code_lines)
            continue

        image_match = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", stripped)
        if image_match:
            alt, relative = image_match.groups()
            add_markdown_image(cell, (source_dir / relative).resolve(), alt or "运行截图")
            index += 1
            continue

        if (
            stripped.startswith("|")
            and index + 1 < len(lines)
            and re.match(r"^\|?\s*:?-{3,}", lines[index + 1].strip())
        ):
            headers = [value.strip() for value in stripped.strip("|").split("|")]
            index += 2
            rows = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                row = [
                    value.strip()
                    for value in lines[index].strip().strip("|").split("|")
                ]
                row += [""] * max(0, len(headers) - len(row))
                rows.append(row[: len(headers)])
                index += 1
            add_markdown_table(cell, headers, rows)
            continue

        subheading = None
        if re.match(r"^#{3,6}\s+", stripped):
            subheading = normalize_heading(re.sub(r"^#{3,6}\s+", "", stripped))
        else:
            subheading = standalone_bold_text(stripped)
            if subheading and HEADING_NUMBER_RE.match(subheading):
                subheading = None
        if subheading:
            paragraph = cell.add_paragraph(style="Normal")
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.keep_with_next = True
            add_inline(paragraph, subheading, size=BODY_SIZE, bold=True)
            index += 1
            continue

        bullet = re.match(r"^[-*]\s+(.*)$", stripped)
        numbered = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if bullet or numbered:
            text = bullet.group(1) if bullet else f"{numbered.group(1)}. {numbered.group(2)}"
            marker = "• " if bullet else ""
            paragraph = cell.add_paragraph(style="Normal")
            paragraph.paragraph_format.left_indent = Inches(0.3)
            paragraph.paragraph_format.first_line_indent = Inches(-0.18)
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.line_spacing = 1.2
            add_inline(paragraph, marker + text, size=BODY_SIZE)
            index += 1
            continue

        if stripped.startswith(">"):
            paragraph = cell.add_paragraph(style="Normal")
            paragraph.paragraph_format.left_indent = Inches(0.18)
            paragraph.paragraph_format.right_indent = Inches(0.12)
            paragraph.paragraph_format.space_after = Pt(3)
            add_inline(paragraph, stripped.lstrip("> ").strip(), size=BODY_SIZE)
            for run in paragraph.runs:
                run.font.italic = True
            index += 1
            continue

        paragraph = cell.add_paragraph(style="Normal")
        paragraph.paragraph_format.first_line_indent = Inches(0.25)
        paragraph.paragraph_format.line_spacing = 1.25
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(paragraph, stripped, size=BODY_SIZE)
        index += 1


def iter_paragraphs(container):
    yield from container.paragraphs
    for table in container.tables:
        for row in table.rows:
            for cell in unique_cells(row):
                yield from iter_paragraphs(cell)


def find_cell_for_paragraph(document: Document, paragraph):
    tc = paragraph._p.getparent()

    def walk(container):
        for table in container.tables:
            for row in table.rows:
                for cell in unique_cells(row):
                    if cell._tc is tc:
                        return cell
                    found = walk(cell)
                    if found is not None:
                        return found
        return None

    cell = walk(document)
    if cell is None:
        raise ReportError("placeholder paragraph is not inside a table cell")
    return cell


def find_token_paragraph(document: Document, token: str):
    for paragraph in iter_paragraphs(document):
        if token in paragraph.text:
            return paragraph
    raise ReportError(f"template placeholder is missing: {token}")


def collect_placeholders(document: Document) -> set[str]:
    found: set[str] = set()
    for paragraph in iter_paragraphs(document):
        found.update(re.findall(r"\{\{([A-Z0-9_]+)\}\}", paragraph.text))
    return found


def replace_tokens_in_paragraph(
    paragraph,
    values: dict[str, str],
    *,
    remove_last_break: bool = False,
) -> bool:
    nodes = list(paragraph._p.iter(qn("w:t")))
    full_text = "".join(node.text or "" for node in nodes)
    matches = list(re.finditer(r"\{\{([A-Z0-9_]+)\}\}", full_text))
    changed = False

    for match in reversed(matches):
        key = match.group(1)
        if key not in values:
            continue
        replacement = values[key]
        start, end = match.span()
        offset = 0
        first = True
        for node in nodes:
            node_text = node.text or ""
            node_start = offset
            node_end = offset + len(node_text)
            offset = node_end
            if node_end <= start or node_start >= end:
                continue
            local_start = max(start, node_start) - node_start
            local_end = min(end, node_end) - node_start
            if first:
                node.text = (
                    node_text[:local_start] + replacement + node_text[local_end:]
                )
                first = False
            else:
                node.text = node_text[:local_start] + node_text[local_end:]
            node.set(qn("xml:space"), "preserve")
        changed = True

    if changed and remove_last_break:
        breaks = list(paragraph._p.iter(qn("w:br")))
        if breaks:
            last_break = breaks[-1]
            last_break.getparent().remove(last_break)
    return changed


def replace_metadata_placeholders(document: Document, metadata: dict[str, str]) -> None:
    values = {
        "COURSE": metadata["course_name"],
        "EXPERIMENT": metadata["experiment_name"],
        "Y": metadata["experiment_date"].split("-")[0],
        "M": metadata["experiment_date"].split("-")[1],
        "D": metadata["experiment_date"].split("-")[2],
        "YEAR": metadata["academic_year"],
        "TERM": metadata["term"],
        "DEPT": metadata["department"],
        "TEACHER": metadata["instructor"],
        "NAME": metadata["student_name"],
        "COLLEGE": metadata.get("college", metadata["department"]),
        "MAJOR": metadata["major"],
        "ID": metadata["student_id"],
        "TYPE": metadata["experiment_type"],
        "HOURS": metadata["experiment_hours"],
        "TIME": metadata["experiment_time"],
    }
    for paragraph in iter_paragraphs(document):
        if "{{" in paragraph.text:
            replace_tokens_in_paragraph(paragraph, values)


def split_markdown_subsections(lines: list[str], label: str):
    groups: list[tuple[str, list[str]]] = []
    current_title = None
    current_lines: list[str] = []
    for line in lines:
        stripped = line.strip()
        title = None
        if re.match(r"^###\s+", stripped):
            title = normalize_heading(re.sub(r"^###\s+", "", stripped))
        else:
            bold = standalone_bold_text(stripped)
            if bold and not HEADING_NUMBER_RE.match(bold):
                title = bold
        if title:
            if current_title is not None:
                groups.append((current_title, current_lines))
            current_title = title
            current_lines = []
        elif current_title is not None:
            current_lines.append(line)
        elif stripped:
            raise ReportError(f"{label}: content appears before the first third-level heading")
    if current_title is not None:
        groups.append((current_title, current_lines))
    if len(groups) != 3:
        raise ReportError(f"{label}: expected exactly three third-level subsections")
    return groups


def insert_content_after_anchor(
    document: Document,
    paragraph,
    lines: list[str],
    source_dir: Path,
    remove_anchor: bool,
) -> None:
    cell = find_cell_for_paragraph(document, paragraph)
    before = list(cell._tc)
    render_section_content(cell, lines, source_dir)
    after = list(cell._tc)
    added = after[len(before) :]
    for element in added:
        cell._tc.remove(element)

    cursor = paragraph._p
    for element in added:
        cursor.addnext(element)
        cursor = element
    if remove_anchor:
        paragraph._p.getparent().remove(paragraph._p)


def fill_body_placeholders(
    document: Document,
    sections: list[tuple[str, list[str]]],
    source_dir: Path,
) -> None:
    if len(sections) != 6:
        raise ReportError("Markdown body must contain exactly six second-level sections")
    if any(line.strip() for line in sections[5][1]):
        raise ReportError(
            "第六节正文没有 B6 占位符；请清空该节正文，或先在模板中增加 {{B6}}"
        )

    # Sections 1-3 use a placeholder embedded after the fixed/custom heading.
    for index, token in ((0, "{{B1}}"), (1, "{{B2}}"), (2, "{{B3}}")):
        paragraph = find_token_paragraph(document, token)
        if token == "{{B3}}":
            replace_tokens_in_paragraph(
                paragraph,
                {"T3": sections[2][0], "B3": ""},
                remove_last_break=True,
            )
        else:
            replace_tokens_in_paragraph(
                paragraph,
                {token[2:-2]: ""},
                remove_last_break=True,
            )
        insert_content_after_anchor(
            document,
            paragraph,
            sections[index][1],
            source_dir,
            remove_anchor=False,
        )

    section4 = split_markdown_subsections(sections[3][1], sections[3][0])
    section5 = split_markdown_subsections(sections[4][1], sections[4][0])
    for token, (_, lines) in zip(
        ("{{B4_1}}", "{{B4_2}}", "{{B4_3}}"),
        section4,
    ):
        paragraph = find_token_paragraph(document, token)
        replace_tokens_in_paragraph(paragraph, {token[2:-2]: ""})
        insert_content_after_anchor(
            document, paragraph, lines, source_dir, remove_anchor=True
        )
    for token, (_, lines) in zip(
        ("{{B5_1}}", "{{B5_2}}", "{{B5_3}}"),
        section5,
    ):
        paragraph = find_token_paragraph(document, token)
        replace_tokens_in_paragraph(paragraph, {token[2:-2]: ""})
        insert_content_after_anchor(
            document, paragraph, lines, source_dir, remove_anchor=True
        )


def build_report(markdown_path: Path, output_path: Path, template_path: Path) -> None:
    source = markdown_path.read_text(encoding="utf-8")
    metadata, body = parse_front_matter(source, markdown_path)
    sections = parse_sections(body, markdown_path)

    document = Document(template_path)
    expected = {
        "COURSE",
        "EXPERIMENT",
        "Y",
        "M",
        "D",
        "YEAR",
        "TERM",
        "DEPT",
        "TEACHER",
        "NAME",
        "ID",
        "COLLEGE",
        "MAJOR",
        "TYPE",
        "HOURS",
        "TIME",
        "B1",
        "B2",
        "T3",
        "B3",
        "B4_1",
        "B4_2",
        "B4_3",
        "B5_1",
        "B5_2",
        "B5_3",
    }
    placeholders = collect_placeholders(document)
    missing = sorted(expected - placeholders)
    if missing:
        raise ReportError(f"template is missing placeholders: {', '.join(missing)}")

    replace_metadata_placeholders(document, metadata)
    fill_body_placeholders(document, sections, markdown_path.parent)

    remaining = collect_placeholders(document)
    if remaining:
        raise ReportError(
            "unresolved placeholders remain: " + ", ".join(sorted(remaining))
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output_path.with_name(output_path.stem + ".build.tmp.docx")
    document.save(temp_output)
    try:
        with ZipFile(template_path) as template_zip, ZipFile(temp_output) as built_zip:
            template_names = set(template_zip.namelist())
            built_names = set(built_zip.namelist())
            added_parts = built_names - template_names
            allowed_changed = {"word/document.xml"}
            if added_parts:
                allowed_changed.update(
                    {"word/_rels/document.xml.rels", "[Content_Types].xml"}
                )

            with ZipFile(output_path, "w", ZIP_DEFLATED) as final_zip:
                for info in template_zip.infolist():
                    data = (
                        built_zip.read(info.filename)
                        if info.filename in allowed_changed
                        else template_zip.read(info.filename)
                    )
                    final_zip.writestr(info, data)
                for name in sorted(added_parts):
                    final_zip.writestr(name, built_zip.read(name))
    finally:
        temp_output.unlink(missing_ok=True)
    print(f"created: {output_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path, help="source Markdown with YAML front matter")
    parser.add_argument("-o", "--output", type=Path, help="output DOCX path")
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help=f"DOCX template (default: {DEFAULT_TEMPLATE})",
    )
    args = parser.parse_args()

    output = args.output or args.markdown.with_suffix(".docx")
    try:
        build_report(args.markdown.resolve(), output.resolve(), args.template.resolve())
    except (OSError, ReportError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
