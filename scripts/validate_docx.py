#!/usr/bin/env python3
"""Validate structural and typography rules for generated NJUPT reports."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


BODY_SIZE = 10.5
SECOND_LEVEL_SIZE = 14
COVER_TITLE_SIZE = 42
TAIL_HEADINGS = (
    "四、实验小结",
    "五、支撑毕业要求指标点",
    "六、指导教师评语",
)


def unique_cells(row):
    cells = []
    seen = set()
    for cell in row.cells:
        marker = id(cell._tc)
        if marker not in seen:
            seen.add(marker)
            cells.append(cell)
    return cells


def visible_runs(paragraph):
    return [run for run in paragraph.runs if run.text.strip()]


def collect_errors(path: Path) -> tuple[list[str], dict]:
    document = Document(path)
    errors: list[str] = []

    used_styles = {
        paragraph.style.name
        for table in document.tables
        for row in table.rows
        for cell in unique_cells(row)
        for paragraph in cell.paragraphs
    }
    used_styles.update(paragraph.style.name for paragraph in document.paragraphs)
    forbidden_used = sorted(
        name
        for name in used_styles
        if name in {"Title", "Subtitle"} or name.startswith("Heading")
    )
    if forbidden_used:
        errors.append(f"forbidden paragraph styles are used: {forbidden_used}")

    forbidden_defined = sorted(
        style.name
        for style in document.styles
        if style.name in {"Title", "Subtitle"} or style.name.startswith("Heading")
    )
    if forbidden_defined:
        errors.append(f"forbidden paragraph styles are defined: {forbidden_defined}")

    all_text = "\n".join(
        paragraph.text
        for table in document.tables
        for row in table.rows
        for cell in unique_cells(row)
        for paragraph in cell.paragraphs
    )
    leftover = re.findall(r"\{\{[A-Z0-9_]+\}\}", all_text)
    if leftover:
        errors.append(f"unresolved placeholders remain: {sorted(set(leftover))}")

    if len(document.tables) < 3:
        errors.append("expected cover tables and one body table")
        return errors, {"used_styles": sorted(used_styles)}

    section = document.sections[0]
    if (
        abs(section.page_width.inches - 8.27) > 0.02
        or abs(section.page_height.inches - 11.69) > 0.02
    ):
        errors.append("page size must be A4 portrait")

    cover_title = next(
        (
            paragraph
            for paragraph in document.paragraphs
            if re.sub(r"\s+", "", paragraph.text) == "实验报告"
        ),
        None,
    )
    if cover_title is None:
        errors.append("cover title 实验报告 is missing")
    elif not any(
        run.font.size and abs(run.font.size.pt - COVER_TITLE_SIZE) < 0.01
        for run in visible_runs(cover_title)
    ):
        errors.append("cover title must use 42 pt (初号)")

    body = document.tables[2]
    if len(body.columns) != 13:
        errors.append(f"body table must use the template's 13-column grid, got {len(body.columns)}")

    row_texts = []
    for row in body.rows:
        cells = unique_cells(row)
        row_texts.append(cells[0].text.strip() if cells else "")

    if not row_texts or re.sub(r"\s+", "", row_texts[0]) != "实验报告":
        errors.append("body table must start with 实验报告")
    upper_rows = "\n".join(
        "\n".join(cell.text for cell in unique_cells(row))
        for row in body.rows[:3]
    )
    for label in ("实验名称", "指导教师", "实验类型", "实验学时", "实验时间"):
        if label not in upper_rows:
            errors.append(f"body metadata label is missing: {label}")

    scoring_index = next(
        (index for index, row in enumerate(body.rows) if "评分细则" in unique_cells(row)[0].text),
        None,
    )
    if scoring_index is None:
        errors.append("评分细则 section is missing")
        scoring_index = len(body.rows)

    section_rows = body.rows[3:scoring_index]
    section_titles = [unique_cells(row)[0].text.strip().splitlines()[0] for row in section_rows]
    if len(section_titles) != 6:
        errors.append(f"expected exactly six report sections, found {len(section_titles)}")
    if section_titles[:2] not in (
        ["一、实验目的和要求", "二、实验环境"],
        ["一、实验目的和要求", "二、实验环境(实验设备)"],
    ):
        errors.append(f"unexpected first two headings: {section_titles[:2]}")
    if tuple(section_titles[-3:]) != TAIL_HEADINGS:
        errors.append(f"unexpected final headings: {section_titles[-3:]}")

    for row, title in zip(section_rows, section_titles):
        paragraph = unique_cells(row)[0].paragraphs[0]
        size = next(
            (
                run.font.size.pt
                for run in visible_runs(paragraph)
                if run.font.size is not None
            ),
            None,
        )
        bold = all(bool(run.bold) for run in visible_runs(paragraph))
        if size is None or abs(size - SECOND_LEVEL_SIZE) > 0.01:
            errors.append(f"second-level heading is not 14 pt: {title}")
        if not bold:
            errors.append(f"second-level heading is not bold: {title}")

    body_sizes = set()
    for row in section_rows:
        cell = unique_cells(row)[0]
        for paragraph in cell.paragraphs[1:]:
            for run in visible_runs(paragraph):
                if run.font.size:
                    body_sizes.add(round(run.font.size.pt, 1))
    unexpected = sorted(size for size in body_sizes if abs(size - BODY_SIZE) > 0.01)
    if unexpected:
        errors.append(f"body/subheading sizes must be 10.5 pt, found {unexpected}")

    summary = {
        "file": str(path),
        "slides": None,
        "page_size_in": [round(section.page_width.inches, 3), round(section.page_height.inches, 3)],
        "cover_tables": 2,
        "body_table_columns": len(body.columns),
        "section_titles": section_titles,
        "used_styles": sorted(used_styles),
        "body_content_sizes": sorted(body_sizes),
        "errors": errors,
    }
    return errors, summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", type=Path, help="optional JSON report path")
    args = parser.parse_args()

    errors, summary = collect_errors(args.docx.resolve())
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    if errors:
        for error in errors:
            print(f"ERROR: {error}", file=sys.stderr)
        return 1
    print(
        "OK: "
        f"{summary['body_table_columns']}-column template, "
        f"{len(summary['section_titles'])} sections, "
        "no heading/title styles"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
