#!/usr/bin/env python3
"""Validate a Markdown report before DOCX conversion."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from build_report import (
    ReportError,
    parse_front_matter,
    parse_sections,
    split_markdown_subsections,
)


def validate(markdown_path: Path) -> dict:
    source = markdown_path.read_text(encoding="utf-8")
    metadata, body = parse_front_matter(source, markdown_path)
    sections = parse_sections(body, markdown_path)

    if len(sections) != 6:
        raise ReportError(f"expected exactly six second-level sections, found {len(sections)}")

    for index, (title, lines) in enumerate(sections[:5], start=1):
        if not any(line.strip() for line in lines):
            raise ReportError(f"section {index} is empty: {title}")

    section4 = split_markdown_subsections(sections[3][1], sections[3][0])
    section5 = split_markdown_subsections(sections[4][1], sections[4][0])
    if any(line.strip() for line in sections[5][1]):
        raise ReportError(
            "section 6 must be empty in the source because the template has no {{B6}} slot"
        )

    fence_count = sum(1 for line in body.splitlines() if line.strip().startswith("```"))
    if fence_count % 2:
        raise ReportError("unclosed fenced code block")

    missing_images = []
    image_paths: set[str] = set()
    for line in body.splitlines():
        match = re.match(r"!\[[^\]]*\]\(([^)]+)\)", line.strip())
        if not match:
            continue
        relative = match.group(1).strip()
        image_paths.add(relative)
        target = (markdown_path.parent / relative).resolve()
        if not target.exists():
            missing_images.append(str(target))

    shot_placeholders = re.findall(
        r"<!--\s*SHOT:\s*([^|>]+?)(?:\s*\|[^>]*)?\s*-->",
        body,
    )
    for relative in shot_placeholders:
        relative = relative.strip()
        if relative not in image_paths:
            raise ReportError(
                f"screenshot marker has no image line for: {relative}"
            )
        target = (markdown_path.parent / relative).resolve()
        if not target.exists() and str(target) not in missing_images:
            missing_images.append(str(target))
    if missing_images:
        raise ReportError(
            "pending screenshots: "
            + ", ".join(missing_images)
            + ". Complete the experiment and save each screenshot at the listed path."
        )

    return {
        "file": str(markdown_path),
        "metadata": metadata,
        "sections": [title for title, _ in sections],
        "section4_subsections": [title for title, _ in section4],
        "section5_subsections": [title for title, _ in section5],
        "screenshot_placeholders": shot_placeholders,
        "errors": [],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()

    try:
        summary = validate(args.markdown.resolve())
    except (OSError, ReportError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    print(
        "OK: six sections, "
        f"{len(summary['section4_subsections'])} section-4 subsections, "
        f"{len(summary['section5_subsections'])} section-5 subsections"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
