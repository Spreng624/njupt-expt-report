#!/usr/bin/env python3
"""Remove personal metadata from a DOCX while preserving the document package."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def sanitize_core(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for path in (
        f"{{{DC}}}title",
        f"{{{DC}}}subject",
        f"{{{DC}}}creator",
        f"{{{DC}}}description",
        f"{{{DC}}}keywords",
        f"{{{CP}}}lastModifiedBy",
        f"{{{CP}}}category",
    ):
        element = root.find(path)
        if element is not None:
            element.text = ""
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def sanitize_custom(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for value in root.findall(f".//{{{VT}}}lpwstr"):
        value.text = ""
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def sanitize(input_path: Path, output_path: Path) -> None:
    temp_path = output_path.with_name(output_path.stem + ".sanitize.tmp")
    with ZipFile(input_path) as source, ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "docProps/core.xml":
                data = sanitize_core(data)
            elif info.filename == "docProps/custom.xml":
                data = sanitize_custom(data)
            target.writestr(info, data)
    os.replace(temp_path, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    args = parser.parse_args()
    output = args.output or args.docx
    sanitize(args.docx.resolve(), output.resolve())
    print(f"sanitized: {output.resolve()}")


if __name__ == "__main__":
    main()
