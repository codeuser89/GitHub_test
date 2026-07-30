#!/usr/bin/env python3
"""
Convert a PowerPoint presentation (.pptx) to a template (.potx).

PPTX and POTX files share the same ZIP/XML structure; the primary difference is
the content type declared for presentation.xml in [Content_Types].xml. This script
copies the source archive, updates those declarations, and writes a valid .potx file.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

CONTENT_TYPES_PATH = "[Content_Types].xml"
APP_XML_PATH = "docProps/app.xml"

PRESENTATION_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)
TEMPLATE_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
)
PRESENTATION_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.presentationml.template"


def _replace_presentation_content_types(xml_text: str) -> tuple[str, int]:
    """Return updated XML and the number of content-type substitutions made."""
    replacements = (
        (PRESENTATION_MAIN_CT, TEMPLATE_MAIN_CT),
        (PRESENTATION_CT, TEMPLATE_CT),
    )
    count = 0
    updated = xml_text
    for old, new in replacements:
        occurrences = updated.count(old)
        if occurrences:
            updated = updated.replace(old, new)
            count += occurrences
    return updated, count


def _ensure_template_flag(app_xml: str) -> str:
    """
    Ensure docProps/app.xml marks the file as a template when possible.

    PowerPoint stores this under the extended-properties namespace. If the file
    already contains a Template element, leave it unchanged.
    """
    if re.search(r"<[^:>]*:?Template\b", app_xml):
        return app_xml

    # Insert before closing Properties tag (works for default and prefixed namespaces).
    return re.sub(
        r"(</[^>]*Properties>)",
        r"<Template>Yes</Template>\1",
        app_xml,
        count=1,
    )


def convert_pptx_to_potx(source: Path, destination: Path) -> None:
    """Convert source .pptx to destination .potx."""
    if not source.is_file():
        raise FileNotFoundError(f"Input file not found: {source}")

    if source.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, got: {source.name}")

    destination.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source, "r") as src_zip, zipfile.ZipFile(
        destination, "w", compression=zipfile.ZIP_DEFLATED
    ) as dst_zip:
        content_types_updated = False
        replacements = 0

        for item in src_zip.infolist():
            data = src_zip.read(item.filename)

            if item.filename == CONTENT_TYPES_PATH:
                text = data.decode("utf-8")
                updated, replacements = _replace_presentation_content_types(text)
                if replacements == 0:
                    raise ValueError(
                        f"{CONTENT_TYPES_PATH} does not contain a presentation "
                        "content type; is this a valid .pptx file?"
                    )
                data = updated.encode("utf-8")
                content_types_updated = True
            elif item.filename == APP_XML_PATH:
                text = data.decode("utf-8")
                data = _ensure_template_flag(text).encode("utf-8")

            dst_zip.writestr(item, data)

        if not content_types_updated:
            raise ValueError(
                f"Archive is missing required part: {CONTENT_TYPES_PATH}"
            )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a PowerPoint presentation (.pptx) to a template (.potx).",
        epilog=(
            "Example:\n"
            "  python pptx_to_potx.py deck.pptx -o company-brand.potx"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Path to the source .pptx file",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path for the output .potx file (default: same name with .potx extension)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    output = args.output or args.input.with_suffix(".potx")
    if output.suffix.lower() != ".potx":
        output = output.with_suffix(".potx")

    try:
        convert_pptx_to_potx(args.input, output)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Created template: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
