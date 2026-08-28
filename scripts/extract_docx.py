#!/usr/bin/env python3
"""
Extract structured content from a .docx file: paragraphs (with heading levels)
and tables, preserving document order. Outputs JSON for downstream analysis.

Usage:
    python extract_docx.py <input.docx> [--output <output.json>] [--min-cell-chars <n>]

Output JSON schema:
{
    "metadata": {"file": str, "paragraph_count": int, "table_count": int},
    "items": [
        {"kind": "paragraph", "text": str, "style": str, "heading_level": int|null,
         "idx": int},
        {"kind": "table", "rows": [[str,...], ...], "idx": int}
    ]
}
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print(
        "ERROR: python-docx is required. Install it first:\n"
        "  pip install python-docx",
        file=sys.stderr,
    )
    sys.exit(1)


HEADING_STYLES = {
    "heading 1": 1, "heading 2": 2, "heading 3": 3, "heading 4": 4,
    "heading 5": 5, "heading 6": 6, "heading 7": 7, "heading 8": 8,
    "heading 9": 9, "标题 1": 1, "标题 2": 2, "标题 3": 3, "标题 4": 4,
    "标题 5": 5, "标题 6": 6, "标题 7": 7, "标题 8": 8, "标题 9": 9,
}


def iter_block_items(parent):
    """Yield paragraphs and tables from the document body in document order."""
    from docx.oxml.ns import qn

    for child in parent.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def table_to_rows(table: Table, min_cell_chars: int = 0):
    """Convert a python-docx Table into a list of rows of cell text."""
    rows = []
    for row in table.rows:
        cells = []
        seen = set()
        for cell in row.cells:
            # Merged cells repeat the same _tc element; deduplicate by element id
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            text = "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
            # Skip cells that are fully empty after cleaning
            if min_cell_chars > 0 and len(text) < min_cell_chars:
                text = ""
            cells.append(text)
        # Skip fully empty rows
        if any(c.strip() for c in cells):
            rows.append(cells)
    return rows


def heading_level_of(paragraph: Paragraph):
    style_name = (paragraph.style.name or "").lower()
    for key, level in HEADING_STYLES.items():
        if style_name == key:
            return level
    # Some docs use direct outline levels
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.outlineLvl is not None:
        try:
            return int(pPr.outlineLvl.val) + 1
        except (ValueError, TypeError):
            pass
    return None


def main():
    parser = argparse.ArgumentParser(description="Extract text and tables from a .docx file")
    parser.add_argument("input", help="Path to the .docx file")
    parser.add_argument("--output", "-o", help="Output JSON file path (default: stdout)")
    parser.add_argument(
        "--min-cell-chars",
        type=int,
        default=1,
        help="Minimum characters for a table cell to be kept (default: 1)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)
    if input_path.suffix.lower() != ".docx":
        print(
            f"WARNING: '{input_path.suffix}' is not .docx; attempting to read anyway",
            file=sys.stderr,
        )

    try:
        doc = Document(str(input_path))
    except Exception as exc:
        print(f"ERROR: Failed to open document: {exc}", file=sys.stderr)
        sys.exit(1)

    items = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            items.append(
                {
                    "kind": "paragraph",
                    "text": text,
                    "style": block.style.name,
                    "heading_level": heading_level_of(block),
                    "idx": len(items),
                }
            )
        elif isinstance(block, Table):
            rows = table_to_rows(block, min_cell_chars=args.min_cell_chars)
            items.append({"kind": "table", "rows": rows, "idx": len(items)})

    result = {
        "metadata": {
            "file": str(input_path),
            "paragraph_count": sum(1 for i in items if i["kind"] == "paragraph"),
            "table_count": sum(1 for i in items if i["kind"] == "table"),
        },
        "items": items,
    }

    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        Path(args.output).write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"Extracted {result['metadata']['paragraph_count']} paragraphs, "
              f"{result['metadata']['table_count']} tables -> {args.output}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
