#!/usr/bin/env python3
"""
normalize.py - CLI entry point for the Security Questionnaire Normalizer.

Usage:
    python normalize.py path/to/file.xlsx > output.json
    python normalize.py path/to/file.docx > output.json
    python normalize.py path/to/file.xlsx --pretty   # human-readable output
    python normalize.py path/to/file.xlsx --validate  # validate after parsing

Outputs a JSON object conforming to schema/output_schema.json.
"""

import argparse
import json
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description="Normalize security questionnaire files (.xlsx, .docx) to JSON."
    )
    parser.add_argument("file", help="Path to the input questionnaire file")
    parser.add_argument(
        "--pretty", action="store_true",
        help="Pretty-print JSON output (default: compact)"
    )
    parser.add_argument(
        "--validate", action="store_true",
        help="Validate output against schema after parsing"
    )
    parser.add_argument(
        "--out", metavar="FILE",
        help="Write output to a file instead of stdout"
    )
    args = parser.parse_args()

    input_path = Path(args.file)

    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    suffix = input_path.suffix.lower()
    if suffix not in (".xlsx", ".docx"):
        print(f"ERROR: Unsupported file type '{suffix}'. Only .xlsx and .docx are supported.", file=sys.stderr)
        sys.exit(1)

    # Parse
    try:
        if suffix == ".xlsx":
            from normalizer.xlsx_parser import parse_xlsx
            result = parse_xlsx(str(input_path))
        else:
            from normalizer.docx_parser import parse_docx
            result = parse_docx(str(input_path))
    except Exception as e:
        print(f"ERROR: Parsing failed: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)

    # Attach top-level metadata
    output = {
        "source_file": str(input_path),
        "questionnaire_name": result.get("questionnaire_name", input_path.stem),
        "version": result.get("version", ""),
        "sections": result.get("sections", []),
    }

    # Serialize
    indent = 2 if args.pretty else None
    json_str = json.dumps(output, indent=indent, ensure_ascii=False)

    if args.out:
        Path(args.out).write_text(json_str, encoding="utf-8")
        print(f"Output written to {args.out}", file=sys.stderr)
    else:
        print(json_str)

    # Optional validation
    if args.validate:
        import tempfile, os
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tmp:
            tmp.write(json_str)
            tmp_path = tmp.name
        try:
            from normalizer.schema_validator import validate_file
            valid = validate_file(tmp_path)
            if not valid:
                sys.exit(1)
        finally:
            os.unlink(tmp_path)

    # Summary to stderr so it doesn't pollute stdout JSON
    total_questions = sum(len(s["questions"]) for s in output["sections"])
    print(
        f"Parsed {len(output['sections'])} section(s), {total_questions} question(s) "
        f"from {input_path.name}",
        file=sys.stderr,
    )


if __name__ == "__main__":
    main()
