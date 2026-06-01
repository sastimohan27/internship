"""
schema_validator.py - Validate a generated JSON output against the schema.

Usage:
    python normalizer/schema_validator.py output.json
"""

import json
import sys
from pathlib import Path

import jsonschema


SCHEMA_PATH = Path(__file__).parent.parent / "schema" / "output_schema.json"


def validate_file(json_path: str) -> bool:
    """
    Load and validate a JSON file against the output schema.
    Returns True if valid, prints errors and returns False otherwise.
    """
    try:
        with open(SCHEMA_PATH) as f:
            schema = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: Schema not found at {SCHEMA_PATH}", file=sys.stderr)
        return False

    try:
        with open(json_path) as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"ERROR: Could not load {json_path}: {e}", file=sys.stderr)
        return False

    validator = jsonschema.Draft7Validator(schema)
    errors = list(validator.iter_errors(data))

    if not errors:
        print("VALID")
        return True

    print(f"INVALID — {len(errors)} error(s) found:\n")
    for err in errors:
        path = " -> ".join(str(p) for p in err.absolute_path) or "(root)"
        print(f"  [{path}] {err.message}")
    return False


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python normalizer/schema_validator.py <output.json>")
        sys.exit(1)
    ok = validate_file(sys.argv[1])
    sys.exit(0 if ok else 1)
