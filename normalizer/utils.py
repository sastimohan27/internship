"""
utils.py - Shared utility functions used across parsers.
"""

import re
import unicodedata


def clean_text(value) -> str:
    """Coerce any cell value to a clean stripped string, preserving newlines."""
    if value is None:
        return ""
    text = str(value).strip()
    # Normalize unicode (e.g. non-breaking spaces from Excel)
    text = unicodedata.normalize("NFKC", text)
    # Collapse multiple horizontal spaces but preserve newlines
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    # Normalize multiple newlines to a single newline
    text = re.sub(r"\n+", "\n", text)
    # Strip whitespace from the beginning and end of each line
    text = "\n".join(line.strip() for line in text.split("\n"))
    return text.strip()



def is_blank(text: str) -> bool:
    return not text.strip()


def make_question_id(section_id: str, index: int) -> str:
    """Generate a question ID like SEC1-Q3."""
    return f"{section_id}-Q{index}"


def make_section_id(index: int) -> str:
    """Fallback section ID when none is detected."""
    return f"SEC{index}"


# Rows that look like metadata/admin noise rather than questions
NOISE_PATTERNS = [
    r"^instructions?",
    r"^overview",
    r"^version\b",
    r"^date\b",
    r"^owner\b",
    r"^comments?\b",
    r"^please (complete|fill|read|answer)",
    r"^this (document|form|sheet|tab|worksheet)",
    r"^note[:\s]",
    r"^legend\b",
    r"^color\b",
    r"^the cells within",
    r"^step \d+",
]

_noise_re = re.compile("|".join(NOISE_PATTERNS), re.IGNORECASE)



def is_noise_row(text: str) -> bool:
    """Return True if this row is metadata/instructions, not a question."""
    return bool(_noise_re.match(text.strip()))
