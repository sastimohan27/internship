"""
detector.py - Heuristics for detecting questions and response types.

Assumption: We can't train a model in a day, so we use deterministic rules.
These rules were written by inspecting HECVAT, CSA CAIQ, and NIST SP 800-53
formatted questionnaires. They will miss edge cases — see RESULTS.md.
"""

import re
from typing import List


# ---------------------------------------------------------------------------
# Question detection
# ---------------------------------------------------------------------------

# Common identifier prefixes seen in security questionnaires
ID_PATTERNS = re.compile(
    r"""
    ^(
        \d+\.\d+        |   # 1.1  2.3
        \d+\.            |   # 1.   2.
        [A-Z]{2,6}-\d+  |   # DCTR-01  AC-2  AIS-01
        [A-Z]{2,6}\d+\.\d+ |  # CC1.1  CC6.2
        [A-Z]-\d+        |   # A-01
        Q\d+             |   # Q1  Q23
        \d+\)                # 1)  2)
    )
    """,
    re.VERBOSE,
)

# Phrases that almost always introduce a question
_QUESTION_PHRASES = re.compile(
    r"\b(does your|do you|is there|are there|describe|provide|explain|"
    r"how does|have you|has your|can you|please describe|please provide|"
    r"what is|what are|when was|who is|list all|identify)\b",
    re.IGNORECASE,
)


def is_question(text: str) -> bool:
    """
    Return True if this line looks like a security questionnaire question.

    We use three signals:
    1. Ends with a question mark
    2. Starts with a known questionnaire ID pattern
    3. Contains a typical question-starter phrase
    """
    stripped = text.strip()
    if not stripped or len(stripped) < 8:
        return False

    if stripped.endswith("?"):
        return True

    if ID_PATTERNS.match(stripped):
        return True

    if _QUESTION_PHRASES.search(stripped):
        return True

    return False


# ---------------------------------------------------------------------------
# Section header detection
# ---------------------------------------------------------------------------

# Known framework sections for better naming
_FRAMEWORK_SECTIONS = {
    # HECVAT
    "DCTR": "Datacenter",
    "PRIV": "Privacy",
    "INCD": "Incident Response",
    "VULN": "Vulnerability Management",
    "BCDR": "Business Continuity and Disaster Recovery",
    "THRD": "Third Party Risk",
    "ACCT": "Access Control",
    "ENPT": "Endpoint Security",
    "NTWK": "Network Security",
    "APPD": "Application Development",
    "CLDS": "Cloud Security",
    "CMPL": "Compliance",
    "RISK": "Risk Management",
    # CSA CAIQ
    "AIS": "Application and Interface Security",
    "CCC": "Change Control and Configuration Management",
    "EKM": "Encryption and Key Management",
    "GRM": "Governance and Risk Management",
    "HRS": "Human Resources",
    "IAM": "Identity and Access Management",
    "IVS": "Infrastructure and Virtualization Security",
    "IPY": "Interoperability and Portability",
    "MOS": "Mobile Security",
    "SEF": "Security Incident Management",
    "STA": "Supply Chain Management",
    "TVM": "Threat and Vulnerability Management",
}

_SECTION_HEADER_RE = re.compile(
    r"""
    ^(
        [A-Z]{2,6}(-\d+)?  |   # DCTR, AIS-01 (but short enough to be a heading)
        Section\s+\d+       |
        Part\s+[A-Z\d]+     |
        \d+\.\s+[A-Z]       |   # "1. Access Control"
        [IVXLC]+\.\s+        |   # Roman numerals I. II.
        [A-Z][a-z]+\s+(Control|Management|Security|Policy|Governance|Risk|Compliance)
    )
    """,
    re.VERBOSE,
)


def is_section_header(text: str) -> bool:
    """Detect if a line is a section heading rather than a question."""
    stripped = text.strip()
    if not stripped or len(stripped) > 120:
        return False

    # All-caps short lines are usually headers
    if stripped.isupper() and 3 <= len(stripped) <= 60:
        return True

    if _SECTION_HEADER_RE.match(stripped):
        return True

    return False


def is_heading_not_question(text: str) -> bool:
    """
    Return True when a line looks like a section heading even if it also
    matches question heuristics (e.g. '1. Network Security').
    Rules:
      - No question mark
      - Short (< 6 words after any leading ID/number)
      - Does not contain a question-starter verb
    """
    stripped = text.strip()
    if "?" in stripped:
        return False

    # Strip leading identifier to get the actual title words
    title_part = re.sub(r"^[\d\w]{1,8}[.)]\s*", "", stripped).strip()
    words = title_part.split()

    if len(words) == 0 or len(words) > 6:
        return False

    # If it contains a question-starter verb, it's a question
    if _QUESTION_PHRASES.search(title_part):
        return False

    return True


def resolve_section_name(raw_id: str) -> str:
    """Map a framework code to a human-readable section name if known."""
    key = raw_id.strip().upper().split("-")[0]
    return _FRAMEWORK_SECTIONS.get(key, raw_id)


# ---------------------------------------------------------------------------
# Response type detection
# ---------------------------------------------------------------------------

_YES_NO_VALUES = {"yes", "no", "n/a", "y", "n", "y/n", "na", "not applicable"}

_MULTI_CHOICE_RE = re.compile(r"^[A-Da-d][.)]\s+\w")  # A) text  or  a. text

# Also recognize the combined string as a yes/no indicator
_YES_NO_COMBINED_RE = re.compile(
    r"\b(yes\s*/\s*no|yes\s*/\s*no\s*/\s*n/?a)\b", re.IGNORECASE
)


def detect_response_type(text: str, nearby_cells: List[str]) -> dict:
    """
    Infer the expected response type from the question text and surrounding cells.

    Returns a dict with keys: response_type, options
    """
    options = _extract_options(nearby_cells)

    # Check for combined "Yes / No / N/A" strings in options
    combined_opts = " ".join(options)
    if _YES_NO_COMBINED_RE.search(combined_opts) or _YES_NO_COMBINED_RE.search(text):
        return {"response_type": "yes_no", "options": ["Yes", "No", "N/A"]}

    # Check for yes/no options (already split)
    if options:
        lower_opts = {o.lower().strip() for o in options}
        if lower_opts <= _YES_NO_VALUES or lower_opts & {"yes", "no"}:
            standard_opts = []
            for o in options:
                o_stripped = o.strip()
                if o_stripped.lower() in ("yes", "y"):
                    standard_opts.append("Yes")
                elif o_stripped.lower() in ("no", "n"):
                    standard_opts.append("No")
                elif o_stripped.lower() in ("n/a", "na", "not applicable"):
                    standard_opts.append("N/A")
                else:
                    standard_opts.append(o_stripped)
            # Make sure we have at least Yes and No
            if "Yes" not in standard_opts:
                standard_opts.insert(0, "Yes")
            if "No" not in standard_opts:
                if "Yes" in standard_opts:
                    idx = standard_opts.index("Yes") + 1
                    standard_opts.insert(idx, "No")
                else:
                    standard_opts.insert(0, "No")
            # De-duplicate Standard options
            seen = set()
            uniq_opts = []
            for o in standard_opts:
                if o.lower() not in seen:
                    seen.add(o.lower())
                    uniq_opts.append(o)
            return {"response_type": "yes_no", "options": uniq_opts}

        # Multiple choice: labelled options like A) B) C)
        if any(_MULTI_CHOICE_RE.match(o) for o in options):
            return {"response_type": "multiple_choice", "options": options}

        # Several distinct options likely means multiple choice too
        if 2 <= len(options) <= 10:
            return {"response_type": "multiple_choice", "options": options}

    # If no options found, try to infer from question text
    text_lower = text.lower().strip()
    
    # Strip leading identifiers like "1.1 ", "AC-01 " to inspect the actual words
    text_clean = re.sub(r"^([A-Z]{2,8}-\d+|[\d.]+)\b\s*", "", text_lower).strip()

    free_text_starters = (
        "describe", "explain", "provide", "detail", "list", "outline", 
        "summarize", "elaborate", "use this area", "state", "how does", 
        "what is", "what are", "who is", "when was", "identify", "please describe",
        "please provide", "please explain", "what details"
    )
    
    yes_no_starters = (
        "do you", "does your", "is there", "are there", "have you", "has your",
        "can you", "could you", "would you", "will", "is ", "are ", "does ",
        "do ", "has ", "have ", "can ", "should ", "was ", "were ", "must ",
        "did ", "has the"
    )

    if text_clean.startswith(free_text_starters):
        return {"response_type": "free_text", "options": []}

    if text_clean.startswith(yes_no_starters) or text_lower.endswith("?"):
        return {"response_type": "yes_no", "options": ["Yes", "No", "N/A"]}

    # Generic free text keyword check
    _free_text_re = re.compile(
        r"\b(describe|explain|provide|detail|list|outline|summarize|elaborate)\b",
        re.IGNORECASE,
    )
    if _free_text_re.search(text):
        return {"response_type": "free_text", "options": []}

    # Fallback — unknown is safer than a wrong guess
    return {"response_type": "unknown", "options": options}


def _extract_options(cells: List[str]) -> List[str]:
    """
    Pull non-empty, short strings from nearby cells as candidate options.
    Also handles slash-separated options like "Yes / No / N/A" and newline-separated options.
    """
    opts = []
    for cell in cells:
        # First split on newlines if present
        subcells = [c.strip() for c in cell.split("\n") if c.strip()]
        for subcell in subcells:
            stripped = subcell
            if not stripped or len(stripped) > 100:
                continue
            # Skip cells that look like question IDs (e.g. "AC-01", "EKM-02")
            if re.match(r"^[A-Z]{1,8}-?\d{1,3}$", stripped):
                continue
            # Skip cells that are clearly question sentences (> 60 chars with multiple words)
            words = stripped.split()
            if len(stripped) > 60 and len(words) > 6:
                continue
            # Expand slash-separated options: "Yes / No / N/A" -> ["Yes", "No", "N/A"]
            if "/" in stripped and len(stripped) < 60:
                temp = re.sub(r"\bN/A\b", "__NA__", stripped, flags=re.IGNORECASE)
                parts = [p.strip() for p in re.split(r"\s*/\s*", temp) if p.strip()]
                parts = [p.replace("__NA__", "N/A") for p in parts]
                if len(parts) > 1:
                    opts.extend(parts)
                    continue
            opts.append(stripped)
            
    # Deduplicate while preserving order
    seen = set()
    result = []
    for o in opts:
        if o not in seen:
            seen.add(o)
            result.append(o)
    return result

