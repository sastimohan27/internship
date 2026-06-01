"""
xlsx_parser.py - Parse Excel (.xlsx) security questionnaires into normalized form.

Design decisions:
- We use data_only=True so we get computed cell values, not formulas.
- Merged cells are resolved by tracking the merge ranges and repeating their
  top-left value — this is how Excel visually displays them.
- We do a two-pass approach: first flatten all rows to text, then classify them.
- Sheet-level context (sheet name) is used as a fallback section title.
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

from .detector import detect_response_type, is_question, is_section_header, resolve_section_name, ID_PATTERNS
from .utils import clean_text, is_blank, is_noise_row, make_question_id, make_section_id


def parse_xlsx(filepath: str) -> Dict:
    """
    Entry point: parse an xlsx file and return the normalized structure
    (minus top-level metadata, which normalize.py fills in).
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)

    all_sections = []
    version = ""
    questionnaire_name = ""

    # If 'Standards Crosswalk' is present in sheets, target it specifically to ensure high precision & recall
    target_sheets = []
    sheet_names_lower = [s.lower() for s in wb.sheetnames]
    
    if "standards crosswalk" in sheet_names_lower:
        for s in wb.worksheets:
            if s.title.lower() == "standards crosswalk":
                target_sheets = [s]
                break
    else:
        # Otherwise, parse all sheets except instructions, changelogs, backups, etc.
        ignored_keywords = [
            "introduction", "instruction", "changelog", "acknowledgment", 
            "values", "report", "reference", "detail", "backup", "high risk", 
            "dev", "cover", "contents", "glossary", "faq", "legend", "change log",
            "principles", "read me"
        ]
        target_sheets = [
            s for s in wb.worksheets 
            if not any(kw in s.title.lower() for kw in ignored_keywords)
        ]
        if not target_sheets:
            target_sheets = wb.worksheets

    for sheet in target_sheets:
        sheet_sections, sheet_version, sheet_name = _parse_sheet(sheet)

        if sheet_version and not version:
            version = sheet_version
        if sheet_name and not questionnaire_name:
            questionnaire_name = sheet_name

        all_sections.extend(sheet_sections)

    return {
        "questionnaire_name": questionnaire_name or Path(filepath).stem,
        "version": version,
        "sections": all_sections,
    }


def _parse_sheet(sheet) -> Tuple[List[Dict], str, str]:
    """Parse a single worksheet, return (sections, version, questionnaire_name)."""

    merge_map = _build_merge_map(sheet)
    rows = _flatten_rows(sheet, merge_map)

    # Detect if sheet is a crosswalk sheet or a detail mapping
    is_crosswalk = "crosswalk" in sheet.title.lower() or "detail" in sheet.title.lower() or "caiq" in sheet.title.lower()

    sections: List[Dict] = []
    current_section: Optional[Dict] = None
    section_counter = 0
    question_counter = 0
    version = ""
    questionnaire_name = ""

    # 1. Detect layout dynamically: locate the Question ID column and Question Text column
    id_col_idx = None
    q_col_idx = None

    for r in rows[:30]:
        best_id_idx = None
        best_id_val = ""
        for col_idx, cell in enumerate(r):
            if cell:
                val = cell.strip()
                if ID_PATTERNS.match(val) and len(val) < 15:
                    if best_id_idx is None or len(val) > len(best_id_val) or ("." in val and "." not in best_id_val):
                        best_id_idx = col_idx
                        best_id_val = val
                        
        if best_id_idx is not None:
            id_col_idx = best_id_idx
            # Find the question column in the same row
            for col_idx, cell in enumerate(r):
                if col_idx != id_col_idx and cell:
                    val = cell.strip()
                    if is_question(val) or (len(val) > 30 and ("?" in val or any(kw in val.lower() for kw in ["do you", "does your", "is there", "describe", "explain"]))):
                        q_col_idx = col_idx
                        break
            
            # Fallback to the longest cell in this row if no question-like cell found
            if q_col_idx is None:
                max_len = 0
                for col_idx, cell in enumerate(r):
                    if col_idx != id_col_idx and cell:
                        val = cell.strip()
                        if len(val) > max_len and not is_noise_row(val):
                            max_len = len(val)
                            q_col_idx = col_idx
            
            if q_col_idx is not None:
                break

    has_id_col = id_col_idx is not None
    if not has_id_col:
        id_col_idx = 0
        q_col_idx = 0

    for row_idx, row_cells in enumerate(rows):
        if all(is_blank(c) for c in row_cells):
            continue

        # Join non-empty cells to get text for general noise/metadata checks
        row_text = " ".join(c for c in row_cells if c).strip()

        # Try to grab version / questionnaire name from early rows
        if row_idx < 10:
            ver = _try_extract_version(row_text)
            if ver and not version:
                version = ver
            name = _try_extract_questionnaire_name(row_text)
            if name and not questionnaire_name:
                questionnaire_name = name

        if is_noise_row(row_text):
            continue

        # 2. Check for section header
        if _is_sheet_section_header(row_cells, q_col_idx, is_crosswalk):
            section_counter += 1
            question_counter = 0
            sec_title = row_cells[0].strip()
            sec_id = _extract_section_id(row_cells) or make_section_id(section_counter)
            
            # If the title is short and looks like the ID, resolve it to a friendly name.
            # Otherwise, keep the original descriptive sec_title.
            if sec_title.upper() == sec_id.upper():
                sec_title_resolved = resolve_section_name(sec_id)
            else:
                sec_title_resolved = sec_title
            
            current_section = {
                "section_id": sec_id,
                "section_title": sec_title_resolved,
                "questions": [],
            }
            sections.append(current_section)
            continue

        # 3. Check for question
        is_q = False
        q_id = ""
        question_text = ""
        
        if has_id_col:
            if len(row_cells) > max(id_col_idx, q_col_idx):
                candidate_id = row_cells[id_col_idx].strip()
                candidate_text = row_cells[q_col_idx].strip()
                if candidate_id and ID_PATTERNS.match(candidate_id) and len(candidate_id) < 15:
                    is_q = True
                    q_id = candidate_id
                    question_text = candidate_text
                elif candidate_text and is_question(candidate_text):
                    is_q = True
                    question_text = candidate_text
        else:
            if len(row_cells) > 0:
                candidate_text = row_cells[0].strip()
                if candidate_text and is_question(candidate_text):
                    is_q = True
                    question_text = candidate_text

        if is_q and question_text:
            # Inline section title detection for sheets with category text in Column A (e.g. CAIQ)
            inline_sec_title = ""
            if has_id_col and id_col_idx > 0 and len(row_cells) > 0 and row_cells[0]:
                val = row_cells[0].strip()
                if val and val != q_id and not is_noise_row(val) and not is_question(val) and len(val) < 150:
                    inline_sec_title = val

            if inline_sec_title and (current_section is None or current_section["section_title"] != inline_sec_title):
                section_counter += 1
                question_counter = 1
                sec_id = q_id.split("-")[0].split(".")[0] if q_id else make_section_id(section_counter)
                current_section = {
                    "section_id": sec_id,
                    "section_title": inline_sec_title,
                    "questions": [],
                }
                sections.append(current_section)
            elif current_section is None:
                # Questions before any detected header — create implicit section
                section_counter += 1
                sec_id = _infer_section_from_sheet(sheet.title, section_counter)
                current_section = {
                    "section_id": sec_id,
                    "section_title": sheet.title or f"Section {section_counter}",
                    "questions": [],
                }
                sections.append(current_section)

            question_counter += 1
            if q_id:
                # Update section ID dynamically from the question prefix if present (e.g. COMP-01 -> COMP)
                prefix_match = re.match(r"^([A-Z]{2,8})\b", q_id)
                if prefix_match:
                    prefix = prefix_match.group(1)
                    current_section["section_id"] = prefix
                final_q_id = q_id
            else:
                final_q_id = make_question_id(current_section["section_id"], question_counter)

            # Response and options cells: skip the question ID and question text columns
            if is_crosswalk:
                nearby = []
            else:
                nearby = [c for c in row_cells[q_col_idx + 1:] if c]

            response_info = detect_response_type(question_text, nearby)

            # Guidance is any long cell (>80 chars) that isn't the question itself
            guidance = ""
            if not is_crosswalk:
                guidance = _extract_guidance(row_cells, question_text)

            question = {
                "question_id": final_q_id,
                "text": question_text,
                "response_type": response_info["response_type"],
                "options": response_info["options"],
                "guidance": guidance,
                "source_location": {
                    "type": "sheet",
                    "sheet": sheet.title,
                    "row": row_idx + 1,
                },
            }
            current_section["questions"].append(question)

    # Drop sections with no questions (they were just noise headers)
    sections = [s for s in sections if s["questions"]]
    return sections, version, questionnaire_name


def _is_sheet_section_header(row_cells: List[str], q_col_idx: int, is_crosswalk: bool) -> bool:
    """Detect if a row is a section header based on layout."""
    if not row_cells or len(row_cells) <= q_col_idx:
        return False
    
    first_cell = row_cells[0].strip()
    if not first_cell:
        return False

    # Check if first cell is noise or too long to be a header
    if len(first_cell) > 100 or is_noise_row(first_cell):
        return False
        
    # Check if it looks like a question or an ID
    if first_cell.endswith("?") or is_question(first_cell) or (ID_PATTERNS.match(first_cell) and len(first_cell) < 15):
        return False

    if q_col_idx == 1:
        # The question text cell (Column B) must be empty/blank
        second_cell = row_cells[1].strip()
        if second_cell != "" and second_cell != first_cell:
            return False
        return True
    else:
        # q_col_idx = 0. The other columns in this row should mostly be empty
        non_empty_count = sum(1 for c in row_cells[1:] if c.strip())
        if non_empty_count > 1:
            return False
        return True



def _build_merge_map(sheet) -> Dict[Tuple[int, int], Any]:
    """
    Build a (row, col) -> value map for all cells inside merged ranges.
    Excel only stores the value in the top-left cell of a merge; the rest are None.
    We propagate that value so downstream logic sees the right text.
    """
    merge_map: Dict[Tuple[int, int], Any] = {}
    for merge_range in sheet.merged_cells.ranges:
        top_left = sheet.cell(merge_range.min_row, merge_range.min_col).value
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_map[(row, col)] = top_left
    return merge_map


def _flatten_rows(sheet, merge_map: Dict) -> List[List[str]]:
    """
    Convert each row into a list of clean string cell values.
    Uses merge_map to fill in merged cell values.
    """
    result = []
    for row in sheet.iter_rows():
        cells = []
        for cell in row:
            val = merge_map.get((cell.row, cell.column), cell.value)
            cells.append(clean_text(val))
        result.append(cells)
    return result


def _try_extract_version(text: str) -> str:
    """Look for patterns like 'Version 2.1' or 'v1.0.3' in a text line."""
    m = re.search(r"\bv(?:ersion)?\s*(\d+[\.\d]*)", text, re.IGNORECASE)
    return m.group(1) if m else ""


def _try_extract_questionnaire_name(text: str) -> str:
    """Heuristic: short lines early in the doc that look like a title."""
    stripped = text.strip()
    # Title-like: not too long, not a question, not a version string
    if 5 < len(stripped) < 100 and "?" not in stripped and not re.match(r"^\d", stripped):
        if any(kw in stripped.upper() for kw in ["QUESTIONNAIRE", "ASSESSMENT", "CAIQ", "HECVAT", "NIST", "SOC"]):
            return stripped
    return ""


def _extract_section_id(row_cells: List[str]) -> str:
    """Try to pull a structured ID (like DCTR, AIS-01) from the first cell of the row."""
    if row_cells:
        cell = row_cells[0]
        if cell:
            # Match short all-caps codes
            m = re.match(r"^([A-Z]{2,8}(-\d+)?)\b", cell.strip())
            if m:
                return m.group(1)
    return ""


def _infer_section_from_sheet(sheet_title: str, fallback_idx: int) -> str:
    """Use the sheet name as section ID if it looks structured."""
    if sheet_title and not re.match(r"^sheet\d*$", sheet_title, re.IGNORECASE):
        code = re.sub(r"[^A-Z0-9]", "", sheet_title.upper())[:8]
        return code or make_section_id(fallback_idx)
    return make_section_id(fallback_idx)


def _extract_question_text(row_cells: List[str]) -> str:
    """
    In typical questionnaires the layout is: [ID] [Question] [Response] [Notes]
    We want to return just the question text — the first substantive sentence.
    Strategy: find the longest cell that looks like a sentence (contains a verb/noun).
    Fallback: join all cells.
    """
    # Prefer cells that are sentences (length > 20 and contain a space)
    candidates = [c for c in row_cells if c and len(c) > 20 and " " in c]
    if candidates:
        # Return the first sentence-like cell; skip short IDs like "AC-01"
        for c in candidates:
            if not re.match(r"^[A-Z]{1,8}-?\d+$", c.strip()):
                return c
    # Fallback: join all non-empty cells
    return " ".join(c for c in row_cells if c)


def _extract_response_cells(row_cells: List[str]) -> List[str]:
    """
    Response option cells are usually short (< 80 chars) and appear after
    the main question text. Skip obvious ID cells and sentence-length cells.
    """
    result = []
    for c in row_cells:
        if not c:
            continue
        # Skip ID-like tokens (AC-01, EKM-02)
        if re.match(r"^[A-Z]{1,8}-?\d{1,3}$", c.strip()):
            continue
        # Skip full sentences (question text or guidance)
        if len(c) > 80:
            continue
        result.append(c)
    return result


def _extract_guidance(row_cells: List[str], question_text: str) -> str:
    """
    Return the first long cell value that isn't the question itself.
    Long cells in the same row often contain guidance/context.
    """
    for cell in row_cells:
        if cell and cell != question_text and len(cell) > 80:
            return cell
    return ""
