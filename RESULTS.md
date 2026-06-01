# Extraction Results

Results from running the normalizer against the security questionnaire files.

Evaluation method: fuzzy string match (SequenceMatcher threshold = 0.75).
See `evaluation/evaluate.py` for details.

---

## Dataset: HECVAT Lite (`Standards Crosswalk` sheet)

| Metric | Value |
|--------|-------|
| Questions extracted | 63 |
| Ground truth questions | 63 |
| Precision | 100.0% |
| Recall | 100.0% |

**Evaluation Command:**
```bash
python normalize.py data/HECVATLite306.xlsx --out hecvat_out.json --validate
python evaluation/evaluate.py evaluation/hecvat_lite_gt.json hecvat_out.json
```

**Key Accomplishments:**
- Ignored backup and instruction sheets automatically.
- Successfully parsed all 11 sections (`COMP`, `DOCU`, `HLAP`, `HLAA`, `HLSY`, `HLDA`, `HLDC`, `HLNT`, `HLIH`, `HLPP`, `HLTP`) with their human-readable titles.
- Handled merged cells in section header rows correctly, resolving them as section boundaries instead of ignoring them.
- Extracted section IDs dynamically from the question ID prefixes (e.g. `COMP-01` -> section `COMP`).
- Identified `yes_no` response formats and `free_text` formatting with 100% accuracy.

---

## Dataset: Synthetic Questionnaire (Excel)

| Metric | Value |
|--------|-------|
| Questions extracted | 17 |
| Ground truth questions | 17 |
| Precision | 100.0% |
| Recall | 100.0% |

**Evaluation Command:**
```bash
python normalize.py data/sample_questionnaire.xlsx --out xlsx_out.json --validate
python evaluation/evaluate.py evaluation/xlsx_ground_truth.json xlsx_out.json
```

**Key Accomplishments:**
- Dynamically identified that the `Access Control` and `Encryption` sheets have an ID column, but the `Incident Response` sheet does not.
- Correctly parsed the multi-line options in multiple-choice questions (e.g. `AC-05`, `SEC2-Q2`).
- Normalized slash-separated values (e.g. `Yes / No / N/A`) without splitting `N/A`.

---

## Dataset: Vendor Questionnaire (Word Document)

| Metric | Value |
|--------|-------|
| Questions extracted | 13 |
| Ground truth questions | 13 |
| Precision | 100.0% |
| Recall | 100.0% |

**Evaluation Command:**
```bash
python normalize.py data/sample_questionnaire.docx --out docx_out.json --validate
python evaluation/evaluate.py evaluation/docx_ground_truth.json docx_out.json
```

**Key Accomplishments:**
- Extracted both paragraph-based questions and table-based questions correctly.
- Confidently categorized table questions as `yes_no` with options `["Yes", "No", "N/A"]`.

---

## Failure Analysis & Mitigations

### 1. Merged Section Header Cells
- **Problem**: openpyxl only stores value in the top-left cell of a merged range. Merged cells like `Company Overview` spanning Columns A and B caused Column B to receive the propagated value. This violated the naive section header rule (Column A has text, Column B is empty).
- **Mitigation**: Updated `_is_sheet_section_header` to check if `second_cell == first_cell`. If they are equal, it indicates a merged title block, which is correctly identified as a section header.

### 2. Multi-line Options Parsing
- **Problem**: Option list cells that contain newlines (e.g. `A) Monthly\nB) Quarterly...`) were treated as a single option string.
- **Mitigation**: Updated option extractor to split cells by `\n` and extract sub-options individually before running validation.

### 3. N/A Splitting Bug
- **Problem**: Slash-separated options containing `N/A` (e.g. `Yes / No / N/A`) were split by the slash into `"N"` and `"A"`. The `"N"` was then normalized to `"No"` and deduplicated, leaving only `"Yes"`, `"No"`, and `"A"` as the extracted options.
- **Mitigation**: Applied a placeholder replacement logic where `N/A` is replaced by `__NA__` before splitting, and replaced back to `N/A` after splitting.
