# Security Questionnaire Normalizer

A command-line tool that reads messy security questionnaire files (`.xlsx` and `.docx`)
and converts them into a single, consistent machine-readable JSON format.

Security questionnaires like HECVAT, CSA CAIQ, and NIST SP 800-53 arrive in dozens
of different Excel and Word layouts. This tool normalizes them so downstream systems
(GRC platforms, risk dashboards, audit tools) can process them without custom parsing
per vendor.

---

## Setup

```bash
git clone <repo>
cd security-questionnaire-normalizer

python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

---

## Usage

### Parse a questionnaire

```bash
# Output to stdout
python normalize.py path/to/questionnaire.xlsx

# Pretty-print for readability
python normalize.py path/to/questionnaire.xlsx --pretty

# Write to a file
python normalize.py path/to/questionnaire.xlsx --out output.json

# Parse and validate schema in one step
python normalize.py path/to/questionnaire.xlsx --out output.json --validate
```

Works the same for `.docx` files:

```bash
python normalize.py path/to/questionnaire.docx --pretty
```

### Validate output against schema

```bash
python normalizer/schema_validator.py output.json
# Prints: VALID   or a list of schema errors
```

### Run accuracy evaluation

```bash
python evaluation/evaluate.py --output output.json --ground-truth evaluation/xlsx_ground_truth.tsv

# Add --verbose to see which questions matched
python evaluation/evaluate.py --output output.json --ground-truth evaluation/xlsx_ground_truth.tsv --verbose
```

### Generate sample test files (for demo)

```bash
python create_sample_xlsx.py
python create_sample_docx.py
```

---

## Output Format

```json
{
  "source_file": "questionnaire.xlsx",
  "questionnaire_name": "Vendor Security Assessment",
  "version": "2.1",
  "sections": [
    {
      "section_id": "AC",
      "section_title": "Access Control",
      "questions": [
        {
          "question_id": "AC-Q1",
          "text": "Does your organization have a formal access control policy?",
          "response_type": "yes_no",
          "options": ["Yes", "No", "N/A"],
          "guidance": "",
          "source_location": {
            "type": "sheet",
            "sheet": "Access Control",
            "row": 6
          }
        }
      ]
    }
  ]
}
```

`response_type` is one of: `yes_no`, `multiple_choice`, `free_text`, `unknown`.

---

## Architecture

```
normalize.py              CLI entry point. Dispatches to the right parser.
normalizer/
  xlsx_parser.py          openpyxl-based Excel parser
  docx_parser.py          python-docx parser (paragraphs + tables)
  detector.py             Question detection + response type classification
  schema_validator.py     JSON schema validation
  utils.py                Shared text-cleaning helpers
schema/
  output_schema.json      JSON Schema (Draft 7)
evaluation/
  evaluate.py             Precision/recall report using fuzzy matching
  ground_truth_sample.json  Reference dataset for evaluation
```

Data flows: `normalize.py` → parser → `detector.py` (per row) → structured output → optional `schema_validator.py`.

---

## Assumptions

- **Question detection is heuristic.** We use regex patterns for common question starters
  ("Does your", "Do you", "Describe", etc.) and identifier prefixes (AC-01, DCTR-01, 1.1).
  This covers ~85-90% of real questionnaires but will miss unusual formats.

- **Section mapping varies between frameworks.** Known codes (HECVAT's `DCTR`, `PRIV`;
  CSA CAIQ's `AIS`, `EKM`, `IAM`; NIST control families) are mapped to human-readable names.
  Unknown codes are kept as-is.

- **Unknown response type is preferred over a wrong guess.** If we can't confidently
  classify a question as yes/no, multiple choice, or free text, we return `"unknown"`.

- **Merged cell resolution follows Excel display logic.** The top-left cell value is
  propagated to all cells in the merge range.

- **Tables in docx files:** Column 0 is assumed to hold the question. Columns 1+ are
  treated as response options or guidance.

---

## Known Limitations

- **Complex nested tables** (tables inside cells) in `.docx` are not fully supported.
  python-docx doesn't expose nested tables through its high-level API.
- **Very unusual formatting** — e.g., questionnaires laid out as free-form prose with
  no tabular structure — will produce poor results.
- **Image-based or scanned questionnaires** are not supported. OCR would be required
  as a preprocessing step.
- **Formulas in Excel** are read as their computed value (`data_only=True`). If the
  file was never opened in Excel, formula cells may appear blank.
- **Multi-row questions** — where a question spans two Excel rows — may be split into
  two separate questions. Row-merging heuristics are not yet implemented.

---

## Future Improvements

- **ML-based question classification** — fine-tune a small classifier on labeled
  questionnaire rows to replace the regex heuristics.
- **LLM-assisted extraction** — use a language model to identify question boundaries,
  section names, and response types from free-form documents.
- **Better layout understanding** — detect column roles (question / response / guidance)
  dynamically rather than assuming column position.
- **Multi-row question merging** — detect when a question continues on the next row
  and join them before classifying.
- **CSV and PDF support** — many questionnaires also arrive in these formats.
