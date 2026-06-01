#!/usr/bin/env python3
"""
evaluate.py - Honest precision/recall measurement for the questionnaire normalizer.

USAGE
-----
  # 1. Run the normalizer to produce JSON output
  python normalize.py data/hecvat_lite.xlsx --pretty --out /tmp/hecvat_out.json

  # 2. Create a ground-truth CSV (see create_ground_truth_template below)
  python evaluate.py --create-template /tmp/hecvat_out.json --section DCTR > gt_dctr.csv

  # 3. Fill in the CSV by hand (mark is_question=1/0 for each row, correct the text)
  #    Then evaluate:
  python evaluate.py --output /tmp/hecvat_out.json --ground-truth gt_dctr.csv

WHY THIS APPROACH
-----------------
The bugs that cause 100% precision/recall are:
  (A) Comparing extracted text against the *same* file used to tune heuristics
      (ground truth should be labeled independently, not generated from your output).
  (B) Using strict string equality: "Do you have a documented policy?" vs
      "Do you have a documented data center physical security policy?" are different
      questions — but also note that merged-cell bleed can give you a truncated copy
      that *exactly matches* a short ground-truth label.
  (C) Index-leaking: if your loop generates both the prediction set and the GT set
      from the same parsed output, every TP is guaranteed.
  (D) Empty ground-truth: if gt_questions is an empty list, precision = 1/1 = 1.0
      and recall = 0/0 = 1.0 (or you get a ZeroDivisionError that's silently caught).

The design here fixes all four:
  - GT is loaded from a CSV you fill in by hand from the raw file (not the JSON).
  - Matching uses token-overlap (Jaccard similarity) with a tunable threshold,
    so small wording differences don't inflate TP counts.
  - A "skipped rows" log makes false negatives *visible* rather than silent.
  - Empty-GT is caught explicitly.
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple


# ---------------------------------------------------------------------------
# Matching strategy
# ---------------------------------------------------------------------------

def _tokenize(text: str) -> set:
    """Lowercase word tokens, stripping punctuation."""
    return set(re.findall(r"[a-z0-9]+", text.lower()))


def token_overlap(a: str, b: str) -> float:
    """
    Jaccard similarity over word tokens.
    Returns 0.0–1.0.  1.0 = identical token sets.
    """
    ta, tb = _tokenize(a), _tokenize(b)
    if not ta and not tb:
        return 1.0
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


def best_match(candidate: str, pool: List[str], threshold: float = 0.55) -> Tuple[int, float]:
    """
    Find the best-matching string in *pool* for *candidate*.
    Returns (index, score) or (-1, 0.0) if no match exceeds threshold.

    WHY 0.55?  At this threshold, "Do you have a data center security policy?"
    matches "Do you have a documented data center physical security policy?"
    (7 shared tokens / 12 union tokens = 0.58) but does NOT match
    "Describe your incident response process" (0.08).
    Tune lower if your GT labels are abbreviated, higher if you want stricter matching.
    """
    best_idx, best_score = -1, 0.0
    for i, p in enumerate(pool):
        score = token_overlap(candidate, p)
        if score > best_score:
            best_idx, best_score = i, score
    if best_score >= threshold:
        return best_idx, best_score
    return -1, best_score


# ---------------------------------------------------------------------------
# Ground-truth CSV format
# ---------------------------------------------------------------------------
# Columns (tab-separated for easy Excel/Sheets editing):
#   section_id  question_text  is_question  response_type  notes
#
# - section_id: the section you're sampling (e.g. "DCTR")
# - question_text: COPY FROM THE RAW FILE — not from your JSON output
# - is_question: 1 if it's a real question, 0 if it's a header / noise row
# - response_type: yes_no | multiple_choice | free_text | unknown
# - notes: optional, for your own reference

GT_COLUMNS = ["section_id", "question_text", "is_question", "response_type", "notes"]


def load_ground_truth(csv_path: str) -> List[Dict]:
    """Load hand-labeled ground truth CSV. Returns only rows where is_question == 1."""
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for i, row in enumerate(reader, 2):  # row 2 = first data row
            if not row.get("question_text", "").strip():
                continue
            try:
                is_q = int(row.get("is_question", "0").strip())
            except ValueError:
                print(f"[WARN] Row {i}: is_question is not 0/1, skipping.", file=sys.stderr)
                continue
            rows.append({
                "question_text": row["question_text"].strip(),
                "is_question": is_q,
                "response_type": row.get("response_type", "").strip(),
                "section_id": row.get("section_id", "").strip(),
            })

    # BUG GUARD: if GT is empty, bail out explicitly rather than silently returning 100%
    real_questions = [r for r in rows if r["is_question"] == 1]
    if not real_questions:
        print(
            "ERROR: Ground truth file has no rows with is_question=1. "
            "This would silently produce 100% precision and recall. "
            "Please fill in the CSV properly.",
            file=sys.stderr,
        )
        sys.exit(1)

    return rows  # return all rows so we can also test for false positives


def load_extracted(json_path: str, section_filter: str = "") -> List[Dict]:
    """Load extracted questions from the normalizer JSON output."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    
    questions = []
    for sec in data.get("sections", []):
        if section_filter and sec["section_id"].upper() != section_filter.upper():
            continue
        for q in sec.get("questions", []):
            questions.append({
                "question_id": q["question_id"],
                "text": q["text"],
                "response_type": q.get("response_type", "unknown"),
                "section_id": sec["section_id"],
            })
    return questions


# ---------------------------------------------------------------------------
# Precision / Recall computation
# ---------------------------------------------------------------------------

def compute_metrics(
    extracted: List[Dict],
    ground_truth: List[Dict],
    threshold: float = 0.55,
    verbose: bool = False,
) -> Dict:
    """
    Compare extracted questions against hand-labeled ground truth.

    Precision = TP / (TP + FP)
        Of everything we extracted, how much was a real question?
    Recall    = TP / (TP + FN)
        Of all real questions in the GT, how many did we find?

    Response-type accuracy is computed separately only over matched pairs.
    """
    gt_real = [r for r in ground_truth if r["is_question"] == 1]
    gt_texts = [r["question_text"] for r in gt_real]
    extracted_texts = [q["text"] for q in extracted]

    if not gt_real:
        raise ValueError("Ground truth contains no questions (is_question=1).")

    # --- Find TPs and FPs from the extracted set ---
    tp_list = []   # (extracted_idx, gt_idx, score)
    fp_list = []   # extracted_idx
    matched_gt_indices = set()

    for ext_idx, ext_text in enumerate(extracted_texts):
        gt_idx, score = best_match(ext_text, gt_texts, threshold)
        if gt_idx >= 0 and gt_idx not in matched_gt_indices:
            tp_list.append((ext_idx, gt_idx, score))
            matched_gt_indices.add(gt_idx)
        else:
            fp_list.append((ext_idx, score))

    # --- Find FNs: GT questions we never matched ---
    fn_indices = [i for i in range(len(gt_real)) if i not in matched_gt_indices]

    tp = len(tp_list)
    fp = len(fp_list)
    fn = len(fn_indices)

    precision = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    recall    = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    f1        = (2 * precision * recall) / (precision + recall) if (precision + recall) > 0 else 0.0

    # Response-type accuracy over matched pairs
    rt_matches = sum(
        1 for ext_idx, gt_idx, _ in tp_list
        if extracted[ext_idx]["response_type"] == gt_real[gt_idx]["response_type"]
        and gt_real[gt_idx]["response_type"]  # skip if GT has no label
    )
    rt_labeled = sum(1 for ext_idx, gt_idx, _ in tp_list if gt_real[gt_idx]["response_type"])
    rt_accuracy = rt_matches / rt_labeled if rt_labeled else None

    if verbose:
        print("\n=== TRUE POSITIVES ===")
        for ext_idx, gt_idx, score in tp_list:
            ext_rt = extracted[ext_idx]["response_type"]
            gt_rt  = gt_real[gt_idx]["response_type"]
            rt_ok  = "✓" if ext_rt == gt_rt else f"✗ (got {ext_rt}, expected {gt_rt})"
            print(f"  [{score:.2f}] {extracted[ext_idx]['text'][:80]}")
            print(f"        RT: {rt_ok}")

        print("\n=== FALSE POSITIVES (extracted but not in GT) ===")
        for ext_idx, score in fp_list:
            print(f"  [best={score:.2f}] {extracted[ext_idx]['text'][:80]}")

        print("\n=== FALSE NEGATIVES (in GT but not extracted) ===")
        for gt_idx in fn_indices:
            print(f"  {gt_real[gt_idx]['question_text'][:80]}")

    return {
        "total_extracted": len(extracted),
        "total_gt_questions": len(gt_real),
        "tp": tp,
        "fp": fp,
        "fn": fn,
        "precision": round(precision, 4),
        "recall": round(recall, 4),
        "f1": round(f1, 4),
        "response_type_accuracy": round(rt_accuracy, 4) if rt_accuracy is not None else "N/A (no RT labels)",
    }


# ---------------------------------------------------------------------------
# Ground-truth template generator
# ---------------------------------------------------------------------------

def create_ground_truth_template(json_path: str, section: str) -> None:
    """
    Print a TSV template pre-filled with extracted questions from *section*.

    IMPORTANT: You should ALSO manually scan the raw source file for questions
    that your tool missed and add them to the CSV with is_question=1.
    This is the only way to capture false negatives.
    """
    extracted = load_extracted(json_path, section_filter=section)
    print("\t".join(GT_COLUMNS))
    for q in extracted:
        row = [
            q["section_id"],
            q["text"].replace("\t", " ").replace("\n", " "),
            "1",          # you'll change to 0 for non-questions you spot
            q["response_type"],
            "",
        ]
        print("\t".join(row))
    print(
        f"\n# NOTE: The above was generated from your *extracted* output for section '{section}'.\n"
        f"# Before using this as ground truth you MUST:\n"
        f"#   1. Open the raw source file and find questions your tool MISSED.\n"
        f"#      Add them as new rows with is_question=1.\n"
        f"#   2. Mark any non-questions your tool extracted as is_question=0.\n"
        f"#   3. Delete this comment block from the CSV.",
        file=sys.stderr,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Evaluate normalizer precision/recall")
    parser.add_argument("--output", help="Path to normalizer JSON output file")
    parser.add_argument("--ground-truth", help="Path to hand-labeled TSV ground truth")
    parser.add_argument("--section", default="", help="Filter to a specific section_id")
    parser.add_argument("--threshold", type=float, default=0.55,
                        help="Jaccard similarity threshold for matching (default 0.55)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Print TP/FP/FN details")
    parser.add_argument("--create-template", metavar="JSON_FILE",
                        help="Generate a ground-truth TSV template from this output file")
    args = parser.parse_args()

    if args.create_template:
        if not args.section:
            print("ERROR: --section required with --create-template", file=sys.stderr)
            sys.exit(1)
        create_ground_truth_template(args.create_template, args.section)
        return

    if not args.output or not args.ground_truth:
        parser.print_help()
        sys.exit(1)

    extracted = load_extracted(args.output, section_filter=args.section)
    ground_truth = load_ground_truth(args.ground_truth)

    if not extracted:
        print(f"ERROR: No questions extracted (section filter: '{args.section}'). "
              f"Check your output file or section name.", file=sys.stderr)
        sys.exit(1)

    metrics = compute_metrics(extracted, ground_truth, threshold=args.threshold, verbose=args.verbose)

    print(json.dumps(metrics, indent=2))


if __name__ == "__main__":
    main()
