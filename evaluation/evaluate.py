"""
evaluate.py - Measure extraction accuracy against a ground truth JSON file.

Usage:
    python evaluation/evaluate.py ground_truth.json extracted_output.json

We use fuzzy string matching (difflib.SequenceMatcher) because question wording
sometimes differs slightly between the source doc and what we extract.
A match threshold of 0.75 was chosen empirically — it catches minor whitespace
and punctuation differences without accepting false positives.

Metrics:
  Precision = matched / extracted  (how much of our output is correct)
  Recall    = matched / ground_truth (how much of the real data we found)
"""

import json
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import List, Tuple


MATCH_THRESHOLD = 0.75  # tune this if needed


def load_questions(filepath: str) -> List[str]:
    """Extract all question texts from a normalized JSON file or flat list."""
    with open(filepath) as f:
        data = json.load(f)

    questions = []
    if "questions" in data:
        for q in data["questions"]:
            text = q.get("text", "").strip()
            if text:
                questions.append(text)
        return questions

    for section in data.get("sections", []):
        for q in section.get("questions", []):
            text = q.get("text", "").strip()
            if text:
                questions.append(text)
    return questions


def fuzzy_match(q1: str, q2: str) -> float:
    """Return a similarity score between 0 and 1."""
    return SequenceMatcher(None, q1.lower(), q2.lower()).ratio()


def find_matches(ground_truth: List[str], extracted: List[str]) -> Tuple[int, List[Tuple]]:
    """
    Greedily match each ground truth question to the best extracted question
    above the threshold. Each extracted question can only be matched once.

    Returns (match_count, list_of_matched_pairs)
    """
    used = set()
    matched_pairs = []

    for gt_q in ground_truth:
        best_score = 0.0
        best_idx = -1
        for i, ex_q in enumerate(extracted):
            if i in used:
                continue
            score = fuzzy_match(gt_q, ex_q)
            if score > best_score:
                best_score = score
                best_idx = i

        if best_score >= MATCH_THRESHOLD and best_idx >= 0:
            used.add(best_idx)
            matched_pairs.append((gt_q, extracted[best_idx], best_score))

    return len(matched_pairs), matched_pairs


def evaluate(ground_truth_path: str, extracted_path: str):
    ground_truth = load_questions(ground_truth_path)
    extracted = load_questions(extracted_path)

    if not ground_truth:
        print("ERROR: Ground truth file has no questions.", file=sys.stderr)
        sys.exit(1)

    if not extracted:
        print("WARNING: Extracted file has no questions.", file=sys.stderr)
        precision = 0.0
        recall = 0.0
        match_count = 0
    else:
        match_count, matched_pairs = find_matches(ground_truth, extracted)
        precision = match_count / len(extracted) if extracted else 0.0
        recall = match_count / len(ground_truth) if ground_truth else 0.0

    # Print report
    print("=" * 50)
    print("  EXTRACTION ACCURACY REPORT")
    print("=" * 50)
    print(f"  Ground truth questions : {len(ground_truth)}")
    print(f"  Extracted questions    : {len(extracted)}")
    print(f"  Matched (threshold={MATCH_THRESHOLD}) : {match_count}")
    print()
    print(f"  Precision : {precision:.1%}")
    print(f"  Recall    : {recall:.1%}")
    print("=" * 50)

    if match_count > 0 and "--verbose" in sys.argv:
        print("\nMatched pairs (ground truth -> extracted):\n")
        for gt, ex, score in matched_pairs:
            print(f"  [{score:.2f}] GT : {gt[:80]}")
            print(f"         EX : {ex[:80]}")
            print()

    return precision, recall


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python evaluation/evaluate.py <ground_truth.json> <extracted.json> [--verbose]")
        sys.exit(1)

    evaluate(sys.argv[1], sys.argv[2])
