#!/usr/bin/env python
"""
format_doc.py — single-command CLI wrapping the full pipeline.

=== REVERT NOTICE (this release) ===
The previous release attempted three fixes in one pass:
  (a) nested-table (table-inside-a-table-cell) font/size normalization,
  (b) usage-based FONT NAME correction for body-like roles, and
  (c) demoting numbered/lettered list items whose style resolves to a
      heading role, if they carry Word list-numbering.

Reported result: destructive regressions to previously-correct text
color and table shading/representation. Root-caused as follows:

  - (c) is the prime suspect: many organizational documents apply
    section-heading numbering (1, 1.1, 1.1.1) via Word's AUTOMATIC
    multi-level list numbering rather than literal typed digits. When
    that's the case, `paragraph.text` does NOT contain the visible
    number at all (Word renders it from the numbering definition,
    invisible to python-docx) - so the safety check meant to protect
    real headings ("does the literal text look like a decimal heading
    pattern?") can never recognize them, and they get wrongly demoted,
    stripping their correct heading color. This exactly matches the
    reported symptom.
  - (a)/(b) are lower-risk in isolation but were shipped together with
    (c) and have not been independently re-verified against the user's
    real document structure since the regression was reported.

THIS RELEASE REVERTS (a), (b), and (c) COMPLETELY, restoring the exact
last known-good behavior (top-level tables only; heading-only usage
voting for font+color; no list-numbering-based demotion at all).

A safer, narrower, OPT-IN redesign of the three original fixes (Aptos in
nested tables, Cambria in body text, list items inheriting heading
color) is planned as a SEPARATE, follow-up change - each one isolated,
tested against a reproduction of the EXACT failure mode above (auto-
numbered real headings must keep their color), and shipped only after
independent verification. See the project's change log / conversation
history for the specific plan discussed with the user.
"""
import argparse
import json
import os
import sys
import time

import policy_extractor
import structure_classifier
import formatting_engine
import integrity_validator
import toc_card_builder


def run(reference_path, input_path, output_path, workdir=".", body_baseline_pt=11.0,
        confidence_threshold=0.7, fail_on_content_drift=True):
    os.makedirs(workdir, exist_ok=True)
    policy_path = os.path.join(workdir, "formatting_policy_v1.json")
    classification_path = os.path.join(workdir, "classification_report.json")
    format_result_path = os.path.join(workdir, "format_result.json")
    pre_toc_path = os.path.join(workdir, "pre_toc_formatted.docx")
    toc_result_path = os.path.join(workdir, "toc_result.json")
    integrity_json_path = os.path.join(workdir, "integrity_result.json")
    report_path = os.path.join(workdir, "formatting_change_report.md")

    t0 = time.time()

    policy = policy_extractor.build_policy(reference_path)
    with open(policy_path, "w", encoding="utf-8") as f:
        json.dump(policy, f, indent=2, ensure_ascii=False)
    print(f"[1/5] Policy extracted from {reference_path} -> {policy_path}")
    print(f"      Theme fonts: {policy['theme_fonts']}")
    print(f"      Table header (uniform, top-level tables only): {policy['table_contexts']['data_table']}")
    print(f"      Table typography (header, uniform): {policy['table_typography']['header']}")
    print(f"      Table typography (body): {policy['table_typography']['body']}")
    print(f"      Footer has PAGE field: {policy['footer_design'].get('has_page_field')}")
    for role in ("H1", "H2", "H3"):
        t = policy["typography"].get(role, {})
        print(f"      {role} color: #{t.get('font', {}).get('color_hex')} "
              f"(usage samples: {t.get('usage_sample_size', 0)})")
    body_font = policy["typography"].get("Body", {}).get("font", {})
    print(f"      Body color (style-def only, NEVER usage-voted): #{body_font.get('color_hex')}")

    structure_classifier.CONFIDENCE_THRESHOLD = confidence_threshold
    classification = structure_classifier.classify_document(
        input_path, body_baseline_pt=body_baseline_pt,
        special_headings=policy.get("special_headings", {}))
    with open(classification_path, "w", encoding="utf-8") as f:
        json.dump(classification, f, indent=2, ensure_ascii=False)
    summary = classification["summary"]
    print(f"[2/5] Classified {summary['total_paragraphs']} paragraphs "
          f"({summary['low_confidence_count']} low-confidence, "
          f"{summary['low_confidence_pct']}%) -> {classification_path}")

    format_result = formatting_engine.format_document(
        input_path, policy_path, classification_path, pre_toc_path)
    with open(format_result_path, "w", encoding="utf-8") as f:
        json.dump(format_result, f, indent=2, ensure_ascii=False)
    print(f"[3/5] Base formatting applied -> {pre_toc_path}")
    print(f"      Table typography result: {format_result['table_typography_result']}")
    print(f"      Table contexts touched: {format_result['tables_touched_by_context']}")
    print(f"      Footer result: {format_result['footer_result']}")

    base_integrity_result = integrity_validator.validate(input_path, pre_toc_path)
    if not base_integrity_result["content_integrity_passed"]:
        with open(integrity_json_path, "w", encoding="utf-8") as f:
            json.dump(base_integrity_result, f, indent=2, ensure_ascii=False)
        print("HALTING: content drift detected in base formatting phases (before TOC "
              "restructuring). Formatted output NOT considered safe to ship.", file=sys.stderr)
        sys.exit(2)
    print(f"[4/5] Base content-integrity check: PASS (0 replacements, 0 relocations)")

    toc_result = toc_card_builder.try_build_toc_card_table(pre_toc_path, policy, output_path)
    with open(toc_result_path, "w", encoding="utf-8") as f:
        json.dump(toc_result, f, indent=2, ensure_ascii=False)
    print(f"[5/5] TOC card table restructuring: applied={toc_result['applied']} "
          f"({toc_result['reason']})")

    final_integrity_result = integrity_validator.validate(input_path, output_path)
    with open(integrity_json_path, "w", encoding="utf-8") as f:
        json.dump(final_integrity_result, f, indent=2, ensure_ascii=False)
    report_md = integrity_validator.build_change_report_markdown(
        final_integrity_result, format_result, summary, policy, toc_result)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_md)

    elapsed = time.time() - t0
    passed = final_integrity_result["content_integrity_passed"]
    print(f"[FINAL] Content-integrity check: {'PASS' if passed else 'FAIL'} "
          f"(replacements={final_integrity_result['replace_mismatch_count']}, "
          f"relocations={final_integrity_result['relocatable_mismatch_count']}, "
          f"word-reconciliation={'PASS' if final_integrity_result['word_reconciliation_passed'] else 'FAIL'}) "
          f"-> {report_path}")
    print(f"Done in {elapsed:.2f}s.")

    if fail_on_content_drift and not passed:
        print("HALTING: content drift detected. Formatted output NOT considered safe to ship.",
              file=sys.stderr)
        sys.exit(2)

    return {
        "policy_path": policy_path, "classification_path": classification_path,
        "output_path": output_path, "report_path": report_path,
        "integrity_passed": passed, "elapsed_seconds": elapsed,
    }


def main():
    ap = argparse.ArgumentParser(
        description="Fast prototype: format an organizational document to match a 4-page reference template.")
    ap.add_argument("--reference", required=True)
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--workdir", default="run_artifacts")
    ap.add_argument("--body-baseline-pt", type=float, default=11.0)
    ap.add_argument("--confidence-threshold", type=float, default=0.7)
    ap.add_argument("--allow-content-drift", action="store_true")
    args = ap.parse_args()

    run(args.reference, args.input, args.output, workdir=args.workdir,
        body_baseline_pt=args.body_baseline_pt,
        confidence_threshold=args.confidence_threshold,
        fail_on_content_drift=not args.allow_content_drift)


if __name__ == "__main__":
    main()
