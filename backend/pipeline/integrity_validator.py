"""Phase 0 — Content-Integrity Validator (v1.11 baseline, reverted)."""
import argparse
import difflib
import json
import re
from collections import Counter
from docx import Document


def _normalize_ws(s):
    return re.sub(r"\s+", " ", s or "").strip()


_EXPECTED_RELABELING_STOPWORDS = {"section"}


def _tokenize(text):
    raw_tokens = re.findall(r"[a-z0-9]+", (text or "").lower())
    tokens = []
    for tok in raw_tokens:
        if tok in _EXPECTED_RELABELING_STOPWORDS:
            continue
        if tok.isdigit():
            tok = str(int(tok))
        tokens.append(tok)
    return tokens


def extract_paragraph_text(doc):
    return [p.text for p in doc.paragraphs]


def extract_table_text(doc):
    cells = []
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cells.append(((t_idx, r_idx, c_idx), cell.text))
    return cells


def _whole_document_word_bag(doc):
    bag = Counter()
    for p in doc.paragraphs:
        bag.update(_tokenize(p.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                bag.update(_tokenize(cell.text))
    return bag


def diff_text_lists(original, formatted, label):
    mismatches = []
    orig_norm = [_normalize_ws(t) for t in original]
    fmt_norm = [_normalize_ws(t) for t in formatted]
    sm = difflib.SequenceMatcher(a=orig_norm, b=fmt_norm, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        mismatches.append({
            "location": label, "type": tag,
            "original_index_range": [i1, i2], "formatted_index_range": [j1, j2],
            "original_text": " | ".join(original[i1:i2])[:200],
            "formatted_text": " | ".join(formatted[j1:j2])[:200],
        })
    return mismatches


def validate(original_path, formatted_path):
    orig_doc = Document(original_path)
    fmt_doc = Document(formatted_path)

    orig_paras = extract_paragraph_text(orig_doc)
    fmt_paras = extract_paragraph_text(fmt_doc)
    para_mismatches = diff_text_lists(orig_paras, fmt_paras, "paragraph_text")

    orig_cells = [text for _, text in extract_table_text(orig_doc)]
    fmt_cells = [text for _, text in extract_table_text(fmt_doc)]
    cell_mismatches = diff_text_lists(orig_cells, fmt_cells, "table_cell_text")

    all_mismatches = para_mismatches + cell_mismatches
    replace_mismatches = [m for m in all_mismatches if m["type"] == "replace"]
    relocatable_mismatches = [m for m in all_mismatches if m["type"] in ("delete", "insert")]

    orig_bag = _whole_document_word_bag(orig_doc)
    fmt_bag = _whole_document_word_bag(fmt_doc)
    missing_words = {}
    for word, count in orig_bag.items():
        fmt_count = fmt_bag.get(word, 0)
        if fmt_count < count:
            missing_words[word] = {"original_count": count, "formatted_count": fmt_count}

    word_reconciliation_passed = len(missing_words) == 0
    passed = (len(replace_mismatches) == 0) and word_reconciliation_passed

    return {
        "content_integrity_passed": passed,
        "total_paragraphs_original": len(orig_paras),
        "total_paragraphs_formatted": len(fmt_paras),
        "total_table_cells_original": len(orig_cells),
        "total_table_cells_formatted": len(fmt_cells),
        "replace_mismatch_count": len(replace_mismatches),
        "replace_mismatches": replace_mismatches,
        "relocatable_mismatch_count": len(relocatable_mismatches),
        "relocatable_mismatches": relocatable_mismatches[:30],
        "word_reconciliation_passed": word_reconciliation_passed,
        "missing_words_sample": dict(list(missing_words.items())[:20]),
        "missing_words_count": len(missing_words),
        "mismatch_count": len(replace_mismatches) + (0 if word_reconciliation_passed else len(missing_words)),
        "mismatches": replace_mismatches,
        "allowed_table_insertions_count": len([m for m in relocatable_mismatches if m["type"] == "insert"]),
    }


def build_change_report_markdown(integrity_result, format_result=None, classification_summary=None,
                                  policy=None, toc_result=None):
    lines = []
    lines.append("# Formatting Change Report\n")
    banner = "✅ PASS — zero content loss/alteration detected" if integrity_result["content_integrity_passed"] \
        else "❌ FAIL — content drift detected, review flagged locations below"
    lines.append(f"**Content integrity: {banner}**\n")
    lines.append(f"- Paragraphs compared: {integrity_result['total_paragraphs_original']} "
                 f"(original) vs {integrity_result['total_paragraphs_formatted']} (formatted)")
    lines.append(f"- Table cells compared: {integrity_result['total_table_cells_original']} "
                 f"(original) vs {integrity_result['total_table_cells_formatted']} (formatted)")
    lines.append(f"- Text REPLACED in place (never allowed): {integrity_result['replace_mismatch_count']}")
    lines.append(f"- Paragraphs/cells relocated or added (allowed ONLY if word-reconciliation passes): "
                 f"{integrity_result['relocatable_mismatch_count']}")
    word_ok = "✅ PASS" if integrity_result['word_reconciliation_passed'] else "❌ FAIL"
    lines.append(f"- Whole-document word-reconciliation check: {word_ok} "
                 f"({integrity_result['missing_words_count']} word(s) missing)\n")

    if integrity_result["replace_mismatches"]:
        lines.append("## Content Replacement Detail (must be zero for sign-off)\n")
        for m in integrity_result["replace_mismatches"][:50]:
            lines.append(f"- **{m['location']}** ({m['type']}): "
                         f"`{m['original_text']}` -> `{m['formatted_text']}`")
        lines.append("")

    if not integrity_result['word_reconciliation_passed']:
        lines.append("## Missing Words Detail (must be empty for sign-off)\n")
        for word, counts in integrity_result["missing_words_sample"].items():
            lines.append(f"- '{word}': appeared {counts['original_count']}x before, "
                         f"only {counts['formatted_count']}x after")
        lines.append("")

    if toc_result:
        lines.append("## TOC Chapter-Card Table Restructuring\n")
        lines.append(f"- Attempted: {toc_result.get('attempted')}")
        lines.append(f"- Applied: {toc_result.get('applied')}")
        lines.append(f"- Reason/status: {toc_result.get('reason')}")
        if toc_result.get("chapters_found"):
            lines.append(f"- Chapter entries parsed: {toc_result['chapters_found']}")
            lines.append(f"- Top-level 'SECTION N...' summary lines relocated into the table: "
                         f"{toc_result.get('top_level_lines_removed', 0)}")
            lines.append(f"- Subsection lines (1.1, 2.1.1, etc.) left completely unchanged: "
                         f"{toc_result.get('child_lines_retained_unchanged', 0)}")
        lines.append("")

    if policy and policy.get("theme_fonts"):
        tf = policy["theme_fonts"]
        lines.append("## Reference Theme Fonts Detected\n")
        lines.append(f"- Major (headings) theme font: {tf.get('major_latin') or 'not set / literal font used'}")
        lines.append(f"- Minor (body) theme font: {tf.get('minor_latin') or 'not set / literal font used'}\n")

    if policy and policy.get("typography"):
        lines.append("## Heading Typography (usage-based majority extraction, HEADINGS ONLY)\n")
        for role in ("Title", "H1", "H2", "H3", "H4"):
            t = policy["typography"].get(role, {})
            f = t.get("font", {})
            lines.append(f"- {role}: font={f.get('name')} (source: {f.get('name_source', 'style_default')}), "
                         f"size={f.get('size_pt')}pt, bold={f.get('bold')}, color=#{f.get('color_hex')} "
                         f"(usage samples examined: {t.get('usage_sample_size', 0)})")
        body_font = policy["typography"].get("Body", {}).get("font", {})
        lines.append(f"- Body (never subject to usage-based color voting - style-definition only): "
                     f"font={body_font.get('name')}, size={body_font.get('size_pt')}pt, "
                     f"color=#{body_font.get('color_hex')}")
        lines.append("")

    if policy and policy.get("table_typography"):
        tt = policy["table_typography"]
        lines.append("## Table Typography Detected in Reference (data_table context, top-level tables only)\n")
        for section, label in (("header", "Header (all columns, all header rows)"),
                                ("body", "Body rows")):
            prof = tt.get(section, {})
            if prof.get("apply"):
                lines.append(f"- {label}: font={prof.get('font_name')} "
                             f"(source: {prof.get('font_name_source')}), size={prof.get('size_pt')}pt "
                             f"(source: {prof.get('size_source')}), sample size={prof.get('sample_size')}")
            else:
                lines.append(f"- {label}: NOT applied ({prof.get('reason')})")
        lines.append("")

    if policy and policy.get("table_contexts"):
        lines.append("## Table Header Shading Detected in Reference (top-level tables only)\n")
        tc = policy["table_contexts"]
        for ctx_name, label in (("banner", "Banner/divider tables"),
                                 ("data_table", "Data tables (all header columns/rows)")):
            prof = tc.get(ctx_name, {})
            if prof.get("apply"):
                lines.append(f"- {label}: APPLIED (confidence "
                             f"{prof.get('shading_confidence')}, shading #{prof.get('shading_hex')}, "
                             f"{prof.get('total_header_cells_examined', 0)} header cells examined "
                             f"across {prof.get('sample_size', 0)} table(s))")
            else:
                lines.append(f"- {label}: NOT applied ({prof.get('reason', 'insufficient confidence')})")
        lines.append("")

    if policy and policy.get("footer_design"):
        fd = policy["footer_design"]
        lines.append("## Footer / Page-Number Design Detected in Reference\n")
        lines.append(f"- Reference footer contains a PAGE field: {fd.get('has_page_field')}")
        lines.append("")

    if format_result and format_result.get("footer_result"):
        fr = format_result["footer_result"]
        lines.append("## Footer Applied to Target Document\n")
        lines.append(f"- Applied: {fr.get('applied')} — {fr.get('reason')}")
        lines.append("")

    if classification_summary:
        lines.append("## Structure Classification Summary\n")
        lines.append(f"- Total paragraphs classified: {classification_summary['total_paragraphs']}")
        lines.append(f"- Low-confidence (< 70%) paragraphs flagged for human review: "
                     f"{classification_summary['low_confidence_count']} "
                     f"({classification_summary['low_confidence_pct']}%)")
        exit_ok = "✅ met (< 5%)" if classification_summary["exit_criterion_met_below_5pct"] else "❌ not met (>= 5%)"
        lines.append(f"- Phase-1 rule-tuning exit criterion: {exit_ok}\n")

    if format_result:
        lines.append("## Style Reassignment Summary\n")
        style_counts = {}
        for c in format_result["paragraph_change_log"]:
            style_counts[c["new_style"]] = style_counts.get(c["new_style"], 0) + 1
        for style, count in sorted(style_counts.items(), key=lambda kv: -kv[1]):
            lines.append(f"- {style}: {count} paragraphs")

        if format_result.get("special_headings_applied"):
            lines.append(f"\n- Special front-matter headings matched & styled in target document: "
                         f"{format_result['special_headings_applied']}")

        ttr = format_result.get("table_typography_result", {})
        lines.append(f"\n- Table header cells (all rows/columns, uniform) with font/size normalized: "
                     f"{ttr.get('header_cells_touched', 0)}")
        lines.append(f"- Table body cells with font/size normalized: {ttr.get('body_cells_touched', 0)}")

        touched = format_result.get("tables_touched_by_context", {})
        lines.append(f"- Banner-context tables header-styled: {touched.get('banner', 0)}")
        lines.append(f"- Data-table headers styled (all rows/columns, uniform): {touched.get('data_table', 0)}")
        lines.append(f"- Tables left untouched (context not confidently supported): "
                     f"{touched.get('other_untouched', 0)}")

        lines.append(f"\n- Named style roles redefined from policy: {len(format_result['applied_styles'])}")
        for s in format_result["applied_styles"]:
            f = s.get("font") or {}
            dup = s.get("duplicate_style_objects_updated", 1)
            lines.append(f"  - {s['style_name']} ({dup} style object(s) updated): font={f.get('name')} "
                         f"(source: {f.get('name_source', 'explicit')}), "
                         f"size={f.get('size_pt')}pt, bold={f.get('bold')}, "
                         f"color={f.get('color_hex')}")

    lines.append("\n---\n*Style and layout are edited programmatically; content text is never "
                 "edited in place.*")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="Phase 0 - Validate content integrity and build the change report.")
    ap.add_argument("--original", required=True)
    ap.add_argument("--formatted", required=True)
    ap.add_argument("--classification")
    ap.add_argument("--format-result")
    ap.add_argument("--policy")
    ap.add_argument("--toc-result")
    ap.add_argument("--report-out", default="formatting_change_report.md")
    ap.add_argument("--json-out", default="integrity_result.json")
    args = ap.parse_args()

    result = validate(args.original, args.formatted)
    with open(args.json_out, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    classification_summary = None
    if args.classification:
        with open(args.classification, encoding="utf-8") as f:
            classification_summary = json.load(f)["summary"]

    format_result = None
    if args.format_result:
        with open(args.format_result, encoding="utf-8") as f:
            format_result = json.load(f)

    policy = None
    if args.policy:
        with open(args.policy, encoding="utf-8") as f:
            policy = json.load(f)

    toc_result = None
    if args.toc_result:
        with open(args.toc_result, encoding="utf-8") as f:
            toc_result = json.load(f)

    report_md = build_change_report_markdown(result, format_result, classification_summary, policy, toc_result)
    with open(args.report_out, "w", encoding="utf-8") as f:
        f.write(report_md)

    print(f"[integrity_validator] content_integrity_passed = {result['content_integrity_passed']}")
    print(f"[integrity_validator] Report -> {args.report_out}")
    if not result["content_integrity_passed"]:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
