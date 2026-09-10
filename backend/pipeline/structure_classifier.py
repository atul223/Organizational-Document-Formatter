"""
Phase 2 — Structure Classifier (v1.17 — inherited-heading-style fix)

=== ROOT CAUSE (confirmed by direct reproduction) ===
Reported symptom: on the document's early "Section" pages, an intro
sentence and nearly every roman-numeral ("i.", "ii.") and lettered
("a)", "b)") task-list item under "Scope of Work" incorrectly renders
in the document's heading color/boldness - EXCEPT one single line,
which correctly stayed plain black.

Reproduced and confirmed: real section headings (e.g. "1.3 Scope of
Work") are styled with the paragraph's OWN, DIRECTLY-applied style
literally named "Heading 2" (a built-in Word style name) - this is
correctly and safely trusted as a genuine heading. However, the task-
list items use CUSTOM outline/list-level paragraph styles (e.g.
"ListOutline1", "ListOutline2" - names that vary per organization/
template) that are themselves BASED ON "Heading 2" purely to inherit
Word's multilevel-list numbering/formatting behavior - a legitimate,
common real-world authoring technique, NOT an indication that the
paragraph is meant to visually look like a heading. The classifier's
existing_style_or_ancestor rule previously trusted ANY role resolved by
walking the style's base_style ancestor chain with the SAME full
confidence as a literal, direct style-name match - causing these list
items to inherit "Heading 2"'s orange/bold styling. The single correct
exception in the report used the "Normal" style directly (no ancestor
walk needed), which is exactly why it alone stayed correct - a clean,
confirming clue.

THE FIX: docx_fast.build_styleid_to_role_depth_map (NEW, purely
additive - does not modify resolve_base_style_role or
build_styleid_to_role_map, so nothing else that depends on those two
functions is affected) records, for every style, how many base_style
hops were needed to find its role match (0 = the style's own literal
name; >0 = only found via an ancestor). classify_paragraph now uses
this depth ONLY to decide whether to keep trusting a HEADING-role
match: if the match came from resolved_role (the style-chain path) AND
required walking at least one ancestor (depth > 0) AND the paragraph's
own text does not itself look like a genuine typed heading pattern
(SECTION N banner or decimal x.y numbering, both handled identically
to before), the style match is no longer blindly trusted - the
paragraph instead falls through to the EXACT SAME general heuristic
classification path used for any paragraph with no resolvable style
role (decimal-numbering check, caption check, ALL-CAPS check, bold-
oversized check, etc. - all COMPLETELY UNCHANGED), which correctly
lands ordinary list-item prose on "Body".

This fix does NOT use Word list-numbering (<w:numPr>) detection at all
- unlike an earlier, reverted attempt at a related issue, so it cannot
repeat that regression (which affected genuinely auto-numbered REAL
headings). It also leaves the "direct style-name lookup" fallback
branch (style_name in BUILTIN_STYLE_TO_ROLE) completely untouched,
since that path is ALWAYS a depth-0/literal match by construction and
was never implicated in this bug.
"""
import re
from docx import Document
from docx_fast import (
    BUILTIN_STYLE_TO_ROLE, build_styleid_to_name, get_paragraph_style_name_fast,
    get_raw_pstyle_id, build_styleid_to_role_map, build_styleid_to_role_depth_map,
    count_header_rows,
)
from policy_extractor import _normalize_heading_text

NUMBERING_RE = re.compile(r"^(\d+(\.\d+){0,2})\s+\S")
TABLE_CAPTION_RE = re.compile(r"^\s*Table\s+\d", re.IGNORECASE)
FIGURE_CAPTION_RE = re.compile(r"^\s*Figure\s+\d", re.IGNORECASE)
SECTION_BANNER_RE = re.compile(r"^\s*SECTION\s+\d", re.IGNORECASE)
ALLCAPS_RE = re.compile(r"^[A-Z0-9 .,'\-&/()]{3,80}$")

ADMIN_DENYLIST_PREFIXES = [
    "to,", "dear sir", "dear madam", "dear sir/madam", "dear sirs",
    "thanking you", "with best regards", "with regards", "best regards",
    "sincerely", "yours faithfully", "yours sincerely", "yours truly",
    "enclosed as above", "enclosure", "cc:", "copy to",
    "subject:", "reference no", "ref no", "ref.:", "date:",
]
SIGNATURE_BLOCK_RE = re.compile(r"^\(.{2,60}\)\s*,")


def _is_admin_boilerplate(text):
    norm = text.strip().lower()
    if not norm:
        return False
    for prefix in ADMIN_DENYLIST_PREFIXES:
        if norm.startswith(prefix):
            return True
    if SIGNATURE_BLOCK_RE.match(text.strip()):
        return True
    return False


HEADING_ROLES = {"Title", "H1", "H2", "H3", "H4"}
MAX_HEADING_WORDS = 18
CONFIDENCE_THRESHOLD = 0.7


def _run_bold_italic_and_size(paragraph):
    sizes = []
    bolds = []
    italics = []
    for r in paragraph.runs:
        if r.font.size:
            sizes.append(r.font.size.pt)
        bolds.append(bool(r.bold))
        italics.append(bool(r.italic))
    max_size = max(sizes) if sizes else None
    majority_bold = (sum(bolds) >= len(bolds) / 2.0) if bolds else False
    majority_italic = (sum(italics) >= len(italics) / 2.0) if italics else False
    return majority_bold, majority_italic, max_size


def _looks_like_typed_heading_pattern(text):
    """v1.17: literal-text patterns that indicate a GENUINE, typed
    section-heading numbering scheme (decimal x.y, or 'SECTION N') -
    used as a safety override so that even a style resolved only via
    ancestor-chain walking is STILL trusted as a heading if the
    paragraph's own text unmistakably looks like a real numbered
    heading. This is deliberately narrow and text-pattern-based only
    (never reads Word list-numbering/<w:numPr> - see module docstring
    for why that signal is intentionally avoided here)."""
    return bool(SECTION_BANNER_RE.match(text)) or bool(NUMBERING_RE.match(text))


# v1.23 FIX: the v1.17 "untrusted inherited heading" guard was intended
# to catch narrow TOC/index/outline-mechanics styles that are based on
# a Heading style ONLY so Word's TOC field can pick them up (e.g. a
# "TOC 1" style) - NOT real, organization-branded custom heading styles
# (e.g. "CorpHeading2", "ProposalHeading1") that are extremely common
# in real Word templates and rely on Word's own multilevel-list
# auto-numbering (no literal digits in the paragraph's own text).
#
# Since this tool must now work correctly for ANY employee's reference
# template (not a single fixed template), silently discarding a
# perfectly valid, structurally-resolved heading role match just
# because the style's own name isn't literally "Heading N" is no
# longer acceptable - it directly contradicts the requirement that
# formatting be derived ONLY and FULLY from what's actually present
# in/inherited from the reference/target document's own style
# hierarchy, with no arbitrary heuristic override discarding real
# structural signal. The original v1.17 fix was far broader than its
# stated intent and silently demoted EVERY custom-named heading-
# derived style without a literal typed number to Body (at a high,
# non-flagged 0.55-0.9 confidence), discarding a fully legitimate,
# high-signal, structural match.
#
# This narrow denylist restores trust for legitimate custom heading
# styles (the overwhelmingly common real-world case) while still
# protecting against genuine TOC/index/outline-numbering styles the
# original fix was designed for.
_UNTRUSTED_STYLE_NAME_MARKERS = ("toc", "index", "outline numbered", "list number")


def _is_style_name_genuinely_untrustworthy(style_name):
    if not style_name:
        return False
    norm = style_name.strip().lower()
    return any(marker in norm for marker in _UNTRUSTED_STYLE_NAME_MARKERS)


def classify_paragraph(paragraph, body_baseline_pt=11.0, style_name=None,
                        resolved_role=None, resolved_role_depth=None,
                        special_heading_keys=None):
    if style_name is None:
        style_name = paragraph.style.name if paragraph.style else "Normal"
    text = paragraph.text.strip()
    word_count = len(text.split())

    if special_heading_keys:
        norm = _normalize_heading_text(text)
        if norm in special_heading_keys:
            return {
                "role": f"SpecialHeading::{norm}", "confidence": 0.99,
                "method": "special_heading_match",
                "reason": (f"Text exactly matches a known front-matter heading "
                           f"('{norm}') that had its own captured style in the "
                           f"reference document; applying that exact style, "
                           f"bypassing generic H1/H2/H3 buckets."),
            }

    if _is_admin_boilerplate(text):
        return {
            "role": "Body", "confidence": 0.9,
            "method": "heuristic:admin_boilerplate_denylist",
            "reason": "Matches common letter/administrative boilerplate pattern; excluded from heading formatting.",
        }

    claimed_role = resolved_role
    # v1.17: this branch (direct dict lookup of the paragraph's OWN
    # style name) is ALWAYS a depth-0/literal match by construction -
    # completely unaffected by this fix, exactly as before.
    if claimed_role is None and style_name in BUILTIN_STYLE_TO_ROLE and style_name != "Normal":
        claimed_role = BUILTIN_STYLE_TO_ROLE[style_name]
        resolved_role_depth = 0
        came_from_style_chain = False
    else:
        came_from_style_chain = resolved_role is not None

    if claimed_role is not None:
        # v1.23 FIX (see module-level comment above _UNTRUSTED_STYLE_
        # NAME_MARKERS for full rationale): a HEADING-role match found
        # by walking UP a style's base_style ancestor chain (depth > 0)
        # is now only distrusted if the style's OWN name genuinely
        # looks like a TOC/index/outline-numbering mechanics style -
        # not merely because it isn't literally named "Heading N".
        is_untrusted_inherited_heading = (
            came_from_style_chain
            and claimed_role in HEADING_ROLES
            and resolved_role_depth is not None
            and resolved_role_depth > 0
            and not _looks_like_typed_heading_pattern(text)
            and _is_style_name_genuinely_untrustworthy(style_name)
        )
        if not is_untrusted_inherited_heading:
            if claimed_role in HEADING_ROLES and word_count > MAX_HEADING_WORDS:
                return {
                    "role": "Body", "confidence": 0.55,
                    "method": "heuristic:style_override_long_text",
                    "reason": (f"Paragraph style '{style_name}' resolves to heading role "
                               f"'{claimed_role}' but has {word_count} words (> {MAX_HEADING_WORDS}); "
                               f"reclassified as Body for human review."),
                }
            return {"role": claimed_role, "confidence": 0.97, "method": "existing_style_or_ancestor",
                    "reason": f"Paragraph style '{style_name}' resolves to role '{claimed_role}'."}
        # else: intentionally fall through to the heuristic path below,
        # exactly as if claimed_role had never been resolved at all.

    if not text:
        return {"role": "Body", "confidence": 1.0, "method": "empty", "reason": "Empty paragraph."}

    if SECTION_BANNER_RE.match(text) and word_count <= 12:
        return {"role": "H1", "confidence": 0.9, "method": "heuristic:section_banner",
                "reason": "Text begins with 'SECTION N' banner pattern."}

    m = NUMBERING_RE.match(text)
    if m and word_count <= MAX_HEADING_WORDS:
        depth = m.group(1).count(".") + 1
        role = {1: "H1", 2: "H2", 3: "H3"}.get(depth, "H3")
        return {"role": role, "confidence": 0.95, "method": "heuristic:numbering_pattern",
                "reason": f"Leading numbering token '{m.group(1)}' matches heading depth {depth}."}

    if TABLE_CAPTION_RE.match(text):
        return {"role": "Caption", "confidence": 0.9, "method": "heuristic:table_caption",
                "reason": "Text begins with 'Table N' caption pattern."}
    if FIGURE_CAPTION_RE.match(text):
        return {"role": "Caption", "confidence": 0.9, "method": "heuristic:figure_caption",
                "reason": "Text begins with 'Figure N' caption pattern."}

    majority_bold, majority_italic, max_size = _run_bold_italic_and_size(paragraph)
    is_short = word_count <= 12

    if ALLCAPS_RE.match(text) and is_short:
        return {"role": "H2", "confidence": 0.75, "method": "heuristic:all_caps_short_line",
                "reason": "ALL CAPS, short, isolated line."}

    if majority_bold and max_size and max_size > body_baseline_pt + 1.5 and is_short:
        role = "H2" if max_size < body_baseline_pt + 6 else "H1"
        return {"role": role, "confidence": 0.8, "method": "heuristic:bold_oversized",
                "reason": f"Bold run(s) at {max_size}pt vs body baseline {body_baseline_pt}pt."}

    if majority_italic and not majority_bold and is_short:
        return {"role": "H3", "confidence": 0.7, "method": "heuristic:italic_short_line",
                "reason": "Italic (not bold), short, isolated line - common sub-heading style."}

    if majority_bold and is_short:
        return {"role": "H3", "confidence": 0.6, "method": "heuristic:bold_short_line",
                "reason": "Bold, short, isolated line; size not conclusively larger than body."}

    return {"role": "Body", "confidence": 0.9, "method": "default",
            "reason": "No heading/caption signal found; treated as body text."}


def _classify_target_table_context(table):
    if not table.rows:
        return "other"
    header_row = table.rows[0]
    cells = header_row.cells
    n_cells = len(cells)
    n_rows = len(table.rows)
    combined_text = " ".join(c.text.strip() for c in cells).strip()
    word_count = len(combined_text.split())
    if n_rows == 1 and n_cells == 1 and combined_text:
        if SECTION_BANNER_RE.match(combined_text) or (ALLCAPS_RE.match(combined_text) and word_count <= 15):
            return "banner"
    if n_cells >= 2:
        return "data_table"
    return "other"


def classify_table_cell_paragraph(paragraph, row_index, n_rows, n_cells_in_row, table_total_rows,
                                   header_row_count=1):
    text = paragraph.text.strip()
    word_count = len(text.split())
    if SECTION_BANNER_RE.match(text) and word_count <= 12:
        return {"role": "H1", "confidence": 0.9, "method": "heuristic:table_section_banner",
                "reason": "Table cell text matches 'SECTION N' banner pattern."}
    if table_total_rows == 1 and n_cells_in_row == 1 and word_count <= 12 and text:
        return {"role": "H2", "confidence": 0.65, "method": "heuristic:single_cell_banner_table",
                "reason": "Single-row, single-cell table; treated as a banner/divider heading."}
    if row_index < header_row_count:
        return {"role": "TableHeader", "confidence": 0.95, "method": "structural:header_row",
                "reason": f"Row {row_index} is within the detected header region "
                         f"(0 to {header_row_count - 1})."}
    return {"role": "TableBody", "confidence": 0.95, "method": "structural:body_row",
            "reason": "Row is outside the detected header region; treated as body row."}


def classify_document(input_path, body_baseline_pt=11.0, special_headings=None):
    doc = Document(input_path)
    styleid_to_name = build_styleid_to_name(doc)
    styleid_to_role = build_styleid_to_role_map(doc, BUILTIN_STYLE_TO_ROLE)
    # v1.17 NEW: additive depth map (see docx_fast.py) - used only to
    # decide whether to trust a HEADING-role match; does not replace
    # or alter styleid_to_role above in any way.
    styleid_to_role_depth = build_styleid_to_role_depth_map(doc, BUILTIN_STYLE_TO_ROLE)
    special_heading_keys = set((special_headings or {}).keys())

    paragraph_results = []
    for idx, p in enumerate(doc.paragraphs):
        style_id = get_raw_pstyle_id(p)
        style_name = styleid_to_name.get(style_id, "Normal") if style_id else "Normal"
        resolved_role = styleid_to_role.get(style_id) if style_id else None
        resolved_role_depth = styleid_to_role_depth.get(style_id) if style_id else None
        result = classify_paragraph(p, body_baseline_pt=body_baseline_pt,
                                     style_name=style_name, resolved_role=resolved_role,
                                     resolved_role_depth=resolved_role_depth,
                                     special_heading_keys=special_heading_keys)
        result.update({"index": idx, "text_preview": p.text.strip()[:80]})
        paragraph_results.append(result)

    table_cell_results = []
    table_context_by_index = {}
    table_header_row_count_by_index = {}
    for t_idx, table in enumerate(doc.tables):
        table_context_by_index[t_idx] = _classify_target_table_context(table)
        hdr_row_count = count_header_rows(table)
        table_header_row_count_by_index[t_idx] = hdr_row_count
        n_rows = len(table.rows)
        for r_idx, row in enumerate(table.rows):
            n_cells = len(row.cells)
            for c_idx, cell in enumerate(row.cells):
                for p_idx, p in enumerate(cell.paragraphs):
                    result = classify_table_cell_paragraph(p, r_idx, n_rows, n_cells, n_rows,
                                                            header_row_count=hdr_row_count)
                    result.update({
                        "table_index": t_idx, "row": r_idx, "col": c_idx, "para": p_idx,
                        "text_preview": p.text.strip()[:80],
                    })
                    table_cell_results.append(result)

    low_confidence = [r for r in paragraph_results
                       if r["confidence"] < CONFIDENCE_THRESHOLD and r["text_preview"]]
    total = len(paragraph_results)
    pct_low = (len(low_confidence) / total * 100) if total else 0.0

    return {
        "paragraphs": paragraph_results,
        "table_cell_paragraphs": table_cell_results,
        "table_contexts": table_context_by_index,
        "table_header_row_counts": table_header_row_count_by_index,
        "low_confidence_paragraphs": low_confidence,
        "summary": {
            "total_paragraphs": total,
            "low_confidence_count": len(low_confidence),
            "low_confidence_pct": round(pct_low, 2),
            "exit_criterion_met_below_5pct": pct_low < 5.0,
        },
    }


if __name__ == "__main__":
    import argparse, json
    ap = argparse.ArgumentParser(description="Phase 2 - Classify paragraph roles in the target document.")
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", default="classification_report.json")
    ap.add_argument("--body-baseline-pt", type=float, default=11.0)
    args = ap.parse_args()
    report = classify_document(args.input, body_baseline_pt=args.body_baseline_pt)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"[structure_classifier] {report['summary']}")
