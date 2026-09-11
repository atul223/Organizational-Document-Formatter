"""
Phase 1 — Policy Extractor (v1.22 — per-table-shape banding detection,
fixing the "no colors picked up at all" regression)

=== ROOT CAUSE OF THIS BUG (confirmed by direct reproduction) ===
Reported symptom: after the v1.21 cnfStyle fix, generated tables STILL
showed NO row-banding color at all - a regression even from the
earlier "patchy" state to a complete absence of color.

Directly reproduced and confirmed: v1.19/v1.20/v1.21 all shared one
fundamental architectural assumption - that a document has exactly ONE
single, uniform, alternating body-row color pattern, extracted via a
WHOLE-DOCUMENT majority vote combining every data table's body rows
together indiscriminately. This assumption is FALSE for real, complex
documents (like the user's actual 8-page work plan): it contains SEVERAL
DIFFERENT KINDS of tables that legitimately have DIFFERENT color
schemes - e.g. a family of small "Phase A/B/C/D/E" activity tables
(all sharing the identical header "No. | Activity | Description |
Timing") that share ONE simple, subtle alternating band pattern, AND a
completely different, ONE-OFF "Monthly Progress Report" table whose
rows are each given their OWN distinct, semantically-meaningful color
(not a two-value alternation at all - a per-role color-coding scheme).
When the whole-document vote combines samples from BOTH of these
fundamentally different tables together, the one-off table's wildly
different colors dilute/contradict the votes for BOTH parities,
dragging the computed confidence below the (deliberately conservative)
threshold - causing the ENTIRE extraction to report `apply: False`
GLOBALLY, and therefore applying NO banding AT ALL to ANY table in the
target document, even ones (like the Phase A/B/C/D/E family) that DO
have a perfectly clean, confidently-detectable pattern of their own.
This was directly reproduced with a synthetic document mirroring this
exact structure and confirmed to fail identically.

THE FIX: `docx_fast.table_header_signature()` computes a stable "shape
signature" for a table from its own header cell text (normalized) plus
column count. `extract_body_row_banding()` is now rewritten to GROUP
all reference data tables by this signature FIRST, and run the
parity-vote algorithm SEPARATELY, PER GROUP - so the 5 "Phase" tables
(which all share the identical header signature) are correctly combined
into one confident vote (now backed by many more total body-row samples
than any single one of them alone could offer), while the one-off
"Monthly Progress Report" table forms its OWN, separate group of one
table - which, lacking enough of its own consistent samples, correctly
and safely reports `apply: False` FOR THAT GROUP ONLY, without
poisoning or suppressing the Phase-table family's correctly-detected
pattern in any way.

A single additional "default" fallback profile (the OLD whole-document
vote, computed exactly as before) is ALSO still produced and stored
alongside the per-signature groups, for the rare case where a TARGET
document contains a table whose header signature does not match ANY
signature seen in the reference at all - preserving a reasonable,
best-effort fallback rather than leaving such an unmatched table with
absolutely no banding decision available.

This is fully backward-compatible: any reference document containing
only ONE table "shape" throughout (as in every synthetic test run
against this pipeline so far) reduces to a SINGLE signature group
whose combined vote is mathematically IDENTICAL to the old whole-
document vote - so no previously-verified, already-correct behavior on
those simpler documents changes in any way.

The APPLICATION side (formatting_engine.py) is updated correspondingly:
for each target table, its own header signature is computed the SAME
way, looked up against the reference's per-signature profiles first,
and only falls back to the "default" whole-document profile if no
signature match is found - applying whichever profile is found (if
`apply: True`) exactly as before (uniform per-parity direct shading,
with defensive cnfStyle stripping, both completely unchanged from
v1.21).

All other logic (header shading, table typography, footer table-layout
replication, TOC card builder) is completely UNCHANGED.
"""
import argparse
import json
import re
from collections import Counter
from docx import Document
from docx.oxml.ns import qn
from ooxml_fonts import extract_theme_fonts
from docx_fast import (
    BUILTIN_STYLE_TO_ROLE, build_styleid_to_name, build_styleid_to_role_map,
    get_raw_pstyle_id, count_header_rows, header_row_cells, dedupe_row_cells,
    read_row_cnf_band, read_cell_cnf_band, table_header_signature,
    iter_tables_recursive,
)

ROLE_STYLE_CANDIDATES = {
    "Title":       ["Title"],
    "H1":          ["Heading 1", "Heading1"],
    "H2":          ["Heading 2", "Heading2"],
    "H3":          ["Heading 3", "Heading3"],
    "H4":          ["Heading 4", "Heading4"],
    "Body":        ["Normal", "Body Text", "BodyText"],
    "Caption":     ["Caption"],
    "Quote":       ["Quote", "Intense Quote"],
    "ListBullet":  ["List Bullet", "ListParagraph", "List Paragraph"],
}

MAJOR_FONT_ROLES = {"Title", "H1", "H2", "H3", "H4"}
MINOR_FONT_ROLES = {"Body", "Caption", "Quote", "ListBullet"}

ALL_TYPOGRAPHY_USAGE_OVERRIDE_ROLES = MAJOR_FONT_ROLES | MINOR_FONT_ROLES
HEADING_USAGE_OVERRIDE_ROLES = MAJOR_FONT_ROLES

BODY_LIKE_FONT_NAME_USAGE_ROLES = {"Body", "Caption", "Quote", "ListBullet"}
MIN_MINOR_FONT_NAME_SAMPLES = 1
MIN_MINOR_FONT_NAME_MAJORITY_RATIO = 0.5

# v1.25 FIX (issues #2 and #4 - "yellow/inconsistent body text color"):
# Word ships several BUILT-IN styles (e.g. "Subtitle") that structurally
# inherit from "Normal" (so docx_fast.resolve_base_style_role correctly
# walks up their base_style chain and lands on "Body"), but which are
# semantically DISTINCT, one-off editorial/decorative elements - NOT
# representative body prose. A reference document typically contains
# only ONE such paragraph (e.g. a single italic/colored subtitle line
# under the cover title), so if it happens to carry its own explicit
# color (a very common real-world authoring choice, since subtitles are
# deliberately styled differently from body text), that single sample
# becomes the ENTIRE usage-vote "majority" for Body typography under the
# v1.24 fix - incorrectly painting all real body text with a color/
# font meant only for that one decorative line. This is a genuine,
# confirmed root cause (traced directly to the reference document's own
# "Subtitle"-styled line), not a hypothetical edge case.
#
# The fix below EXCLUDES known non-prose, semantically-distinct built-in
# Word style names from contributing to Body-role usage voting, while
# leaving ALL other usage-voting behavior (including the "one real
# sample is enough evidence" principle established for the "any
# reference template" requirement) completely untouched. This is
# intentionally scoped ONLY to the Body role and ONLY to well-known,
# non-generic Word style names - it does not affect heading roles
# (which already have their own, separately-vetted resolution logic)
# and does not reintroduce any statistical-sample-size threshold.
DISTINCT_NON_BODY_STYLE_NAMES = {
    "subtitle", "date", "signature", "salutation", "closing",
    "e-mail signature", "envelope address", "envelope return",
    "message header",
}


def _is_distinct_non_body_style(style_name):
    if not style_name:
        return False
    return style_name.strip().lower() in DISTINCT_NON_BODY_STYLE_NAMES


SECTION_BANNER_RE = re.compile(r"^\s*SECTION\s+\d", re.IGNORECASE)
ALLCAPS_RE = re.compile(r"^[A-Z0-9 .,'\-&/()]{3,80}$")

NO_FILL_SENTINEL = "__NO_FILL__"

MIN_BANDING_SAMPLES_PER_PARITY = 1
MIN_BANDING_CONFIDENCE = 0.5

SPECIAL_HEADING_TEXTS = [
    "TABLE OF CONTENTS",
    "LIST OF TABLES",
    "LIST OF FIGURES",
    "LIST OF ABBREVIATIONS",
    "ABBREVIATIONS",
    "EXECUTIVE SUMMARY",
    "LIST OF ANNEXES",
    "ANNEXES",
    "APPENDICES",
    "LIST OF APPENDICES",
    "GLOSSARY",
    "ACRONYMS",
    "LIST OF ACRONYMS",
]

TOC_HEADING_NORM = "TABLE OF CONTENTS"

MAX_USAGE_SAMPLE_WORDS = 20
ROLES_WITH_USAGE_WORD_COUNT_LIMIT = MAJOR_FONT_ROLES


def _normalize_heading_text(text):
    return re.sub(r"\s+", " ", text or "").strip().upper()


def _font_props(style):
    font = style.font
    props = {}
    try:
        props["name"] = font.name
    except Exception:
        props["name"] = None
    try:
        props["size_pt"] = font.size.pt if font.size else None
    except Exception:
        props["size_pt"] = None
    try:
        props["bold"] = font.bold
    except Exception:
        props["bold"] = None
    try:
        props["italic"] = font.italic
    except Exception:
        props["italic"] = None
    try:
        if font.color and font.color.type is not None and font.color.rgb:
            props["color_hex"] = str(font.color.rgb)
        else:
            props["color_hex"] = None
    except Exception:
        props["color_hex"] = None
    try:
        props["underline"] = bool(font.underline)
    except Exception:
        props["underline"] = None
    return props


def _paragraph_props(style):
    pf = style.paragraph_format
    out = {}
    try:
        out["alignment"] = str(pf.alignment) if pf.alignment is not None else None
    except Exception:
        out["alignment"] = None
    try:
        out["space_before_pt"] = pf.space_before.pt if pf.space_before else None
    except Exception:
        out["space_before_pt"] = None
    try:
        out["space_after_pt"] = pf.space_after.pt if pf.space_after else None
    except Exception:
        out["space_after_pt"] = None
    try:
        out["line_spacing"] = pf.line_spacing
    except Exception:
        out["line_spacing"] = None
    try:
        out["first_line_indent_in"] = (pf.first_line_indent.inches
                                        if pf.first_line_indent else None)
    except Exception:
        out["first_line_indent_in"] = None
    try:
        out["left_indent_in"] = pf.left_indent.inches if pf.left_indent else None
    except Exception:
        out["left_indent_in"] = None
    return out


# ---------------------------------------------------------------------
# v1.25 NEW (issue #1 - "Aptos font appearing"): resolves the EFFECTIVE
# font/paragraph properties Word would actually RENDER for a given
# style, by walking that style's own base_style ancestor chain and
# taking the first non-None value found for each property. This is
# needed because many real-world reference documents leave certain
# properties (very commonly the font NAME) completely unset at the
# style level, relying entirely on the style's ancestor chain -> the
# document's THEME to render the visible font. A direct, literal-only
# read of a single style's own definition can therefore come back
# entirely empty even though Word renders a perfectly well-defined,
# visible font for it - and leaving that property as None caused
# downstream formatting code to skip setting it entirely, silently
# leaving whatever font the TARGET document happened to already have
# (e.g. Word's newer default "Aptos") completely untouched.
# ---------------------------------------------------------------------
def _style_own_theme_font_ref(style):
    """python-docx's Font.name property ONLY ever reads the LITERAL
    <w:rFonts w:ascii="..."> attribute - it returns None whenever a
    style's run-properties set only a THEME font reference (e.g.
    <w:rFonts w:asciiTheme="majorHAnsi"/>), even though Word visually
    renders that theme font perfectly well. Without checking for this
    directly at the raw-XML level, a style-chain walk that treats
    "font.name is None" as "this style sets no font at all" would
    incorrectly skip PAST a style that legitimately sets a theme font
    (extremely common for "Title"/heading styles specifically) and
    instead pick up a less-specific ANCESTOR style's literal font -
    which is the WRONG font per Word's actual cascade priority (the
    nearest/most-specific style in the chain always wins for whatever
    it defines, whether literal or theme-based).

    Returns "major", "minor", or None."""
    try:
        rPr = style.element.rPr
    except Exception:
        return None
    if rPr is None:
        return None
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        return None
    theme_val = rFonts.get(qn('w:asciiTheme')) or rFonts.get(qn('w:hAnsiTheme'))
    if not theme_val:
        return None
    theme_val_lower = theme_val.lower()
    if "major" in theme_val_lower:
        return "major"
    if "minor" in theme_val_lower:
        return "minor"
    return None


def _resolve_effective_style_font(style, theme_fonts=None, max_depth=15):
    theme_fonts = theme_fonts or {}
    result = {"name": None, "size_pt": None, "bold": None, "italic": None,
              "color_hex": None, "underline": None}
    seen_ids = set()
    current = style
    depth = 0
    while current is not None and depth < max_depth and any(v is None for v in result.values()):
        f = getattr(current, "font", None)
        if f is not None:
            if result["name"] is None:
                try:
                    if f.name:
                        result["name"] = f.name
                        result["name_source"] = "style_chain_literal"
                except Exception:
                    pass
                # v1.25: if THIS level's own rPr sets a THEME font
                # reference (and no literal font was found above), that
                # theme reference takes priority over any ancestor's
                # literal font - resolve it now and stop walking for
                # "name" specifically.
                if result["name"] is None:
                    theme_ref = _style_own_theme_font_ref(current)
                    if theme_ref == "major" and theme_fonts.get("major_latin"):
                        result["name"] = theme_fonts["major_latin"]
                        result["name_source"] = "style_chain_theme_major"
                    elif theme_ref == "minor" and theme_fonts.get("minor_latin"):
                        result["name"] = theme_fonts["minor_latin"]
                        result["name_source"] = "style_chain_theme_minor"
            if result["size_pt"] is None:
                try:
                    if f.size:
                        result["size_pt"] = f.size.pt
                except Exception:
                    pass
            if result["bold"] is None:
                try:
                    if f.bold is not None:
                        result["bold"] = f.bold
                except Exception:
                    pass
            if result["italic"] is None:
                try:
                    if f.italic is not None:
                        result["italic"] = f.italic
                except Exception:
                    pass
            if result["color_hex"] is None:
                try:
                    if f.color and f.color.type is not None and f.color.rgb:
                        result["color_hex"] = str(f.color.rgb)
                except Exception:
                    pass
            if result["underline"] is None:
                try:
                    if f.underline is not None:
                        result["underline"] = bool(f.underline)
                except Exception:
                    pass
        style_id = getattr(current, "style_id", None)
        if style_id is not None:
            if style_id in seen_ids:
                break
            seen_ids.add(style_id)
        try:
            current = current.base_style
        except Exception:
            break
        depth += 1
    return result


def extract_typography_from_style_defs(doc, reference_path):
    """Baseline pass: reads each role's STYLE DEFINITION (e.g. what
    Word's 'Normal' style itself declares). This is later treated only
    as a fallback baseline - see extract_typography_from_usage() below,
    which overrides these values with what's ACTUALLY, visibly used in
    the document whenever real usage data is available."""
    styles_by_name = {s.name: s for s in doc.styles}
    typography = {}
    for role, candidates in ROLE_STYLE_CANDIDATES.items():
        found = None
        for cand in candidates:
            if cand in styles_by_name:
                found = styles_by_name[cand]
                break
        if found is None:
            typography[role] = {"source_style": None, "font": {}, "paragraph": {}}
            continue
        typography[role] = {
            "source_style": found.name,
            "font": _font_props(found),
            "paragraph": _paragraph_props(found),
        }

    theme_fonts = extract_theme_fonts(reference_path)

    for role in MAJOR_FONT_ROLES:
        block = typography[role]
        if block.get("font", {}).get("name"):
            continue
        if theme_fonts.get("major_latin"):
            block.setdefault("font", {})["name"] = theme_fonts["major_latin"]
            block["font"]["name_source"] = "theme_major"

    body_block = typography.get("Body", {})
    if not body_block.get("font", {}).get("name") and theme_fonts.get("minor_latin"):
        body_block.setdefault("font", {})["name"] = theme_fonts["minor_latin"]
        body_block["font"]["name_source"] = "theme_minor"
    body_font_name = typography.get("Body", {}).get("font", {}).get("name")
    body_size_pt = typography.get("Body", {}).get("font", {}).get("size_pt")

    for role in MINOR_FONT_ROLES:
        if role == "Body":
            continue
        block = typography[role]
        if not block.get("font", {}).get("name"):
            if body_font_name:
                block.setdefault("font", {})["name"] = body_font_name
                block["font"]["name_source"] = "body_fallback_priority"
            elif theme_fonts.get("minor_latin"):
                block.setdefault("font", {})["name"] = theme_fonts["minor_latin"]
                block["font"]["name_source"] = "theme_minor"
        if not block.get("font", {}).get("size_pt") and body_size_pt:
            block.setdefault("font", {})["size_pt"] = body_size_pt
            block["font"]["size_source"] = "body_fallback_priority"

    return typography


def extract_typography_from_usage(doc, base_typography):
    """Examines EVERY typography role - Title/H1-H4 AND Body/Caption/
    Quote/ListBullet - and votes on the font/paragraph properties
    ACTUALLY applied via direct/manual run-level formatting across
    every matching paragraph in the reference document, overriding a
    (possibly stale) style-definition baseline. A single genuine usage
    example is sufficient evidence per the "any employee, any reference
    template" requirement.

    v1.25 FIX: paragraphs whose OWN literal style is a known, distinct,
    non-generic-prose built-in Word style (e.g. "Subtitle") are EXCLUDED
    from contributing to Body-role usage voting, even though such
    styles structurally inherit the Body role via their base_style
    ancestor chain - see DISTINCT_NON_BODY_STYLE_NAMES above for the
    full, directly-confirmed rationale (a single decorative/editorial
    "Subtitle" line was incorrectly hijacking the entire document's
    Body text color).
    """
    styleid_to_role = build_styleid_to_role_map(doc, BUILTIN_STYLE_TO_ROLE)
    styleid_to_name = build_styleid_to_name(doc)

    votes = {role: {
        "name": Counter(), "size_pt": Counter(), "color_hex": Counter(),
        "bold": [], "italic": [],
        "alignment": Counter(), "space_before_pt": Counter(), "space_after_pt": Counter(),
    } for role in ALL_TYPOGRAPHY_USAGE_OVERRIDE_ROLES}
    sample_counts = Counter()

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style_id = get_raw_pstyle_id(p)
        if style_id is None:
            role = "Body"
            own_style_name = "Normal"
        else:
            role = styleid_to_role.get(style_id)
            own_style_name = styleid_to_name.get(style_id, "")
        if role is None or role not in ALL_TYPOGRAPHY_USAGE_OVERRIDE_ROLES:
            continue
        # v1.25 FIX: exclude distinct non-prose styles from Body voting.
        if role == "Body" and _is_distinct_non_body_style(own_style_name):
            continue
        if role in ROLES_WITH_USAGE_WORD_COUNT_LIMIT and len(text.split()) > MAX_USAGE_SAMPLE_WORDS:
            continue
        rep_run = next((r for r in p.runs if r.text.strip()), None)
        if rep_run is None:
            continue
        sample_counts[role] += 1
        v = votes[role]
        try:
            if rep_run.font.name:
                v["name"][rep_run.font.name] += 1
        except Exception:
            pass
        try:
            if rep_run.font.size:
                v["size_pt"][round(rep_run.font.size.pt, 1)] += 1
        except Exception:
            pass
        try:
            if rep_run.font.color and rep_run.font.color.type is not None and rep_run.font.color.rgb:
                v["color_hex"][str(rep_run.font.color.rgb)] += 1
        except Exception:
            pass
        try:
            if rep_run.font.bold is not None:
                v["bold"].append(rep_run.font.bold)
        except Exception:
            pass
        try:
            if rep_run.font.italic is not None:
                v["italic"].append(rep_run.font.italic)
        except Exception:
            pass
        pf = p.paragraph_format
        try:
            if pf.alignment is not None:
                v["alignment"][str(pf.alignment)] += 1
        except Exception:
            pass
        try:
            if pf.space_before is not None:
                v["space_before_pt"][round(pf.space_before.pt, 1)] += 1
        except Exception:
            pass
        try:
            if pf.space_after is not None:
                v["space_after_pt"][round(pf.space_after.pt, 1)] += 1
        except Exception:
            pass

    result = {}
    for role, base in base_typography.items():
        if role not in ALL_TYPOGRAPHY_USAGE_OVERRIDE_ROLES:
            result[role] = base
            continue

        font = dict(base.get("font", {}))
        para = dict(base.get("paragraph", {}))
        v = votes.get(role)
        n_samples = sample_counts.get(role, 0)
        if v and n_samples >= 1:
            if v["name"]:
                font["name"] = v["name"].most_common(1)[0][0]
                font["name_source"] = "usage_majority"
            if v["size_pt"]:
                font["size_pt"] = v["size_pt"].most_common(1)[0][0]
                font["size_source"] = "usage_majority"
            if v["bold"]:
                font["bold"] = sum(v["bold"]) >= len(v["bold"]) / 2.0
            if v["italic"]:
                font["italic"] = sum(v["italic"]) >= len(v["italic"]) / 2.0
            if v["color_hex"]:
                top_color, top_count = v["color_hex"].most_common(1)[0]
                total_color_votes = sum(v["color_hex"].values())
                if total_color_votes > 0 and (top_count / total_color_votes) >= 0.5:
                    font["color_hex"] = top_color
            if v["alignment"]:
                para["alignment"] = v["alignment"].most_common(1)[0][0]
            if v["space_before_pt"]:
                para["space_before_pt"] = v["space_before_pt"].most_common(1)[0][0]
            if v["space_after_pt"]:
                para["space_after_pt"] = v["space_after_pt"].most_common(1)[0][0]
        result[role] = {
            "source_style": base.get("source_style"),
            "font": font, "paragraph": para,
            "usage_sample_size": n_samples,
        }
    return result


def extract_minor_role_font_name_from_usage(doc, typography):
    """SECONDARY safety net only - fires only if a role's font name is
    STILL empty after the primary usage pass above (i.e. the role has
    literally zero qualifying paragraphs anywhere in the reference
    document)."""
    roles_needing_fallback = {
        role for role in BODY_LIKE_FONT_NAME_USAGE_ROLES
        if not typography.get(role, {}).get("font", {}).get("name")
    }
    if not roles_needing_fallback:
        return typography

    styleid_to_role = build_styleid_to_role_map(doc, BUILTIN_STYLE_TO_ROLE)
    styleid_to_name = build_styleid_to_name(doc)
    votes = {role: Counter() for role in roles_needing_fallback}
    total_runs_examined = Counter()

    for p in doc.paragraphs:
        style_id = get_raw_pstyle_id(p)
        if style_id is None:
            role = "Body"
            own_style_name = "Normal"
        else:
            role = styleid_to_role.get(style_id)
            own_style_name = styleid_to_name.get(style_id, "")
        if role is None or role not in roles_needing_fallback:
            continue
        # v1.25 FIX: same exclusion as the primary pass above.
        if role == "Body" and _is_distinct_non_body_style(own_style_name):
            continue
        for r in p.runs:
            if not r.text.strip():
                continue
            total_runs_examined[role] += 1
            try:
                if r.font.name:
                    votes[role][r.font.name] += 1
            except Exception:
                pass

    for role in roles_needing_fallback:
        total = total_runs_examined[role]
        v = votes[role]
        if total < MIN_MINOR_FONT_NAME_SAMPLES or not v:
            continue
        top_name, top_count = v.most_common(1)[0]
        ratio = top_count / total
        if ratio >= MIN_MINOR_FONT_NAME_MAJORITY_RATIO:
            block = typography.setdefault(role, {"font": {}, "paragraph": {}})
            block.setdefault("font", {})["name"] = top_name
            block["font"]["name_source"] = "minor_role_usage_fallback"
            block["font"]["name_usage_sample_size"] = total
            block["font"]["name_usage_confidence"] = round(ratio, 2)

    return typography


def extract_typography(doc, reference_path):
    base = extract_typography_from_style_defs(doc, reference_path)
    with_usage = extract_typography_from_usage(doc, base)
    final = extract_minor_role_font_name_from_usage(doc, with_usage)
    return final


def extract_special_headings(doc, theme_fonts=None, fallback_font_name=None, fallback_size_pt=None):
    """v1.25 FIX (issue #1 - Aptos font appearing): for each matched
    special front-matter heading (e.g. "TABLE OF CONTENTS"), the
    representative run's OWN literal font properties are read first
    (unchanged). But whenever a property comes back None - which is
    very common for the font NAME specifically, since many real-world
    templates leave headings' fonts entirely theme-driven rather than
    literally set - the property is now resolved via the paragraph's
    OWN style ancestor chain (_resolve_effective_style_font), which is
    exactly what Word itself renders. If the font NAME is still
    unresolved after that (i.e. neither the run nor any style in its
    ancestor chain ever sets one - fully theme-only), it falls back to
    the reference document's theme MAJOR font (front-matter special
    headings are heading-like elements), then finally to the already-
    resolved Body font name as a last resort. This ensures a concrete,
    correct font name is always captured - never left as None, which
    previously caused the formatting engine to silently skip setting
    the font at all, leaving the TARGET document's own (possibly
    unrelated, e.g. "Aptos") font completely untouched."""
    theme_fonts = theme_fonts or {}
    known_set = {_normalize_heading_text(t) for t in SPECIAL_HEADING_TEXTS}
    found = {}
    for p in doc.paragraphs:
        norm = _normalize_heading_text(p.text)
        if norm in known_set and norm not in found and p.runs:
            rep_run = next((r for r in p.runs if r.text.strip()), p.runs[0])
            font_props = {}
            try:
                font_props["name"] = rep_run.font.name
            except Exception:
                font_props["name"] = None
            try:
                font_props["size_pt"] = rep_run.font.size.pt if rep_run.font.size else None
            except Exception:
                font_props["size_pt"] = None
            try:
                font_props["bold"] = rep_run.font.bold
            except Exception:
                font_props["bold"] = None
            try:
                font_props["italic"] = rep_run.font.italic
            except Exception:
                font_props["italic"] = None
            try:
                if rep_run.font.color and rep_run.font.color.type is not None and rep_run.font.color.rgb:
                    font_props["color_hex"] = str(rep_run.font.color.rgb)
                else:
                    font_props["color_hex"] = None
            except Exception:
                font_props["color_hex"] = None
            try:
                font_props["underline"] = bool(rep_run.font.underline)
            except Exception:
                font_props["underline"] = None

            # v1.25: fill any still-empty properties via the paragraph's
            # own style ancestor chain (what Word actually renders).
            if any(font_props.get(k) is None for k in ("name", "size_pt", "bold", "italic", "color_hex")):
                try:
                    effective = _resolve_effective_style_font(p.style, theme_fonts=theme_fonts)
                except Exception:
                    effective = {}
                for k in ("name", "size_pt", "bold", "italic", "color_hex", "underline"):
                    if font_props.get(k) is None and effective.get(k) is not None:
                        font_props[k] = effective[k]
                        font_props[f"{k}_source"] = "style_chain_effective"

            # v1.25: final fallback for font NAME only - theme major,
            # then the already-resolved Body font name.
            if not font_props.get("name"):
                if theme_fonts.get("major_latin"):
                    font_props["name"] = theme_fonts["major_latin"]
                    font_props["name_source"] = "theme_major_fallback"
                elif fallback_font_name:
                    font_props["name"] = fallback_font_name
                    font_props["name_source"] = "body_font_fallback"
            if not font_props.get("size_pt") and fallback_size_pt:
                font_props["size_pt"] = fallback_size_pt
                font_props["size_source"] = "body_size_fallback"

            pf = p.paragraph_format
            para_props = {}
            try:
                para_props["alignment"] = str(pf.alignment) if pf.alignment is not None else None
            except Exception:
                para_props["alignment"] = None
            try:
                para_props["space_before_pt"] = pf.space_before.pt if pf.space_before else None
            except Exception:
                para_props["space_before_pt"] = None
            try:
                para_props["space_after_pt"] = pf.space_after.pt if pf.space_after else None
            except Exception:
                para_props["space_after_pt"] = None

            found[norm] = {
                "font": font_props,
                "paragraph": para_props,
                "matched_text_sample": p.text.strip(),
                "paragraph_index": next((i for i, pp in enumerate(doc.paragraphs) if pp._p is p._p), None),
            }

    return found


def _get_paragraph_bottom_border(paragraph):
    pPr = paragraph._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        return None
    bottom = pBdr.find(qn('w:bottom'))
    if bottom is None:
        return None
    return {
        "val": bottom.get(qn('w:val')),
        "sz": bottom.get(qn('w:sz')),
        "space": bottom.get(qn('w:space')),
        "color": bottom.get(qn('w:color')),
    }


_COND_FORMAT_TYPE_HEADER = "firstRow"
_COND_FORMAT_TYPE_BAND1 = "band1Horz"
_COND_FORMAT_TYPE_BAND2 = "band2Horz"


def _get_table_style_conditional_tcPr_rPr(table, cond_type=_COND_FORMAT_TYPE_HEADER, max_depth=15):
    try:
        style = table.style
    except Exception:
        style = None
    if style is None:
        return None, None

    seen_style_ids = set()
    current = style
    depth = 0
    while current is not None and depth < max_depth:
        style_element = getattr(current, "element", None)
        if style_element is not None:
            for tblStylePr in style_element.findall(qn('w:tblStylePr')):
                if tblStylePr.get(qn('w:type')) == cond_type:
                    tcPr = tblStylePr.find(qn('w:tcPr'))
                    rPr = tblStylePr.find(qn('w:rPr'))
                    if tcPr is not None or rPr is not None:
                        return tcPr, rPr
        style_id = getattr(current, "style_id", None)
        if style_id is not None:
            if style_id in seen_style_ids:
                break
            seen_style_ids.add(style_id)
        try:
            current = current.base_style
        except Exception:
            break
        depth += 1
    return None, None


def _table_row_band_size(table):
    try:
        tblPr = table._tbl.tblPr
        if tblPr is not None:
            band_size_el = tblPr.find(qn('w:tblStyleRowBandSize'))
            if band_size_el is not None:
                val = band_size_el.get(qn('w:val'))
                if val:
                    size = int(val)
                    if size >= 1:
                        return size
    except Exception:
        pass
    return 1


def _direct_cell_shading_hex(cell):
    try:
        tcPr = cell._tc.tcPr
        if tcPr is not None:
            shd = tcPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                if fill and fill.upper() != "AUTO":
                    return fill.upper()
    except Exception:
        pass
    return None


def _style_conditional_shading_hex(table, cond_type=_COND_FORMAT_TYPE_HEADER):
    tcPr, _ = _get_table_style_conditional_tcPr_rPr(table, cond_type)
    if tcPr is None:
        return None
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        return None
    fill = shd.get(qn('w:fill'))
    if fill and fill.upper() != "AUTO":
        return fill.upper()
    return None


def _effective_cell_shading(cell, table, cond_type=_COND_FORMAT_TYPE_HEADER):
    direct = _direct_cell_shading_hex(cell)
    if direct:
        return direct
    return _style_conditional_shading_hex(table, cond_type)


def _effective_cell_shading_for_band(cell, table, band_cond_type):
    direct = _direct_cell_shading_hex(cell)
    if direct:
        return direct
    return _style_conditional_shading_hex(table, band_cond_type)


def _style_conditional_font_props(table, cond_type=_COND_FORMAT_TYPE_HEADER):
    _, rPr = _get_table_style_conditional_tcPr_rPr(table, cond_type)
    if rPr is None:
        return {}
    props = {}
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        name = rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))
        if name:
            props["name"] = name
    sz = rPr.find(qn('w:sz'))
    if sz is not None and sz.get(qn('w:val')):
        try:
            props["size_pt"] = int(sz.get(qn('w:val'))) / 2.0
        except Exception:
            pass
    b = rPr.find(qn('w:b'))
    if b is not None:
        val = b.get(qn('w:val'))
        props["bold"] = (val not in ("false", "0", "off"))
    i = rPr.find(qn('w:i'))
    if i is not None:
        val = i.get(qn('w:val'))
        props["italic"] = (val not in ("false", "0", "off"))
    color = rPr.find(qn('w:color'))
    if color is not None:
        val = color.get(qn('w:val'))
        if val and val.upper() != "AUTO":
            props["color_hex"] = val.upper()
    return props


def _cell_effective_run_props(cell, table, cond_type=_COND_FORMAT_TYPE_HEADER):
    names, sizes, colors, bolds = Counter(), Counter(), Counter(), []
    for p in cell.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            try:
                if r.font.name:
                    names[r.font.name] += 1
            except Exception:
                pass
            try:
                if r.font.size:
                    sizes[round(r.font.size.pt, 1)] += 1
            except Exception:
                pass
            try:
                if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                    colors[str(r.font.color.rgb)] += 1
            except Exception:
                pass
            bolds.append(bool(r.bold))

    style_fallback = _style_conditional_font_props(table, cond_type)
    result = {
        "name": names.most_common(1)[0][0] if names else style_fallback.get("name"),
        "size_pt": sizes.most_common(1)[0][0] if sizes else style_fallback.get("size_pt"),
        "color_hex": colors.most_common(1)[0][0] if colors else style_fallback.get("color_hex"),
        "bold": (sum(bolds) >= len(bolds) / 2.0) if bolds else style_fallback.get("bold"),
    }
    return result


def _cell_run_props(cell):
    names, sizes, colors, bolds = Counter(), Counter(), Counter(), []
    for p in cell.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            try:
                if r.font.name:
                    names[r.font.name] += 1
            except Exception:
                pass
            try:
                if r.font.size:
                    sizes[round(r.font.size.pt, 1)] += 1
            except Exception:
                pass
            try:
                if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                    colors[str(r.font.color.rgb)] += 1
            except Exception:
                pass
            bolds.append(bool(r.bold))
    return names, sizes, colors, bolds


def _looks_like_card_number(text):
    t = text.strip()
    return bool(re.match(r"^\d{1,3}$", t))


def _find_toc_card_table(doc, toc_heading_para_index, search_window=25):
    if toc_heading_para_index is None:
        return None

    body = doc.element.body
    para_elements = [p._p for p in doc.paragraphs]
    try:
        anchor_elm = para_elements[toc_heading_para_index]
    except IndexError:
        anchor_elm = None

    body_children = list(body)
    if anchor_elm is not None and anchor_elm in body_children:
        anchor_pos = body_children.index(anchor_elm)
    else:
        anchor_pos = 0

    candidate_tables = []
    for tbl in doc.tables:
        if tbl._tbl in body_children:
            pos = body_children.index(tbl._tbl)
            if pos >= anchor_pos:
                candidate_tables.append((pos, tbl))
    candidate_tables.sort(key=lambda x: x[0])

    for _, tbl in candidate_tables:
        if not tbl.rows or len(tbl.columns) < 2:
            continue
        matches = 0
        total = 0
        for row in tbl.rows:
            total += 1
            first_cell_text = row.cells[0].text
            if _looks_like_card_number(first_cell_text):
                matches += 1
        if total > 0 and (matches / total) >= 0.5:
            return tbl
    return None


def _extract_table_border_style(table):
    try:
        tblPr = table._tbl.tblPr
        if tblPr is not None:
            borders = tblPr.find(qn('w:tblBorders'))
            if borders is not None:
                vals = [child.get(qn('w:val')) for child in borders]
                if all(v in (None, 'nil', 'none') for v in vals):
                    return "none"
                return "default"
    except Exception:
        pass
    try:
        style_id = table.style.style_id if table.style else None
    except Exception:
        style_id = None
    if style_id and 'grid' not in style_id.lower() and 'table' == style_id.lower():
        return "none"
    return "default"


def extract_toc_card_design(doc, special_headings):
    result = {
        "detected": False,
        "title_underline": None,
        "subtitle_style": None,
        "n_cols": None,
        "columns": [],
        "col_widths_emu": None,
        "sample_rows": 0,
        "border_style": "none",
    }

    toc_entry = special_headings.get(TOC_HEADING_NORM)
    if not toc_entry:
        return result
    toc_idx = toc_entry.get("paragraph_index")
    if toc_idx is None:
        return result

    toc_paragraph = doc.paragraphs[toc_idx]
    result["title_underline"] = _get_paragraph_bottom_border(toc_paragraph)

    for p in doc.paragraphs[toc_idx + 1: toc_idx + 6]:
        text = p.text.strip()
        if not text:
            continue
        word_count = len(text.split())
        majority_italic = False
        rep_font = None
        if p.runs:
            italics = [bool(r.italic) for r in p.runs if r.text.strip()]
            majority_italic = bool(italics) and sum(italics) >= len(italics) / 2.0
            rep_run = next((r for r in p.runs if r.text.strip()), None)
            if rep_run is not None:
                rep_font = {}
                try:
                    rep_font["name"] = rep_run.font.name
                except Exception:
                    rep_font["name"] = None
                try:
                    rep_font["size_pt"] = rep_run.font.size.pt if rep_run.font.size else None
                except Exception:
                    rep_font["size_pt"] = None
                try:
                    rep_font["bold"] = rep_run.font.bold
                except Exception:
                    rep_font["bold"] = None
                try:
                    rep_font["italic"] = rep_run.font.italic
                except Exception:
                    rep_font["italic"] = None
                try:
                    if rep_run.font.color and rep_run.font.color.type is not None and rep_run.font.color.rgb:
                        rep_font["color_hex"] = str(rep_run.font.color.rgb)
                    else:
                        rep_font["color_hex"] = None
                except Exception:
                    rep_font["color_hex"] = None
        if majority_italic and word_count <= 8:
            result["subtitle_style"] = {"font": rep_font, "text_sample": text}
        break

    card_table = _find_toc_card_table(doc, toc_idx)
    if card_table is None:
        return result

    result["border_style"] = _extract_table_border_style(card_table)

    n_cols = len(card_table.columns)
    result["n_cols"] = n_cols
    result["sample_rows"] = len(card_table.rows)

    col_shading = [Counter() for _ in range(n_cols)]
    col_fonts_max = [[] for _ in range(n_cols)]
    col_fonts_min = [[] for _ in range(n_cols)]

    for row in card_table.rows:
        for c_idx, cell in enumerate(row.cells):
            shade = _effective_cell_shading(cell, card_table)
            if shade:
                col_shading[c_idx][shade] += 1
            candidates = []
            for p in cell.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    try:
                        size = r.font.size.pt if r.font.size else None
                    except Exception:
                        size = None
                    candidates.append((size or 0, r))
            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                max_r = candidates[0][1]
                candidates.sort(key=lambda x: x[0], reverse=False)
                min_r = candidates[0][1]

                def _props(r):
                    props = {}
                    try:
                        props["name"] = r.font.name
                    except Exception:
                        props["name"] = None
                    try:
                        props["size_pt"] = r.font.size.pt if r.font.size else None
                    except Exception:
                        props["size_pt"] = None
                    try:
                        props["bold"] = r.font.bold
                    except Exception:
                        props["bold"] = None
                    try:
                        props["italic"] = r.font.italic
                    except Exception:
                        props["italic"] = None
                    try:
                        if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                            props["color_hex"] = str(r.font.color.rgb)
                        else:
                            props["color_hex"] = None
                    except Exception:
                        props["color_hex"] = None
                    return props

                col_fonts_max[c_idx].append(_props(max_r))
                col_fonts_min[c_idx].append(_props(min_r))

    def _majority_font(font_list, key):
        vals = Counter(f.get(key) for f in font_list if f.get(key) is not None)
        return vals.most_common(1)[0][0] if vals else None

    for c_idx in range(n_cols):
        shading_hex = col_shading[c_idx].most_common(1)[0][0] if col_shading[c_idx] else None
        title_font = {
            "name": _majority_font(col_fonts_max[c_idx], "name"),
            "size_pt": _majority_font(col_fonts_max[c_idx], "size_pt"),
            "bold": _majority_font(col_fonts_max[c_idx], "bold"),
            "italic": _majority_font(col_fonts_max[c_idx], "italic"),
            "color_hex": _majority_font(col_fonts_max[c_idx], "color_hex"),
        }
        subtitle_font = {
            "name": _majority_font(col_fonts_min[c_idx], "name"),
            "size_pt": _majority_font(col_fonts_min[c_idx], "size_pt"),
            "bold": _majority_font(col_fonts_min[c_idx], "bold"),
            "italic": _majority_font(col_fonts_min[c_idx], "italic"),
            "color_hex": _majority_font(col_fonts_min[c_idx], "color_hex"),
        }
        result["columns"].append({
            "shading_hex": shading_hex,
            "primary_font": title_font,
            "secondary_font": subtitle_font,
        })

    try:
        result["col_widths_emu"] = [col.width for col in card_table.columns]
    except Exception:
        result["col_widths_emu"] = None

    result["detected"] = bool(result["columns"]) and any(c["shading_hex"] for c in result["columns"])
    return result


def extract_page_setup(doc):
    sec = doc.sections[0]
    return {
        "page_width_in": round(sec.page_width.inches, 3) if sec.page_width else None,
        "page_height_in": round(sec.page_height.inches, 3) if sec.page_height else None,
        "orientation": str(sec.orientation),
        "margins_in": {
            "top": round(sec.top_margin.inches, 3) if sec.top_margin else None,
            "bottom": round(sec.bottom_margin.inches, 3) if sec.bottom_margin else None,
            "left": round(sec.left_margin.inches, 3) if sec.left_margin else None,
            "right": round(sec.right_margin.inches, 3) if sec.right_margin else None,
        },
        "gutter_in": round(sec.gutter.inches, 3) if sec.gutter else 0,
    }


def extract_header_footer(doc):
    sec = doc.sections[0]
    def _text(container):
        return [p.text for p in container.paragraphs if p.text.strip()]
    return {
        "header_text": _text(sec.header),
        "footer_text": _text(sec.footer),
        "different_first_page": sec.different_first_page_header_footer,
    }


def _paragraph_contains_page_field(paragraph):
    p_elm = paragraph._p
    for fldSimple in p_elm.findall(qn('w:fldSimple')):
        instr = fldSimple.get(qn('w:instr')) or ""
        if 'PAGE' in instr.upper():
            return True
    instr_texts = p_elm.findall(f".//{qn('w:instrText')}")
    for it in instr_texts:
        if it.text and 'PAGE' in it.text.upper():
            return True
    return False


def _representative_footer_run_font(paragraph):
    rep_run = next((r for r in paragraph.runs if r.text.strip()), None)
    if rep_run is None and paragraph.runs:
        rep_run = paragraph.runs[0]
    if rep_run is None:
        return {}
    props = {}
    try:
        props["name"] = rep_run.font.name
    except Exception:
        props["name"] = None
    try:
        props["size_pt"] = rep_run.font.size.pt if rep_run.font.size else None
    except Exception:
        props["size_pt"] = None
    try:
        props["bold"] = rep_run.font.bold
    except Exception:
        props["bold"] = None
    try:
        props["italic"] = rep_run.font.italic
    except Exception:
        props["italic"] = None
    try:
        if rep_run.font.color and rep_run.font.color.type is not None and rep_run.font.color.rgb:
            props["color_hex"] = str(rep_run.font.color.rgb)
        else:
            props["color_hex"] = None
    except Exception:
        props["color_hex"] = None
    return props


def find_page_field_display_runs(root_element):
    display_runs = []
    state = "idle"
    current_is_page = False

    for run_el in root_element.iter(qn('w:r')):
        fld_type_in_run = None
        instr_text_in_run = None
        for child in run_el.iter():
            if child is run_el:
                continue
            if child.tag == qn('w:fldChar'):
                fld_type_in_run = child.get(qn('w:fldCharType'))
            elif child.tag == qn('w:instrText'):
                instr_text_in_run = child.text or ""

        if fld_type_in_run == 'begin':
            state = "in_instr"
            current_is_page = False
            continue
        if fld_type_in_run == 'separate':
            state = "in_display" if (state == "in_instr" and current_is_page) else "idle"
            continue
        if fld_type_in_run == 'end':
            state = "idle"
            current_is_page = False
            continue
        if instr_text_in_run is not None:
            if state == "in_instr" and 'PAGE' in instr_text_in_run.upper():
                current_is_page = True
            continue
        if state == "in_display":
            display_runs.append(run_el)

    for fld_simple in root_element.iter(qn('w:fldSimple')):
        instr = fld_simple.get(qn('w:instr')) or ""
        if 'PAGE' in instr.upper():
            for r in fld_simple.iter(qn('w:r')):
                display_runs.append(r)

    seen_ids = set()
    unique_runs = []
    for r in display_runs:
        rid = id(r)
        if rid in seen_ids:
            continue
        seen_ids.add(rid)
        unique_runs.append(r)
    return unique_runs


def _run_font_props_from_element(run_element):
    props = {"name": None, "size_pt": None, "bold": None, "italic": None, "color_hex": None}
    rPr = run_element.find(qn('w:rPr'))
    if rPr is None:
        return props
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        name = rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))
        if name:
            props["name"] = name
    sz = rPr.find(qn('w:sz'))
    if sz is not None and sz.get(qn('w:val')):
        try:
            props["size_pt"] = int(sz.get(qn('w:val'))) / 2.0
        except Exception:
            pass
    b = rPr.find(qn('w:b'))
    if b is not None:
        props["bold"] = (b.get(qn('w:val')) not in ("false", "0", "off"))
    i = rPr.find(qn('w:i'))
    if i is not None:
        props["italic"] = (i.get(qn('w:val')) not in ("false", "0", "off"))
    color = rPr.find(qn('w:color'))
    if color is not None:
        val = color.get(qn('w:val'))
        if val and val.upper() != "AUTO":
            props["color_hex"] = val.upper()
    return props


def _element_ancestor_tags(element, stop_at):
    tags = set()
    current = element.getparent()
    while current is not None and current is not stop_at:
        tags.add(current.tag)
        current = current.getparent()
    return tags


TABLE_TAGS = {qn('w:tbl')}
SHAPE_TAGS = {qn('w:drawing'), qn('w:pict'), qn('w:txbxContent')}


def _cell_paragraph_alignment(cell):
    try:
        p = cell.paragraphs[0]
        align = p.paragraph_format.alignment
        return str(align) if align is not None else None
    except Exception:
        return None


def _cell_representative_font(cell):
    names, sizes, colors = Counter(), Counter(), Counter()
    bolds, italics = [], []
    for p in cell.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            try:
                if r.font.name:
                    names[r.font.name] += 1
            except Exception:
                pass
            try:
                if r.font.size:
                    sizes[round(r.font.size.pt, 1)] += 1
            except Exception:
                pass
            try:
                if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                    colors[str(r.font.color.rgb)] += 1
            except Exception:
                pass
            try:
                if r.font.bold is not None:
                    bolds.append(r.font.bold)
            except Exception:
                pass
            try:
                if r.font.italic is not None:
                    italics.append(r.font.italic)
            except Exception:
                pass
    return {
        "name": names.most_common(1)[0][0] if names else None,
        "size_pt": sizes.most_common(1)[0][0] if sizes else None,
        "color_hex": colors.most_common(1)[0][0] if colors else None,
        "bold": (sum(bolds) >= len(bolds) / 2.0) if bolds else None,
        "italic": (sum(italics) >= len(italics) / 2.0) if italics else None,
    }


def extract_footer_table_design(doc, body_font_name=None):
    sec = doc.sections[0]
    footer = sec.footer
    if not footer.tables:
        return None
    table = footer.tables[0]
    if not table.rows:
        return None

    footer_root = footer._element
    page_field_run_ids = {id(r) for r in find_page_field_display_runs(footer_root)}

    n_rows = len(table.rows)
    n_cols = len(table.columns)
    cells_design = []
    for r_idx, row in enumerate(table.rows):
        seen_tc_ids = set()
        for c_idx, cell in enumerate(row.cells):
            tc_id = id(cell._tc)
            if tc_id in seen_tc_ids:
                continue
            seen_tc_ids.add(tc_id)

            cell_run_ids = {id(r._element) for p in cell.paragraphs for r in p.runs}
            is_page_field_cell = bool(cell_run_ids & page_field_run_ids)

            if is_page_field_cell:
                pf_run_el = next((r for r in cell._tc.iter(qn('w:r')) if id(r) in page_field_run_ids), None)
                font = _run_font_props_from_element(pf_run_el) if pf_run_el is not None else {}
                text = ""
            else:
                font = _cell_representative_font(cell)
                text = cell.text.strip()

            if not font.get("name"):
                font["name"] = body_font_name

            cells_design.append({
                "row": r_idx, "col": c_idx,
                "text": text,
                "alignment": _cell_paragraph_alignment(cell),
                "shading_hex": _direct_cell_shading_hex(cell),
                "font": font,
                "is_page_field": is_page_field_cell,
            })

    try:
        col_widths_emu = [col.width for col in table.columns]
    except Exception:
        col_widths_emu = None

    return {
        "n_rows": n_rows,
        "n_cols": n_cols,
        "cells": cells_design,
        "col_widths_emu": col_widths_emu,
        "border_style": _extract_table_border_style(table),
        "has_page_field_cell": any(c["is_page_field"] for c in cells_design),
    }


def extract_footer_design(doc, theme_fonts=None, body_font_name=None):
    sec = doc.sections[0]
    footer = sec.footer
    result = {"has_page_field": False, "paragraphs": []}

    for p in footer.paragraphs:
        text = p.text.strip()
        has_field = _paragraph_contains_page_field(p)
        if not text and not has_field:
            continue
        if has_field:
            result["has_page_field"] = True

        prefix_parts = []
        suffix_parts = []
        state = "prefix"
        p_elm = p._p
        for child in p_elm.iter():
            tag = child.tag
            if tag == qn('w:fldChar'):
                fld_type = child.get(qn('w:fldCharType'))
                if fld_type == 'begin':
                    state = "in_field"
                elif fld_type == 'end':
                    state = "suffix"
                continue
            if tag == qn('w:fldSimple'):
                state = "suffix"
                continue
            if tag == qn('w:t'):
                if state == "prefix":
                    prefix_parts.append(child.text or "")
                elif state == "suffix":
                    suffix_parts.append(child.text or "")

        font = _representative_footer_run_font(p)
        if not font.get("name"):
            font["name"] = body_font_name

        try:
            alignment = str(p.paragraph_format.alignment) if p.paragraph_format.alignment is not None else None
        except Exception:
            alignment = None

        result["paragraphs"].append({
            "prefix_text": "".join(prefix_parts),
            "has_page_field": has_field,
            "suffix_text": "".join(suffix_parts),
            "alignment": alignment,
            "font": font,
            "full_text_sample": text,
        })

    footer_root = footer._element
    display_run_elements = find_page_field_display_runs(footer_root)
    page_field_runs = []
    page_field_in_table = False
    page_field_in_shape = False
    for run_el in display_run_elements:
        ancestor_tags = _element_ancestor_tags(run_el, footer_root)
        in_table = bool(ancestor_tags & TABLE_TAGS)
        in_shape = bool(ancestor_tags & SHAPE_TAGS)
        if in_table:
            page_field_in_table = True
        if in_shape:
            page_field_in_shape = True
        font = _run_font_props_from_element(run_el)
        if not font.get("name"):
            font["name"] = body_font_name
        page_field_runs.append({"font": font, "in_table": in_table, "in_shape": in_shape})

    if display_run_elements:
        result["has_page_field"] = True
    result["page_field_runs"] = page_field_runs
    result["page_field_in_table"] = page_field_in_table
    result["page_field_in_shape"] = page_field_in_shape

    result["table_design"] = extract_footer_table_design(doc, body_font_name=body_font_name)

    return result


def _cell_shading_hex(cell):
    return _direct_cell_shading_hex(cell)


def _classify_reference_table_context(table):
    if not table.rows:
        return None
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
    return None


def _vote_cell_group(cells, table):
    shading_votes = Counter()
    color_votes = Counter()
    bold_votes = []
    examined = 0
    for cell in cells:
        if not cell.text.strip():
            continue
        examined += 1
        shade = _effective_cell_shading(cell, table)
        if shade:
            shading_votes[shade] += 1
        props = _cell_effective_run_props(cell, table)
        if props.get("color_hex"):
            color_votes[props["color_hex"]] += 1
        if props.get("bold") is not None:
            bold_votes.append(props["bold"])
    return shading_votes, color_votes, bold_votes, examined


def extract_table_contexts(doc):
    """v1.25: now traverses NESTED tables too (via iter_tables_recursive),
    so a reference document's own per-candidate sub-tables (e.g. a
    "Team Leader / Sanitary Engineer / ..." staffing table nested
    inside a larger project cell) contribute to - and correctly
    receive - the same header-shading pattern as top-level data tables."""
    banner_shading = Counter()
    banner_color = Counter()
    banner_bold = []
    banner_table_count = 0
    banner_cell_count = 0

    data_shading = Counter()
    data_color = Counter()
    data_bold = []
    data_table_count = 0
    data_cell_count = 0

    for _path, tbl in iter_tables_recursive(doc):
        ctx = _classify_reference_table_context(tbl)
        if ctx is None:
            continue
        if ctx == "banner":
            header_row = tbl.rows[0]
            banner_table_count += 1
            s, c, b, n = _vote_cell_group(header_row.cells, tbl)
            banner_shading.update(s); banner_color.update(c); banner_bold.extend(b)
            banner_cell_count += n
        elif ctx == "data_table":
            hdr_row_count = count_header_rows(tbl)
            cells = header_row_cells(tbl, hdr_row_count)
            if cells:
                data_table_count += 1
                s, c, b, n = _vote_cell_group(cells, tbl)
                data_shading.update(s); data_color.update(c); data_bold.extend(b)
                data_cell_count += n

    def _build_profile(shading_votes, color_votes, bold_votes, table_count, cell_count, threshold=0.5):
        if cell_count == 0:
            return {"apply": False, "reason": "no matching tables/cells found in reference document "
                                               "(genuinely no data to extract for this element)",
                     "sample_size": table_count}
        top_shade, top_shade_count = (shading_votes.most_common(1)[0] if shading_votes else (None, 0))
        confidence = (top_shade_count / cell_count) if cell_count else 0.0
        top_color, _ = (color_votes.most_common(1)[0] if color_votes else (None, 0))
        majority_bold = (sum(bold_votes) >= len(bold_votes) / 2.0) if bold_votes else None
        apply_shading = bool(top_shade) and confidence >= threshold
        return {
            "apply": apply_shading,
            "sample_size": table_count,
            "total_header_cells_examined": cell_count,
            "shading_hex": top_shade if apply_shading else None,
            "shading_confidence": round(confidence, 2),
            "text_color_hex": top_color,
            "bold": majority_bold,
        }

    banner_profile = _build_profile(banner_shading, banner_color, banner_bold, banner_table_count, banner_cell_count)
    data_profile = _build_profile(data_shading, data_color, data_bold, data_table_count, data_cell_count)

    return {
        "banner": banner_profile,
        "data_table": data_profile,
    }


def extract_table_typography(doc, theme_fonts=None, body_font_props=None):
    """v1.25: now traverses nested tables too - see extract_table_
    contexts() docstring for the full rationale."""
    theme_fonts = theme_fonts or {}
    body_font_props = body_font_props or {}
    fallback_font_name = body_font_props.get("name") or theme_fonts.get("minor_latin")
    fallback_size_pt = body_font_props.get("size_pt")

    def _header_profile():
        names, sizes = Counter(), Counter()
        sample_size = 0
        bare_cells = 0
        for _path, tbl in iter_tables_recursive(doc):
            ctx = _classify_reference_table_context(tbl)
            if ctx != "data_table":
                continue
            hdr_row_count = count_header_rows(tbl)
            for cell in header_row_cells(tbl, hdr_row_count):
                if not cell.text.strip():
                    continue
                props = _cell_effective_run_props(cell, tbl)
                if props.get("name"):
                    names[props["name"]] += 1
                if props.get("size_pt"):
                    sizes[round(props["size_pt"], 1)] += 1
                else:
                    bare_cells += 1
                sample_size += 1
        if sample_size == 0:
            return {"apply": False, "reason": "no header rows/cells found in reference"}
        font_name = names.most_common(1)[0][0] if names else fallback_font_name
        font_name_source = "usage_or_table_style_majority" if names else (
            "body_or_theme_fallback" if fallback_font_name else None)
        size_pt = sizes.most_common(1)[0][0] if sizes else fallback_size_pt
        size_source = "usage_or_table_style_majority" if sizes else (
            "body_fallback" if fallback_size_pt else None)
        return {
            "apply": bool(font_name or size_pt),
            "sample_size": sample_size,
            "font_name": font_name, "font_name_source": font_name_source,
            "size_pt": size_pt, "size_source": size_source,
            "color_hex": None,
            "bare_cells_with_no_direct_size": bare_cells,
        }

    def _body_profile():
        names, sizes, colors = Counter(), Counter(), Counter()
        sample_size = 0
        for _path, tbl in iter_tables_recursive(doc):
            ctx = _classify_reference_table_context(tbl)
            if ctx != "data_table":
                continue
            hdr_row_count = count_header_rows(tbl)
            for row in tbl.rows[hdr_row_count:]:
                for cell in row.cells:
                    names_c, sizes_c, colors_c, _ = _cell_run_props(cell)
                    names.update(names_c)
                    sizes.update(sizes_c)
                    colors.update(colors_c)
                    sample_size += 1
        if sample_size == 0:
            return {"apply": False, "reason": "no body rows found in reference"}
        font_name = names.most_common(1)[0][0] if names else fallback_font_name
        font_name_source = "usage_majority" if names else ("body_or_theme_fallback" if fallback_font_name else None)
        size_pt = sizes.most_common(1)[0][0] if sizes else fallback_size_pt
        size_source = "usage_majority" if sizes else ("body_fallback" if fallback_size_pt else None)
        color_hex = colors.most_common(1)[0][0] if colors and colors.most_common(1)[0][0] == "000000" else None
        return {
            "apply": bool(font_name or size_pt),
            "sample_size": sample_size,
            "font_name": font_name, "font_name_source": font_name_source,
            "size_pt": size_pt, "size_source": size_source,
            "color_hex": color_hex,
        }

    return {
        "header": _header_profile(),
        "body": _body_profile(),
    }


def _row_effective_shading_or_none(row, table, body_row_parity):
    cnf_band = read_row_cnf_band(row)
    if cnf_band is None:
        cells_for_cnf_check = dedupe_row_cells(row)
        for cell in cells_for_cnf_check:
            cnf_band = read_cell_cnf_band(cell)
            if cnf_band is not None:
                break

    effective_parity = cnf_band if cnf_band is not None else body_row_parity
    band_cond_type = _COND_FORMAT_TYPE_BAND1 if effective_parity == 0 else _COND_FORMAT_TYPE_BAND2

    votes = Counter()
    for cell in dedupe_row_cells(row):
        if not cell.text.strip():
            continue
        shade = _effective_cell_shading_for_band(cell, table, band_cond_type)
        votes[shade if shade else NO_FILL_SENTINEL] += 1
    if not votes:
        return None, effective_parity
    return votes.most_common(1)[0][0], effective_parity


def _vote_banding_for_table_group(tables):
    parity_votes = [Counter(), Counter()]
    parity_sample_counts = [0, 0]
    cnf_driven_rows = 0
    tables_with_body_rows = 0

    for tbl in tables:
        hdr_row_count = count_header_rows(tbl)
        body_rows = tbl.rows[hdr_row_count:]
        if not body_rows:
            continue
        tables_with_body_rows += 1
        row_band_size = _table_row_band_size(tbl)
        for body_idx, row in enumerate(body_rows):
            position_group = (body_idx // row_band_size) % 2
            value, effective_parity = _row_effective_shading_or_none(row, tbl, position_group)
            if effective_parity != position_group:
                cnf_driven_rows += 1
            if value is None:
                continue
            parity_votes[effective_parity][value] += 1
            parity_sample_counts[effective_parity] += 1

    if parity_sample_counts[0] < MIN_BANDING_SAMPLES_PER_PARITY or \
       parity_sample_counts[1] < MIN_BANDING_SAMPLES_PER_PARITY:
        return {
            "apply": False,
            "reason": (f"insufficient body-row samples to detect banding for this table shape "
                       f"(parity 0: {parity_sample_counts[0]}, parity 1: {parity_sample_counts[1]}, "
                       f"need >= {MIN_BANDING_SAMPLES_PER_PARITY} each)"),
            "tables_in_group": tables_with_body_rows,
            "cnf_style_corrections": cnf_driven_rows,
        }

    top0, count0 = parity_votes[0].most_common(1)[0]
    top1, count1 = parity_votes[1].most_common(1)[0]
    confidence0 = count0 / parity_sample_counts[0]
    confidence1 = count1 / parity_sample_counts[1]

    if confidence0 < MIN_BANDING_CONFIDENCE or confidence1 < MIN_BANDING_CONFIDENCE:
        return {
            "apply": False,
            "reason": (f"body-row shading not consistent enough for this table shape to confirm a "
                       f"banding pattern (confidence0={confidence0:.2f}, confidence1={confidence1:.2f}, "
                       f"need >= {MIN_BANDING_CONFIDENCE})"),
            "tables_in_group": tables_with_body_rows,
            "cnf_style_corrections": cnf_driven_rows,
        }

    if top0 == top1:
        return {
            "apply": False,
            "reason": "body rows are uniformly shaded for this table shape (no alternation detected)",
            "tables_in_group": tables_with_body_rows,
            "uniform_value": (None if top0 == NO_FILL_SENTINEL else top0),
            "cnf_style_corrections": cnf_driven_rows,
        }

    band_colors = [
        (None if top0 == NO_FILL_SENTINEL else top0),
        (None if top1 == NO_FILL_SENTINEL else top1),
    ]
    return {
        "apply": True,
        "band_colors": band_colors,
        "confidence": [round(confidence0, 2), round(confidence1, 2)],
        "sample_size": [parity_sample_counts[0], parity_sample_counts[1]],
        "tables_in_group": tables_with_body_rows,
        "cnf_style_corrections": cnf_driven_rows,
        "reason": "confirmed alternating body-row background shading for this table shape",
    }


def extract_body_row_banding(doc):
    """v1.25: now traverses nested tables too via iter_tables_recursive
    (path is unused here since grouping is by header signature, which
    is already location-independent)."""
    groups = {}
    for _path, tbl in iter_tables_recursive(doc):
        ctx = _classify_reference_table_context(tbl)
        if ctx != "data_table":
            continue
        hdr_row_count = count_header_rows(tbl)
        sig = table_header_signature(tbl, hdr_row_count)
        if sig is None:
            continue
        groups.setdefault(sig, []).append(tbl)

    by_signature = {}
    for sig, tables in groups.items():
        by_signature[sig] = _vote_banding_for_table_group(tables)

    all_data_tables = [tbl for _path, tbl in iter_tables_recursive(doc)
                        if _classify_reference_table_context(tbl) == "data_table"]
    default_profile = _vote_banding_for_table_group(all_data_tables)
    default_profile["reason"] = (default_profile.get("reason", "") +
                                 " [whole-document fallback profile, used only when a target table's "
                                 "header does not match any known reference table shape]")

    return {
        "by_signature": by_signature,
        "default": default_profile,
    }


def extract_numbering_hint(doc):
    heading_pattern_seen = None
    caption_pattern_seen = None
    for p in doc.paragraphs:
        t = p.text.strip()
        if not heading_pattern_seen and re.match(r"^\d+(\.\d+){0,2}\s+\S", t):
            heading_pattern_seen = re.match(r"^(\d+(\.\d+){0,2})\s", t).group(1)
        if not caption_pattern_seen and re.match(r"^(Table|Figure)\s+\d", t, re.I):
            caption_pattern_seen = t.split()[0] + " " + t.split()[1].rstrip(":.")
    return {
        "heading_numbering_example": heading_pattern_seen,
        "caption_pattern_example": caption_pattern_seen,
    }


def build_policy(reference_path):
    doc = Document(reference_path)
    theme_fonts = extract_theme_fonts(reference_path)
    typography = extract_typography(doc, reference_path)
    # v1.25: extract_special_headings now needs theme_fonts + the
    # already-resolved Body font name as fallback inputs - see that
    # function's docstring for the full rationale (fix for issue #1,
    # "Aptos font appearing").
    special_headings = extract_special_headings(
        doc, theme_fonts=theme_fonts,
        fallback_font_name=typography.get("Body", {}).get("font", {}).get("name"),
        fallback_size_pt=typography.get("Body", {}).get("font", {}).get("size_pt"))
    policy = {
        "policy_version": "1.25",
        "source_reference_document": reference_path,
        "theme_fonts": theme_fonts,
        "typography": typography,
        "special_headings": special_headings,
        "toc_card_design": extract_toc_card_design(doc, special_headings),
        "page_setup": extract_page_setup(doc),
        "header_footer": extract_header_footer(doc),
        "footer_design": extract_footer_design(
            doc, theme_fonts=theme_fonts,
            body_font_name=typography.get("Body", {}).get("font", {}).get("name")),
        "table_contexts": extract_table_contexts(doc),
        "table_typography": extract_table_typography(
            doc, theme_fonts=theme_fonts, body_font_props=typography.get("Body", {}).get("font", {})),
        "table_row_banding": extract_body_row_banding(doc),
        "numbering_hint": extract_numbering_hint(doc),
    }
    return policy


def main():
    ap = argparse.ArgumentParser(description="Phase 1 - Extract formatting policy from a reference template.")
    ap.add_argument("--reference", required=True)
    ap.add_argument("--out", default="formatting_policy_v1.json")
    args = ap.parse_args()

    policy = build_policy(args.reference)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(policy, f, indent=2, ensure_ascii=False)
    print(f"[policy_extractor] Wrote {args.out}")
    typ = policy["typography"]
    for role in ("Title", "H1", "H2", "H3", "H4", "Body", "Caption", "Quote", "ListBullet"):
        f = typ.get(role, {}).get("font", {})
        print(f"  {role}: font={f.get('name')} (source: {f.get('name_source', 'style_default')}), "
              f"size={f.get('size_pt')}pt, color=#{f.get('color_hex')}, "
              f"usage_samples={typ.get(role, {}).get('usage_sample_size', 0)}")


if __name__ == "__main__":
    main()




