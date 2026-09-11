"""
Phase 3 — Formatting Engine (v1.22 — per-table-shape banding application)

=== v1.22 FIX: applying the correct, per-table-shape banding pattern
    (fixing the "no colors picked up at all" regression) ===
See policy_extractor.py's module docstring for the full, directly-
confirmed root cause: a single whole-document banding pattern is the
wrong model when a document contains multiple, structurally different
table "shapes" with genuinely different color schemes - one such
table's colors can dilute/contradict another's in a combined vote,
suppressing detection of BOTH.

`apply_table_body_row_banding` now computes each TARGET table's own
header "shape signature" (docx_fast.table_header_signature - the SAME
function used on the reference side, guaranteeing identical, stable
matching logic) and looks up the corresponding profile from the
reference's NEW `table_row_banding["by_signature"]` dict FIRST. Only if
no signature match is found does it fall back to `table_row_banding
["default"]` (the old whole-document vote, preserved for exactly this
fallback case). Whichever profile is selected (if `apply: True`) is
then applied using the EXACT SAME per-parity direct-shading logic as
before (v1.21's defensive cnfStyle-stripping included, completely
unchanged) - so a target table whose shape exactly matches a confident
reference family (e.g. every "No. | Activity | Description | Timing"
table) now correctly receives that family's OWN specific banding
pattern, while a target table matching a reference table that has its
own distinct, non-alternating per-row coloring (and therefore reports
`apply: False` for its OWN signature) is correctly left untouched,
without that decision affecting any OTHER table shape in the document.

This is backward-compatible: a reference document with only ONE table
shape throughout produces a `by_signature` dict with exactly one entry,
whose profile is mathematically identical to the OLD whole-document
vote - so every previously-verified simple/synthetic test case behaves
identically.

All other logic (header shading, table typography, footer table-layout
replication, TOC card builder) is completely UNCHANGED.
"""
import argparse
import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx_fast import (
    build_styleid_to_name, get_paragraph_style_name_fast,
    count_header_rows, header_row_cells, dedupe_row_cells,
    clear_cnf_style, table_header_signature, iter_tables_recursive,
)
from ooxml_fonts import (
    set_run_font_safe, set_style_font_safe,
    sync_complex_script_size_and_emphasis, set_literal_font_no_theme_override,
)
from policy_extractor import find_page_field_display_runs

ROLE_TO_STYLE_NAME = {
    "Title": "Title",
    "H1": "Heading 1",
    "H2": "Heading 2",
    "H3": "Heading 3",
    "H4": "Heading 4",
    "Body": "Normal",
    "Caption": "Caption",
    "Quote": "Quote",
    "ListBullet": "List Bullet",
}

HEADING_ROLES = {"Title", "H1", "H2", "H3", "H4", "Caption"}
BODY_LIKE_ROLES = {"Body", "Caption", "Quote", "ListBullet"}
SPECIAL_HEADING_PREFIX = "SpecialHeading::"

ALIGNMENT_MAP = {
    "LEFT (0)": WD_ALIGN_PARAGRAPH.LEFT,
    "CENTER (1)": WD_ALIGN_PARAGRAPH.CENTER,
    "RIGHT (2)": WD_ALIGN_PARAGRAPH.RIGHT,
    "JUSTIFY (3)": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

MAX_SAFE_BODY_LUMINANCE = 140


def _hex_luminance(hex_color):
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return 0.299 * r + 0.587 * g + 0.114 * b
    except Exception:
        return None


def _is_safe_text_color(hex_color, role):
    if role not in BODY_LIKE_ROLES:
        return True
    if not hex_color:
        return True
    luminance = _hex_luminance(hex_color)
    if luminance is None:
        return True
    return luminance <= MAX_SAFE_BODY_LUMINANCE


def _apply_font_to_style(style, font_policy, role=None):
    if not font_policy:
        return
    f = style.font
    if font_policy.get("name"):
        set_style_font_safe(style, font_policy["name"])
    if font_policy.get("size_pt"):
        f.size = Pt(font_policy["size_pt"])
    if font_policy.get("bold") is not None:
        f.bold = font_policy["bold"]
    if font_policy.get("italic") is not None:
        f.italic = font_policy["italic"]
    if font_policy.get("underline") is not None:
        f.underline = font_policy["underline"]
    color_hex = font_policy.get("color_hex")
    if color_hex and _is_safe_text_color(color_hex, role):
        try:
            f.color.rgb = RGBColor.from_string(color_hex)
        except Exception:
            pass


def _apply_paragraph_format_to_style(style, para_policy):
    if not para_policy:
        return
    pf = style.paragraph_format
    if para_policy.get("space_before_pt") is not None:
        pf.space_before = Pt(para_policy["space_before_pt"])
    if para_policy.get("space_after_pt") is not None:
        pf.space_after = Pt(para_policy["space_after_pt"])
    if para_policy.get("line_spacing") is not None:
        pf.line_spacing = para_policy["line_spacing"]
    if para_policy.get("first_line_indent_in") is not None:
        pf.first_line_indent = Inches(para_policy["first_line_indent_in"])
    if para_policy.get("left_indent_in") is not None:
        pf.left_indent = Inches(para_policy["left_indent_in"])
    align = para_policy.get("alignment")
    if align in ALIGNMENT_MAP:
        pf.alignment = ALIGNMENT_MAP[align]


def redefine_styles(doc, policy):
    typography = policy.get("typography", {})
    by_name = {}
    for s in doc.styles:
        by_name.setdefault(s.name, []).append(s)

    applied = []
    for role, style_name in ROLE_TO_STYLE_NAME.items():
        role_policy = typography.get(role, {})
        matches = by_name.get(style_name, [])
        if not matches:
            continue
        for style in matches:
            _apply_font_to_style(style, role_policy.get("font"), role=role)
            _apply_paragraph_format_to_style(style, role_policy.get("paragraph"))
        applied.append({"role": role, "style_name": style_name,
                         "duplicate_style_objects_updated": len(matches),
                         "font": role_policy.get("font"),
                         "paragraph": role_policy.get("paragraph")})
    return applied


def apply_page_setup(doc, policy):
    page = policy.get("page_setup", {})
    margins = page.get("margins_in", {})
    for sec in doc.sections:
        if page.get("page_width_in"):
            sec.page_width = Inches(page["page_width_in"])
        if page.get("page_height_in"):
            sec.page_height = Inches(page["page_height_in"])
        if margins.get("top") is not None:
            sec.top_margin = Inches(margins["top"])
        if margins.get("bottom") is not None:
            sec.bottom_margin = Inches(margins["bottom"])
        if margins.get("left") is not None:
            sec.left_margin = Inches(margins["left"])
        if margins.get("right") is not None:
            sec.right_margin = Inches(margins["right"])


def _apply_special_heading_formatting(paragraph, entry):
    font_policy = entry.get("font") or {}
    para_policy = entry.get("paragraph") or {}
    for run in paragraph.runs:
        if font_policy.get("name"):
            set_run_font_safe(run, font_policy["name"])
        if font_policy.get("size_pt"):
            run.font.size = Pt(font_policy["size_pt"])
        if font_policy.get("bold") is not None:
            run.font.bold = font_policy["bold"]
        if font_policy.get("italic") is not None:
            run.font.italic = font_policy["italic"]
        if font_policy.get("underline") is not None:
            run.font.underline = font_policy["underline"]
        if font_policy.get("color_hex"):
            try:
                run.font.color.rgb = RGBColor.from_string(font_policy["color_hex"])
            except Exception:
                pass
        sync_complex_script_size_and_emphasis(run)

    pf = paragraph.paragraph_format
    align = para_policy.get("alignment")
    if align in ALIGNMENT_MAP:
        pf.alignment = ALIGNMENT_MAP[align]
    if para_policy.get("space_before_pt") is not None:
        pf.space_before = Pt(para_policy["space_before_pt"])
    if para_policy.get("space_after_pt") is not None:
        pf.space_after = Pt(para_policy["space_after_pt"])


def _normalize_runs_direct_formatting(paragraph, role, typography, method=None):
    font_policy = typography.get(role, {}).get("font") or {}

    if method == "heuristic:style_override_long_text":
        for run in paragraph.runs:
            run.font.bold = False
            run.font.italic = False

    if method == "heuristic:admin_boilerplate_denylist":
        for run in paragraph.runs:
            rPr = run._element.find(qn('w:rPr'))
            if rPr is not None:
                color_el = rPr.find(qn('w:color'))
                if color_el is not None:
                    rPr.remove(color_el)

    if not font_policy:
        return
    force_emphasis = role in HEADING_ROLES
    for run in paragraph.runs:
        if font_policy.get("name"):
            set_run_font_safe(run, font_policy["name"])
        if font_policy.get("size_pt"):
            run.font.size = Pt(font_policy["size_pt"])
        if force_emphasis:
            if font_policy.get("bold") is not None:
                run.font.bold = font_policy["bold"]
            if font_policy.get("italic") is not None:
                run.font.italic = font_policy["italic"]
            if font_policy.get("underline") is not None:
                run.font.underline = font_policy["underline"]
            color_hex = font_policy.get("color_hex")
            if color_hex:
                try:
                    run.font.color.rgb = RGBColor.from_string(color_hex)
                except Exception:
                    pass
        else:
            if font_policy.get("color_hex") in ("000000",):
                try:
                    run.font.color.rgb = RGBColor.from_string("000000")
                except Exception:
                    pass
        sync_complex_script_size_and_emphasis(run)


def reassign_paragraph_styles(doc, classification, typography, special_headings=None):
    special_headings = special_headings or {}
    style_cache = {}
    for style_name in set(ROLE_TO_STYLE_NAME.values()):
        try:
            style_cache[style_name] = doc.styles[style_name]
        except KeyError:
            style_cache[style_name] = None

    styleid_to_name = build_styleid_to_name(doc)

    change_log = []
    paragraphs = doc.paragraphs
    for result in classification["paragraphs"]:
        idx = result["index"]
        if idx >= len(paragraphs):
            continue
        role = result["role"]
        p = paragraphs[idx]
        old_style = get_paragraph_style_name_fast(p, styleid_to_name)

        if role.startswith(SPECIAL_HEADING_PREFIX):
            key = role[len(SPECIAL_HEADING_PREFIX):]
            entry = special_headings.get(key)
            if entry:
                _apply_special_heading_formatting(p, entry)
            change_log.append({
                "index": idx, "role": role,
                "old_style": old_style, "new_style": old_style,
                "confidence": result["confidence"], "method": result["method"],
                "text_preview": result["text_preview"],
            })
            continue

        style_name = ROLE_TO_STYLE_NAME.get(role, "Normal")
        target_style = style_cache.get(style_name)
        if old_style != style_name and target_style is not None:
            p.style = target_style
        _normalize_runs_direct_formatting(p, role, typography, method=result["method"])
        change_log.append({
            "index": idx, "role": role,
            "old_style": old_style, "new_style": style_name,
            "confidence": result["confidence"], "method": result["method"],
            "text_preview": result["text_preview"],
        })
    return change_log


_TCPR_ELEMENTS_AFTER_SHD = [
    'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText',
    'w:vAlign', 'w:hideMark', 'w:cnfStyle',
]


def _set_cell_shading(cell, hex_color):
    if not hex_color:
        return
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        insert_before = None
        for tag in _TCPR_ELEMENTS_AFTER_SHD:
            found = tcPr.find(qn(tag))
            if found is not None:
                insert_before = found
                break
        if insert_before is not None:
            insert_before.addprevious(shd)
        else:
            tcPr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)


def _clear_cell_shading(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        insert_before = None
        for tag in _TCPR_ELEMENTS_AFTER_SHD:
            found = tcPr.find(qn(tag))
            if found is not None:
                insert_before = found
                break
        if insert_before is not None:
            insert_before.addprevious(shd)
        else:
            tcPr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'auto')


def _is_light_color(hex_color):
    lum = _hex_luminance(hex_color)
    return lum is None or lum > 150


def _force_set_run_size_and_font(run, font_name=None, size_pt=None):
    if font_name:
        set_run_font_safe(run, font_name)
    if size_pt:
        run.font.size = Pt(size_pt)
    sync_complex_script_size_and_emphasis(run)


def _apply_typography_profile_to_cell(cell, profile):
    if not profile or not profile.get("apply"):
        return 0
    font_name = profile.get("font_name")
    size_pt = profile.get("size_pt")
    color_hex = profile.get("color_hex")
    touched = 0
    for p in cell.paragraphs:
        for run in p.runs:
            if not run.text.strip():
                continue
            _force_set_run_size_and_font(run, font_name=font_name, size_pt=size_pt)
            if color_hex == "000000":
                try:
                    run.font.color.rgb = RGBColor.from_string(color_hex)
                except Exception:
                    pass
            touched += 1
    return touched


def _enforce_uniform_header_sizes(doc, header_profile, classification):
    if not header_profile or not header_profile.get("apply"):
        return {"cells_checked": 0, "runs_corrected": 0}
    size_pt = header_profile.get("size_pt")
    font_name = header_profile.get("font_name")
    if not size_pt:
        return {"cells_checked": 0, "runs_corrected": 0}

    table_context_by_index = classification.get("table_contexts", {})
    table_header_row_counts = classification.get("table_header_row_counts", {})

    cells_checked = 0
    runs_corrected = 0
    target_size_emu = Pt(size_pt)

    for path, table in iter_tables_recursive(doc):
        ctx = table_context_by_index.get(path, "other")
        if ctx != "data_table" or not table.rows:
            continue
        hdr_row_count = table_header_row_counts.get(path) or count_header_rows(table)
        for cell in header_row_cells(table, hdr_row_count):
            cells_checked += 1
            for p in cell.paragraphs:
                for run in p.runs:
                    if not run.text.strip():
                        continue
                    current_size = None
                    try:
                        current_size = run.font.size
                    except Exception:
                        current_size = None
                    if current_size != target_size_emu:
                        _force_set_run_size_and_font(run, font_name=font_name, size_pt=size_pt)
                        runs_corrected += 1

    return {"cells_checked": cells_checked, "runs_corrected": runs_corrected}


def apply_table_typography(doc, table_typography, classification):
    """v1.25: iterates ALL tables (top-level + nested at any depth) via
    iter_tables_recursive, keyed by the SAME structural path strings
    produced during classification - see docx_fast.iter_tables_
    recursive() and structure_classifier.classify_document() for the
    full rationale (fix for issue #3, inconsistent nested-table header
    formatting)."""
    table_typography = table_typography or {}
    header_profile = table_typography.get("header")
    body_profile = table_typography.get("body")
    table_context_by_index = classification.get("table_contexts", {})
    table_header_row_counts = classification.get("table_header_row_counts", {})
    header_cells_touched = 0
    body_cells_touched = 0

    for path, table in iter_tables_recursive(doc):
        ctx = table_context_by_index.get(path, "other")
        if ctx != "data_table":
            continue
        if not table.rows:
            continue
        hdr_row_count = table_header_row_counts.get(path) or count_header_rows(table)
        for cell in header_row_cells(table, hdr_row_count):
            header_cells_touched += _apply_typography_profile_to_cell(cell, header_profile)
        for row in table.rows[hdr_row_count:]:
            for cell in row.cells:
                body_cells_touched += _apply_typography_profile_to_cell(cell, body_profile)

    safety_net_result = _enforce_uniform_header_sizes(doc, header_profile, classification)

    return {
        "header_cells_touched": header_cells_touched,
        "body_cells_touched": body_cells_touched,
        "safety_net_cells_checked": safety_net_result["cells_checked"],
        "safety_net_runs_corrected": safety_net_result["runs_corrected"],
    }


def apply_table_context_styles(doc, classification, table_contexts_policy):
    """v1.25: nested-table-aware (see apply_table_typography docstring)."""
    table_context_by_index = classification.get("table_contexts", {})
    table_header_row_counts = classification.get("table_header_row_counts", {})
    tables_touched = {"banner": 0, "data_table": 0, "other_untouched": 0}

    data_profile = table_contexts_policy.get("data_table")
    banner_profile = table_contexts_policy.get("banner")

    def _apply_profile_to_cells(cells, profile):
        if not profile or not profile.get("apply"):
            return False
        shading_hex = profile.get("shading_hex")
        if not shading_hex:
            return False
        text_color_hex = profile.get("text_color_hex")
        bold = profile.get("bold")
        text_color = text_color_hex or ("000000" if _is_light_color(shading_hex) else "FFFFFF")
        seen_tc_ids = set()
        for cell in cells:
            if id(cell._tc) in seen_tc_ids:
                continue
            seen_tc_ids.add(id(cell._tc))
            _set_cell_shading(cell, shading_hex)
            clear_cnf_style(cell, is_row=False)
            for p in cell.paragraphs:
                for run in p.runs:
                    if bold is not None:
                        run.font.bold = bold
                    try:
                        run.font.color.rgb = RGBColor.from_string(text_color)
                    except Exception:
                        pass
                    sync_complex_script_size_and_emphasis(run)
        return True

    for path, table in iter_tables_recursive(doc):
        ctx = table_context_by_index.get(path, "other")

        if ctx == "banner":
            if _apply_profile_to_cells(table.rows[0].cells if table.rows else [], banner_profile):
                tables_touched["banner"] += 1
            else:
                tables_touched["other_untouched"] += 1
            continue

        if ctx != "data_table" or not table.rows:
            tables_touched["other_untouched"] += 1
            continue

        hdr_row_count = table_header_row_counts.get(path) or count_header_rows(table)
        cells = header_row_cells(table, hdr_row_count)

        if _apply_profile_to_cells(cells, data_profile):
            tables_touched["data_table"] += 1
        else:
            tables_touched["other_untouched"] += 1

    return tables_touched


def _apply_band_colors_to_table(table, hdr_row_count, band_colors):
    body_rows = table.rows[hdr_row_count:]
    if not body_rows:
        return 0, 0

    rows_touched = 0
    cells_touched = 0
    for body_idx, row in enumerate(body_rows):
        parity = body_idx % 2
        band_color = band_colors[parity]
        clear_cnf_style(row, is_row=True)
        row_touched_any = False
        for cell in dedupe_row_cells(row):
            if band_color:
                _set_cell_shading(cell, band_color)
            else:
                _clear_cell_shading(cell)
            clear_cnf_style(cell, is_row=False)
            cells_touched += 1
            row_touched_any = True
        if row_touched_any:
            rows_touched += 1
    return rows_touched, cells_touched


def apply_table_body_row_banding(doc, row_banding_policy, classification):
    """v1.25: nested-table-aware (see apply_table_typography docstring).
    Signature-based matching is already location-independent, so nested
    tables benefit automatically once they're included in the traversal."""
    result = {"applied": False, "tables_touched": 0, "rows_touched": 0, "cells_touched": 0,
              "tables_matched_by_signature": 0, "tables_matched_by_default_fallback": 0,
              "tables_with_no_applicable_profile": 0}

    if not row_banding_policy:
        result["reason"] = "no row-banding policy available"
        return result

    by_signature = row_banding_policy.get("by_signature") or {}
    default_profile = row_banding_policy.get("default") or {}

    table_context_by_index = classification.get("table_contexts", {})
    table_header_row_counts = classification.get("table_header_row_counts", {})

    tables_touched = 0
    rows_touched_total = 0
    cells_touched_total = 0
    matched_by_signature = 0
    matched_by_default = 0
    no_applicable_profile = 0

    for path, table in iter_tables_recursive(doc):
        ctx = table_context_by_index.get(path, "other")
        if ctx != "data_table" or not table.rows:
            continue
        hdr_row_count = table_header_row_counts.get(path) or count_header_rows(table)

        sig = table_header_signature(table, hdr_row_count)
        profile = by_signature.get(sig) if sig is not None else None
        matched_via = "signature"
        if profile is None or not profile.get("apply"):
            if profile is not None and not profile.get("apply"):
                no_applicable_profile += 1
                continue
            profile = default_profile
            matched_via = "default_fallback"

        if not profile or not profile.get("apply"):
            no_applicable_profile += 1
            continue

        band_colors = profile.get("band_colors")
        if not band_colors or len(band_colors) != 2:
            no_applicable_profile += 1
            continue

        rows_touched, cells_touched = _apply_band_colors_to_table(table, hdr_row_count, band_colors)
        if rows_touched > 0:
            tables_touched += 1
            rows_touched_total += rows_touched
            cells_touched_total += cells_touched
            if matched_via == "signature":
                matched_by_signature += 1
            else:
                matched_by_default += 1

    result["applied"] = tables_touched > 0
    result["tables_touched"] = tables_touched
    result["rows_touched"] = rows_touched_total
    result["cells_touched"] = cells_touched_total
    result["tables_matched_by_signature"] = matched_by_signature
    result["tables_matched_by_default_fallback"] = matched_by_default
    result["tables_with_no_applicable_profile"] = no_applicable_profile
    result["reason"] = (f"applied per-table-shape banding to {tables_touched} data table(s) "
                        f"({matched_by_signature} matched a specific reference table shape, "
                        f"{matched_by_default} used the whole-document fallback pattern); "
                        f"{no_applicable_profile} table(s) had no applicable banding pattern")
    return result


def apply_special_table_cell_formatting(doc, classification, typography):
    """v1.25: nested-table-aware (see apply_table_typography docstring)."""
    style_cache = {}
    for style_name in set(ROLE_TO_STYLE_NAME.values()):
        try:
            style_cache[style_name] = doc.styles[style_name]
        except KeyError:
            style_cache[style_name] = None

    change_log = []
    by_key = {}
    for r in classification.get("table_cell_paragraphs", []):
        by_key[(r["table_index"], r["row"], r["col"], r["para"])] = r

    for path, table in iter_tables_recursive(doc):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                for p_idx, p in enumerate(cell.paragraphs):
                    key = (path, r_idx, c_idx, p_idx)
                    result = by_key.get(key)
                    if not result:
                        continue
                    role = result["role"]
                    if role in ("TableHeader", "TableBody"):
                        continue
                    style_name = ROLE_TO_STYLE_NAME.get(role)
                    target_style = style_cache.get(style_name) if style_name else None
                    if target_style is not None:
                        p.style = target_style
                    _normalize_runs_direct_formatting(p, role, typography)
                    change_log.append({
                        "table_index": path, "row": r_idx, "col": c_idx, "para": p_idx,
                        "role": role, "method": result["method"],
                        "text_preview": result["text_preview"],
                    })
    return change_log


def _fix_zoom_element(doc):
    settings = doc.settings.element
    zoom = settings.find(qn('w:zoom'))
    if zoom is not None and zoom.get(qn('w:percent')) is None:
        zoom.set(qn('w:percent'), '100')


_TAGS_AFTER_UPDATE_FIELDS = [
    'w:hdrShapeDefaults', 'w:footnotePr', 'w:endnotePr', 'w:compat',
    'w:docVars', 'w:rsids', 'm:mathPr', 'w:attachedSchema',
    'w:themeFontLang', 'w:clrSchemeMapping', 'w:doNotIncludeSubdocsInStats',
    'w:doNotAutoCompressPictures', 'w:forceUpgrade', 'w:captions',
    'w:readModeInkLockDown', 'w:smartTagType', 'w:shapeDefaults',
    'w:doNotEmbedSmartTags', 'w:decimalSymbol', 'w:listSeparator',
]


def _force_update_fields_on_open(doc):
    settings = doc.settings.element
    uf = settings.find(qn('w:updateFields'))
    if uf is not None:
        uf.set(qn('w:val'), 'true')
        return
    uf = OxmlElement('w:updateFields')
    uf.set(qn('w:val'), 'true')
    insert_before = None
    for tag in _TAGS_AFTER_UPDATE_FIELDS:
        found = settings.find(qn(tag))
        if found is not None:
            insert_before = found
            break
    if insert_before is not None:
        insert_before.addprevious(uf)
    else:
        settings.append(uf)


def _clear_paragraph_runs(paragraph):
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)


def _insert_page_field_run(paragraph, font_props, cached_display_text="1"):
    def _new_run():
        r = paragraph.add_run()
        if font_props:
            if font_props.get("name"):
                set_run_font_safe(r, font_props["name"])
            if font_props.get("size_pt"):
                r.font.size = Pt(font_props["size_pt"])
            if font_props.get("bold") is not None:
                r.font.bold = font_props["bold"]
            if font_props.get("italic") is not None:
                r.font.italic = font_props["italic"]
            if font_props.get("color_hex"):
                try:
                    r.font.color.rgb = RGBColor.from_string(font_props["color_hex"])
                except Exception:
                    pass
        return r

    r1 = _new_run()
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    r1._element.append(fldChar_begin)

    r2 = _new_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    r2._element.append(instrText)

    r3 = _new_run()
    fldChar_sep = OxmlElement('w:fldChar')
    fldChar_sep.set(qn('w:fldCharType'), 'separate')
    r3._element.append(fldChar_sep)

    r4 = _new_run()
    r4.text = cached_display_text

    r5 = _new_run()
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    r5._element.append(fldChar_end)


def _restyle_existing_page_fields_in_place(doc, footer_design):
    from collections import Counter

    page_field_runs_ref = footer_design.get("page_field_runs", [])
    if not page_field_runs_ref:
        return {"applied": False, "reason": "Reference footer's page-number design is embedded in a "
                                             "table or shape, but no page-field display run could be "
                                             "found in the reference to copy styling from."}

    name_votes = Counter(r["font"].get("name") for r in page_field_runs_ref if r["font"].get("name"))
    size_votes = Counter(r["font"].get("size_pt") for r in page_field_runs_ref if r["font"].get("size_pt"))
    bold_votes = [r["font"].get("bold") for r in page_field_runs_ref if r["font"].get("bold") is not None]
    italic_votes = [r["font"].get("italic") for r in page_field_runs_ref if r["font"].get("italic") is not None]
    color_votes = Counter(r["font"].get("color_hex") for r in page_field_runs_ref if r["font"].get("color_hex"))

    font = {
        "name": name_votes.most_common(1)[0][0] if name_votes else None,
        "size_pt": size_votes.most_common(1)[0][0] if size_votes else None,
        "bold": (sum(bold_votes) >= len(bold_votes) / 2.0) if bold_votes else None,
        "italic": (sum(italic_votes) >= len(italic_votes) / 2.0) if italic_votes else None,
        "color_hex": color_votes.most_common(1)[0][0] if color_votes else None,
    }

    sections_touched = 0
    runs_restyled = 0
    seen_footer_ids = set()
    for footer in _all_footer_instances(doc):
        footer_root = footer._element
        if id(footer_root) in seen_footer_ids:
            continue
        seen_footer_ids.add(id(footer_root))

        display_run_elements = find_page_field_display_runs(footer_root)
        if not display_run_elements:
            continue
        sections_touched += 1
        for run_el in display_run_elements:
            rPr = run_el.get_or_add_rPr()
            if font.get("name"):
                set_literal_font_no_theme_override(rPr, font["name"])
            if font.get("size_pt"):
                sz = rPr.find(qn('w:sz'))
                if sz is None:
                    sz = OxmlElement('w:sz')
                    rPr.append(sz)
                sz.set(qn('w:val'), str(int(round(font["size_pt"] * 2))))
                szCs = rPr.find(qn('w:szCs'))
                if szCs is None:
                    szCs = OxmlElement('w:szCs')
                    rPr.append(szCs)
                szCs.set(qn('w:val'), str(int(round(font["size_pt"] * 2))))
            if font.get("bold") is not None:
                b = rPr.find(qn('w:b'))
                if b is None:
                    b = OxmlElement('w:b')
                    rPr.append(b)
                if not font["bold"]:
                    b.set(qn('w:val'), '0')
                elif b.get(qn('w:val')) is not None:
                    del b.attrib[qn('w:val')]
            if font.get("italic") is not None:
                it = rPr.find(qn('w:i'))
                if it is None:
                    it = OxmlElement('w:i')
                    rPr.append(it)
                if not font["italic"]:
                    it.set(qn('w:val'), '0')
                elif it.get(qn('w:val')) is not None:
                    del it.attrib[qn('w:val')]
            if font.get("color_hex"):
                color_el = rPr.find(qn('w:color'))
                if color_el is None:
                    color_el = OxmlElement('w:color')
                    rPr.append(color_el)
                color_el.set(qn('w:val'), font["color_hex"])
            runs_restyled += 1

    if runs_restyled == 0:
        return {"applied": False, "reason": "Reference footer's page-number design is embedded in a "
                                             "table or shape, but no corresponding page-field run could "
                                             "be found in the target document's own footer to restyle."}

    return {"applied": True, "sections_touched": sections_touched, "runs_restyled": runs_restyled,
            "reason": "Existing page-number field(s) restyled in place (font/size/bold/italic/color "
                      "only) to match the reference; full table layout was not available to rebuild."}


def _footer_instances_for_section(sec, odd_even=False):
    candidates = [sec.footer]
    try:
        if sec.different_first_page_header_footer:
            candidates.append(sec.first_page_footer)
    except Exception:
        pass
    if odd_even:
        try:
            candidates.append(sec.even_page_footer)
        except Exception:
            pass
    for footer in candidates:
        if footer is not None:
            yield footer


def _all_footer_instances(doc):
    seen_ids = set()
    odd_even = False
    try:
        odd_even = bool(doc.settings.odd_and_even_pages_header_footer)
    except Exception:
        odd_even = False

    for sec in doc.sections:
        for footer in _footer_instances_for_section(sec, odd_even=odd_even):
            fid = id(footer._element)
            if fid in seen_ids:
                continue
            seen_ids.add(fid)
            yield footer


_TBLPR_ELEMENTS_AFTER_BORDERS = [
    'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook',
    'w:tblCaption', 'w:tblDescription',
]
_TCPR_ELEMENTS_AFTER_BORDERS = [
    'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText',
    'w:vAlign', 'w:hideMark', 'w:cnfStyle',
]


def _insert_respecting_order(parent, new_el, successor_tags):
    for tag in successor_tags:
        found = parent.find(qn(tag))
        if found is not None:
            found.addprevious(new_el)
            return
    parent.append(new_el)


def _apply_table_border_style(table, border_style):
    if border_style != "none":
        return
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    borders = OxmlElement('w:tblBorders')
    for tag in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:val'), 'nil')
        borders.append(el)
    _insert_respecting_order(tblPr, borders, _TBLPR_ELEMENTS_AFTER_BORDERS)

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            existing_tc = tcPr.find(qn('w:tcBorders'))
            if existing_tc is not None:
                tcPr.remove(existing_tc)
            tcBorders = OxmlElement('w:tcBorders')
            for tag in ('top', 'left', 'bottom', 'right'):
                el = OxmlElement(f'w:{tag}')
                el.set(qn('w:val'), 'nil')
                tcBorders.append(el)
            _insert_respecting_order(tcPr, tcBorders, _TCPR_ELEMENTS_AFTER_BORDERS)


def _rebuild_footer_table(footer, table_design, cached_page_display="1", usable_width_emu=None):
    for child in list(footer._element):
        footer._element.remove(child)
    footer.add_paragraph()

    n_rows = table_design["n_rows"]
    n_cols = table_design["n_cols"]
    widths_emu = table_design.get("col_widths_emu")
    if widths_emu and len(widths_emu) == n_cols:
        ref_total = sum(widths_emu)
        if usable_width_emu and ref_total and abs(ref_total - usable_width_emu) > 1000:
            scale = usable_width_emu / ref_total
            widths_emu = [int(round(w * scale)) for w in widths_emu]
        total_width_emu = sum(widths_emu)
    else:
        widths_emu = None
        total_width_emu = usable_width_emu or 6000000

    table = footer.add_table(rows=n_rows, cols=n_cols, width=Emu(int(total_width_emu)))

    if widths_emu:
        try:
            for i, col in enumerate(table.columns):
                col.width = Emu(widths_emu[i])
                for cell in col.cells:
                    cell.width = Emu(widths_emu[i])
        except Exception:
            pass

    cells_by_pos = {(c["row"], c["col"]): c for c in table_design["cells"]}
    seen_tc_ids = set()
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            tc_id = id(cell._tc)
            if tc_id in seen_tc_ids:
                continue
            seen_tc_ids.add(tc_id)

            cell_design = cells_by_pos.get((r_idx, c_idx))
            if cell_design is None:
                continue

            cell.text = ""
            p = cell.paragraphs[0]
            align = cell_design.get("alignment")
            if align in ALIGNMENT_MAP:
                p.alignment = ALIGNMENT_MAP[align]

            font = cell_design.get("font") or {}
            if cell_design.get("is_page_field"):
                _insert_page_field_run(p, font, cached_display_text=cached_page_display)
            else:
                text = cell_design.get("text") or ""
                if text:
                    run = p.add_run(text)
                    if font.get("name"):
                        set_run_font_safe(run, font["name"])
                    if font.get("size_pt"):
                        run.font.size = Pt(font["size_pt"])
                    if font.get("bold") is not None:
                        run.font.bold = font["bold"]
                    if font.get("italic") is not None:
                        run.font.italic = font["italic"]
                    if font.get("color_hex"):
                        try:
                            run.font.color.rgb = RGBColor.from_string(font["color_hex"])
                        except Exception:
                            pass
                    sync_complex_script_size_and_emphasis(run)

            shading_hex = cell_design.get("shading_hex")
            if shading_hex:
                _set_cell_shading(cell, shading_hex)

    _apply_table_border_style(table, table_design.get("border_style", "none"))
    return table


def apply_footer_design(doc, footer_design, cached_page_display="1"):
    if not footer_design or not footer_design.get("has_page_field"):
        return {"applied": False, "reason": "Reference footer did not contain a detectable PAGE field; "
                                             "target footer left unchanged."}

    table_design = footer_design.get("table_design")
    if table_design and table_design.get("has_page_field_cell"):
        instances_touched = 0
        seen_ids = set()
        odd_even = False
        try:
            odd_even = bool(doc.settings.odd_and_even_pages_header_footer)
        except Exception:
            odd_even = False
        for sec in doc.sections:
            usable_width_emu = None
            try:
                if sec.page_width and sec.left_margin is not None and sec.right_margin is not None:
                    usable_width_emu = int(sec.page_width - sec.left_margin - sec.right_margin)
            except Exception:
                usable_width_emu = None
            for footer in _footer_instances_for_section(sec, odd_even=odd_even):
                fid = id(footer._element)
                if fid in seen_ids:
                    continue
                seen_ids.add(fid)
                footer.is_linked_to_previous = False
                _rebuild_footer_table(footer, table_design, cached_page_display=cached_page_display,
                                       usable_width_emu=usable_width_emu)
                instances_touched += 1
        return {"applied": True, "sections_touched": instances_touched,
                "reason": "Reference footer's full table layout was replicated exactly into every "
                          "footer instance in the target document."}

    if footer_design.get("page_field_in_table") or footer_design.get("page_field_in_shape"):
        return _restyle_existing_page_fields_in_place(doc, footer_design)

    ref_paragraphs = footer_design.get("paragraphs", [])
    page_field_para_design = next((p for p in ref_paragraphs if p.get("has_page_field")), None)
    if page_field_para_design is None:
        return _restyle_existing_page_fields_in_place(doc, footer_design)

    sections_touched = 0
    for footer in _all_footer_instances(doc):
        footer.is_linked_to_previous = False

        for p in footer.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        first_p = footer.paragraphs[0]
        _clear_paragraph_runs(first_p)

        font = page_field_para_design.get("font") or {}
        align = page_field_para_design.get("alignment")
        if align in ALIGNMENT_MAP:
            first_p.paragraph_format.alignment = ALIGNMENT_MAP[align]

        prefix_text = page_field_para_design.get("prefix_text") or ""
        suffix_text = page_field_para_design.get("suffix_text") or ""

        if prefix_text:
            r_prefix = first_p.add_run(prefix_text)
            if font.get("name"):
                set_run_font_safe(r_prefix, font["name"])
            if font.get("size_pt"):
                r_prefix.font.size = Pt(font["size_pt"])
            if font.get("bold") is not None:
                r_prefix.font.bold = font["bold"]
            if font.get("color_hex"):
                try:
                    r_prefix.font.color.rgb = RGBColor.from_string(font["color_hex"])
                except Exception:
                    pass

        _insert_page_field_run(first_p, font, cached_display_text=cached_page_display)

        if suffix_text:
            r_suffix = first_p.add_run(suffix_text)
            if font.get("name"):
                set_run_font_safe(r_suffix, font["name"])
            if font.get("size_pt"):
                r_suffix.font.size = Pt(font["size_pt"])
            if font.get("bold") is not None:
                r_suffix.font.bold = font["bold"]
            if font.get("color_hex"):
                try:
                    r_suffix.font.color.rgb = RGBColor.from_string(font["color_hex"])
                except Exception:
                    pass

        sections_touched += 1

    return {"applied": True, "sections_touched": sections_touched,
            "reason": "Footer rebuilt with a genuine, auto-updating PAGE field matching the reference's "
                      "captured font/alignment/static-text design."}


def format_document(input_path, policy_path, classification_path, output_path):
    with open(policy_path, encoding="utf-8") as f:
        policy = json.load(f)
    with open(classification_path, encoding="utf-8") as f:
        classification = json.load(f)

    typography = policy.get("typography", {})
    special_headings = policy.get("special_headings", {})
    doc = Document(input_path)
    applied_styles = redefine_styles(doc, policy)
    apply_page_setup(doc, policy)
    change_log = reassign_paragraph_styles(doc, classification, typography, special_headings)
    table_cell_change_log = apply_special_table_cell_formatting(doc, classification, typography)

    table_typography = policy.get("table_typography", {})
    table_typography_result = apply_table_typography(doc, table_typography, classification)

    table_contexts_policy = policy.get("table_contexts", {})
    tables_touched = apply_table_context_styles(doc, classification, table_contexts_policy)

    row_banding_policy = policy.get("table_row_banding", {})
    row_banding_result = apply_table_body_row_banding(doc, row_banding_policy, classification)

    footer_design = policy.get("footer_design", {})
    footer_result = apply_footer_design(doc, footer_design)

    _fix_zoom_element(doc)
    _force_update_fields_on_open(doc)

    doc.save(output_path)

    return {
        "applied_styles": applied_styles,
        "paragraph_change_log": change_log,
        "table_cell_change_log": table_cell_change_log,
        "table_typography_result": table_typography_result,
        "tables_touched_by_context": tables_touched,
        "table_row_banding_result": row_banding_result,
        "table_contexts_policy": table_contexts_policy,
        "table_typography_used": table_typography,
        "special_headings_applied": [k for k in special_headings.keys()],
        "footer_result": footer_result,
        "output_path": output_path,
    }


def main():
    ap = argparse.ArgumentParser(description="Phase 3 - Apply formatting policy deterministically.")
    ap.add_argument("--input", required=True)
    ap.add_argument("--policy", required=True)
    ap.add_argument("--classification", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    result = format_document(args.input, args.policy, args.classification, args.output)
    print(f"[formatting_engine] Table typography result: {result['table_typography_result']}")
    print(f"[formatting_engine] Table row banding result: {result['table_row_banding_result']}")
    print(f"[formatting_engine] Saved -> {result['output_path']}")


if __name__ == "__main__":
    main()

