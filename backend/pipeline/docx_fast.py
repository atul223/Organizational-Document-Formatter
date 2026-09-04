"""Shared performance + style-resolution helpers."""
from docx.oxml.ns import qn

BUILTIN_STYLE_TO_ROLE = {
    "Title": "Title",
    "Heading 1": "H1",
    "Heading 2": "H2",
    "Heading 3": "H3",
    "Heading 4": "H4",
    "Caption": "Caption",
    "Quote": "Quote",
    "Intense Quote": "Quote",
    "List Bullet": "ListBullet",
    "List Paragraph": "ListBullet",
    "Normal": "Body",
    "Body Text": "Body",
}


def build_styleid_to_name(doc):
    mapping = {}
    for s in doc.styles:
        try:
            if s.style_id:
                mapping[s.style_id] = s.name
        except Exception:
            continue
    return mapping


def get_raw_pstyle_id(paragraph_or_xml):
    p_elm = getattr(paragraph_or_xml, "_p", paragraph_or_xml)
    pPr = p_elm.find(qn('w:pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        return None
    return pStyle.get(qn('w:val'))


def get_paragraph_style_name_fast(paragraph, styleid_to_name, default_name="Normal"):
    style_id = get_raw_pstyle_id(paragraph)
    if style_id is None:
        return default_name
    return styleid_to_name.get(style_id, default_name)


def resolve_base_style_role(style, builtin_style_to_role, max_depth=15):
    seen_ids = set()
    current = style
    depth = 0
    while current is not None and depth < max_depth:
        try:
            name = current.name
        except Exception:
            break
        if name in builtin_style_to_role:
            return builtin_style_to_role[name]
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
    return None


def build_styleid_to_role_map(doc, builtin_style_to_role):
    try:
        from docx.enum.style import WD_STYLE_TYPE
        paragraph_type = WD_STYLE_TYPE.PARAGRAPH
    except Exception:
        paragraph_type = None

    mapping = {}
    for s in doc.styles:
        try:
            if paragraph_type is not None and s.type != paragraph_type:
                continue
        except Exception:
            pass
        style_id = getattr(s, "style_id", None)
        if not style_id:
            continue
        mapping[style_id] = resolve_base_style_role(s, builtin_style_to_role)
    return mapping


def resolve_base_style_role_with_depth(style, builtin_style_to_role, max_depth=15):
    seen_ids = set()
    current = style
    depth = 0
    while current is not None and depth < max_depth:
        try:
            name = current.name
        except Exception:
            break
        if name in builtin_style_to_role:
            return builtin_style_to_role[name], depth
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
    return None, None


def build_styleid_to_role_depth_map(doc, builtin_style_to_role):
    try:
        from docx.enum.style import WD_STYLE_TYPE
        paragraph_type = WD_STYLE_TYPE.PARAGRAPH
    except Exception:
        paragraph_type = None

    mapping = {}
    for s in doc.styles:
        try:
            if paragraph_type is not None and s.type != paragraph_type:
                continue
        except Exception:
            pass
        style_id = getattr(s, "style_id", None)
        if not style_id:
            continue
        role, depth = resolve_base_style_role_with_depth(s, builtin_style_to_role)
        if role is not None:
            mapping[style_id] = depth
    return mapping


def dedupe_row_cells(row):
    """Returns a row's cells with merged-cell duplicates removed."""
    seen = set()
    result = []
    for c in row.cells:
        tc_id = id(c._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        result.append(c)
    return result


def _row_cell_texts(row):
    return [c.text.strip() for c in dedupe_row_cells(row)]


def _row0_has_merge(row):
    tc_ids = [id(c._tc) for c in row.cells]
    return len(set(tc_ids)) < len(tc_ids)


def _row_is_all_short(row, max_words=4):
    texts = [t for t in _row_cell_texts(row) if t]
    if not texts:
        return False
    return all(len(t.split()) <= max_words for t in texts)


def count_header_rows(table, max_header_rows=2):
    """Detects whether a table has a single header row (the default) or
    two (a primary label row plus a short sub-header row)."""
    if not table.rows:
        return 0
    if len(table.rows) < 2:
        return 1
    row0 = table.rows[0]
    row1 = table.rows[1]
    if _row0_has_merge(row0) and _row_is_all_short(row1, max_words=4):
        return min(2, max_header_rows)
    return 1


def header_row_cells(table, header_row_count):
    """Returns a single, deduplicated list of cells across ALL header
    rows [0, header_row_count) of a table."""
    cells = []
    seen = set()
    for r_idx in range(min(header_row_count, len(table.rows))):
        for c in table.rows[r_idx].cells:
            tc_id = id(c._tc)
            if tc_id in seen:
                continue
            seen.add(tc_id)
            cells.append(c)
    return cells


def table_header_signature(table, header_row_count):
    """v1.22 NEW: returns a stable, JSON-serializable "shape signature"
    for a table, based on its NORMALIZED header cell texts (upper-
    cased, whitespace-collapsed) joined across all header rows, plus
    its column count. Two tables sharing the exact same header text
    (e.g. every "No. | Activity | Description | Timing" activity table
    that recurs throughout a document) are considered the SAME shape
    for the purposes of row-banding pattern detection/matching - see
    policy_extractor.extract_body_row_banding for the full rationale.
    Returns None if the table has no header cells to build a signature
    from at all (e.g. an empty table)."""
    import re
    cells = header_row_cells(table, header_row_count)
    if not cells:
        return None
    texts = tuple(re.sub(r"\s+", " ", c.text or "").strip().upper() for c in cells)
    if not any(texts):
        return None
    return "col{}::{}".format(len(texts), "|".join(texts))


# ---------------------------------------------------------------------
# v1.21: reads Word's own per-row `<w:cnfStyle>` conditional-formatting
# bit-flag.
# ---------------------------------------------------------------------

_CNF_STYLE_BIT_ORDER = [
    "firstRow", "lastRow", "firstColumn", "lastColumn",
    "oddVBand", "evenVBand", "oddHBand", "evenHBand",
    "firstRowFirstColumn", "firstRowLastColumn",
    "lastRowFirstColumn", "lastRowLastColumn",
]
_CNF_BAND1_HORZ_BIT_INDEX = 6
_CNF_BAND2_HORZ_BIT_INDEX = 7


def _parse_cnf_style_element(cnf_element):
    if cnf_element is None:
        return None
    val = cnf_element.get(qn('w:val'))
    if not val or len(val) < 8:
        return None
    return val


def read_row_cnf_band(row):
    """Returns 0 if this row's OWN <w:cnfStyle> (on its <w:trPr>)
    explicitly marks it as the FIRST horizontal band (band1Horz), 1 if
    it explicitly marks the SECOND (band2Horz), or None if no cnfStyle
    is present on this row at all."""
    try:
        trPr = row._tr.find(qn('w:trPr'))
    except Exception:
        return None
    if trPr is None:
        return None
    cnf = trPr.find(qn('w:cnfStyle'))
    val = _parse_cnf_style_element(cnf)
    if val is None:
        return None
    if val[_CNF_BAND1_HORZ_BIT_INDEX] == '1':
        return 0
    if val[_CNF_BAND2_HORZ_BIT_INDEX] == '1':
        return 1
    return None


def read_cell_cnf_band(cell):
    """Same as `read_row_cnf_band`, but reads a per-CELL <w:cnfStyle>
    (on the cell's own <w:tcPr>) instead of the row's."""
    try:
        tcPr = cell._tc.tcPr
    except Exception:
        return None
    if tcPr is None:
        return None
    cnf = tcPr.find(qn('w:cnfStyle'))
    val = _parse_cnf_style_element(cnf)
    if val is None:
        return None
    if val[_CNF_BAND1_HORZ_BIT_INDEX] == '1':
        return 0
    if val[_CNF_BAND2_HORZ_BIT_INDEX] == '1':
        return 1
    return None


def clear_cnf_style(element_with_tcpr_or_trpr, is_row=False):
    """Removes any existing <w:cnfStyle> from a row's <w:trPr> or a
    cell's <w:tcPr> entirely."""
    try:
        if is_row:
            container = element_with_tcpr_or_trpr._tr.find(qn('w:trPr'))
        else:
            container = element_with_tcpr_or_trpr._tc.tcPr
    except Exception:
        return
    if container is None:
        return
    cnf = container.find(qn('w:cnfStyle'))
    if cnf is not None:
        container.remove(cnf)
