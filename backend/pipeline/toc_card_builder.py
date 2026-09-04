"""
Phase 3.5 — TOC Chapter-Card Table Builder

Rebuilds the reference document's "chapter card" Table of Contents design
in the target document - using ONLY text that already exists in the
target's own linear TOC listing. Nothing is invented. Deletes only the
redundant top-level "SECTION N. Title ... page" summary line for each
chapter card built (its content is now verbatim inside the card); every
subsection line (1.1, 2.1.1, etc.) is left completely untouched. A
whole-document word-reconciliation check runs before saving; on any
failure the document is passed through unchanged.
"""
import re
import shutil
from collections import Counter
from docx import Document
from docx.shared import Pt, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from ooxml_fonts import set_run_font_safe
from policy_extractor import _normalize_heading_text, TOC_HEADING_NORM

CHAPTER_LINE_RE = re.compile(r"^SECTION\s+(\d+)\.?\s*(.+?)\s+(\d{1,4})\s*$", re.IGNORECASE)
CHILD_LINE_RE = re.compile(r"^(\d+)\.(\d+)\s+(.+?)\s+(\d{1,4})\s*$")
ANY_NUMBERED_TOC_LINE_RE = re.compile(r"^(\d+(?:\.\d+)+)\s+(.+?)\s+(\d{1,4})\s*$")

SUBTITLE_SEPARATOR = " | "
MAX_SUBTITLE_ITEMS = 3

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


def _document_word_bag(doc):
    bag = Counter()
    for p in doc.paragraphs:
        bag.update(_tokenize(p.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                bag.update(_tokenize(cell.text))
    return bag


def parse_toc_chapters(doc, toc_idx, max_scan=500):
    chapters = []
    current = None
    consecutive_unmatched = 0
    front_matter_indices = []
    seen_first_chapter = False
    end_idx = min(len(doc.paragraphs), toc_idx + 1 + max_scan)

    for idx in range(toc_idx + 1, end_idx):
        paragraph = doc.paragraphs[idx]
        text = paragraph.text.strip()
        if not text:
            continue

        m_top = CHAPTER_LINE_RE.match(text)
        if m_top:
            consecutive_unmatched = 0
            seen_first_chapter = True
            current = {
                "number_raw": m_top.group(1),
                "title": m_top.group(2).strip(),
                "page": m_top.group(3),
                "children_titles": [],
                "chapter_line_element": paragraph._p,
                "chapter_line_full_text": text,
            }
            chapters.append(current)
            continue

        m_child = CHILD_LINE_RE.match(text)
        if m_child and current is not None and m_child.group(1) == current["number_raw"]:
            consecutive_unmatched = 0
            current["children_titles"].append(m_child.group(3).strip())
            continue

        if not seen_first_chapter:
            front_matter_indices.append(idx)
            continue

        if ANY_NUMBERED_TOC_LINE_RE.match(text):
            consecutive_unmatched = 0
            continue

        consecutive_unmatched += 1
        if consecutive_unmatched >= 2:
            break

    return front_matter_indices, chapters


def _apply_font_dict(run, font_dict, force_bold_default=False):
    font_dict = font_dict or {}
    name = font_dict.get("name")
    if name:
        set_run_font_safe(run, name)
    size = font_dict.get("size_pt")
    if size:
        run.font.size = Pt(size)
    bold = font_dict.get("bold")
    if bold is not None:
        run.font.bold = bold
    elif force_bold_default:
        run.font.bold = True
    italic = font_dict.get("italic")
    if italic is not None:
        run.font.italic = italic
    color = font_dict.get("color_hex")
    if color:
        try:
            run.font.color.rgb = RGBColor.from_string(color)
        except Exception:
            pass


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


def _remove_all_table_borders(table):
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


def _fill_simple_cell(cell, text, font_dict, shading_hex, align_right=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if align_right:
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    _apply_font_dict(run, font_dict, force_bold_default=True)
    _set_cell_shading(cell, shading_hex)


def _fill_title_subtitle_cell(cell, title, subtitle, column_design):
    cell.text = ""
    p_title = cell.paragraphs[0]
    run_title = p_title.add_run(title)
    _apply_font_dict(run_title, column_design.get("primary_font"), force_bold_default=True)
    if subtitle:
        p_sub = cell.add_paragraph()
        run_sub = p_sub.add_run(subtitle)
        _apply_font_dict(run_sub, column_design.get("secondary_font"))
    _set_cell_shading(cell, column_design.get("shading_hex"))


def _fill_title_subtitle_page_cell(cell, title, subtitle, page, column_design):
    cell.text = ""
    p_title = cell.paragraphs[0]
    run_title = p_title.add_run(title)
    _apply_font_dict(run_title, column_design.get("primary_font"), force_bold_default=True)
    run_page = p_title.add_run("\t" + page)
    _apply_font_dict(run_page, column_design.get("primary_font"), force_bold_default=True)
    if subtitle:
        p_sub = cell.add_paragraph()
        run_sub = p_sub.add_run(subtitle)
        _apply_font_dict(run_sub, column_design.get("secondary_font"))
    _set_cell_shading(cell, column_design.get("shading_hex"))


def build_card_table(doc, anchor_paragraph, chapters, design):
    n_cols = design.get("n_cols") or 3
    n_cols = 3 if n_cols not in (2, 3) else n_cols
    columns = list(design.get("columns") or [])
    while len(columns) < n_cols:
        columns.append({"shading_hex": None, "primary_font": {}, "secondary_font": {}})

    table = doc.add_table(rows=len(chapters), cols=n_cols)
    try:
        table.autofit = False
    except Exception:
        pass
    anchor_paragraph._p.addnext(table._tbl)

    widths_emu = design.get("col_widths_emu")
    if widths_emu and len(widths_emu) == n_cols:
        try:
            for i, col in enumerate(table.columns):
                col.width = Emu(widths_emu[i])
                for cell in col.cells:
                    cell.width = Emu(widths_emu[i])
        except Exception:
            pass

    for row_idx, chapter in enumerate(chapters):
        row = table.rows[row_idx]
        num_str = chapter["number_raw"]
        try:
            num_str = f"{int(num_str):02d}"
        except Exception:
            pass
        title = chapter["title"]
        subtitle = SUBTITLE_SEPARATOR.join(chapter["children_titles"][:MAX_SUBTITLE_ITEMS])
        page = chapter["page"]

        if n_cols == 3:
            _fill_simple_cell(row.cells[0], num_str, columns[0].get("primary_font"), columns[0].get("shading_hex"))
            _fill_title_subtitle_cell(row.cells[1], title, subtitle, columns[1])
            _fill_simple_cell(row.cells[2], page, columns[2].get("primary_font"), columns[2].get("shading_hex"),
                               align_right=True)
        else:
            _fill_simple_cell(row.cells[0], num_str, columns[0].get("primary_font"), columns[0].get("shading_hex"))
            _fill_title_subtitle_page_cell(row.cells[1], title, subtitle, page, columns[1])

    _remove_all_table_borders(table)

    return table


def apply_toc_heading_underline(paragraph, border_spec):
    if not border_spec:
        return
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    bottom = pBdr.find(qn('w:bottom'))
    if bottom is None:
        bottom = OxmlElement('w:bottom')
        pBdr.append(bottom)
    if border_spec.get("val"):
        bottom.set(qn('w:val'), border_spec["val"])
    if border_spec.get("sz"):
        bottom.set(qn('w:sz'), border_spec["sz"])
    if border_spec.get("space"):
        bottom.set(qn('w:space'), border_spec["space"])
    if border_spec.get("color"):
        bottom.set(qn('w:color'), border_spec["color"])


def style_front_matter_lines(doc, front_matter_indices, subtitle_style):
    if not subtitle_style or not subtitle_style.get("font"):
        return
    font = subtitle_style["font"]
    for idx in front_matter_indices:
        p = doc.paragraphs[idx]
        for run in p.runs:
            _apply_font_dict(run, font)


def _delete_paragraph_element(p_element):
    parent = p_element.getparent()
    if parent is not None:
        parent.remove(p_element)


def try_build_toc_card_table(input_path, policy, output_path):
    design = policy.get("toc_card_design", {}) or {}
    result = {
        "attempted": False, "applied": False, "reason": None,
        "chapters_found": 0, "chapters": [],
        "top_level_lines_removed": 0,
        "child_lines_retained_unchanged": 0,
        "word_reconciliation_passed": None,
        "word_reconciliation_detail": None,
    }

    if not design.get("detected"):
        result["reason"] = ("Reference document's TABLE OF CONTENTS did not contain a "
                            "detectable chapter-card table design; skipping restructuring.")
        shutil.copyfile(input_path, output_path)
        return result

    doc = Document(input_path)
    toc_idx = None
    for i, p in enumerate(doc.paragraphs):
        if _normalize_heading_text(p.text) == TOC_HEADING_NORM:
            toc_idx = i
            break

    if toc_idx is None:
        result["reason"] = "Could not find a 'TABLE OF CONTENTS' heading paragraph in the target document."
        shutil.copyfile(input_path, output_path)
        return result

    result["attempted"] = True
    original_word_bag = _document_word_bag(doc)

    front_matter_indices, chapters = parse_toc_chapters(doc, toc_idx)
    valid_chapters = [c for c in chapters if c.get("title") and c.get("page")]
    result["chapters_found"] = len(valid_chapters)
    result["chapters"] = [
        {"number": c["number_raw"], "title": c["title"], "page": c["page"],
         "subtitle_sources": c["children_titles"][:MAX_SUBTITLE_ITEMS],
         "total_children_retained": len(c["children_titles"])}
        for c in valid_chapters
    ]
    result["child_lines_retained_unchanged"] = sum(len(c["children_titles"]) for c in valid_chapters)

    if len(valid_chapters) < 2:
        result["reason"] = (f"Only {len(valid_chapters)} chapter-level TOC entries were confidently "
                            f"parsed (need >= 2 for a meaningful card table); skipping restructuring.")
        shutil.copyfile(input_path, output_path)
        return result

    all_text = "\n".join(p.text for p in doc.paragraphs)
    for ch in valid_chapters:
        if ch["title"] not in all_text or ch["page"] not in all_text:
            result["reason"] = f"Traceability pre-check failed for chapter {ch['title']!r}; aborting."
            shutil.copyfile(input_path, output_path)
            return result

    anchor_paragraph = doc.paragraphs[front_matter_indices[-1]] if front_matter_indices else doc.paragraphs[toc_idx]

    style_front_matter_lines(doc, front_matter_indices, design.get("subtitle_style"))
    apply_toc_heading_underline(doc.paragraphs[toc_idx], design.get("title_underline"))
    build_card_table(doc, anchor_paragraph, valid_chapters, design)

    for ch in valid_chapters:
        _delete_paragraph_element(ch["chapter_line_element"])
    result["top_level_lines_removed"] = len(valid_chapters)

    final_word_bag = _document_word_bag(doc)
    missing = {}
    for word, count in original_word_bag.items():
        final_count = final_word_bag.get(word, 0)
        if final_count < count:
            missing[word] = (count, final_count)

    if missing:
        sample = list(missing.items())[:10]
        result["word_reconciliation_passed"] = False
        result["word_reconciliation_detail"] = (
            f"{len(missing)} word(s) had a lower count after restructuring than before - "
            f"aborting and leaving the document unchanged. Sample: {sample}"
        )
        result["reason"] = result["word_reconciliation_detail"]
        shutil.copyfile(input_path, output_path)
        return result

    result["word_reconciliation_passed"] = True
    result["word_reconciliation_detail"] = (
        "Every word present before TOC restructuring is still present (at least as "
        "many times) after restructuring - verified via whole-document word count "
        "reconciliation."
    )

    doc.save(output_path)
    result["applied"] = True
    result["reason"] = ("TOC chapter-card table successfully built. Top-level chapter summary "
                        "lines were relocated into the table; all subsection lines remain "
                        "completely unchanged directly below it.")
    return result
