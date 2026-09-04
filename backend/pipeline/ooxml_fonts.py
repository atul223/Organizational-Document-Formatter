"""Low-level OOXML font helpers. Fixes Word's theme-attribute rendering
precedence bug: setting a literal font name via python-docx does not
remove pre-existing theme attributes on the same <w:rFonts> element, and
Word renders the theme value instead of the literal one when both are
present. These helpers explicitly clear theme attributes whenever a
literal font is set."""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_THEME_ATTRS = ['w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme']
_LITERAL_ATTRS = ['w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs']


def set_literal_font_no_theme_override(rPr_element, font_name):
    if rPr_element is None or not font_name:
        return
    rFonts = rPr_element.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr_element.insert(0, rFonts)
    for attr in _LITERAL_ATTRS:
        rFonts.set(qn(attr), font_name)
    for attr in _THEME_ATTRS:
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]


def set_run_font_safe(run, font_name):
    if not font_name:
        return
    rPr = run._element.get_or_add_rPr()
    set_literal_font_no_theme_override(rPr, font_name)


def set_style_font_safe(style, font_name):
    if not font_name:
        return
    rPr = style.element.get_or_add_rPr()
    set_literal_font_no_theme_override(rPr, font_name)


def sync_complex_script_size_and_emphasis(run):
    rPr = run._element.rPr
    if rPr is None:
        return
    sz = rPr.find(qn('w:sz'))
    if sz is not None:
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = OxmlElement('w:szCs')
            sz.addnext(szCs)
        szCs.set(qn('w:val'), sz.get(qn('w:val')))
    b = rPr.find(qn('w:b'))
    bCs = rPr.find(qn('w:bCs'))
    if b is not None and bCs is None:
        bCs = OxmlElement('w:bCs')
        b.addnext(bCs)
    i = rPr.find(qn('w:i'))
    iCs = rPr.find(qn('w:iCs'))
    if i is not None and iCs is None:
        iCs = OxmlElement('w:iCs')
        i.addnext(iCs)


def extract_theme_fonts(docx_path):
    import zipfile
    from lxml import etree
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    result = {"major_latin": None, "minor_latin": None}
    try:
        with zipfile.ZipFile(docx_path) as z:
            theme_names = sorted(
                n for n in z.namelist()
                if n.startswith('word/theme/theme') and n.endswith('.xml')
            )
            if not theme_names:
                return result
            with z.open(theme_names[0]) as f:
                tree = etree.parse(f)
    except Exception:
        return result

    major = tree.find('.//a:fontScheme/a:majorFont/a:latin', ns)
    minor = tree.find('.//a:fontScheme/a:minorFont/a:latin', ns)
    if major is not None:
        result["major_latin"] = major.get('typeface') or None
    if minor is not None:
        result["minor_latin"] = minor.get('typeface') or None
    for k, v in list(result.items()):
        if not v or v.strip().lower() == "none":
            result[k] = None
    return result
