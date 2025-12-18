import re
import html
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional, Iterable, Union

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# =========================
# CONFIG
# =========================

OUTPUT_META_LABELS = ["Title", "Description", "URL", "Territories", "Target Keyword"]

# Varianti input -> chiave output (case-insensitive)
INPUT_KEY_MAP = {
    "title": "Title",
    "seo title": "Title",
    "meta title": "Title",

    "description": "Description",
    "meta description": "Description",

    "url": "URL",

    "territories": "Territories",
    "territory": "Territories",

    "target keyword": "Target Keyword",
    "target keywords": "Target Keyword",
    "kw": "Target Keyword",
    "keyword": "Target Keyword",
    "primary keyword": "Target Keyword",

    "h1": "H1",
}

# riconosci “Testo:” con varianti
TESTO_KEYS = {"testo", "content", "article", "body", "testo articolo"}


# =========================
# HTML helpers
# =========================

def html_entities(s: str) -> str:
    """
    Converte caratteri speciali in HTML entities.
    NOTA: da usare SOLO per l'output HTML (H1/H2/H3/body).
    """
    if not s:
        return ""

    # escape di & < >
    s = html.escape(s, quote=False)

    replacements = {
        "’": "&rsquo;",
        "‘": "&lsquo;",
        "“": "&ldquo;",
        "”": "&rdquo;",
        "–": "&ndash;",
        "—": "&mdash;",
        "…": "&hellip;",
        "à": "&agrave;",
        "è": "&egrave;",
        "é": "&eacute;",
        "ì": "&igrave;",
        "ò": "&ograve;",
        "ù": "&ugrave;",
        "À": "&Agrave;",
        "È": "&Egrave;",
        "É": "&Eacute;",
        "Ì": "&Igrave;",
        "Ò": "&Ograve;",
        "Ù": "&Ugrave;",
        "ô": "&ocirc;",
        "Ô": "&Ocirc;",
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    return s


# =========================
# DOCX helpers
# =========================

def shade_cell(cell, fill_hex: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)

def set_cell_text(cell, text: str, bold: bool = False, color: Optional[RGBColor] = None, size_pt: int = 10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text or "")
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = color


def _normalize_key(k: str) -> str:
    k = k.strip().lower()
    k = re.sub(r"\s+", " ", k)
    return k

def _is_document_like(obj) -> bool:
    return hasattr(obj, "element") and hasattr(obj.element, "body")

def _is_cell_like(obj) -> bool:
    return hasattr(obj, "_tc")

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    """
    Itera paragraph e tabelle in ordine di apparizione.
    Supporta Document e celle.
    """
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    if _is_document_like(parent):
        parent_elm = parent.element.body
    elif _is_cell_like(parent):
        parent_elm = parent._tc
    else:
        raise TypeError("Parent non supportato: {}".format(type(parent)))

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =========================
# Hyperlink extraction (HTML channel)
# =========================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _runs_text(run_elm) -> str:
    out = []
    for node in run_elm.iter():
        if node.tag.endswith("}t") and node.text:
            out.append(node.text)
    return "".join(out)

def _extract_url_from_instr(instr: str) -> str:
    if not instr:
        return "#"
    instr = instr.strip()

    m = re.search(r'HYPERLINK\s+"([^"]+)"', instr, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    m = re.search(r"HYPERLINK\s+(\S+)", instr, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    return "#"

def paragraph_to_html(paragraph: Paragraph) -> str:
    """
    Converte un Par
