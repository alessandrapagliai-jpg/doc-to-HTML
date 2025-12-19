import re
import html
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Iterable, Union

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

TESTO_KEYS = {"testo", "content", "article", "body", "testo articolo"}


# =========================
# HTML entities (SOLO OUTPUT)
# =========================

def html_entities(s: str) -> str:
    if not s:
        return ""

    anchors = []

    def stash_anchor(m):
        anchors.append(m.group(0))
        return f"__ANCHOR_{len(anchors)-1}__"

    s = re.sub(r"<a\s+href=\"[^\"]+\">.*?</a>", stash_anchor, s, flags=re.I | re.S)

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
    }

    for k, v in replacements.items():
        s = s.replace(k, v)

    for i, a in enumerate(anchors):
        s = s.replace(f"__ANCHOR_{i}__", a)

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

def set_cell_text(cell, text: str, bold=False, color=None, size_pt=10):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(text or "")
    r.bold = bold
    r.font.size = Pt(size_pt)
    if color:
        r.font.color.rgb = color


# =========================
# Iterazione DOCX
# =========================

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    parent_elm = parent.element.body if hasattr(parent, "element") else parent._tc
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =========================
# Hyperlink-aware paragraph
# =========================

def _iter_text_runs(node) -> str:
    texts = []
    for t in node.findall(".//w:t", namespaces=node.nsmap):
        if t.text:
            texts.append(t.text)
    return "".join(texts)

def paragraph_to_text_with_links(paragraph: Paragraph) -> str:
    out = []
    part = paragraph.part

    for child in paragraph._p.iterchildren():
        tag = child.tag

        # hyperlink
        if tag.endswith("}hyperlink"):
            r_id = child.get(qn("r:id"))
            text = _iter_text_runs(child)

            if r_id and r_id in part.rels and text.strip():
                href = part.rels[r_id].target_ref
                out.append(f'<a href="{href}">{text}</a>')

        # run normale (NON dentro hyperlink)
        elif tag.endswith("}r"):
            if child.getparent() is not paragraph._p:
                continue

            text = _iter_text_runs(child)
            if text.strip():
                out.append(text)

    return "".join(out).strip()


# =========================
# Estrazione RAW
# =========================

def extra
