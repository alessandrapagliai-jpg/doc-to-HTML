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
# HTML ENTITIES (SOLO OUTPUT)
# =========================

def html_entities(s: str) -> str:
    if not s:
        return ""

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

    return s


# =========================
# DOCX HELPERS
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
# ITERAZIONE DOCX
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


def extract_lines_raw(doc: Document) -> List[str]:
    lines = []

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = (block.text or "").strip()
            if t:
                lines.append(t)

        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        t = (p.text or "").strip()
                        if t:
                            lines.append(t)

    return lines


# =========================
# PARSING INPUT DOCX
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    doc = Document(str(path))
    lines = extract_lines_raw(doc)

    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""
    intro_paras: List[str] = []
    sections: List[Dict[str, Any]] = []

    in_testo = False
    testo_re = re.compile(r"^({})\s*:\s*$".format("|".join(TESTO_KEYS)), re.I)

    current_section = None

    for line in lines:
        # start body
        if testo_re.match(line):
            in_testo = True
            continue

        # META
        if not in_testo:
            m = re.match(r"^([^:]{1,80})\s*:\s*(.*)$", line)
            if m:
                key = m.group(1).strip().lower()
                val = m.group(2).strip()
                if key in INPUT_KEY_MAP:
                    out = INPUT_KEY_MAP[key]
                    if out == "H1":
                        h1 = val
                    elif out in meta:
                        meta[out] = val
            continue

        # BODY — HEADER
        m = re.match(r"^(.*)\s*\((h2|h3)\)\s*$", line, re.I)
        if m:
            # flush sezione precedente
            if current_section:
                sections.append(current_section)

            tag = m.group(2).lower()
            title = html_entities(m.group(1).strip())

            current_section = {
                "block": "✏️ S3",
                "items": [f"<{tag}>{title}</{tag}>"]
            }
            continue

        # BODY — PARAGRAFI
        if current_section is None:
            intro_paras.append(
                f'<p class="h-text-size-14 h-font-primary">{html_entities(line)}</p>'
            )
        else:
            current_section["items"].append(
                f'<p class="h-text-size-14 h-font-primary">{html_entities(line)}</p>'
            )

    if current_section:
        sections.append(current_section)

    if not h1:
        h1 = meta.get("Title") or "Untitled"

    return {
        "meta": meta,
        "h1": h1,
        "intro": intro_paras,
        "sections": sections
    }


# =========================
# STRUCTURE OF CONTENT
# =========================

def build_structure(parsed: Dict[str, Any]) -> List[str]:
    s = ["H1", "Intro"]
    s.extend(["✏️ S3"] * len(parsed["sections"]))
    return s


# =========================
# OUTPUT DOCX
# =========================

def write_output_docx(parsed: Dict[str, Any], out: Path):
    doc = Document()
    meta = parsed["meta"]

    # Titolo
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(meta.get("Title") or parsed["h1"])
    r.bold = True
    r.font.size = Pt(20)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Meta table
    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"

    for i, k in enumerate(OUTPUT_META_LABELS):
        shade_cell(table.cell(i, 0), "000000")
        set_cell_text(table.cell(i, 0), k, bold=True, color=RGBColor(255, 255, 255))
        set_cell_text(table.cell(i, 1), meta.get(k, ""))

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Structure of content
    p = doc.add_paragraph("Structure of content:")
    p.runs[0].bold = True
    for l in build_structure(parsed):
        doc.add_paragraph(l)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # HTML Output table
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"
    set_cell_text(t2.cell(0, 0), "Block", bold=True)
    set_cell_text(t2.cell(0, 1), "⭐ HTML Output ⭐", bold=True)

    # H1
    row = t2.add_row().cells
    row[0].text = "H1"
    row[1].add_paragraph(f"<h1>{html_entities(parsed['h1'])}</h1>")

    # Intro
    for p_html in parsed["intro"]:
        row = t2.add_row().cells
        row[0].text = "Intro"
        row[1].add_paragraph(p_html)

    # S3 SECTIONS
    for sec in parsed["sections"]:
        row = t2.add_row().cells
        row[0].text = sec["block"]

        cell = row[1]
        cell.text = ""

        for html_block in sec["items"]:
            cell.add_paragraph(html_block)

    doc.save(str(out))


# =========================
# STREAMLIT HELPER
# =========================

def convert_uploaded_file(uploaded_file):
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        inp = d / uploaded_file.name
        out = d / f"output_{uploaded_file.name}"
        inp.write_bytes(uploaded_file.read())
        parsed = parse_input_docx(inp)
        write_output_docx(parsed, out)
        final = Path(tempfile.gettempdir()) / out.name
        final.write_bytes(out.read_bytes())
        return final
