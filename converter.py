import re
import html
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Iterable, Union, Optional

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
        "ô": "&ocirc;",
        "Ô": "&Ocirc;",
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
    lines: List[str] = []

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
# PARSING INPUT
# =========================

Heading = Dict[str, Any]
# heading item structure:
# {
#   "level": 2 or 3,
#   "block": "S2" or "✏️ S3",
#   "tag": "h2" or "h3",
#   "title": "<h2>..</h2>" (html),
#   "paras": ["<p>..</p>", ...]
# }

def parse_input_docx(path: Path) -> Dict[str, Any]:
    doc = Document(str(path))
    lines = extract_lines_raw(doc)

    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""

    in_testo = False
    testo_re = re.compile(r"^({})\s*:\s*$".format("|".join(TESTO_KEYS)), re.I)

    intro: List[str] = []
    headings: List[Heading] = []

    current: Optional[Heading] = None

    def flush_current():
        nonlocal current
        if current is not None:
            headings.append(current)
            current = None

    for line in lines:
        if testo_re.match(line):
            in_testo = True
            continue

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

        # BODY: heading?
        m = re.match(r"^(.*)\s*\((h2|h3)\)\s*$", line, re.I)
        if m:
            # nuovo heading => chiudi il precedente
            flush_current()

            level = 2 if m.group(2).lower() == "h2" else 3
            tag = "h2" if level == 2 else "h3"
            block = "S2" if level == 2 else "✏️ S3"

            title_txt = html_entities(m.group(1).strip())
            title_html = f"<{tag}>{title_txt}</{tag}>"

            current = {
                "level": level,
                "block": block,
                "tag": tag,
                "title": title_html,
                "paras": []
            }
            continue

        # paragraph
        p_html = f'<p class="h-text-size-14 h-font-primary">{html_entities(line)}</p>'

        if current is None:
            intro.append(p_html)
        else:
            current["paras"].append(p_html)

    flush_current()

    if not h1:
        h1 = meta.get("Title") or "Untitled"

    return {
        "meta": meta,
        "h1": h1,
        "intro": intro,
        "headings": headings,
    }


# =========================
# STRUCTURE OF CONTENT
# =========================

def build_structure(parsed: Dict[str, Any]) -> List[str]:
    s = ["H1", "Intro"]
    for h in parsed["headings"]:
        s.append(h["block"])
    return s


# =========================
# OUTPUT DOCX
# =========================

def write_output_docx(parsed: Dict[str, Any], out: Path):
    doc = Document()
    meta = parsed["meta"]

    # Titolo grande (RAW)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(meta.get("Title") or parsed["h1"])
    r.bold = True
    r.font.size = Pt(20)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Tabella metadati
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

    # Tabella Block | HTML Output
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"
    set_cell_text(t2.cell(0, 0), "Block", bold=True)
    set_cell_text(t2.cell(0, 1), "⭐ HTML Output ⭐", bold=True)

    # H1
    row = t2.add_row().cells
    row[0].text = "H1"
    row[1].add_paragraph(f"<h1>{html_entities(parsed['h1'])}</h1>")

    # Intro (una riga per paragrafo)
    for p_html in parsed["intro"]:
        row = t2.add_row().cells
        row[0].text = "Intro"
        row[1].add_paragraph(p_html)

    # Headings: UNA RIGA PER HEADING (S2 o S3) con header + paragrafi dentro la stessa cella
    for h in parsed["headings"]:
        row = t2.add_row().cells
        row[0].text = h["block"]

        cell = row[1]
        cell.text = ""
        cell.add_paragraph(h["title"])
        for p_html in h["paras"]:
            cell.add_paragraph(p_html)

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
