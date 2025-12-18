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
# HTML helpers
# =========================

def html_entities(s: str) -> str:
    """HTML entities SOLO per body / headings"""
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
    run = p.add_run(text or "")
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color


# =========================
# Iter DOCX blocks
# =========================

def _normalize_key(k: str) -> str:
    return re.sub(r"\s+", " ", k.strip().lower())

def _is_document_like(obj) -> bool:
    return hasattr(obj, "element") and hasattr(obj.element, "body")

def _is_cell_like(obj) -> bool:
    return hasattr(obj, "_tc")

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    if _is_document_like(parent):
        parent_elm = parent.element.body
    elif _is_cell_like(parent):
        parent_elm = parent._tc
    else:
        raise TypeError

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =========================
# Hyperlink parsing
# =========================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _runs_text(run_elm) -> str:
    return "".join(
        node.text for node in run_elm.iter()
        if node.tag.endswith("}t") and node.text
    )

def _extract_url_from_instr(instr: str) -> str:
    if not instr:
        return "#"
    m = re.search(r'HYPERLINK\s+"([^"]+)"', instr, re.I)
    if m:
        return m.group(1)
    m = re.search(r"HYPERLINK\s+(\S+)", instr, re.I)
    if m:
        return m.group(1)
    return "#"


def paragraph_to_html(paragraph: Paragraph) -> str:
    parts = []
    p_elm = paragraph._p

    for child in p_elm.iterchildren():
        tag = child.tag.split("}")[-1]

        if tag == "hyperlink":
            r_id = child.get(qn("r:id"))
            url = "#"
            if r_id:
                rel = paragraph.part.rels.get(r_id)
                if rel:
                    url = rel.target_ref

            text = "".join(_runs_text(r) for r in child if r.tag.endswith("}r"))
            if text:
                parts.append(f'<a href="{url}">{html_entities(text)}</a>')
            continue

        if tag == "r":
            txt = _runs_text(child)
            if txt:
                parts.append(html_entities(txt))

    return "".join(parts).strip()


def extract_all_lines_as_html(doc_obj) -> List[str]:
    lines = []

    def add(t):
        t = (t or "").strip()
        if t:
            lines.append(t)

    for block in iter_block_items(doc_obj):
        if isinstance(block, Paragraph):
            add(paragraph_to_html(block))
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for sub in iter_block_items(cell):
                        if isinstance(sub, Paragraph):
                            add(paragraph_to_html(sub))

    return lines


# =========================
# Parsing input DOCX
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    src = Document(str(path))
    lines = extract_all_lines_as_html(src)

    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""
    testo_lines = []

    in_testo = False
    testo_re = re.compile(r"^({})\s*:\s*$".format("|".join(TESTO_KEYS)), re.I)

    for line in lines:
        if testo_re.match(line):
            in_testo = True
            continue

        if not in_testo:
            m = re.match(r"^([^:]{1,80})\s*:\s*(.*)$", line)
            if m:
                k = _normalize_key(m.group(1))
                v = m.group(2).strip()

                if k in INPUT_KEY_MAP:
                    out_k = INPUT_KEY_MAP[k]

                    if out_k == "H1":
                        h1 = html.unescape(v)

                    elif out_k in meta:
                        if out_k == "Description":
                            meta[out_k] = v           # entities OK
                        else:
                            meta[out_k] = html.unescape(v)
            continue

        testo_lines.append(line)

    if not h1:
        h1 = meta.get("Title") or testo_lines[0]

    # --- parsing sezioni ---
    intro = []
    sections = []
    current_title = None
    current_paras = []

    def flush():
        nonlocal current_title, current_paras
        if current_title:
            sections.append({
                "title": current_title,
                "paras": current_paras[:]
            })
        current_title = None
        current_paras = []

    for t in testo_lines:
        m = re.match(r"^(.*)\s*\((h2|h3)\)$", t, re.I)
        if m:
            flush()
            current_title = m.group(1).strip()
            continue

        if current_title is None:
            intro.append(t)
        else:
            current_paras.append(t)

    flush()

    return {
        "meta": meta,
        "h1": h1.strip(),
        "intro_paras": intro,
        "sections": sections,
    }


# =========================
# HTML blocks
# =========================

def build_html_rows(parsed) -> List[Tuple[str, str]]:
    rows = []

    rows.append(("H1", f"<h1>{html_entities(parsed['h1'])}</h1>"))

    intro_html = "\n\n".join(
        f'<p class="h-text-size-14 h-font-primary">{p}</p>'
        for p in parsed["intro_paras"]
    )
    rows.append(("Intro", intro_html))

    for sec in parsed["sections"]:
        parts = [
            f'<h2><strong>{html_entities(sec["title"])}</strong></h2>'
        ]
        for p in sec["paras"]:
            parts.append(
                f'<p class="h-text-size-14 h-font-primary">{p}</p>'
            )
        rows.append(("S3", "\n\n".join(parts)))

    return rows


def build_structure_of_content(html_rows):
    s = ["H1", "Intro"]
    s3_count = sum(1 for b, _ in html_rows if b == "S3")
    s.extend(["✏️ S3"] * s3_count)
    return s


# =========================
# Output DOCX
# =========================

def write_output_docx(parsed: Dict[str, Any], output_path: Path):
    doc = Document()
    meta = parsed["meta"]

    # Titolo DOCX — RAW
    title = meta.get("Title") or parsed["h1"]

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(20)

    doc.add_paragraph("")

    # --- tabella metadati (IDENTICA) ---
    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"

    for i, key in enumerate(OUTPUT_META_LABELS):
        shade_cell(table.cell(i, 0), "000000")
        set_cell_text(
            table.cell(i, 0),
            key,
            bold=True,
            color=RGBColor(255, 255, 255)
        )
        set_cell_text(
            table.cell(i, 1),
            meta.get(key, "")
        )

    doc.add_paragraph("")
    doc.add_paragraph("")

    # --- Structure of content ---
    p = doc.add_paragraph("Structure of content:")
    p.runs[0].bold = True

    html_rows = build_html_rows(parsed)
    for line in build_structure_of_content(html_rows):
        doc.add_paragraph(line)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # --- Block | HTML Output ---
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"

    shade_cell(t2.cell(0, 0), "D9D9D9")
    shade_cell(t2.cell(0, 1), "D9D9D9")
    set_cell_text(t2.cell(0, 0), "Block", bold=True)
    set_cell_text(t2.cell(0, 1), "⭐ HTML Output ⭐", bold=True)

    for block, html_block in html_rows:
        row = t2.add_row().cells
        row[0].text = block
        row[1].text = html_block

    doc.save(str(output_path))


# =========================
# Streamlit helper
# =========================

def convert_one(input_docx: Path, output_docx: Path):
    parsed = parse_input_docx(input_docx)
    write_output_docx(parsed, output_docx)


def convert_uploaded_file(uploaded_file) -> Path:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        inp = tmpdir / uploaded_file.name
        out = tmpdir / f"output_{uploaded_file.name}"

        inp.write_bytes(uploaded_file.read())
        convert_one(inp, out)

        final = Path(tempfile.gettempdir()) / out.name
        final.write_bytes(out.read_bytes())
        return final
