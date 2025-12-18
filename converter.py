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
# HTML helpers (SOLO BODY)
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
# RAW + HTML paragraph
# =========================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def paragraph_to_text(paragraph: Paragraph) -> str:
    return (paragraph.text or "").strip()

def _runs_text(run_elm) -> str:
    return "".join(
        n.text for n in run_elm.iter()
        if n.tag.endswith("}t") and n.text
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

            txt = "".join(_runs_text(r) for r in child if r.tag.endswith("}r"))
            if txt:
                parts.append(f'<a href="{html.escape(url)}">{html_entities(txt)}</a>')
            continue

        if tag == "r":
            txt = _runs_text(child)
            if txt:
                parts.append(html_entities(txt))

    return "".join(parts).strip()


def extract_all_lines_raw_and_html(doc_obj) -> List[Tuple[str, str]]:
    out = []

    def add(raw, html_):
        raw = (raw or "").strip()
        html_ = (html_ or "").strip()
        if raw or html_:
            out.append((raw, html_))

    for block in iter_block_items(doc_obj):
        if isinstance(block, Paragraph):
            add(paragraph_to_text(block), paragraph_to_html(block))

        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for sub in iter_block_items(cell):
                        if isinstance(sub, Paragraph):
                            add(paragraph_to_text(sub), paragraph_to_html(sub))

    return out


# =========================
# Parsing input DOCX
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    src = Document(str(path))
    lines = extract_all_lines_raw_and_html(src)

    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""
    testo_pairs = []

    in_testo = False
    testo_re = re.compile(r"^({})\s*:\s*$".format("|".join(TESTO_KEYS)), re.I)

    for raw, html_ in lines:
        if testo_re.match(raw):
            in_testo = True
            continue

        if not in_testo:
            m = re.match(r"^([^:]{1,80})\s*:\s*(.*)$", raw)
            if m:
                k = _normalize_key(m.group(1))
                v = m.group(2).strip()
                if k in INPUT_KEY_MAP:
                    out_k = INPUT_KEY_MAP[k]
                    if out_k == "H1":
                        h1 = v
                    elif out_k in meta:
                        meta[out_k] = v
            continue

        testo_pairs.append((raw, html_))

    if not testo_pairs:
        for raw, html_ in lines:
            if not re.match(r"^([^:]{1,80})\s*:", raw):
                testo_pairs.append((raw, html_))

    if not h1:
        h1 = meta.get("Title") or (testo_pairs[0][0] if testo_pairs else "Untitled")

    intro_html: List[str] = []
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

    for raw, html_ in testo_pairs:
        m = re.match(r"^(.*)\s*\((h2|h3)\)\s*$", raw, re.I)
        if m:
            flush()
            current_title = m.group(1).strip()
            continue

        if current_title is None:
            intro_html.append(html_)
        else:
            current_paras.append(html_)

    flush()

    return {
        "meta": meta,
        "h1": h1,
        "intro_paras": intro_html,
        "sections": sections,
    }


# =========================
# HTML blocks (IDENTICI)
# =========================

def build_html_rows(parsed: Dict[str, Any]) -> List[Tuple[str, str]]:
    rows = []

    rows.append(("H1", f"<h1>{html_entities(parsed['h1'])}</h1>"))

    intro_html = "\n\n".join(
        f'<p class="h-text-size-14 h-font-primary">{p}</p>'
        for p in parsed["intro_paras"]
    )
    rows.append(("Intro", intro_html))

    for sec in parsed["sections"]:
        parts = [f'<h2><strong>{html_entities(sec["title"])}</strong></h2>']
        parts.extend(
            f'<p class="h-text-size-14 h-font-primary">{p}</p>'
            for p in sec["paras"]
        )
        rows.append(("S3", "\n\n".join(parts)))

    return rows


def build_structure_of_content(html_rows):
    s = ["H1", "Intro"]
    s.extend(["✏️ S3"] * sum(1 for b, _ in html_rows if b == "S3"))
    return s


# =========================
# Output DOCX
# =========================

def write_output_docx(parsed: Dict[str, Any], output_path: Path):
    doc = Document()

    meta = parsed.get("meta", {})

    # =========================
    # Titolo top DOCX (RAW)
    # =========================
    title = (meta.get("Title") or "").strip() or (parsed.get("h1") or "").strip() or "Untitled"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(20)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # =========================
    # Tabella metadati
    # =========================
    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"

    for i, key in enumerate(OUTPUT_META_LABELS):
        left = table.cell(i, 0)
        right = table.cell(i, 1)

        shade_cell(left, "000000")
        set_cell_text(
            left,
            key,
            bold=True,
            color=RGBColor(255, 255, 255),
            size_pt=10
        )

        # META = RAW (NO html_entities)
        set_cell_text(
            right,
            (meta.get(key, "") or ""),
            bold=False,
            size_pt=10
        )

    doc.add_paragraph("")
    doc.add_paragraph("")

    # =========================
    # Structure of content
    # =========================
    p = doc.add_paragraph("Structure of content:")
    if p.runs:
        p.runs[0].bold = True
    else:
        p.add_run("Structure of content:").bold = True

    html_rows = build_html_rows(parsed)
    for line in build_structure_of_content(html_rows):
        doc.add_paragraph(line)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # =========================
    # Tabella Block | HTML Output
    # =========================
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"

    hdr0 = t2.cell(0, 0)
    hdr1 = t2.cell(0, 1)

    shade_cell(hdr0, "D9D9D9")
    shade_cell(hdr1, "D9D9D9")

    set_cell_text(hdr0, "Block", bold=True, size_pt=10)
    set_cell_text(hdr1, "⭐ HTML Output ⭐", bold=True, size_pt=10)

    # =========================
    # RIGHE CON PARAGRAFI VERI
    # =========================
    for block, html_block in html_rows:
        row_cells = t2.add_row().cells

        # Colonna Block
        row_cells[0].text = block
        for r in row_cells[0].paragraphs[0].runs:
            r.font.size = Pt(10)

        # Colonna HTML Output — QUI IL FIX
        cell = row_cells[1]
        cell.text = ""  # rimuove il paragrafo automatico

        # ogni \n\n = nuovo paragrafo Word
        chunks = [c for c in html_block.split("\n\n") if c.strip()]

        for chunk in chunks:
            p = cell.add_paragraph(chunk)
            for r in p.runs:
                r.font.size = Pt(10)

    # =========================
    # Save
    # =========================
    doc.save(str(output_path))

# =========================
# Streamlit helper
# =========================

def convert_uploaded_file(uploaded_file) -> Path:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        inp = tmpdir / uploaded_file.name
        out = tmpdir / f"output_{uploaded_file.name}"
        inp.write_bytes(uploaded_file.read())
        parsed = parse_input_docx(inp)
        write_output_docx(parsed, out)
        final = Path(tempfile.gettempdir()) / out.name
        final.write_bytes(out.read_bytes())
        return final

