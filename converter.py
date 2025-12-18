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
    """
    Converte caratteri speciali in HTML entities.
    USARE SOLO PER:
    - H1 HTML
    - H2 / H3
    - Corpo del testo
    NON usare per:
    - Titolo DOCX
    - Meta Title
    - Meta Description
    """
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

def set_cell_text(cell, text: str, bold: bool = False, color: Optional[RGBColor] = None, size_pt: int = 10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text or "")
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = color


# =========================
# DOCX parsing helpers
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
        raise TypeError(f"Parent non supportato: {type(parent)}")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =========================
# Hyperlink extraction
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
        return m.group(1).strip()

    m = re.search(r"HYPERLINK\s+(\S+)", instr, re.I)
    if m:
        return m.group(1).strip()

    return "#"


def paragraph_to_html(paragraph: Paragraph) -> str:
    parts: List[str] = []
    p_elm = paragraph._p

    in_field = False
    field_instr = ""
    field_url = None
    after_separate = False
    field_display_parts: List[str] = []

    def flush_field():
        nonlocal in_field, field_instr, field_url, after_separate, field_display_parts
        if field_url and field_display_parts:
            visible = "".join(field_display_parts).strip()
            if visible:
                parts.append(
                    f'<a href="{html.escape(field_url, quote=True)}">{html_entities(visible)}</a>'
                )
        in_field = False
        field_instr = ""
        field_url = None
        after_separate = False
        field_display_parts = []

    for child in p_elm.iterchildren():
        tag = child.tag.split("}")[-1]

        if tag == "hyperlink":
            r_id = child.get(qn("r:id"))
            url = "#"
            if r_id:
                rel = paragraph.part.rels.get(r_id)
                if rel and getattr(rel, "target_ref", None):
                    url = rel.target_ref

            text = "".join(_runs_text(r) for r in child if r.tag.endswith("}r")).strip()
            if text:
                parts.append(f'<a href="{html.escape(url, quote=True)}">{html_entities(text)}</a>')
            continue

        if tag == "fldSimple":
            instr = child.get(qn("w:instr")) or ""
            url = _extract_url_from_instr(instr)
            text = "".join(node.text for node in child.iter() if node.tag.endswith("}t") and node.text)
            if text:
                parts.append(f'<a href="{html.escape(url, quote=True)}">{html_entities(text.strip())}</a>')
            continue

        if tag == "r":
            fldChar = child.find(".//w:fldChar", NS)
            if fldChar is not None:
                fld_type = fldChar.get(qn("w:fldCharType"))
                if fld_type == "begin":
                    flush_field()
                    in_field = True
                    continue
                if fld_type == "separate":
                    field_url = _extract_url_from_instr(field_instr)
                    after_separate = True
                    continue
                if fld_type == "end":
                    flush_field()
                    continue

            instrText = child.find(".//w:instrText", NS)
            if instrText is not None and in_field and not after_separate:
                field_instr += instrText.text or ""
                continue

            txt = _runs_text(child)
            if not txt:
                continue

            if in_field and after_separate:
                field_display_parts.append(txt)
            else:
                parts.append(html_entities(txt))

    if in_field:
        flush_field()

    return "".join(parts).strip()


# =========================
# Parsing input DOCX
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    src = Document(str(path))
    lines = extract_all_lines_as_html(src)

    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""
    testo_lines: List[str] = []

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
                        h1 = v
                    elif out_k in meta:
                        meta[out_k] = v
            continue

        testo_lines.append(line)

    if not testo_lines:
        testo_lines = [
            l for l in lines
            if not re.match(r"^([^:]{1,80})\s*:", l)
        ]

    if not h1:
        h1 = meta.get("Title") or testo_lines[0] if testo_lines else "Untitled"

    return {
        "meta": meta,
        "h1": h1.strip(),
        "intro_paras": testo_lines,
        "sections": [],
    }


# =========================
# HTML blocks
# =========================

def build_html_rows(parsed: Dict[str, Any]) -> List[Tuple[str, str]]:
    rows: List[Tuple[str, str]] = []

    rows.append(("H1", f"<h1>{html_entities(parsed['h1'])}</h1>"))

    intro_html = "\n\n".join(
        f'<p class="h-text-size-14 h-font-primary">{p}</p>'
        for p in parsed.get("intro_paras", [])
    )
    rows.append(("Intro", intro_html))

    return rows


# =========================
# Output DOCX
# =========================

def write_output_docx(parsed: Dict[str, Any], output_path: Path):
    doc = Document()
    meta = parsed["meta"]

    # Titolo DOCX (NO conversione)
    title = meta.get("Title") or parsed["h1"] or "Untitled"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(20)

    doc.add_paragraph("")

    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"

    for i, key in enumerate(OUTPUT_META_LABELS):
        shade_cell(table.cell(i, 0), "000000")
        set_cell_text(table.cell(i, 0), key, bold=True, color=RGBColor(255, 255, 255))
        set_cell_text(table.cell(i, 1), meta.get(key, ""))  # NO html_entities

    doc.add_paragraph("")
    doc.add_paragraph("")

    for block, html_block in build_html_rows(parsed):
        doc.add_paragraph(block)
        doc.add_paragraph(html_block)

    doc.save(str(output_path))


# =========================
# Runner helpers
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
