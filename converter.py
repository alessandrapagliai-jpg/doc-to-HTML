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
    Converte un Paragraph in HTML-safe, convertendo hyperlink Word in <a>.
    """
    parts: List[str] = []
    p_elm = paragraph._p

    # Stato per campi complessi
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
                    '<a href="{href}">{text}</a>'.format(
                        href=html.escape(field_url, quote=True),
                        text=html_entities(visible)
                    )
                )
        in_field = False
        field_instr = ""
        field_url = None
        after_separate = False
        field_display_parts = []

    for child in p_elm.iterchildren():
        tag = child.tag.split("}")[-1]

        # (A) Hyperlink standard
        if tag == "hyperlink":
            r_id = child.get(qn("r:id"))
            url = "#"
            if r_id:
                rel = paragraph.part.rels.get(r_id)
                if rel and getattr(rel, "target_ref", None):
                    url = rel.target_ref

            link_text_parts = []
            for grand in child.iterchildren():
                if grand.tag.split("}")[-1] == "r":
                    link_text_parts.append(_runs_text(grand))

            link_text = "".join(link_text_parts).strip()
            if link_text:
                parts.append(
                    '<a href="{href}">{text}</a>'.format(
                        href=html.escape(url, quote=True),
                        text=html_entities(link_text)
                    )
                )
            continue

        # (B) fldSimple
        if tag == "fldSimple":
            instr = child.get(qn("w:instr")) or child.get("{%s}instr" % W_NS) or ""
            url = _extract_url_from_instr(instr)

            display = []
            for node in child.iter():
                if node.tag.endswith("}t") and node.text:
                    display.append(node.text)
            visible = "".join(display).strip()

            if visible:
                parts.append(
                    '<a href="{href}">{text}</a>'.format(
                        href=html.escape(url, quote=True),
                        text=html_entities(visible)
                    )
                )
            continue

        # (C) Run: campi complessi
        if tag == "r":
            fldChar = child.find(".//w:fldChar", NS)
            if fldChar is not None:
                fld_type = fldChar.get(qn("w:fldCharType")) or fldChar.get("{%s}fldCharType" % W_NS)

                if fld_type == "begin":
                    if in_field:
                        flush_field()
                    in_field = True
                    field_instr = ""
                    field_url = None
                    after_separate = False
                    field_display_parts = []
                    continue

                if fld_type == "separate" and in_field:
                    field_url = _extract_url_from_instr(field_instr)
                    after_separate = True
                    continue

                if fld_type == "end" and in_field:
                    flush_field()
                    continue

            instrText = child.find(".//w:instrText", NS)
            if instrText is not None and in_field and not after_separate:
                if instrText.text:
                    field_instr += instrText.text
                continue

            txt = _runs_text(child)
            if not txt:
                continue

            if in_field:
                if after_separate:
                    field_display_parts.append(txt)
            else:
                parts.append(html_entities(txt))
            continue

    if in_field:
        flush_field()

    return "".join(parts).strip()


# =========================
# RAW extraction (for meta/title parsing)
# =========================

def paragraph_to_text(paragraph: Paragraph) -> str:
    """
    Estrae testo RAW (Unicode) dal paragraph.
    """
    # paragraph.text include il testo "visibile" (senza url dei link)
    return (paragraph.text or "").strip()


# =========================
# Dual extraction: RAW + HTML (same order)
# =========================

def extract_all_lines_raw_and_html(doc_obj) -> List[Tuple[str, str]]:
    """
    Estrae righe mantenendo ordine.
    Per ogni riga: (raw_text, html_text)
    - raw_text: Unicode "pulito" per parsare meta e label
    - html_text: HTML-safe con entities + <a href="..."> per il body
    """
    out: List[Tuple[str, str]] = []

    def add_pair(raw: str, h: str):
        raw = (raw or "").strip()
        h = (h or "").strip()
        if raw or h:
            out.append((raw, h))

    def walk(parent):
        for block in iter_block_items(parent):
            if isinstance(block, Paragraph):
                add_pair(paragraph_to_text(block), paragraph_to_html(block))
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        # contenuto cella in ordine (para + tabelle annidate)
                        for sub in iter_block_items(cell):
                            if isinstance(sub, Paragraph):
                                add_pair(paragraph_to_text(sub), paragraph_to_html(sub))
                            elif isinstance(sub, Table):
                                # tabelle nidificate (1 livello extra)
                                for r2 in sub.rows:
                                    for c2 in r2.cells:
                                        for sub2 in iter_block_items(c2):
                                            if isinstance(sub2, Paragraph):
                                                add_pair(paragraph_to_text(sub2), paragraph_to_html(sub2))

    walk(doc_obj)
    return out


# =========================
# Parsing input
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    src = Document(str(path))
    lines = extract_all_lines_raw_and_html(src)

    meta: Dict[str, str] = {k: "" for k in OUTPUT_META_LABELS}
    h1: str = ""
    testo_pairs: List[Tuple[str, str]] = []

    in_testo = False
    testo_re = re.compile(r"^({})\s*:\s*$".format("|".join(TESTO_KEYS)), re.IGNORECASE)

    for raw_line, html_line in lines:
        # start body marker: controlla sul RAW
        if testo_re.match(raw_line):
            in_testo = True
            continue

        if not in_testo:
            # meta line sul RAW
            m = re.match(r"^([^:]{1,80})\s*:\s*(.*)$", raw_line)
            if m:
                k_raw = _normalize_key(m.group(1))
                v_raw = (m.group(2) or "").strip()

                if k_raw in INPUT_KEY_MAP:
                    out_k = INPUT_KEY_MAP[k_raw]

                    # META e H1 devono restare RAW (niente entities)
                    if out_k == "H1":
                        h1 = v_raw
                    elif out_k in meta:
                        meta[out_k] = v_raw
            continue

        # body: conserva coppia (raw, html) per parsing sezioni ma output html
        testo_pairs.append((raw_line, html_line))

    # fallback: se "Testo:" manca, prendi tutto ciò che non è meta
    if not testo_pairs:
        tmp: List[Tuple[str, str]] = []
        for raw_line, html_line in lines:
            m = re.match(r"^([^:]{1,80})\s*:\s*(.*)$", raw_line)
            if m:
                k_raw = _normalize_key(m.group(1))
                if k_raw in INPUT_KEY_MAP or k_raw in TESTO_KEYS:
                    continue
            tmp.append((raw_line, html_line))
        testo_pairs = tmp

    # fallback H1
    if not h1:
        h1 = (meta.get("Title") or "").strip()
    if not h1 and testo_pairs:
        h1 = (testo_pairs[0][0] or "").strip()
    if not h1:
        h1 = "Untitled"

    # Parsing corpo: intro + sezioni (h2/h3) usando RAW per riconoscere (h2)/(h3)
    intro_paras_html: List[str] = []
    sections: List[Dict[str, Any]] = []
    current_title_raw: Optional[str] = None
    current_paras_html: List[str] = []

    def flush_section():
        nonlocal current_title_raw, current_paras_html, sections
        if current_title_raw is not None:
            sections.append({"title_raw": current_title_raw, "paras_html": current_paras_html[:]})
        current_title_raw = None
        current_paras_html = []

    for raw_t, html_t in testo_pairs:
        raw_t = (raw_t or "").strip()
        html_t = (html_t or "").strip()
        if not raw_t and not html_t:
            continue

        m = re.match(r"^(.*)\s*\((h2|h3)\)\s*$", raw_t, re.IGNORECASE)
        if m:
            flush_section()
            current_title_raw = m.group(1).strip()
            continue

        if current_title_raw is None:
            if html_t:
                intro_paras_html.append(html_t)
        else:
            if html_t:
                current_paras_html.append(html_t)

    flush_section()

    return {
        "meta": meta,                      # RAW
        "h1_raw": h1.strip(),              # RAW
        "intro_paras_html": [p for p in intro_paras_html if p],  # HTML-safe
        "sections": sections,              # titles RAW, paras HTML-safe
    }


# =========================
# HTML generation (final blocks)
# =========================

def build_html_rows(parsed: Dict[str, Any]) -> List[Tuple[str, str]]:
    rows: List[Tuple[str, str]] = []

    # H1: converti SOLO qui
    h1_raw = parsed.get("h1_raw") or ""
    rows.append(("H1", "<h1>{}</h1>".format(html_entities(h1_raw))))

    intro_paras = parsed.get("intro_paras_html") or []
    if intro_paras:
        intro_html = "\n\n".join(
            '<p class="h-text-size-14 h-font-primary">{}</p>'.format(p)
            for p in intro_paras
        )
    else:
        intro_html = '<p class="h-text-size-14 h-font-primary"></p>'
    rows.append(("Intro", intro_html))

    for sec in parsed.get("sections", []):
        title_raw = sec.get("title_raw", "")
        paras_html = sec.get("paras_html", [])
        parts = ['<h2><strong>{}</strong></h2>'.format(html_entities(title_raw))]
        parts.extend(
            '<p class="h-text-size-14 h-font-primary">{}</p>'.format(p)
            for p in paras_html if p
        )
        rows.append(("S3", "\n\n".join(parts).strip()))

    return rows

def build_structure_of_content(html_rows: List[Tuple[str, str]]) -> List[str]:
    s = ["H1", "Intro"]
    s3_count = sum(1 for b, _ in html_rows if b == "S3")
    s.extend(["✏️ S3"] * s3_count)
    return s


# =========================
# DOCX writer
# =========================

def write_output_docx(parsed: Dict[str, Any], output_path: Path):
    doc = Document()

    meta = parsed.get("meta", {})
    # Titolo top DOCX: RAW (NO entities)
    title = (meta.get("Title") or "").strip() or (parsed.get("h1_raw") or "").strip() or "Untitled"

    # Titolo top
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(20)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Tabella metadati (IDENTICA al tuo script originale)
    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"

    for i, key in enumerate(OUTPUT_META_LABELS):
        left = table.cell(i, 0)
        right = table.cell(i, 1)

        shade_cell(left, "000000")
        set_cell_text(left, key, bold=True, color=RGBColor(255, 255, 255), size_pt=10)

        # META: RAW (NO entities)
        set_cell_text(right, (meta.get(key, "") or ""), bold=False, size_pt=10)

    doc.add_paragraph("")
    doc.add_paragraph("")

    # Structure of content
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

    # Tabella Block | HTML Output (identica)
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"

    hdr0 = t2.cell(0, 0)
    hdr1 = t2.cell(0, 1)
    shade_cell(hdr0, "D9D9D9")
    shade_cell(hdr1, "D9D9D9")
    set_cell_text(hdr0, "Block", bold=True, size_pt=10)
    set_cell_text(hdr1, "⭐ HTML Output ⭐", bold=True, size_pt=10)

    for block, html_block in html_rows:
        row_cells = t2.add_row().cells
        row_cells[0].text = block
        row_cells[1].text = html_block

        for cell in row_cells:
            for para in cell.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(10)

    doc.save(str(output_path))


# =========================
# Runner
# =========================

def convert_one(input_docx: Path, output_docx: Path):
    parsed = parse_input_docx(input_docx)
    write_output_docx(parsed, output_docx)


# =========================
# Streamlit helper
# =========================

def convert_uploaded_file(uploaded_file) -> Path:
    """
    Converte un UploadedFile di Streamlit e restituisce un Path scaricabile.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        input_path = tmpdir / uploaded_file.name
        output_path = tmpdir / f"output_{uploaded_file.name}"

        input_path.write_bytes(uploaded_file.read())
        convert_one(input_path, output_path)

        final_output = Path(tempfile.gettempdir()) / output_path.name
        final_output.write_bytes(output_path.read_bytes())
        return final_output
