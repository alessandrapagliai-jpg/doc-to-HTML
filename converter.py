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
    "target keyword": "Target Keyword",
    "kw": "Target Keyword",
    "keyword": "Target Keyword",
    "h1": "H1",
}

# =========================
# HELPERS FORMALIZZAZIONE HTML
# =========================

def get_html_text(paragraph: Paragraph) -> str:
    """Estrae il testo da un paragrafo preservando il grassetto (strong)."""
    full_html = ""
    for run in paragraph.runs:
        text = html.escape(run.text)
        if run.bold:
            full_html += f"<strong>{text}</strong>"
        else:
            full_html += text
    
    # Pulizia entità comuni (opzionale se già usato html.escape)
    replacements = {"’": "&rsquo;", "à": "&agrave;", "è": "&egrave;", "é": "&eacute;", "ì": "&igrave;", "ò": "&ograve;", "ù": "&ugrave;"}
    for k, v in replacements.items():
        full_html = full_html.replace(k, v)
    return full_html

# =========================
# PARSING LOGIC
# =========================

def parse_input_docx(path: Path) -> Dict[str, Any]:
    doc = Document(str(path))
    meta = {k: "" for k in OUTPUT_META_LABELS}
    h1 = ""
    body_elements = []
    
    # 1. Estrazione Metadati dalla Tabella
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 2:
                key = row.cells[0].text.strip().lower().replace(":", "")
                val = row.cells[1].text.strip()
                if key in INPUT_KEY_MAP:
                    mapped_key = INPUT_KEY_MAP[key]
                    if mapped_key == "H1":
                        h1 = val
                    else:
                        meta[mapped_key] = val

    # 2. Estrazione Contenuto (Paragrafi fuori dalle tabelle o dopo la tabella meta)
    # Saltiamo i paragrafi che sono già stati usati nei metadati se necessario
    found_testo_start = False
    
    for para in doc.paragraphs:
        text_raw = para.text.strip()
        if not text_raw: continue
        
        # Identifica se è un header (h2) o (h3)
        header_match = re.search(r"\((h2|h3)\)$", text_raw, re.I)
        
        if header_match:
            tag = header_match.group(1).lower()
            clean_text = re.sub(r"\s*\(h[23]\)$", "", text_raw, flags=re.I)
            body_elements.append({
                "block": "✏️ S3",
                "html": f"<{tag}><strong>{clean_text}</strong></{tag}>"
            })
        else:
            # È un paragrafo standard
            # Se il paragrafo contiene solo il nome del tag (es. "Testo:"), lo ignoriamo
            if text_raw.lower().startswith("testo:"):
                continue
                
            html_content = get_html_text(para)
            # Determina il nome del blocco
            block_name = "Intro" if not any(b["block"] == "✏️ S3" for b in body_elements) else "✏️ S3"
            
            body_elements.append({
                "block": block_name,
                "html": f'<p class="h-text-size-14 h-font-primary">{html_content}</p>'
            })

    if not h1: h1 = meta.get("Title", "Senza Titolo")

    return {"meta": meta, "h1": h1, "body": body_elements}

# =========================
# OUTPUT GENERATION (Mantenendo la tua struttura)
# =========================

def write_output_docx(parsed: Dict[str, Any], out: Path):
    doc = Document()
    
    # Header Titolo
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(parsed["h1"])
    r.bold = True
    r.font.size = Pt(18)

    # Tabella Meta
    table = doc.add_table(rows=len(OUTPUT_META_LABELS), cols=2)
    table.style = "Table Grid"
    for i, k in enumerate(OUTPUT_META_LABELS):
        # Header cella (Nero)
        cell_key = table.cell(i, 0)
        tc_pr = cell_key._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), "000000")
        tc_pr.append(shd)
        
        rk = cell_key.paragraphs[0].add_run(k)
        rk.bold = True
        rk.font.color.rgb = RGBColor(255, 255, 255)
        
        table.cell(i, 1).text = parsed["meta"].get(k, "")

    doc.add_paragraph("\nStructure of content:")
    for b in ["H1", "Intro"] + [item["block"] for item in parsed["body"] if "h2" in item["html"]]:
        doc.add_paragraph(b, style='List Bullet')

    doc.add_paragraph("\n")
    
    # Tabella HTML
    t_html = doc.add_table(rows=1, cols=2)
    t_html.style = "Table Grid"
    t_html.cell(0,0).text = "Block"
    t_html.cell(0,1).text = "⭐ HTML Output ⭐"
    
    for item in parsed["body"]:
        row = t_html.add_row().cells
        row[0].text = item["block"]
        row[1].text = item["html"] # Inseriamo il codice HTML come testo piano

    doc.save(str(out))
