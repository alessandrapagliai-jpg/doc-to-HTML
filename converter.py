from docx import Document
from pathlib import Path

P_CLASS = 'class="h-text-size-14 h-font-primary"'

def convert_uploaded_file(uploaded_file):
    src = Document(uploaded_file)
    out = Document()

    in_body = False
    html_lines = []

    # ---- PARSE INPUT ----
    for p in src.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        # intercettiamo l’H1 editoriale
        if text.startswith("H1:"):
            h1 = text.replace("H1:", "").strip()
            html_lines.append(f"<h1><strong>{h1}</strong></h1>")
            in_body = True
            continue

        if not in_body:
            # prima dell’H1: metadati, li copiamo così come sono
            out.add_paragraph(text)
            continue

        # h2
        if text.lower().startswith("(h2)"):
            html_lines.append(f"<h2>{text[4:].strip()}</h2>")
            continue

        # h3
        if text.lower().startswith("(h3)"):
            html_lines.append(f"<h3>{text[4:].strip()}</h3>")
            continue

        # paragrafo
        html_lines.append(
            f'<p {P_CLASS}>{text}</p>'
        )

    # ---- SCRITTURA OUTPUT ----

    out.add_paragraph("")  # spazio
    out.add_heading("⭐ HTML Output ⭐", level=1)
    out.add_paragraph("")  # spazio

    for line in html_lines:
        out.add_paragraph(line)

    out_path = Path("/tmp") / uploaded_file.name.replace(".docx", "_OUTPUT.docx")
    out.save(out_path)

    return out_path
