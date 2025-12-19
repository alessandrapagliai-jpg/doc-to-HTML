from docx import Document
from pathlib import Path
import html

P_CLASS = 'class="h-text-size-14 h-font-primary"'

def convert_uploaded_file(uploaded_file):
    doc = Document(uploaded_file)

    output = []
    started = False

    for p in doc.paragraphs:
        raw = p.text.strip()
        if not raw:
            continue

        # --- H1 ---
        if raw.startswith("H1:"):
            h1 = raw.replace("H1:", "").strip()
            output.append(f"<h1><strong>{html.escape(h1)}</strong></h1>")
            started = True
            continue

        # ignora tutto prima dell’H1
        if not started:
            continue

        # --- H2 ---
        if raw.lower().startswith("(h2)"):
            h2 = raw[4:].strip()
            output.append(f"<h2>{html.escape(h2)}</h2>")
            continue

        # --- H3 ---
        if raw.lower().startswith("(h3)"):
            h3 = raw[4:].strip()
            output.append(f"<h3>{html.escape(h3)}</h3>")
            continue

        # --- PARAGRAFO ---
        # NB: qui NON usiamo escape totale perché il testo
        # può già contenere <strong> o <a>
        output.append(
            f'<p {P_CLASS}>{raw}</p>'
        )

    out_path = Path("/tmp") / uploaded_file.name.replace(".docx", ".html")
    out_path.write_text("\n\n".join(output), encoding="utf-8")

    return out_path
