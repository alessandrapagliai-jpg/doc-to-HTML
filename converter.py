# =========================
# Streamlit helper
# =========================

import tempfile
from pathlib import Path

def convert_uploaded_file(uploaded_file) -> Path:
    """
    Converte un file DOCX caricato via Streamlit
    e restituisce il Path del file DOCX convertito.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        input_path = tmpdir / uploaded_file.name
        output_path = tmpdir / f"output_{uploaded_file.name}"

        # Salva il file caricato
        with open(input_path, "wb") as f:
            f.write(uploaded_file.read())

        # Conversione
        convert_one(input_path, output_path)

        # Copia l'output in una posizione accessibile a Streamlit
        final_output = Path(tempfile.gettempdir()) / output_path.name
        final_output.write_bytes(output_path.read_bytes())

        return final_output