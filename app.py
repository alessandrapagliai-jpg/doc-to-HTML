import streamlit as st
from pathlib import Path

from converter import convert_uploaded_file

st.set_page_config(
    page_title="DOCX → HTML SEO Converter",
    layout="centered"
)

st.title("📄 DOCX → HTML SEO Converter")
st.write(
    "Carica uno o più file DOCX. "
    "Il sistema li convertirà nel formato HTML SEO-ready."
)

uploaded_files = st.file_uploader(
    "Upload file DOCX",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} file caricati")

    if st.button("🚀 Converti"):
        with st.spinner("Conversione in corso..."):
            outputs = []

            for uf in uploaded_files:
                try:
                    out_path = convert_uploaded_file(uf)
                    outputs.append(out_path)
                except Exception as e:
                    st.error(f"Errore su {uf.name}: {e}")

        st.success("Conversione completata!")

        for out in outputs:
            with open(out, "rb") as f:
                st.download_button(
                    label=f"⬇️ Scarica {out.name}",
                    data=f,
                    file_name=out.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
