import os
import tempfile

import streamlit as st

from pdf_to_excel_universal import extract_pdf_to_excel


st.set_page_config(page_title="Universal PDF Table Extractor", page_icon="📄", layout="wide")

st.title("📄 Universal PDF Table Extractor")
st.write("Upload a PDF and convert detected tables into Excel.")

with st.sidebar:
    st.header("Options")

    force_type = st.selectbox(
        "PDF type",
        ["auto", "digital", "scanned", "mixed"],
        index=0
    )

    dpi = st.slider("DPI (for scanned/mixed PDFs)", min_value=150, max_value=400, value=250, step=50)

    lang = st.text_input("OCR language", value="eng")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Uploaded: {uploaded_file.name}")

    if st.button("Extract to Excel", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, uploaded_file.name)
            output_path = os.path.join(
                tmpdir,
                f"{os.path.splitext(uploaded_file.name)[0]}_tables.xlsx"
            )

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            try:
                with st.spinner("Processing PDF..."):
                    result = extract_pdf_to_excel(
                        pdf_path=pdf_path,
                        output_path=output_path,
                        force_type=None if force_type == "auto" else force_type,
                        dpi=dpi,
                        lang=lang,
                    )

                if result and os.path.exists(result):
                    with open(result, "rb") as f:
                        st.success("Extraction completed.")
                        st.download_button(
                            label="Download Excel",
                            data=f.read(),
                            file_name=os.path.basename(result),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.warning("No tables were found in the PDF.")

            except Exception as e:
                st.error(f"Error: {e}")
