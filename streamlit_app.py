import streamlit as st
from processor import configure, process_and_fill

# Configure backend with secrets
configure(
    st.secrets["AZURE_API_KEY"],
    st.secrets["AZURE_API_ENDPOINT"]
)

st.set_page_config(page_title="TP Template Filler", layout="wide")
st.title("Transfer Pricing Document Filler")

st.download_button(
    "Download variables excel",
    data=open("Variales To Fill In.xlsx", "rb").read(),
    file_name="Variales To Fill In.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.info(
    """
**How this tool works**

1. Update your **Variables (.xlsx)**:
   - **Last year value (Col E)** will be **replaced by This year value (Col F)**.
   - If you want the AI to suggest **This year value**, write guidance in **AI prompt (Col D)**.  
     When **Col F** is empty and **Col D** has a prompt, the AI will read your uploaded files and propose a value.

2. Upload your files below — **Variables** and **Template** are required; the others are optional.

3. Click **Generate filled file**. When it finishes, you can download:
   - the filled template (DOCX/PPTX), and
   - **variables_filled.xlsx** (showing this year’s values, including those the AI generated).
"""
)

# --- File uploaders (variables + template required) ---
guidelines_file = st.file_uploader("OECD Transfer Pricing Guideline (.txt) — optional", type=["txt"], key="u_guidelines")
transcript_file = st.file_uploader("Client Meeting Transcript (.docx) — optional", type=["docx"], key="u_transcript")
analysis_file   = st.file_uploader("Financial Documents (.pdf) — optional", type=["pdf"], key="u_pdf")
variables_file  = st.file_uploader("Variables (.xlsx) — REQUIRED", type=["xlsx"], key="u_excel")
template_file   = st.file_uploader("Last year local file (.pptx/.docx) — REQUIRED", type=["pptx", "docx"], key="u_template")

# Show what we detected
with st.expander("Selected files"):
    st.write(f"- **guidelines** → {guidelines_file.name if guidelines_file else '(none)'}")
    st.write(f"- **transcript** → {transcript_file.name if transcript_file else '(none)'}")
    st.write(f"- **analysis** → {analysis_file.name if analysis_file else '(none)'}")
    st.write(f"- **variables** → {variables_file.name if variables_file else '(none)'}")
    st.write(f"- **template** → {template_file.name if template_file else '(none)'}")

# Build the file_map used by the backend; pass None for missing optionals
file_map = {
    "guidelines": guidelines_file or None,
    "transcript": transcript_file or None,
    "pdf": analysis_file or None,
    "excel": variables_file or None,
    "template": template_file or None,
}

# Session state bucket for outputs so downloads don't trigger regeneration
if "generated" not in st.session_state:
    st.session_state.generated = None  # dict with bytes & metadata

def generate_outputs():
    if not variables_file or not template_file:
        st.error("Please upload both the Variables (.xlsx) and Template (.pptx/.docx) files.")
        return

    with st.spinner("Processing..."):
        try:
            result = process_and_fill(file_map)

            # Normalize return (support old: str; new: tuple)
            if isinstance(result, tuple):
                doc_path, excel_path = result
            else:
                doc_path, excel_path = result, None

            # Read bytes now and keep in session_state so downloads don't wipe results
            with open(doc_path, "rb") as f:
                doc_bytes = f.read()

            if doc_path.endswith(".pptx"):
                doc_name = "filled.pptx"
                doc_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            else:
                doc_name = "filled.docx"
                doc_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            excel_bytes = None
            if excel_path:
                try:
                    with open(excel_path, "rb") as f:
                        excel_bytes = f.read()
                except Exception:
                    excel_bytes = None

            st.session_state.generated = {
                "doc_bytes": doc_bytes,
                "doc_name": doc_name,
                "doc_mime": doc_mime,
                "excel_bytes": excel_bytes,
                "excel_name": "variables_filled.xlsx",
                "excel_mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
            st.success("Done—download below.")

        except Exception as e:
            st.session_state.generated = None
            st.error(f"Error: {e}")

# Generate button
if st.button("Generate filled file"):
    generate_outputs()

# If we already have generated outputs in this session, show download buttons.
# Clicking these will rerun the script (Streamlit behavior) but we WON'T lose results.
gen = st.session_state.generated
if gen:
    st.download_button(
        f"Download {gen['doc_name']}",
        data=gen["doc_bytes"],
        file_name=gen["doc_name"],
        mime=gen["doc_mime"],
        key="dl_doc",
    )
    if gen["excel_bytes"]:
        st.download_button(
            "Download variables_filled.xlsx",
            data=gen["excel_bytes"],
            file_name=gen["excel_name"],
            mime=gen["excel_mime"],
            key="dl_xlsx",
        )
