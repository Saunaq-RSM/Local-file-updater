import io
import os
import tempfile
import pandas as pd
import streamlit as st
from processor import configure, find_relevant_variables, fill_section_values, generate_doc_from_excel_map

# Configure backend with secrets
configure(
    st.secrets["AZURE_API_KEY"],
    st.secrets["AZURE_API_ENDPOINT"]
)

st.set_page_config(page_title="TP Template Updater", layout="wide")
st.title("Transfer Pricing Document Updater")

st.download_button(
    "Download variables excel",
    data = open("RSM NL - TP Agent 2 Yearly Update Variables - 18.09.2025 V1.xlsx", "rb").read(),
    file_name="RSM NL - TP Agent 2 Yearly Update Variables - 18.09.2025 V1.xlsx",
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

3. Click **Step 1 – Fill & preview variables**. Review/edit the table (we only show columns A, E, F).

4. Click **Step 2 – Generate final document** to produce the DOCX/PPTX using your edited values.
"""
)

# --- File uploaders (variables + template required) ---
guidelines_file = st.file_uploader("OECD Transfer Pricing Guideline (.txt) — optional", type=["txt"], key="u_guidelines")
transcript_file = st.file_uploader("Client Meeting Transcript (.docx) — optional", type=["docx"], key="u_transcript")
analysis_file   = st.file_uploader("Financial Documents (.pdf) — optional", type=["pdf"], key="u_pdf")
variables_file  = st.file_uploader("Variables (.xlsx) — REQUIRED", type=["xlsx"], key="u_excel")
template_file   = st.file_uploader("Last year local file (.pptx/.docx) — REQUIRED for Step 2", type=["pptx", "docx"], key="u_template")

# Prefill option (Boolean)
prefill_last_year = st.checkbox(
    "AI prefill last-year values (fill Column E using prompts in Column D)",
    value=False,
    help="If enabled, the app first asks the AI to fill Column E based on Column D, then performs replacements."
)

# Session state: outputs & dataframes
if "generated" not in st.session_state:
    st.session_state.generated = None
if "full_df" not in st.session_state:
    st.session_state.full_df = None         # full sheet (all columns)
if "view_df" not in st.session_state:
    st.session_state.view_df = None         # only columns A, E, F
if "filled_excel_path" not in st.session_state:
    st.session_state.filled_excel_path = None
if "step2_ready" not in st.session_state:
    st.session_state.step2_ready = False
if "section_excel_path" not in st.session_state:
    st.session_state.section_excel_path = None  # Excel after section-filling step
if "sheet_df" not in st.session_state:
    st.session_state.sheet_df = None
if "sheet_path" not in st.session_state:
    st.session_state.sheet_path = None
if "generated" not in st.session_state:
    st.session_state.generated = None

# Helper: write DF to a Windows-safe temp file and return the PATH
def _df_to_temp_xlsx_path(df: pd.DataFrame) -> str:
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    df.to_excel(path, index=False)
    return path

def _coerce_df(df: pd.DataFrame) -> pd.DataFrame:
    # Keep it simple and editable in Streamlit
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].fillna("")
    # Make 'prompt' definitely editable
    if "prompt" in out.columns:
        out["prompt"] = out["prompt"].astype("string").fillna("")
    return out

def _set_sheet(df: pd.DataFrame):
    df = _coerce_df(df)
    st.session_state.sheet_df = df
    st.session_state.sheet_path = _df_to_temp_xlsx_path(df)

def _download_current_sheet(label="Download current variables.xlsx", key="dl_current_sheet"):
    if st.session_state.sheet_df is None:
        return
    bio = io.BytesIO()
    st.session_state.sheet_df.to_excel(bio, index=False)
    st.download_button(
        label,
        data=bio.getvalue(),
        file_name="variables_current.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        use_container_width=True,
    )

# Build the file_map used by the backend; pass None for missing optionals
base_file_map = {
    "guidelines": guidelines_file or None,
    "transcript": transcript_file or None,
    "pdf": analysis_file or None,
    "excel": variables_file or None,
    "template": template_file or None,
}

with st.expander("Selected files"):
    st.write(f"- **guidelines** → {guidelines_file.name if guidelines_file else '(none)'}")
    st.write(f"- **transcript** → {transcript_file.name if transcript_file else '(none)'}")
    st.write(f"- **analysis** → {analysis_file.name if analysis_file else '(none)'}")
    st.write(f"- **variables** → {variables_file.name if variables_file else '(none)'}")
    st.write(f"- **template** → {template_file.name if template_file else '(none)'}")

# Helper to pick the display columns (A,E,F) robustly
def pick_view_columns(df: pd.DataFrame):
    cols = list(df.columns)

    def by_idx(i, default=None):
        return cols[i] if i < len(cols) else default

    # Try name-based first, then fall back to positional (A=0, E=4, F=5)
    col_A = next((c for c in cols if str(c).strip().lower() in {"variable_name","variable name","variable","name"}), by_idx(0))
    col_E = next((c for c in cols if "last" in str(c).lower() and "year" in str(c).lower()), by_idx(4))
    col_F = next((c for c in cols if "this" in str(c).lower() and "year" in str(c).lower()), by_idx(5))

    # Filter out Nones / duplicates while preserving order
    ordered = []
    for c in [col_A, col_E, col_F]:
        if c is not None and c in cols and c not in ordered:
            ordered.append(c)
    return ordered

def run_step1_fill_and_preview():
    with st.spinner("Preparing variables…"):
        try:
            # If we already have a sheet in memory, keep it (user may be returning from Step 2)
            if st.session_state.sheet_df is not None:
                st.success("Variables loaded from current session. Edit below.")
                return

            file_map_preview = dict(base_file_map)
            result = find_relevant_variables(file_map_preview)
            excel_path = result[1] if isinstance(result, tuple) else result
            if not excel_path:
                st.error("No Excel produced by Step 1.")
                return

            df = pd.read_excel(excel_path, engine="openpyxl")
            _set_sheet(df)
            st.success("Variables filled. Edit below.")

        except Exception as e:
            st.error(f"Error in Step 1: {e}")

st.button("Step 1 – Fill & preview variables", on_click=run_step1_fill_and_preview)

# Single editor (always edits the global sheet)
if st.session_state.sheet_df is not None:
    st.subheader("Edit variables (live)")
    edited_df = st.data_editor(
        st.session_state.sheet_df,
        key="editor_sheet",
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
    )
    # Persist edits and refresh saved path
    _set_sheet(edited_df)

    # Download the exact sheet that will be used by Step 2/3
    _download_current_sheet()

# ---------------- STEP 2 ----------------
def run_step2_fill_sections():
    if st.session_state.sheet_path is None:
        st.error("Please run Step 1 first (or load variables) so we have a current sheet.")
        return

    with st.spinner("Filling section values…"):
        try:
            file_map2 = dict(base_file_map)
            file_map2["excel"] = st.session_state.sheet_path  # pass the current saved path
            result = fill_section_values(file_map2)
            section_excel_path = result[1] if isinstance(result, tuple) else result
            if not section_excel_path:
                st.error("Section-filling step did not return a valid Excel path.")
                return

            df2 = pd.read_excel(section_excel_path, engine="openpyxl")
            _set_sheet(df2)  # update the ONE global sheet
            st.success("Section values filled. You can keep editing below.")

        except Exception as e:
            st.error(f"Error in Step 2: {e}")

st.button("Step 2 – Fill section values", on_click=run_step2_fill_sections)

# If a sheet exists, we already show the editor above and auto-persist + download button

# ---------------- STEP 3 ----------------
def run_step3_generate():
    if st.session_state.sheet_path is None or st.session_state.sheet_df is None:
        st.error("Please prepare the variables in Step 1 (and optionally Step 2) first.")
        return
    if not template_file:
        st.error("Please upload the Template (.pptx/.docx) for Step 3.")
        return

    with st.spinner("Generating final document…"):
        try:
            file_map3 = dict(base_file_map)
            file_map3["excel"] = st.session_state.sheet_path  # always the latest persisted sheet

            result = generate_doc_from_excel_map(file_map3)
            if isinstance(result, tuple):
                doc_path, excel_path = result
            else:
                doc_path, excel_path = result, None

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
            st.error(f"Error in Step 3: {e}")

st.button("Step 3 – Generate final document", on_click=run_step3_generate)

# Downloads
gen = st.session_state.generated
if gen:
    st.download_button(
        f"Download {gen['doc_name']}",
        data=gen["doc_bytes"],
        file_name=gen["doc_name"],
        mime=gen["doc_mime"],
        key="dl_doc",
        use_container_width=True,
    )
    if gen["excel_bytes"]:
        st.download_button(
            "Download variables_filled.xlsx",
            data=gen["excel_bytes"],
            file_name=gen["excel_name"],
            mime=gen["excel_mime"],
            key="dl_xlsx",
            use_container_width=True,
        )
