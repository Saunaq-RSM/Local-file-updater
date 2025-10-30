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
st.title("TP Agent 2 Yearly Update")

st.download_button(
    "Download variables excel",
    data = open("documents2/RSM NL - TP Agent 2 Yearly Update Variables - 18.09.2025 V1.xlsx", "rb").read(),
    file_name="documents2/RSM NL - TP Agent 2 Yearly Update Variables - 18.09.2025 V1.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.info(
    """
**How this tool works**


1. Upload your files below — Only *Last year local file* is mandatory to upload.

2. Click **Step 1 – Fill & preview variables**. Review/edit the table, Make sure the old values are unique in the text and replacable. Rerunning will rewrite the old table

3. Click **Step 2 – Fill Section Variables**. This will answer any prompts you put in the prompts column of the table from the first step. 

4. Click **Step 3 – Generate final document** to produce the DOCX/PPTX using your edited values.
"""
)

# --- File uploaders (variables + template required) ---
guidelines_file = st.file_uploader("OECD Transfer Pricing Guideline (.txt) — optional", type=["txt"], key="u_guidelines")
transcript_file = st.file_uploader("Client Meeting Transcript (.docx) — optional", type=["docx"], key="u_transcript")
analysis_file   = st.file_uploader("Financial Documents (.pdf) — optional", type=["pdf"], key="u_pdf")
variables_file  = st.file_uploader("Variables (.xlsx) — REQUIRED", type=["xlsx"], key="u_excel")
template_file   = st.file_uploader("Last year local file (.pptx/.docx) — REQUIRED for Step 2", type=["pptx", "docx"], key="u_template")

# Prefill option (Boolean)

# Session state: outputs & dataframes
if "generated" not in st.session_state:
    st.session_state.generated = None
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
    
# Keep editor widget state and script state in sync at the start of each run


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
    
def _persist_editor_from_state():
    if "editor_sheet" in st.session_state:
        _set_sheet(st.session_state.editor_sheet)

# _persist_editor_from_state()

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

def run_step1_fill_and_preview():
    with st.spinner("Preparing variables…"):
        try:
            # Clear downstream state so it always reflects the latest Step 1 output
            st.session_state.generated = None
            st.session_state.step2_ready = False
            st.session_state.section_excel_path = None

            file_map_preview = dict(base_file_map)
            result = find_relevant_variables(file_map_preview)
            excel_path = result[1] if isinstance(result, tuple) else result
            if not excel_path:
                st.error("No Excel produced by Step 1.")
                return

            df = pd.read_excel(excel_path, engine="openpyxl")
            _set_sheet(df)  # <- overwrites session_state.sheet_df and sheet_path
            st.success("Variables filled from the latest run. Edit below.")

        except Exception as e:
            st.error(f"Error in Step 1: {e}")


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

# ---------------- Buttons (inline, no on_click) ----------------
st.divider()
col1, col2, col3 = st.columns(3)
with col1:
    if st.button("Step 1 – Fill & preview variables"):
        run_step1_fill_and_preview()
with col2:
    if st.button("Step 2 – Fill section values"):
        run_step2_fill_sections()
with col3:
    if st.button("Step 3 – Generate final document"):
        run_step3_generate()

if st.session_state.sheet_df is not None:
    st.subheader("Edit variables (live)")

    # Render the editor with the current df
    edited_df = st.data_editor(
        st.session_state.sheet_df,
        key="editor_sheet",          # keep the key if you want, but don't read it from session_state
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
    )

    # If the user made a change, persist it and immediately re-run so the editor shows the new base df
    try:
        changed = not st.session_state.sheet_df.equals(edited_df)
    except Exception:
        # equals() can raise if dtypes changed; fall back to a safer check
        changed = True

    if changed:
        _set_sheet(edited_df)   # persists to st.session_state.sheet_df + temp path
        st.rerun()              # refresh now so the UI reflects the new base data this run

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
