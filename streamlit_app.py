import io
import os
import tempfile
import pandas as pd
import streamlit as st
from processor import configure, find_relevant_variables, fill_section_values,generate_doc_from_excel_map

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

# Helper: write DF to a Windows-safe temp file and return the PATH
def _df_to_temp_xlsx_path(df: pd.DataFrame) -> str:
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)  # close OS handle so pandas can write freely on Windows
    df.to_excel(path, index=False)
    return path

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

# -------- STEP 1: Fill & preview variables (no template replacement yet) --------
def run_step1_fill_and_preview():
    with st.spinner("Filling variables…"):
        try:
            file_map_preview = dict(base_file_map)
            result = find_relevant_variables(file_map_preview)

            # Normalize return: (doc_path, excel_path) or just doc_path
            if isinstance(result, tuple):
                _doc_path, excel_path = result
            else:
                _doc_path, excel_path = result, None

            st.session_state.filled_excel_path = excel_path
            full_df = pd.read_excel(excel_path, engine="openpyxl")
            st.session_state.full_df = full_df
            st.session_state.view_df = full_df.copy()

            st.success("Variables filled. Review and edit the selected columns below, then proceed to Step 2.")
        except Exception as e:
            st.error(f"Error in Step 1: {e}")

st.button("Step 1 – Fill & preview variables", on_click=run_step1_fill_and_preview)

# If we have a view_df, show it as an editable table (only A, E, F)
if st.session_state.view_df is not None and not st.session_state.step2_ready:
    st.subheader("Step 1 edits")
    edited_view_df_step1 = st.data_editor(
        st.session_state.view_df,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_step1",
    )
    # persist edits from step 1 table
    merged_df = st.session_state.full_df.copy()
    for col in edited_view_df_step1.columns:
        if col in merged_df.columns:
            merged_df[col] = edited_view_df_step1[col]

    # Save to Windows-safe temp PATH and remember it
    tmp_after_edit_1 = _df_to_temp_xlsx_path(merged_df)
    st.session_state.filled_excel_path = tmp_after_edit_1  # latest SoT for Step 3

    # (optional) download button for step 1 editor
    toexcel1 = io.BytesIO()
    edited_view_df_step1.to_excel(toexcel1, index=False)
    st.download_button(
        "Download current edited view (Step 1)",
        data=toexcel1.getvalue(),
        file_name="variables_view_step1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_step1",
    )

def run_step2_fill_sections():
    if st.session_state.view_df is None:
        st.error("Please run Step 1 and review/edit the variables first.")
        return

    with st.spinner("Filling section values…"):
        try:
            # Merge current edits back into the full sheet
            full_df = st.session_state.full_df.copy()
            view_df = st.session_state.view_df
            for col in view_df.columns:
                if col in full_df.columns:
                    full_df[col] = view_df[col]

            # Save merged DF to PATH and pass PATH to backend
            tmp_path = _df_to_temp_xlsx_path(full_df)

            file_map2 = dict(base_file_map)
            file_map2["excel"] = tmp_path  # pass PATH, not open file

            result = fill_section_values(file_map2)

            section_excel_path = result[1] if isinstance(result, tuple) else result
            if not section_excel_path:
                st.error("Section-filling step did not return a valid Excel path.")
                return

            st.session_state.section_excel_path = section_excel_path
            st.session_state.step2_ready = True

            df2 = pd.read_excel(section_excel_path, engine="openpyxl")
            st.session_state.full_df = df2
            st.session_state.view_df = df2.copy()

            # Seed SoT to the just-produced file (user edits may override this later)
            st.session_state.filled_excel_path = section_excel_path

            st.success("Section values filled. You can now review/edit them below, then proceed to Step 3.")
        except Exception as e:
            st.error(f"Error in Step 2: {e}")

st.button("Step 2 – Fill section values", on_click=run_step2_fill_sections)

if st.session_state.step2_ready and st.session_state.section_excel_path:
    st.subheader("Step 2 edits (section-filled)")
    edited_view_df_step2 = st.data_editor(
        st.session_state.view_df,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_step2",
    )
    # persist edits from step 2 table
    merged_df2 = st.session_state.full_df.copy()
    for col in edited_view_df_step2.columns:
        if col in merged_df2.columns:
            merged_df2[col] = edited_view_df_step2[col]

    # Save to PATH and update SoT for Step 3
    tmp_after_edit_2 = _df_to_temp_xlsx_path(merged_df2)
    st.session_state.filled_excel_path = tmp_after_edit_2  # latest SoT

    # (optional) download button for step 2 editor
    toexcel2 = io.BytesIO()
    edited_view_df_step2.to_excel(toexcel2, index=False)
    st.download_button(
        "Download current edited view (Step 2)",
        data=toexcel2.getvalue(),
        file_name="variables_view_step2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_step2",
    )

# -------- STEP 3: Generate final document using the edited values --------
def run_step3_generate():
    if st.session_state.view_df is None or st.session_state.full_df is None:
        st.error("Please run Step 1 (and Step 2 if desired) before generating.")
        return
    if not template_file:
        st.error("Please upload the Template (.pptx/.docx) for Step 3.")
        return

    with st.spinner("Generating final document…"):
        try:
            # Merge any last edits
            full_df = st.session_state.full_df.copy()
            view_df = st.session_state.view_df
            for col in view_df.columns:
                if col in full_df.columns:
                    full_df[col] = view_df[col]

            # Prefer the latest edited workbook path
            excel_path_for_generation = st.session_state.get("filled_excel_path")
            if not excel_path_for_generation:
                # Fallback: write current merged DF to PATH
                excel_path_for_generation = _df_to_temp_xlsx_path(full_df)
                st.session_state.filled_excel_path = excel_path_for_generation

            file_map2 = dict(base_file_map)
            file_map2["excel"] = excel_path_for_generation  # pass PATH

            # Generate doc (no prefill here)
            result = generate_doc_from_excel_map(file_map2)

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

# Downloads (persist via session_state so no re-generation on click)
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
