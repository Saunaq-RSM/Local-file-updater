import io
import tempfile
import pandas as pd
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
    if not variables_file:
        st.error("Please upload the Variables (.xlsx) file for Step 1.")
        return

    with st.spinner("Filling variables…"):
        try:
            # Call backend but *skip* template replacement by passing template=None
            file_map_preview = dict(base_file_map)
            file_map_preview["template"] = None  # forces fallback DOCX; we'll ignore it

            result = process_and_fill(file_map_preview, prefill_last_year=prefill_last_year)

            # Normalize return: (doc_path, excel_path) or just doc_path
            if isinstance(result, tuple):
                _doc_path, excel_path = result
            else:
                _doc_path, excel_path = result, None

            if not excel_path:
                st.error("Could not create the filled variables sheet.")
                return

            st.session_state.filled_excel_path = excel_path
            full_df = pd.read_excel(excel_path, engine="openpyxl")
            st.session_state.full_df = full_df

            view_cols = pick_view_columns(full_df)
            st.session_state.view_df = full_df[view_cols].copy()

            st.success("Variables filled. Review and edit the selected columns below, then proceed to Step 2.")
        except Exception as e:
            st.error(f"Error in Step 1: {e}")

st.button("Step 1 – Fill & preview variables", on_click=run_step1_fill_and_preview)

# If we have a view_df, show it as an editable table (only A, E, F)
if st.session_state.view_df is not None:
    st.subheader("Edit variables: ")
    edited_view_df = st.data_editor(
        st.session_state.view_df,
        use_container_width=True,
        num_rows="dynamic",
        key="variables_editor_aef",
    )
    # Persist the edited view
    st.session_state.view_df = edited_view_df

    # Optional download of the edited view
    toexcel = io.BytesIO()
    edited_view_df.to_excel(toexcel, index=False)
    st.download_button(
        "Download current edited view (A,E,F)",
        data=toexcel.getvalue(),
        file_name="variables_view_aef.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_current_vars_view",
    )

# -------- STEP 2: Generate final document using the edited values --------
def run_step2_generate():
    if st.session_state.view_df is None or st.session_state.full_df is None:
        st.error("Please run Step 1 and review/edit the variables first.")
        return
    if not template_file:
        st.error("Please upload the Template (.pptx/.docx) for Step 2.")
        return

    with st.spinner("Generating final document…"):
        try:
            # Merge edits (A,E,F) back into the full dataframe
            full_df = st.session_state.full_df.copy()
            view_df = st.session_state.view_df

            # Align on column names present in view_df
            for col in view_df.columns:
                if col in full_df.columns:
                    full_df[col] = view_df[col]

            # Write the merged full sheet to a temp Excel and pass to backend
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            full_df.to_excel(tmp.name, index=False)

            file_map2 = dict(base_file_map)
            file_map2["excel"] = open(tmp.name, "rb")  # pass file-like to backend

            # We already edited final values; avoid prefill at this stage
            result = process_and_fill(file_map2, prefill_last_year=False)

            # Normalize return (support old: str; new: tuple)
            if isinstance(result, tuple):
                doc_path, excel_path = result
            else:
                doc_path, excel_path = result, None

            # Read bytes now and keep in session_state so downloads persist across reruns
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
            st.error(f"Error in Step 2: {e}")

st.button("Step 2 – Generate final document", on_click=run_step2_generate)

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
